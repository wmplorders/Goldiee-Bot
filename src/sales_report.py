"""Conversation workflow and private, company-specific customer town registry."""
import hashlib
import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

from .settings import PROJECT_ROOT
from .tally_job_lock import serialized
from .sales_report_data import collect, parse_command
from .sales_report_output import build, indian

STATE_DIR=PROJECT_ROOT/'data'/'sales_reports'
TOWNS_PATH=PROJECT_ROOT/'data'/'sales_report_towns.json'


def company_key(config):
    c=config['tally']
    return hashlib.sha256(json.dumps([c.get(k,'') for k in ('company_name','data_path','company_folder','http_url')]).encode()).hexdigest()


def _path(scope):return STATE_DIR/(hashlib.sha256(scope.encode()).hexdigest()+'.json')


def _save(path,value):
    path.parent.mkdir(parents=True,exist_ok=True);temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8');temp.replace(path)


def has_pending(scope):
    p=_path(scope)
    if not p.exists():return False
    state=json.loads(p.read_text(encoding='utf-8'))
    return state.get('stage')=='town' and datetime.now(timezone.utc).timestamp()-state['updated']<86400


def cancel(scope):
    p=_path(scope)
    if p.exists():
        state=json.loads(p.read_text(encoding='utf-8'));state['stage']='cancelled';_save(p,state)


def town_registry(config):
    db=json.loads(TOWNS_PATH.read_text(encoding='utf-8')) if TOWNS_PATH.exists() else {}
    return db,db.setdefault(company_key(config),{})


def _norm(name):return re.sub(r'[^a-z0-9]','',re.sub(r'\bC\d+\b','',name,flags=re.I).casefold())


def import_towns(path,payload,registry):
    from openpyxl import load_workbook
    wb=load_workbook(path,read_only=True,data_only=True)
    imported=0
    try:
        for ws in wb:
            values=ws.iter_rows(values_only=True)
            for _ in range(5):
                header=next(values,None)
                if header is None:break
                if len(header)>=2 and str(header[0]).strip().upper()=='PARTY NAME' and str(header[1]).strip().upper()=='POST CODE':break
            else:continue
            if not header or len(header)<2 or str(header[0]).strip().upper()!='PARTY NAME' or str(header[1]).strip().upper()!='POST CODE':continue
            for cells in values:
                if len(cells)<2 or not cells[0] or not cells[1]:continue
                name,town=str(cells[0]).strip(),str(cells[1]).strip()
                candidates=[r for r in payload['rows'] if _norm(r['name'])==_norm(name)]
                if len(candidates)!=1:continue
                if not 1<=len(town)<=80 or any(ord(c)<32 for c in town):raise ValueError('Invalid town in uploaded report.')
                row=candidates[0];old=registry.get(row['id'])
                if old and old['town']!=town:
                    raise ValueError(f"Town conflict for {row['name']}: saved {old['town']}, uploaded {town}.")
                registry[row['id']]={'name':row['name'],'town':town};imported+=1
    finally:wb.close()
    if not imported:raise ValueError('No matching active customer towns found. Upload a sales report with PARTY NAME and POST CODE columns.')
    return imported


def _missing(payload,registry):return [r for r in payload['rows'] if r['id'] not in registry]


def _result(state):
    if state['stage']=='town':
        row=state['missing'][0]
        return {'message':f"Town / Post Code needed for {row['name']}.\nReply with the town, or upload an existing sales-report Excel to import known towns.\n{len(state['missing'])} active customer(s) still need towns. Send refresh to recheck Tally groups, or cancel to stop.",
                'files':[],'prompt':{'stage':'town','type':'sales_town','customer':row['name']}}
    return {'message':state.get('message','Sales report cancelled.'),'files':state.get('files',[]),'prompt':None}


def _finish(state,config,registry):
    payload=state['payload'];state['missing']=_missing(payload,registry)
    if state['missing']:
        state['stage']='town';return
    for row in payload['rows']:row['town']=registry[row['id']]['town']
    payload['company']=config['tally']['company_name']
    stem=f"Sales Report {payload['start']} to {payload['end']} {uuid.uuid4().hex[:8]}"
    xlsx,png=build(payload,config['paths']['processed'],stem)
    state['stage']='done';state['files']=[str(xlsx),str(png)]
    t=payload['totals'];excluded=payload['excluded']
    state['message']=(f"Sales and Payment Report: {payload['start']} to {payload['end']}\n"
        f"{len(payload['rows'])} active customers, alphabetically arranged.\n"
        f"Spices sales: Rs. {indian(t['spices'])}\nDetergent sales: Rs. {indian(t['detergent'])}\nPayments: Rs. {indian(t['payment'])}\n"
        f"Debit balance: Rs. {indian(t['debit'])}; Credit balance: Rs. {indian(t['credit'])}.\n"
        f"{sum(v=='Closed Customers' for v in excluded.values())} Closed Customers excluded. "
        f"{sum(v!='Closed Customers' for v in excluded.values())} non-customer tax ledger(s) excluded.\n"
        f"{payload['counts'].get('credit_notes_ignored_in_sales',0)} credit notes excluded from sales; actual closing balances include all ledger adjustments.\n"
        'Snapshot shows the final date and preceding calendar date. Excel and PNG attached. All customer balances reconciled.')


@serialized
def advance(scope,text,config,attachments=None):
    attachments=attachments or []
    path=_path(scope);fresh=bool(re.match(r'^sales\s+report\s+',text.strip(),re.I))
    db,registry=town_registry(config)
    if fresh:
        try:
            from PIL import Image
        except ImportError as exc:
            raise RuntimeError('Sales snapshots require Pillow. Run .\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt on the bot laptop.') from exc
        start,end=parse_command(text)
        payload=collect(config,start,end)
        state={'key':company_key(config),'start':start.isoformat(),'end':end.isoformat(),'payload':payload,'stage':'town'}
    else:
        if not has_pending(scope):raise ValueError('No active sales report. Send sales report 1/11/26 to 18/11/26.')
        state=json.loads(path.read_text(encoding='utf-8'))
        if text.strip().lower()=='cancel':
            state['stage']='cancelled';_save(path,state);return _result(state)
        if state['key']!=company_key(config):raise ValueError('Tally company or data location changed. Start a new sales report.')
        if not attachments and text.strip().lower()!='refresh':
            town=text.strip()
            if not 1<=len(town)<=80 or any(ord(c)<32 for c in town):raise ValueError('Enter one town / post code, up to 80 characters.')
            row=state['missing'][0];registry[row['id']]={'name':row['name'],'town':town}
    if attachments:
        excels=[Path(p) for p in attachments if Path(p).suffix.lower() in {'.xlsx','.xlsm'}]
        if not excels:raise ValueError('Upload a previous sales-report Excel (.xlsx).')
        for source in excels:import_towns(source,state['payload'],registry)
    _save(TOWNS_PATH,db)
    if not fresh and (text.strip().lower()=='refresh' or not _missing(state['payload'],registry)):
        start,end=parse_command(f"sales report {state['start']} to {state['end']}")
        state['payload']=collect(config,start,end)
    _finish(state,config,registry)
    state['updated']=datetime.now(timezone.utc).timestamp();_save(path,state)
    return _result(state)

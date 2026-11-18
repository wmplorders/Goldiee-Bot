"""Read-only, conversation-scoped Tally / DMS sales invoice comparison."""
import hashlib
import html
import json
import re
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from pathlib import Path

from .settings import PROJECT_ROOT
from .tally_context import get_tally_config, ensure_configured_company, tally_request, clean_tally_xml
from .tally_job_lock import serialized

STATE_DIR = PROJECT_ROOT / 'data' / 'sales_compare'
FIELDS = ('qty', 'mrp', 'rate', 'discount', 'cgst_rate', 'sgst_rate', 'igst_rate', 'taxable')
LABELS = ('Quantity PCS', 'MRP', 'Rate', 'Discount %', 'CGST rate', 'SGST rate', 'IGST rate', 'Taxable value')


def money(value):
    return Decimal(str(value)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)


def number(value, blank=False):
    if value is None or str(value).strip() == '':
        if blank:
            return 0.0
        raise ValueError('Required numeric field is missing.')
    match = re.match(r'^\s*(-?[\d,]+(?:\.\d+)?)', str(value))
    if not match:
        raise ValueError(f'Invalid numeric value: {value}')
    return float(money(match[1].replace(',', '')))


def parse_command(text):
    match = re.fullmatch(r'compare\s+(.+?)\s+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{2}-\d{2})', text.strip(), re.I)
    if not match:
        raise ValueError('Use: compare <customer name> <date>, e.g. compare gauri 14/11/26')
    name, raw = match.groups()
    for fmt in ('%d/%m/%y', '%d/%m/%Y', '%d-%m-%y', '%d-%m-%Y', '%Y-%m-%d'):
        try:
            end = datetime.strptime(raw, fmt).date()
            return name.strip().strip('<>'), (end-timedelta(days=6)).isoformat(), end.isoformat()
        except ValueError:
            pass
    raise ValueError('Invalid comparison date.')


def rank_names(query, names):
    def norm(s):
        return ' '.join(re.sub(r'\bC\d+\b', '', s, flags=re.I).casefold().split())
    q = norm(query)
    def score(name):
        n = norm(name)
        words = n.split()
        partial = max((SequenceMatcher(None, q, ' '.join(words[i:i+len(q.split())])).ratio()
                       for i in range(len(words))), default=0)
        return max(SequenceMatcher(None, q, n).ratio(), partial)
    return sorted(set(names), key=lambda n: (-score(n), n))[:3]


def _export(config, kind, fetch, condition=''):
    cfg = get_tally_config(config)
    ensure_configured_company(cfg)
    formula = f'<FILTER>CompareFilter</FILTER>' if condition else ''
    system = f'<SYSTEM TYPE="Formulae" NAME="CompareFilter">{condition}</SYSTEM>' if condition else ''
    xml = ('<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
           '<TYPE>Collection</TYPE><ID>CompareReadOnly</ID></HEADER><BODY><DESC><STATICVARIABLES>'
           '<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>'
           f'<SVCurrentCompany>{html.escape(cfg.company_name)}</SVCurrentCompany></STATICVARIABLES>'
           '<TDL><TDLMESSAGE><COLLECTION NAME="CompareReadOnly">'
           f'<TYPE>{kind}</TYPE><FETCH>{fetch}</FETCH>{formula}</COLLECTION>{system}'
           '</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>')
    response = tally_request(xml, cfg, timeout=240)
    root = ET.fromstring(clean_tally_xml(response, include_udf_namespace=True))
    errors = [n.text for n in root.findall('.//LINEERROR')]
    if errors:
        raise ValueError('Tally export: ' + '; '.join(errors))
    return root


def _text(node, tag):
    return (node.findtext(tag) or '').strip()


def fetch_tally(config, party, start, end):
    # Always repeat date/party/status filtering locally; TDL filters are not the final boundary.
    literal = html.escape(party.replace('"', '" + $$Char:34 + "'))
    root = _export(config, 'Voucher', '*,AllInventoryEntries.*,BatchAllocations.*,AccountingAllocations.*,LedgerEntries.*',
                   f'$PartyLedgerName = &quot;{literal}&quot;')
    records, excluded = [], 0
    for v in root.findall('.//VOUCHER'):
        date = _text(v, 'DATE')
        if _text(v, 'PARTYLEDGERNAME') != party or not start.replace('-', '') <= date <= end.replace('-', ''):
            continue
        if _text(v, 'ISCANCELLED').lower() == 'yes' or _text(v, 'ISOPTIONAL').lower() == 'yes':
            excluded += 1
            continue
        if 'credit' in _text(v, 'VOUCHERTYPENAME').lower() or 'debit' in _text(v, 'VOUCHERTYPENAME').lower():
            continue
        entries = v.findall('ALLINVENTORYENTRIES.LIST')
        ledgers = [{'name': _text(e, 'LEDGERNAME'), 'amount': number(_text(e, 'AMOUNT'))}
                   for e in v.findall('LEDGERENTRIES.LIST') + v.findall('ALLLEDGERENTRIES.LIST')]
        party_rows = [e for e in ledgers if e['name'] == party]
        if not entries or len(party_rows) != 1 or party_rows[0]['amount'] >= 0:
            continue
        allocations = {_text(a, 'LEDGERNAME') for e in entries for a in e.findall('ACCOUNTINGALLOCATIONS.LIST')}
        if not allocations or not allocations <= {'Sales', 'Exempt Sales'}:
            continue
        invoice = _text(v, 'VOUCHERNUMBER')
        lines = []
        for entry in entries:
            name = _text(entry, 'STOCKITEMNAME')
            rates = {}
            for duty in entry.iter('RATEDETAILS.LIST'):
                head = _text(duty, 'GSTRATEDUTYHEAD')
                if head:
                    rates[head] = number(_text(duty, 'GSTRATE'))
            # Tally exports rate details nested under GSTOVRDN... in some versions.
            if not rates:
                head = None
                for n in entry.iter():
                    if n.tag == 'GSTRATEDUTYHEAD': head = (n.text or '').strip()
                    elif n.tag == 'GSTRATE' and head: rates[head] = number(n.text)
            if not {'CGST', 'SGST/UTGST', 'IGST'} <= rates.keys():
                raise ValueError(f'{invoice}: GST rates unavailable for {name}; cannot assert a match.')
            batches = entry.findall('BATCHALLOCATIONS.LIST')
            for batch in batches or [entry]:
                quantity = _text(batch, 'BILLEDQTY') or _text(entry, 'BILLEDQTY')
                rate = _text(batch, 'BATCHRATE') or _text(entry, 'RATE')
                discount = number(_text(batch, 'BATCHDISCOUNT') or _text(entry, 'DISCOUNT'), blank=True)
                free = discount == 100 or number(rate) == 0
                if len(batches)>1 and (not _text(batch,'BILLEDQTY') or (not _text(batch,'AMOUNT') and not free)):
                    raise ValueError(f'{invoice}: incomplete batch quantity/amount for {name}; cannot duplicate item totals across batches.')
                if not re.match(r'^\s*[\d,.]+\s+PCS\b', quantity, re.I) or not re.search(r'/\s*PCS\s*$', rate, re.I):
                    raise ValueError(f'{invoice}: unsupported quantity/rate unit for {name}: {quantity}, {rate}')
                mrps = [(n.text or '').strip() for n in batch.iter() if n.tag.split('}')[-1] == 'UDF_AA_BATCH_MRP_RATE']
                if not mrps or not mrps[0]:
                    raise ValueError(f'{invoice}: batch MRP unavailable for {name}.')
                lines.append(dict(name=name, invoice=invoice, qty=number(quantity), rate=number(rate),
                    mrp=number(mrps[0]), ptr=_text(batch, 'BATCHNAME'),
                    discount=discount,
                    taxable=number(_text(batch, 'AMOUNT') or _text(entry, 'AMOUNT'), blank=free),
                    cgst_rate=rates['CGST'] if any('CGST' in l['name'].upper() for l in ledgers) else 0,
                    sgst_rate=rates['SGST/UTGST'] if any('SGST' in l['name'].upper() for l in ledgers) else 0,
                    igst_rate=rates['IGST'] if any('IGST' in l['name'].upper() for l in ledgers) else 0))
        records.append(dict(id=_text(v,'GUID') or invoice, invoice=invoice,
                            date=f'{date[:4]}-{date[4:6]}-{date[6:]}', value=-party_rows[0]['amount'],
                            lines=lines, ledgers=ledgers))
    return sorted(records, key=lambda r:(r['date'],r['invoice'])), excluded


def dms_invoices(rows, start, end):
    grouped = {}
    excluded = 0
    for r in rows:
        bill, date = str(r.get('Bill ID') or '').strip(), str(r.get('Bill Date') or '')[:10]
        if not bill or not start <= date <= end:
            continue
        if str(r.get('Is Cancelled') or '').strip().lower() in {'true','yes','1','cancelled'}:
            excluded += 1
            continue
        party = str(r.get('Buyer Name') or '').strip()
        unit = str(r.get('Unit') or '').strip().upper()
        if unit and unit not in {'PCS', 'PC', 'PIECES'}:
            raise ValueError(f'{bill}: unsupported DMS quantity unit {unit}; PCS is required.')
        key = (str(r.get('Buyer ERP ID') or party), bill, date)
        total = number(r.get('Invoice Total Value'))
        g = grouped.setdefault(key,dict(id='|'.join(key),invoice=bill,date=date,party=party,
            party_id=key[0],value=total,lines=[],raw=[]))
        if money(g['value']) != money(total):
            raise ValueError(f'DMS invoice total is inconsistent: {bill}')
        fields = ['GST Discount Percent','Trade Discount Percent','Primary Discount Percent','Secondary Discount Percent','Cash Discount Percent']
        active = [number(r.get(k),blank=True) for k in fields if number(r.get(k),blank=True)]
        discount = active[0] if len(active)==1 else 0
        if len(active)>1:
            gross = number(r.get('Gross Amount'))
            if gross == 0:
                raise ValueError(f'{bill}: cannot derive combined discount on a zero gross amount.')
            discount = float(money(Decimal(str(number(r.get('Total Discount'))))*100/Decimal(str(gross))))
        g['raw'].append(r)
        g['lines'].append(dict(name=str(r.get('Product Name') or ''), invoice=bill,
            qty=number(r.get('Invoiced Total Quantity')),mrp=number(r.get('MRP')),rate=number(r.get('Unit Price')),
            discount=discount,taxable=number(r.get('Gross Amount Less Discount')),
            cgst_rate=number(r.get('CGST Tax Rate')),sgst_rate=number(r.get('SGST Tax Rate')),
            igst_rate=number(r.get('IGST Tax Rate')),cgst=number(r.get('CGST Tax Amount')),
            sgst=number(r.get('SGST Tax Amount')),igst=number(r.get('IGST Tax Amount')),
            net=number(r.get('Net Value')),free_qty=number(r.get('Free Quantity (in Pieces)'),blank=True)))
    return list(grouped.values()), excluded


def consolidate(records, mapping=None):
    groups = {}
    for record in records:
        for line in record['lines']:
            name = (mapping or {}).get(line['name'],line['name'])
            key = (name,)+tuple(money(line[k]) for k in ('mrp','rate','discount','cgst_rate','sgst_rate','igst_rate'))
            if key not in groups:
                groups[key] = dict(line, key=name, qty=0, taxable=0, invoices=[], original_names=[])
            g = groups[key]
            for k in ('qty','taxable'):
                g[k] = float(money(g[k])+money(line[k]))
            g['invoices'] = sorted(set(g['invoices']+[record['invoice']]))
            g['original_names'] = sorted(set(g['original_names']+[line['name']]))
    return groups


def compare_records(tally, dms, mapping=None):
    left, right = consolidate(tally,mapping), consolidate(dms)
    rows = []
    def add(t,d,extra=''):
        diffs = [label for k,label in zip(FIELDS,LABELS) if t is not None and d is not None and money(t[k])!=money(d[k])]
        if t is None: diffs=['Only in DMS']
        if d is None: diffs=['Only in Tally']
        rows.append(dict(t=t,d=d,remarks='; '.join(diffs+([extra] if extra else [])) or 'Exact match',matched=not diffs and not extra))
    for key in list(left):
        if key in right:
            add(left.pop(key),right.pop(key))
    names = sorted({k[0] for k in left}|{k[0] for k in right})
    for name in names:
        tl=[v for k,v in left.items() if k[0]==name]
        dl=[v for k,v in right.items() if k[0]==name]
        if len(tl)==len(dl)==1:
            add(tl[0],dl[0])
        else:
            extra='Multiple price combinations; no arbitrary pairing' if tl and dl else ''
            for t in tl: add(t,None,extra)
            for d in dl: add(None,d,extra)
    return sorted(rows,key=lambda r:((r['d'] or r['t'])['key'], (r['d'] or r['t'])['mrp']))


def selection(text, options, multiple=False):
    if multiple and text.strip().lower() == 'all':
        return list(options)
    parts = re.split(r'[,\s]+',text.strip())
    try:
        indexes = list(dict.fromkeys(int(p)-1 for p in parts))
    except ValueError:
        raise ValueError('Reply with the displayed number(s), e.g. 1 or 1,2.')
    if not indexes or (not multiple and len(indexes)!=1) or any(i<0 or i>=len(options) for i in indexes):
        raise ValueError('Choose only numbers from the displayed list.')
    return [options[i] for i in indexes]


def _path(scope):
    return STATE_DIR/(hashlib.sha256(scope.encode()).hexdigest()+'.json')


def has_pending(scope):
    path=_path(scope)
    if not path.exists(): return False
    state=json.loads(path.read_text(encoding='utf-8'))
    return state.get('stage') not in {'done','cancelled'} and datetime.now(timezone.utc).timestamp()-state['updated']<86400


@serialized
def cancel_pending(scope):
    path=_path(scope)
    if path.exists():
        state=json.loads(path.read_text(encoding='utf-8'));state['stage']='cancelled'
        temporary=path.with_suffix('.tmp')
        temporary.write_text(json.dumps(state),encoding='utf-8');temporary.replace(path)


def _prompt(state):
    stage=state['stage']
    if stage=='party':
        options=[dict(label=n,value=str(i+1)) for i,n in enumerate(state['parties'])]
        message='Select the intended Tally party (top 3 matches):'
    elif stage in {'tally','dms'}:
        records=state['tally_options' if stage=='tally' else 'dms_options']
        options=[dict(label=f"{r['invoice']} | {r['date']} | Rs. {r['value']:,.2f}",value=str(i+1)) for i,r in enumerate(records)]
        message=f"Select one or more {stage.upper()} invoices ({state['start']} to {state['end']}). Excluded: {state.get(stage+'_excluded',0)}."
    elif stage=='dms_party':
        options=[dict(label=n,value=str(i+1)) for i,n in enumerate(state['dms_parties'])]
        message='Select the matching DMS party:'
    elif stage=='mapping':
        pairs=state['proposals']
        options=[dict(label='Accept these pairs for this comparison',value='yes'),dict(label='Keep names separate',value='no')]
        message='Treat these names as the same items for this comparison only?\n'+'\n'.join(f"Tally: {t}\nDMS: {d}" for t,d in pairs)
    elif stage=='confirm':
        options=[dict(label='Create Comparison',value='yes'),dict(label='Cancel',value='cancel')]
        message='Confirm comparison (read only):\nTally: '+', '.join(r['invoice'] for r in state['tally'])+'\nDMS: '+', '.join(r['invoice'] for r in state['dms'])
    else:
        return dict(message=state.get('message','Cancelled.'),files=state.get('files',[]),prompt=None)
    numbered='\n'.join(f"{i+1}. {o['label']}" for i,o in enumerate(options))
    hint='Reply with numbers separated by commas, or all.' if stage in {'tally','dms'} else ('Reply yes or no.' if stage=='mapping' else ('Reply yes to create, or cancel.' if stage=='confirm' else 'Reply with one number.'))
    return dict(message=message+'\n'+numbered+'\n'+hint+' Send cancel to stop.',files=[],
                prompt=dict(stage=stage,options=options,multiple=stage in {'tally','dms'}))


def _dms_party(state,party):
    state['dms_options']=sorted([r for r in state['all_dms'] if r['party']==party],key=lambda r:(r['date'],r['invoice']))
    if not state['dms_options']: raise ValueError('No active DMS invoices found for this party.')
    state['stage']='dms'


def _prepare_mapping(state):
    tn={x['name'] for r in state['tally'] for x in r['lines']}
    dn={x['name'] for r in state['dms'] for x in r['lines']}
    unmatched=dn-tn
    pairs=[]
    for t in sorted(tn-dn):
        candidates=rank_names(t,unmatched)
        if candidates and SequenceMatcher(None,t.casefold(),candidates[0].casefold()).ratio()>=.72:
            pairs.append([t,candidates[0]]);unmatched.remove(candidates[0])
    state['proposals']=pairs
    state['mapping']={}
    state['stage']='mapping' if pairs else 'confirm'


@serialized
def advance(scope, text, config):
    path=_path(scope)
    fresh=text.strip().lower().startswith('compare ')
    if fresh:
        query,start,end=parse_command(text)
        root=_export(config,'Ledger','Name')
        names=[n.attrib.get('NAME') or _text(n,'NAME') for n in root.findall('.//LEDGER')]
        parties=rank_names(query,[n for n in names if n])
        if not parties: raise ValueError('No ledgers returned by Tally.')
        state=dict(stage='party',parties=parties,start=start,end=end,company=config['tally']['company_name'],
                   dms_company=str(config['dms']['company_id']),created=datetime.now(timezone.utc).isoformat())
    else:
        if not has_pending(scope): raise ValueError('No active comparison. Send compare <customer> <date>.')
        state=json.loads(path.read_text(encoding='utf-8'))
        if state['company']!=config['tally']['company_name']:
            raise ValueError('Working company changed. Start a new comparison.')
        if state.get('dms_company', str(config['dms']['company_id']))!=str(config['dms']['company_id']):
            raise ValueError('DMS company changed. Start a new comparison.')
        stage=state['stage']
        if text.strip().lower()=='cancel': state['stage']='cancelled'
        elif stage=='party':
            state['party']=selection(text,state['parties'])[0]
            state['tally_options'],state['tally_excluded']=fetch_tally(config,state['party'],state['start'],state['end'])
            if not state['tally_options']: raise ValueError('No active sales invoices for this party in the selected seven days.')
            state['stage']='tally'
        elif stage=='tally':
            state['tally']=selection(text,state['tally_options'],True)
            from .dms_sales import fetch_sales
            state['all_dms'],state['dms_excluded']=dms_invoices(fetch_sales(config,state['start'],state['end']),state['start'],state['end'])
            names={r['party'] for r in state['all_dms']}
            if state['party'] in names: _dms_party(state,state['party'])
            else:
                state['dms_parties']=rank_names(state['party'],names)
                if not state['dms_parties']: raise ValueError('No DMS sales invoices found in this period.')
                state['stage']='dms_party'
        elif stage=='dms_party': _dms_party(state,selection(text,state['dms_parties'])[0])
        elif stage=='dms':
            state['dms']=selection(text,state['dms_options'],True)
            _prepare_mapping(state)
        elif stage=='mapping':
            if text.strip().lower() not in {'yes','no'}: raise ValueError('Reply yes or no.')
            state['mapping']=dict(state['proposals']) if text.strip().lower()=='yes' else {}
            state['stage']='confirm'
        elif stage=='confirm':
            if text.strip().lower()!='yes': raise ValueError('Reply yes to create the comparison, or cancel.')
            from .sales_compare_report import build_report
            rows=compare_records(state['tally'],state['dms'],state['mapping'])
            state['message'],output=build_report(state,rows,config)
            state['files']=[str(output)]
            state['stage']='done'
    state['updated']=datetime.now(timezone.utc).timestamp()
    path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps(state,ensure_ascii=False),encoding='utf-8');temp.replace(path)
    return _prompt(state)

import json
import sys
from pathlib import Path
from datetime import datetime, timedelta, timezone
from .settings import PROJECT_ROOT, load_config


def execute(data):
    from .web_app import file_path
    c = load_config()
    kind = data['kind']
    outputs = []
    if kind == 'connections':
        from .tally_context import ensure_configured_company
        results = []
        try:
            r = ensure_configured_company()
            results.append('Tally connected: '+r['company_name'])
        except Exception as exc:
            results.append('Tally: '+str(exc))
        # Do not start interactive OAuth from a background process.
        results.append('Gmail credentials: '+('present' if (PROJECT_ROOT/'data/token.json').exists() else 'authorization required'))
        results.append('DMS: use Fetch DMS Stock to verify the saved session')
        return dict(message='\n'.join(results), files=[])
    if kind in {'mail','purchase'}:
        from .purchase_processor import build_summary, default_output_filename
        if kind == 'mail':
            from .gmail_monitor import get_gmail_service, _download_attachments
            from .attachment_rules import select_invoice_excels
            from google.oauth2.credentials import Credentials
            from google.auth.transport.requests import Request
            from .gmail_monitor import SCOPES
            token = PROJECT_ROOT/'data/token.json'
            if not token.exists():
                raise ValueError('Run authorize_gmail.bat on the bot laptop first.')
            creds = Credentials.from_authorized_user_file(str(token), SCOPES)
            if not creds.valid:
                if not creds.refresh_token:
                    raise ValueError('Gmail authorization required; run authorize_gmail.bat.')
                creds.refresh(Request())
            service = get_gmail_service(PROJECT_ROOT)
            day = datetime.strptime(data['date'],'%Y-%m-%d').replace(tzinfo=timezone(timedelta(hours=5,minutes=30)))
            query=f'from:{c["gmail"]["sender"]} after:{int(day.timestamp())} before:{int((day+timedelta(days=1)).timestamp())}'
            sources=[]; page=None
            while True:
                result=service.users().messages().list(userId='me',q=query,maxResults=100,pageToken=page).execute()
                for ref in result.get('messages',[]):
                    msg=service.users().messages().get(userId='me',id=ref['id']).execute()
                    sources += select_invoice_excels(_download_attachments(service,ref['id'],msg,Path(c['paths']['downloads'])))
                page=result.get('nextPageToken')
                if not page: break
        else:
            sources=[file_path(data['file'],'uploads')]
        if not sources:
            return dict(message='No invoice Excel attachments found for this date.',files=[])
        from .tally_item_master import get_item_master
        master=get_item_master(force_refresh=True)
        if master['source'] not in {'tally','today_cache'}:
            raise ValueError('Fresh item master unavailable. Check the selected Tally company before processing.')
        messages=[]
        for source in sources:
            r=build_summary(source,Path(c['paths']['processed'])/default_output_filename(source))
            outputs.append(r['output'])
            note=f"{r['invoice_no']}: {r['total_case']:g} cases; {r['output_rows']} rows." + (' New items: '+', '.join(r['new_items']) if r['new_items'] else '')
            if r.get('scheme_confirmations'):
                note += ' Confirm scheme: ' + '; '.join(r['scheme_confirmations'])
            messages.append(note)
        message='\n'.join(messages)
    elif kind == 'unloading':
        from .unloading_slip import build_unloading_slip, parse_user_date
        from .undispatched import parse_bills
        if data.get('dispatch') not in {'none','bills'}:
            raise ValueError('Choose None or enter undispatched sales bills.')
        bills=parse_bills(data.get('bills','')) if data['dispatch']=='bills' else []
        if data['dispatch']=='bills' and not bills: raise ValueError('Enter sales bill numbers')
        invoices=[x.strip() for x in data.get('invoices','').replace('\n',',').split(',') if x.strip()]
        r=build_unloading_slip(parse_user_date(data['date']),invoices,bills)
        outputs=[r['xlsx'],r['pdf']]
        message=f"{r['rows']} rows; {r['total_cases']:g} incoming cases. Invoices: "+', '.join(r['invoice_numbers'])
        if bills: message+='; Stock add-back: '+', '.join(bills)
    elif kind == 'challans':
        from .delivery_vouchers import build_delivery_vouchers,default_output_filename
        source=file_path(data['file'],'uploads')
        r=build_delivery_vouchers(source,Path(c['paths']['processed'])/default_output_filename(source))
        outputs=[r['output']]; message=f"{r['generated']} challans created."
        if r['errors']: message+='\nRows not generated:\n'+'\n'.join(r['errors'])
    elif kind == 'stock':
        from .tally_batch_stock import build
        path=Path(c['paths']['processed'])/('Stock Status '+datetime.now().strftime('%Y-%m-%d %H-%M-%S')+'.xlsx')
        r=build(path)
        if r['errors']: raise ValueError('Stock export contains errors; review Tally before retrying.')
        outputs=[str(path)];message=f"{r['rows']} batch rows exported."
    elif kind == 'dms':
        from .dms_client import build_inventory_workbook
        import shutil
        r=build_inventory_workbook(c)
        source=Path(r['output']);dest=Path(c['paths']['processed'])/source.name
        if source.resolve()!=dest.resolve(): shutil.copy2(source,dest)
        outputs=[str(dest)];message=f"{r['rows']} DMS rows exported."
    return dict(message=message,files=[Path(p).name for p in outputs])


if __name__=='__main__':
    job=json.loads((PROJECT_ROOT/'data/web_jobs'/ (sys.argv[1]+'.json')).read_text())
    print(json.dumps(execute(job['request'])))

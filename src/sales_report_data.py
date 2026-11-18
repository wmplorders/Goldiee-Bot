"""Read-only sales/payment data, using bounded Tally voucher-entry batches."""
import calendar
import html
import re
import uuid
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal

from .tally_context import get_tally_config, ensure_configured_company, tally_request, clean_tally_xml


def parse_date(value):
    for fmt in ('%d/%m/%Y','%d/%m/%y','%d-%m-%Y','%d-%m-%y','%Y-%m-%d'):
        try:return datetime.strptime(value.strip(),fmt).date()
        except ValueError:pass
    raise ValueError('Invalid date. Use DD/MM/YYYY or DD/MM/YY.')


def parse_command(text):
    match=re.fullmatch(r'sales\s+report\s+(.+?)(?:\s+to\s+(.+))?',text.strip(),re.I)
    if not match:raise ValueError('Use: sales report 1/11/26 to 18/11/26')
    if match[2]:start,end=parse_date(match[1]),parse_date(match[2])
    else:
        end=parse_date(match[1]);start=end.replace(day=1)
    if start>end or (start.year,start.month)!=(end.year,end.month):
        raise ValueError('Choose an ordered date range within a single month.')
    return start,end


def amount(value):
    # Empty native Tally amounts represent zero; non-numeric text is never guessed.
    return Decimal((value or '').strip() or '0').quantize(Decimal('.01'))


def text(node,tag):return (node.findtext(tag) or '').strip()


def export(config,kind,fetch,start,end,condition='',debtors=False,balances=False):
    name='SalesReport'+uuid.uuid4().hex
    fy=date(end.year if end.month>=4 else end.year-1,4,1)
    period=(f'<COMPUTE>ReportClosing : $$FromValue:&quot;{fy:%d-%m-%Y}&quot;:'
            f'$$ToValue:&quot;{end:%d-%m-%Y}&quot;:$ClosingBalance</COMPUTE>') if balances else ''
    request=('<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
             f'<TYPE>Collection</TYPE><ID>{name}</ID></HEADER><BODY><DESC><STATICVARIABLES>'
             '<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>'
             f'<SVCurrentCompany>{html.escape(config["tally"]["company_name"])}</SVCurrentCompany>'
             f'<SVFROMDATE>{start:%Y%m%d}</SVFROMDATE><SVTODATE>{end:%Y%m%d}</SVTODATE>'
             f'</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME="{name}">'
             f'<TYPE>{kind}</TYPE><FETCH>{fetch}</FETCH>{period}'
             +('<CHILDOF>$$GroupSundryDebtors</CHILDOF><BELONGSTO>Yes</BELONGSTO>' if debtors else '')
             +(f'<FILTER>{name}Filter</FILTER>' if condition else '')+'</COLLECTION>'
             +(f'<SYSTEM TYPE="Formulae" NAME="{name}Filter">{condition}</SYSTEM>' if condition else '')
             +'</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>')
    try:
        response=tally_request(request,get_tally_config(config),timeout=60)
    except TimeoutError as exc:
        raise TimeoutError('Tally export timed out. No report was generated. Wait until Tally is responsive before retrying.') from exc
    root=ET.fromstring(clean_tally_xml(response,include_udf_namespace=True))
    errors=[n.text for n in root.findall('.//LINEERROR')]
    if errors:raise ValueError('Tally export: '+'; '.join(errors))
    return root


def customer_sets(groups,ledgers):
    parents={n.attrib['NAME']:text(n,'PARENT') for n in groups.findall('.//GROUP') if n.attrib.get('NAME')}
    closed={name for name in parents if name.casefold()=='closed customers'}
    while True:
        descendants={name for name,parent in parents.items() if parent in closed}
        if descendants<=closed:break
        closed.update(descendants)
    active,excluded={},{}
    for node in ledgers.findall('.//LEDGER'):
        name=node.attrib.get('NAME')
        if not name:continue
        parent=text(node,'PARENT')
        if parent in closed:excluded[name]='Closed Customers'
        elif re.match(r'^(?:TDS|TCS)\s+\d',name,re.I):excluded[name]='Tax adjustment ledger, not a customer'
        else:active[name]={'name':name,'id':text(node,'GUID') or name,'parent':parent}
    return active,excluded


def ledger_balances(root,names):
    values={}
    for node in root.findall('.//LEDGER'):
        name=node.attrib.get('NAME')
        if name not in names:continue
        method=node.find('REPORTCLOSING')
        if method is None:raise ValueError(f'Date-specific closing balance unavailable for {name}.')
        values[name]=amount(method.text)
    if set(values)!=set(names):raise ValueError('Some customer balance records are missing. Retry the report.')
    return values


def collect(config,start,end):
    ensure_configured_company(get_tally_config(config))
    month=start.replace(day=1)
    groups=export(config,'Group','Name,Parent',month,end)
    ledgers=export(config,'Ledger','Name,GUID,Parent,OpeningBalance,ReportClosing',month,end,debtors=True,balances=True)
    customers,excluded=customer_sets(groups,ledgers)
    if not customers:raise ValueError('No active customer ledgers found under Sundry Debtors.')
    closing=ledger_balances(ledgers,customers)
    previous=month-timedelta(days=1)
    old=export(config,'Ledger','Name,ReportClosing',previous.replace(day=1),previous,debtors=True,balances=True)
    opening=ledger_balances(old,customers)
    headers=export(config,'Voucher','Date,VoucherNumber,VoucherTypeName,MasterID,GUID,IsCancelled,IsOptional',month,end)
    refs={}
    for v in headers.findall('.//VOUCHER'):
        raw=text(v,'DATE')
        if not month.strftime('%Y%m%d')<=raw<=end.strftime('%Y%m%d'):continue
        mid=int(text(v,'MASTERID'))
        if mid in refs:raise ValueError('Duplicate Tally voucher ID in header export.')
        refs[mid]=raw
    vouchers=[];ids=sorted(refs)
    for offset in range(0,len(ids),10):
        batch=ids[offset:offset+10]
        condition=' OR '.join(f'$MasterID = {i}' for i in batch)
        root=export(config,'Voucher','Date,VoucherNumber,VoucherTypeName,PartyLedgerName,MasterID,GUID,IsCancelled,IsOptional,LedgerEntries.*',month,end,condition=condition)
        found=[v for v in root.findall('.//VOUCHER') if v.find('DATE') is not None]
        returned=[int(text(v,'MASTERID')) for v in found]
        if len(returned)!=len(batch) or set(returned)!=set(batch):
            raise ValueError('Tally voucher batch was incomplete or changed during export. Retry the report.')
        if any(text(v,'DATE')!=refs[int(text(v,'MASTERID'))] for v in found):
            raise ValueError('A voucher changed date during export. Retry the report.')
        vouchers.extend(found)
    # A second balance read catches edits made while the batches were being downloaded.
    final=export(config,'Ledger','Name,GUID,Parent,OpeningBalance,ReportClosing',month,end,debtors=True,balances=True)
    final_groups=export(config,'Group','Name,Parent',month,end)
    final_customers,final_excluded=customer_sets(final_groups,final)
    if final_customers!=customers or final_excluded!=excluded or ledger_balances(final,customers)!=closing:
        raise ValueError('Customers or balances changed during export. Please retry when Tally is idle.')
    return aggregate(customers,excluded,opening,closing,vouchers,start,end)


def aggregate(customers,excluded,opening,closing,vouchers,start,end):
    daily=defaultdict(lambda:[Decimal(0),Decimal(0),Decimal(0)])
    movement=defaultdict(Decimal);counts=Counter();audit=[]
    for v in vouchers:
        day=datetime.strptime(text(v,'DATE'),'%Y%m%d').date()
        if not start.replace(day=1)<=day<=end:raise ValueError('Voucher lies outside fetched period.')
        if text(v,'ISCANCELLED').casefold()=='yes' or text(v,'ISOPTIONAL').casefold()=='yes':
            counts['cancelled_or_optional']+=1;continue
        entries=v.findall('LEDGERENTRIES.LIST') or v.findall('ALLLEDGERENTRIES.LIST')
        if any(not text(n,'LEDGERNAME') and amount(n.findtext('AMOUNT')) for n in entries):
            raise ValueError('An accounting entry has an amount but no ledger name.')
        ledger_rows=[(text(n,'LEDGERNAME'),amount(n.findtext('AMOUNT'))) for n in entries if text(n,'LEDGERNAME')]
        for name,val in ledger_rows:
            if name in customers:movement[name]+=val
        if day<start:continue
        typ=text(v,'VOUCHERTYPENAME');number=text(v,'VOUCHERNUMBER')
        if 'credit note' in typ.casefold():counts['credit_notes_ignored_in_sales']+=1
        if typ in {'Tax Invoice','Bill of Supply','Detergent Invoice'}:
            party=text(v,'PARTYLEDGERNAME')
            if party in excluded:counts['excluded_customer_invoices']+=1
            else:
                if party not in customers:raise ValueError(f'{number}: sales party {party!r} is not an active customer under Sundry Debtors.')
                vals=[a for n,a in ledger_rows if n==party]
                if len(vals)!=1 or vals[0]>0:raise ValueError(f'{number}: ambiguous customer invoice value.')
                category=1 if typ=='Detergent Invoice' else 0
                daily[(party,day)][category]-=vals[0];counts[typ]+=1
                audit.append({'date':day.isoformat(),'invoice':number,'party':party,'kind':'detergent' if category else 'spices','amount':str(-vals[0])})
        main=sum((a for n,a in ledger_rows if n.casefold()=='main branch'),Decimal(0))
        if main>0 and any(n in customers and a<0 for n,a in ledger_rows):
            raise ValueError(f'{number}: reversal involving Main Branch found; review before generating the report.')
        if main<0:
            for name,val in ledger_rows:
                if name in customers and val>0:
                    daily[(name,day)][2]+=val;counts['payments']+=1
                    audit.append({'date':day.isoformat(),'invoice':number,'party':name,'kind':'payment','amount':str(val)})
    bad={n:str(closing[n]-opening[n]-movement[n]) for n in customers if closing[n]-opening[n]-movement[n]!=0}
    if bad:raise ValueError('Closing balance reconciliation failed: '+str(bad))
    days=calendar.monthrange(end.year,end.month)[1]
    rows=[]
    for name in sorted(customers,key=str.casefold):
        values=[]
        for i in range(1,days+1):
            day=date(end.year,end.month,i)
            values.extend([float(x) if x else None for x in daily[(name,day)]] if start<=day<=end else [None]*3)
        rows.append({'name':name,'id':customers[name]['id'],'values':values,
                     'debit':float(-closing[name]) if closing[name]<0 else None,
                     'credit':float(closing[name]) if closing[name]>0 else None})
    totals={kind:float(sum((Decimal(r['amount']) for r in audit if r['kind']==kind),Decimal(0))) for kind in ('spices','detergent','payment')}
    totals['debit']=float(sum((-v for v in closing.values() if v<0),Decimal(0)))
    totals['credit']=float(sum((v for v in closing.values() if v>0),Decimal(0)))
    return {'company':None,'start':start.isoformat(),'end':end.isoformat(),'days':days,'rows':rows,
            'totals':totals,'counts':dict(counts),'excluded':excluded,'audit':audit}

"""Read-only sales bill add-backs for unloading reports."""
import copy
import html
import re
import xml.etree.ElementTree as ET

from . import tally_batch_stock as t
from .tally_context import get_tally_config, tally_request, clean_tally_xml, current_company_static_variables, ensure_configured_company


PROMPT = ('Are any billed goods still not dispatched? Reply with the full bill numbers '
          'separated by commas or new lines, or reply `none`. Reply `cancel` to cancel.')


def parse_bills(text):
    if text.strip().casefold() in {'none', 'no', 'nil', '0'}:
        return []
    bills = list(dict.fromkeys(re.split(r'[\s,;]+', text.strip())))
    if not bills or any(not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9/._-]*', b) or not any(c.isdigit() for c in b) for b in bills):
        raise ValueError('Send full bill numbers separated by commas, or reply none.')
    return bills


def fetch_additions(bills):
    config = get_tally_config()
    ensure_configured_company(config)
    additions = []
    for bill in dict.fromkeys(bills):
        if parse_bills(bill) != [bill]:
            raise ValueError('Invalid bill number: ' + bill)
        body = ('<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
                '<TYPE>Collection</TYPE><ID>UndispatchedBills</ID></HEADER><BODY><DESC>'
                '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>'
                + current_company_static_variables(config) + '</STATICVARIABLES><TDL><TDLMESSAGE>'
                '<COLLECTION NAME="UndispatchedBills"><TYPE>Voucher</TYPE>'
                '<FETCH>*,AllInventoryEntries.*,BatchAllocations.*</FETCH><FILTER>BillMatch</FILTER>'
                '</COLLECTION><SYSTEM TYPE="Formulae" NAME="BillMatch">$VoucherNumber = "'
                + html.escape(bill) + '"</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>')
        root = ET.fromstring(clean_tally_xml(tally_request(body, config, timeout=60), True))
        matches = [v for v in root.findall('.//VOUCHER') if t.text(v, 'VOUCHERNUMBER') == bill]
        if len(matches) != 1:
            raise ValueError(f'{bill}: expected one voucher; found {len(matches)}. No slip created.')
        v = matches[0]
        if t.text(v, 'ISCANCELLED') != 'No' or t.text(v, 'ISOPTIONAL') != 'No':
            raise ValueError(f'{bill}: cancelled or optional voucher.')
        entries = v.findall('ALLINVENTORYENTRIES.LIST')
        sales_ledgers = [t.text(a, 'LEDGERNAME') for i in entries for a in i.findall('ACCOUNTINGALLOCATIONS.LIST')]
        if not entries or not sales_ledgers or any(x != 'Sales' for x in sales_ledgers):
            raise ValueError(f'{bill}: not a supported sales invoice (Sales ledger required).')
        for item in entries:
            name = t.text(item, 'STOCKITEMNAME')
            qty, unit = t.parse_qty(t.text(item, 'ACTUALQTY'))
            if qty <= 0 or t.text(item, 'ISDEEMEDPOSITIVE') != 'No':
                raise ValueError(f'{bill}: invalid outward quantity for {name}.')
            total = 0
            for batch in item.findall('BATCHALLOCATIONS.LIST'):
                amount, batch_unit = t.parse_qty(t.text(batch, 'ACTUALQTY'))
                mrp = t.udf_value(batch, 'UDF_AA_BATCH_MRP_RATE')
                ptr = t.text(batch, 'BATCHNAME')
                if amount <= 0 or batch_unit != unit or not ptr or not isinstance(mrp, (float, int)) or mrp <= 0:
                    raise ValueError(f'{bill}: missing/invalid batch quantity or MRP for {name}.')
                additions.append(dict(bill=bill, item=name, ptr=ptr, mrp=mrp, qty=amount, unit=unit))
                total += amount
            if abs(total - qty) > .00001:
                raise ValueError(f'{bill}: batch quantities do not reconcile for {name}.')
    return additions


def apply_additions(stock, additions, master_path):
    stock = copy.deepcopy(stock)
    masters = ET.fromstring(t.clean_xml(master_path))
    factors = {}
    for node in masters.findall('.//STOCKITEM'):
        d, c = t.number(t.text(node, 'DENOMINATOR')), t.number(t.text(node, 'CONVERSION'))
        if isinstance(d, (float, int)) and isinstance(c, (float, int)) and c and d > 0:
            factors[node.attrib['NAME']] = (d / c, t.text(node, 'BASEUNITS'))
    audit = []
    for a in additions:
        factor, unit = factors.get(a['item'], (0, ''))
        if factor <= 0 or unit != a['unit']:
            raise ValueError('Missing item-master conversion or unit mismatch: ' + a['item'])
        rows = stock.setdefault(a['item'], [])
        # PTR spelling identifies distinct Tally batches; never merge 15.0 with 15.00.
        hits = [r for r in rows if r['ptr'] == a['ptr'] and r['mrp'] == f"{a['mrp']:.2f}"]
        before = sum(r['qty'] for r in hits)
        if hits:
            hits[0]['qty'] += a['qty']
            hits[0]['case_qty'] = hits[0]['qty'] / factor
        else:
            rows.append(dict(item=a['item'], ptr=a['ptr'], mrp=f"{a['mrp']:.2f}", qty=a['qty'], case_qty=a['qty']/factor, pcs_per_case=factor))
        audit.append([a['bill'], a['item'], a['ptr'], a['mrp'], before, a['qty'], before+a['qty'], factor, a['qty']/factor])
    return stock, audit

"""Exception-first reconciliation workbook. No changes to either source system."""
import re
import uuid
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .sales_compare import FIELDS, LABELS, money


def total(records, key):
    return sum((money(x[key]) for x in records), money(0))


def reconcile(state):
    tally, dms = state['tally'],state['dms']
    tl=[l for r in tally for l in r['lines']]
    dl=[l for r in dms for l in r['lines']]
    result=[['Quantity (PCS)',float(total(tl,'qty')),float(total(dl,'qty'))],
            ['Taxable value',float(total(tl,'taxable')),float(total(dl,'taxable'))]]
    tt,dt=defaultdict(lambda:money(0)),defaultdict(lambda:money(0))
    other_t=money(0)
    for record in tally:
        for ledger in record['ledgers']:
            match=re.search(r'\b(CGST|SGST|IGST)\b.*?(\d+(?:\.\d+)?)\s*%',ledger['name'],re.I)
            if match:
                tt[(match[1].upper(),money(match[2]))]+=money(ledger['amount'])
            elif ledger['name']!=state['party']:
                other_t+=money(ledger['amount'])
    for line in dl:
        for tax in ('cgst','sgst','igst'):
            if money(line[tax]) or money(line[tax+'_rate']):
                dt[(tax.upper(),money(line[tax+'_rate']))]+=money(line[tax])
    for key in sorted(set(tt)|set(dt)):
        result.append([f'{key[0]} @ {key[1]}%',float(tt[key]),float(dt[key])])
    tv,dv=total(tally,'value'),total(dms,'value')
    if total(tl,'taxable')+sum(tt.values(),money(0))+other_t!=tv:
        raise ValueError('Exported Tally item/ledger amounts do not reconcile to selected invoice totals.')
    other_d=dv-total(dl,'taxable')-sum(dt.values(),money(0))
    result += [['Round off / other charges',float(other_t),float(other_d)],['Final invoice value',float(tv),float(dv)]]
    return result


def build_report(state, rows, config):
    totals=reconcile(state)
    bad=[r for r in rows if not r['matched']]
    tv,dv=totals[-1][1:]
    delta=money(dv)-money(tv)
    overview=(f"{state['party']}\nDMS is Rs. {abs(delta):,.2f} {'higher' if delta>0 else 'lower'} than Tally."
              if delta else f"{state['party']}\nFinal invoice totals match.")
    lines=[overview,f"{len(rows)-len(bad)} matching combinations; {len(bad)} with differences.",
           'Tally: '+', '.join(r['invoice'] for r in state['tally']),
           'DMS: '+', '.join(r['invoice'] for r in state['dms'])]
    exceptions=[]
    for row in bad:
        t,d=row['t'],row['d']
        name=(d or t)['key']
        if t and d:
            changes=[]
            for field,label in zip(FIELDS,LABELS):
                if money(t[field])!=money(d[field]):
                    changes.append(f"{label}: Tally {t[field]:,.2f}; DMS {d[field]:,.2f}")
                    exceptions.append([name,label,t[field],d[field]])
            lines.append(name+'\n'+'; '.join(changes))
        else:
            lines.append(name+': '+row['remarks'])
            exceptions.append([name,row['remarks'],t['qty'] if t else None,d['qty'] if d else None])
    warnings=[]
    for r in state['dms']:
        for line in r['lines']:
            if line.get('free_qty'):
                warnings.append(f"DMS free quantity for {line['name']}: {line['free_qty']:g} PCS; review billed versus physical quantity.")
    lines+=warnings
    tax_differences=[f"{label}: Tally {t:,.2f}; DMS {d:,.2f}" for label,t,d in totals
                     if ('GST' in label or label=='Round off / other charges') and money(t)!=money(d)]
    if tax_differences:
        lines.append('Tax / rounding differences:\n'+'\n'.join(tax_differences))
    wb=Workbook();wb.remove(wb.active)
    header_fill=PatternFill('solid',fgColor='DCECE9')
    red=PatternFill('solid',fgColor='FCE4D6');green=PatternFill('solid',fgColor='E2F0D9')
    def sheet(name,headers,values,widths):
        ws=wb.create_sheet(name);ws.append(headers)
        for v in values:ws.append(v)
        for cell in ws[1]:cell.font=Font(bold=True);cell.fill=header_fill
        for i,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(i)].width=w
        ws.freeze_panes='B2';ws.auto_filter.ref=ws.dimensions
        return ws
    snap=sheet('Snapshot',['Item','Difference','Tally','DMS'],exceptions or [['All compared item combinations match','','','']],[58,45,23,23])
    for cells in list(snap.rows)[1:]:
        for cell in cells:cell.fill=red if bad else green
    snap.append([]);snap.append([overview.replace('\n',' - ')]);snap.merge_cells(start_row=snap.max_row,start_column=1,end_row=snap.max_row,end_column=4)
    snap.append([f'{len(rows)-len(bad)} matching combinations; {len(bad)} with differences.'])
    snap.append(['Tally: '+', '.join(r['invoice']+' ('+r['date']+')' for r in state['tally'])])
    snap.append(['DMS: '+', '.join(r['invoice']+' ('+r['date']+')' for r in state['dms'])])
    for r in range(snap.max_row-2,snap.max_row+1):snap.merge_cells(start_row=r,start_column=1,end_row=r,end_column=4)
    snap.append([]);snap.append(['Component','Tally','DMS','Tally minus DMS'])
    for label,t,d in totals:
        r=snap.max_row+1;snap.append([label,t,d,f'=ROUND(B{r}-C{r},2)'])
        snap.cell(r,4).fill=green if money(t)==money(d) else red
    snap.append(['Two decimals; zero tolerance. Differences = Tally minus DMS.'])
    snap.merge_cells(start_row=snap.max_row,start_column=1,end_row=snap.max_row,end_column=4)
    snap.append(['Tax comparison uses actual Tally voucher ledgers versus DMS exported taxes, not inferred per-item Tally taxes.'])
    snap.merge_cells(start_row=snap.max_row,start_column=1,end_row=snap.max_row,end_column=4)
    for warning in warnings:
        snap.append([warning]);snap.merge_cells(start_row=snap.max_row,start_column=1,end_row=snap.max_row,end_column=4)
    headers=['Tally item','DMS item']
    for label in LABELS:headers.extend(['Tally '+label,'DMS '+label])
    headers+=['Taxable difference','Remarks','Tally invoices','DMS invoices']
    detail=[]
    for row in rows:
        t,d=row['t'],row['d']
        values=['; '.join(t['original_names']) if t else '', '; '.join(d['original_names']) if d else '']
        for key in FIELDS:values.extend([t[key] if t else None,d[key] if d else None])
        values += [None,row['remarks'],', '.join(t['invoices']) if t else '',', '.join(d['invoices']) if d else '']
        detail.append(values)
    ws=sheet('Item Comparison',headers,detail,[56,56]+[20]*16+[23,55,35,35])
    for i,row in enumerate(rows,2):
        if row['matched']:
            for cell in ws[i]:cell.fill=green
        elif row['t'] and row['d']:
            for j,key in enumerate(FIELDS):
                if money(row['t'][key])!=money(row['d'][key]):
                    ws.cell(i,3+j*2).fill=red;ws.cell(i,4+j*2).fill=red
        else:
            for cell in ws[i]:cell.fill=red
        if row['t'] and row['d']:ws.cell(i,19,f'=ROUND(Q{i}-R{i},2)')
        if not row['matched']:ws.cell(i,20).fill=red
    keys=['invoice','name','qty','ptr','mrp','rate','discount','taxable','cgst_rate','sgst_rate','igst_rate']
    sheet('Tally Detail',keys,[[l.get(k,'') for k in keys] for r in state['tally'] for l in r['lines']],[29,60]+[19]*9)
    sheet('Tally Ledgers',['Invoice','Ledger','Amount'],[[r['invoice'],l['name'],l['amount']] for r in state['tally'] for l in r['ledgers']],[30,65,23])
    raw=[line for r in state['dms'] for line in r['raw']]
    if raw:
        keys=list(raw[0]);sheet('DMS Raw',keys,[[l.get(k) for k in keys] for l in raw],[max(18,min(60,len(k)+3)) for k in keys])
    meta=[['Tally company',state['company']],['DMS company ID',str(config['dms']['company_id'])],
          ['Period',state['start']+' to '+state['end']],['Created',datetime.now().isoformat(timespec='seconds')],
          ['Source','Read-only Tally XML and fresh DMS Sales Invoice Dump'],
          ['Effective discount','When multiple discount fields are active: Total Discount / Gross Amount x 100, rounded to two decimals. Original values retained in DMS Raw.'],
          ['Name matches','For this comparison only: '+str(state.get('mapping') or 'Exact names only')],
          ['Tally exclusions',state.get('tally_excluded',0)],['DMS excluded rows',state.get('dms_excluded',0)]]
    sheet('Sources',['Field','Value'],meta,[30,120])
    # Prevent names or DMS text from becoming executable Excel formulas.
    for tab in wb:
        for row in tab:
            tab.row_dimensions[row[0].row].height=15
            for cell in row:
                cell.alignment=Alignment(vertical='center',wrap_text=False)
                cell.font=Font(name='Calibri',size=11,bold=cell.row==1)
                if isinstance(cell.value,(float,int)):cell.number_format='#,##0.00;[Red]-#,##0.00'
                if isinstance(cell.value,str) and cell.data_type=='f':
                    if not (tab.title in {'Snapshot','Item Comparison'} and re.fullmatch(r'=ROUND\([A-Z]+\d+-[A-Z]+\d+,2\)',cell.value)):
                        cell.data_type='s'
    destination=Path(config['paths']['processed']);destination.mkdir(parents=True,exist_ok=True)
    name=re.sub(r'[^A-Za-z0-9 ._-]','_',state['party'])[:60]
    output=destination/f'Sales Compare {name} {state["end"]} {uuid.uuid4().hex[:8]}.xlsx'
    wb.save(output)
    return '\n\n'.join(lines),output

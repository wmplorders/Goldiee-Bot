import copy
import json
import tempfile
import unittest
import xml.etree.ElementTree as ET
from datetime import date
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from src import sales_report as flow
from src import sales_report_data as data
from src import sales_report_output as output


def voucher(kind,entries,day='20261114',party='Party',optional=False):
    root=ET.Element('VOUCHER')
    for k,v in {'DATE':day,'VOUCHERTYPENAME':kind,'VOUCHERNUMBER':'1','PARTYLEDGERNAME':party,'ISOPTIONAL':'Yes' if optional else 'No'}.items():ET.SubElement(root,k).text=v
    for name,amount in entries:
        row=ET.SubElement(root,'ALLLEDGERENTRIES.LIST' if kind=='Journal' else 'LEDGERENTRIES.LIST')
        ET.SubElement(row,'LEDGERNAME').text=name;ET.SubElement(row,'AMOUNT').text=str(amount)
    return root


def payload():
    customers={'Party':{'id':'id','name':'Party','parent':'Town'}}
    vs=[voucher('Tax Invoice',[('Party',-105)]),voucher('Bill of Supply',[('Party',-20)]),
        voucher('Detergent Invoice',[('Party',-10)]),voucher('Journal',[('Main Branch',-50),('Party',50)]),
        voucher('GST Credit Note',[('Party',10)])]
    p=data.aggregate(customers,{}, {'Party':Decimal(0)}, {'Party':Decimal(-75)},vs,date(2026,11,1),date(2026,11,18))
    p['company']='Company';p['rows'][0]['town']='Town'
    return p


class SalesReportTests(unittest.TestCase):
    def test_date_range(self):
        self.assertEqual(data.parse_command('sales report 1/11/26 to 18/11/26'),(date(2026,11,1),date(2026,11,18)))
        self.assertEqual(data.parse_command('sales report 2026-11-18'),(date(2026,11,1),date(2026,11,18)))
        for command in ('sales report 31/11/26','sales report 18/11/26 to 1/11/26','sales report 1/10/26 to 18/11/26'):
            with self.assertRaises(ValueError):data.parse_command(command)

    def test_closed_descendants_and_tax_ledger(self):
        groups=ET.fromstring('<R><GROUP NAME="Closed Customers"><PARENT>Sundry Debtors</PARENT></GROUP><GROUP NAME="Old Town"><PARENT>Closed Customers</PARENT></GROUP></R>')
        ledgers=ET.fromstring('<R><LEDGER NAME="Active"><PARENT>Sundry Debtors</PARENT></LEDGER><LEDGER NAME="Closed"><PARENT>Old Town</PARENT></LEDGER><LEDGER NAME="TDS 194Q : Someone"><PARENT>Sundry Debtors</PARENT></LEDGER></R>')
        active,excluded=data.customer_sets(groups,ledgers)
        self.assertEqual(set(active),{'Active'});self.assertEqual(excluded['Closed'],'Closed Customers')

    def test_sales_payments_and_balances(self):
        p=payload();self.assertEqual(p['totals'],{'spices':125.0,'detergent':10.0,'payment':50.0,'debit':75.0,'credit':0.0})
        self.assertEqual(p['counts']['credit_notes_ignored_in_sales'],1)
        self.assertEqual(p['rows'][0]['values'][39:42],[125.0,10.0,50.0])
        self.assertTrue(all(x is None for x in p['rows'][0]['values'][54:]))

    def test_no_sale_from_optional(self):
        p=data.aggregate({'Party':{'id':'p'}},{},{'Party':Decimal(0)},{'Party':Decimal(0)},
            [voucher('Tax Invoice',[('Party',-10)],optional=True)],date(2026,11,1),date(2026,11,18))
        self.assertEqual(p['totals']['spices'],0)

    def test_balances_include_prior_days_and_all_adjustments(self):
        vs=[voucher('Tax Invoice',[('Party',-100)],day='20261102'),voucher('GST Credit Note',[('Party',10)],day='20261115')]
        p=data.aggregate({'Party':{'id':'p'}},{},{'Party':Decimal(0)},{'Party':Decimal(-90)},vs,date(2026,11,10),date(2026,11,18))
        self.assertEqual(p['totals']['spices'],0);self.assertEqual(p['totals']['debit'],90)

    def test_reconciliation_blocks_wrong_balance(self):
        with self.assertRaisesRegex(ValueError,'reconciliation'):
            data.aggregate({'Party':{'id':'p'}},{},{'Party':Decimal(0)},{'Party':Decimal(-1)},[],date(2026,11,1),date(2026,11,18))

    def test_main_branch_reversal_blocks(self):
        with self.assertRaisesRegex(ValueError,'reversal'):
            data.aggregate({'Party':{'id':'p'}},{},{'Party':Decimal(0)},{'Party':Decimal(-50)},
                [voucher('Journal',[('Main Branch',50),('Party',-50)])],date(2026,11,1),date(2026,11,18))

    def test_indian_grouping(self):
        for val,expected in [(0,'0'),(999,'999'),(1000,'1,000'),(100000,'1,00,000'),(9953773,'99,53,773'),(10000000,'1,00,00,000'),(-123456,'-1,23,456'),(10.5,'11')]:
            self.assertEqual(output.indian(val),expected)

    def test_snapshot_final_and_previous_calendar_day(self):
        p=payload();p['rows'][0]['values'][48:54]=[1,2,3,4,5,6]
        previous,end,rows=output.snapshot_rows(p)
        self.assertEqual((previous,end),(date(2026,11,17),date(2026,11,18)))
        self.assertEqual(rows[0][2:8],[1,2,3,4,5,6])

    def test_first_day_snapshot_does_not_borrow_wrong_month(self):
        p=payload();p['start']='2026-11-01';p['end']='2026-11-01'
        previous,end,rows=output.snapshot_rows(p)
        self.assertEqual(previous,date(2026,10,31));self.assertEqual(rows[0][2:5],[None,None,None])

    def test_outputs_styles_formulas_and_png(self):
        from PIL import Image
        p=payload()
        with tempfile.TemporaryDirectory() as tmp:
            xlsx,png=output.build(p,tmp,'Test')
            w=load_workbook(xlsx);s=w.active
            self.assertEqual(s.max_column,96);self.assertEqual(s.freeze_panes,'C3')
            self.assertTrue(all(d.height==15 for d in s.row_dimensions.values()))
            self.assertEqual(s['AY3'].border.left.style,'medium');self.assertEqual(s['BD4'].border.bottom.style,'medium')
            self.assertEqual(s['CO3'].data_type,'f');self.assertEqual(s['CO3'].number_format,output.INDIAN_FORMAT)
            self.assertEqual(s['CQ3'].value,75);self.assertIsNone(s['BE3'].value);w.close()
            with Image.open(png) as im:self.assertGreater(im.width,1500);self.assertGreater(im.height,150)

    def test_compact_png_resolution_and_unchanged_data(self):
        from PIL import Image
        p=payload();before=copy.deepcopy(p)
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'compact.png';output.png(p,path)
            with Image.open(path) as image:
                self.assertEqual(image.width,3989)
                self.assertAlmostEqual(image.info['dpi'][0],300,delta=1)
                self.assertEqual(image.format,'PNG')
            self.assertEqual(p,before)

    def test_full_names_and_long_towns_wrap_without_truncation(self):
        from PIL import Image,ImageDraw,ImageFont
        draw=ImageDraw.Draw(Image.new('RGB',(1,1)))
        font=ImageFont.load_default(size=20)
        for label in ['A VERY LONG CUSTOMER NAME C0012345(ROBERTSGANJ)','CUSTOMERIDENTIFIERWITHOUTSPACES'*4,'Town Name With Several Words']:
            lines=output._wrap_snapshot_text(label,100,font,draw)
            self.assertEqual(''.join(lines).replace(' ',''),label.replace(' ',''))
            self.assertTrue(all(draw.textlength(line,font=font)<=100 for line in lines))

    def test_large_figures_expand_columns_instead_of_clipping(self):
        from PIL import Image
        p=payload();large=copy.deepcopy(p);large['rows'][0]['debit']=1234567890123
        large['rows'][0]['name']='LONGCUSTOMERIDENTIFIER'*8
        large['rows'][0]['town']='LONGTOWNNAME'*5
        with tempfile.TemporaryDirectory() as tmp:
            a,b=Path(tmp)/'small.png',Path(tmp)/'large.png'
            output.png(p,a);output.png(large,b)
            with Image.open(a) as first,Image.open(b) as second:
                self.assertGreater(second.width,first.width)
                self.assertGreater(second.height,first.height)

    def test_month_length_february(self):
        p=payload();p['start']='2028-02-01';p['end']='2028-02-29';p['days']=29;p['rows'][0]['values']=[None]*87
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'feb.xlsx';output.excel(p,path);w=load_workbook(path)
            self.assertEqual(w.active.max_column,93);w.close()

    def test_town_import_and_conflict(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'towns.xlsx';w=Workbook();w.active.append(['Title']);w.active.append(['PARTY NAME','POST CODE']);w.active.append(['Party','Town']);w.save(path)
            registry={};self.assertEqual(flow.import_towns(path,payload(),registry),1);self.assertEqual(registry['id']['town'],'Town')
            with self.assertRaises(ValueError):flow.import_towns(path,payload(),{'id':{'name':'Party','town':'Other'}})

    def test_workflow_missing_town_saved_and_company_scoped(self):
        p=payload();p['rows'][0].pop('town');config={'tally':{'company_name':'Company'},'paths':{'processed':'unused'}}
        with tempfile.TemporaryDirectory() as tmp,patch.object(flow,'STATE_DIR',Path(tmp)/'sessions'),patch.object(flow,'TOWNS_PATH',Path(tmp)/'towns.json'),patch.object(flow,'collect',return_value=copy.deepcopy(p)),patch.object(flow,'build',return_value=(Path('a.xlsx'),Path('a.png'))):
            first=flow.advance('one','sales report 18/11/26',config);self.assertEqual(first['prompt']['stage'],'town')
            other=copy.deepcopy(config);other['tally']['company_name']='Other'
            with self.assertRaises(ValueError):flow.advance('one','Town',other)
            result=flow.advance('one','Town',config);self.assertEqual(result['files'],['a.xlsx','a.png']);self.assertFalse(flow.has_pending('one'))
            db,registry=flow.town_registry(config);self.assertEqual(registry['id']['town'],'Town')
            result=flow.advance('two','sales report 18/11/26',config);self.assertIsNone(result['prompt'])

    def test_cancel_and_scope_isolation(self):
        p=payload();p['rows'][0].pop('town');config={'tally':{'company_name':'Company'},'paths':{'processed':'unused'}}
        with tempfile.TemporaryDirectory() as tmp,patch.object(flow,'STATE_DIR',Path(tmp)/'sessions'),patch.object(flow,'TOWNS_PATH',Path(tmp)/'towns.json'),patch.object(flow,'collect',return_value=p):
            flow.advance('one','sales report 18/11/26',config)
            with self.assertRaises(ValueError):flow.advance('two','Town',config)
            self.assertIsNone(flow.advance('one','cancel',config)['prompt']);self.assertFalse(flow.has_pending('one'))

    def test_web_and_router_use_shared_workflow(self):
        import asyncio
        from src.web_jobs import execute
        from src.command_router import run_command,recognized_command
        self.assertTrue(recognized_command('sales report 18/11/26'))
        response={'message':'Ready','files':['a.xlsx','b.png'],'prompt':None}
        with patch.object(flow,'advance',return_value=response) as advance:
            r=execute({'kind':'sales','text':'sales report 18/11/26','conversation_id':'x'})
            self.assertEqual(r['files'],['a.xlsx','b.png']);self.assertEqual(advance.call_args.args[0],'web:x')
            r=asyncio.run(run_command('sales report 18/11/26',[],{},sales_scope='whatsapp:x'))
            self.assertEqual(r.message,'Ready');self.assertEqual(len(r.attachments),2)

    def test_png_download_allowed_but_not_upload(self):
        from src.web_app import file_path
        self.assertEqual(file_path('sample.png').suffix,'.png')
        with self.assertRaises(ValueError):file_path('sample.png','uploads')

    def test_collect_uses_bounded_batches_and_reconciles(self):
        groups=ET.fromstring('<R><GROUP NAME="Town"><PARENT>Sundry Debtors</PARENT></GROUP></R>')
        current=ET.fromstring('<R><LEDGER NAME="Party"><GUID>id</GUID><PARENT>Town</PARENT><REPORTCLOSING>-105</REPORTCLOSING></LEDGER></R>')
        old=ET.fromstring('<R><LEDGER NAME="Party"><REPORTCLOSING>0</REPORTCLOSING></LEDGER></R>')
        headers=ET.fromstring('<R><VOUCHER><MASTERID>1</MASTERID><DATE>20261114</DATE></VOUCHER></R>')
        v=voucher('Tax Invoice',[('Party',-105)]);ET.SubElement(v,'MASTERID').text='1'
        batch=ET.Element('R');batch.append(v)
        responses=[groups,current,old,headers,batch,current,groups]
        with patch.object(data,'ensure_configured_company'),patch.object(data,'export',side_effect=responses) as export:
            p=data.collect({'tally':{'company_name':'Company'}},date(2026,11,1),date(2026,11,18))
            self.assertEqual(p['totals']['spices'],105)
            self.assertEqual(export.call_args_list[4].kwargs['condition'],'$MasterID = 1')
            self.assertNotIn('AllLedgerEntries.*',export.call_args_list[4].args[2])
        responses=[groups,current,old,headers,ET.Element('R')]
        with patch.object(data,'ensure_configured_company'),patch.object(data,'export',side_effect=responses):
            with self.assertRaisesRegex(ValueError,'incomplete'):
                data.collect({'tally':{'company_name':'Company'}},date(2026,11,1),date(2026,11,18))

    def test_town_import_wrong_header_rejected(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'wrong.xlsx';w=Workbook();w.active.append(['PARTY NAME','Billing Rate']);w.active.append(['Party','9.00']);w.save(path)
            with self.assertRaises(ValueError):flow.import_towns(path,payload(),{})


if __name__=='__main__':unittest.main()

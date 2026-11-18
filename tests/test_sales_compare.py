import asyncio
import copy
import json
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch

from src import sales_compare as sc
from src.sales_compare_report import build_report


def line(name='Item', **updates):
    row=dict(name=name,invoice='1',qty=10,mrp=20,rate=10,discount=0,taxable=100,
             cgst_rate=2.5,sgst_rate=2.5,igst_rate=0,cgst=2.5,sgst=2.5,igst=0,net=105,ptr='7.00')
    row.update(updates)
    return row


def voucher(lines=None, **updates):
    row=dict(id='1',invoice='1',date='2026-11-14',party='Party',value=105,
             lines=lines or [line()],ledgers=[dict(name='Party',amount=-105),dict(name='CGST @ 2.5%',amount=2.5),dict(name='SGST @ 2.5%',amount=2.5)],raw=[])
    row.update(updates)
    return row


class CompareTests(unittest.TestCase):
    def test_window_and_command(self):
        self.assertEqual(sc.parse_command('compare maa vaishnow 14/11/26'),('maa vaishnow','2026-11-08','2026-11-14'))
        self.assertEqual(sc.parse_command('COMPARE gauri 2026-11-14')[1],'2026-11-08')
        with self.assertRaises(ValueError):sc.parse_command('compare gauri 31/02/26')

    def test_top_three(self):
        names=['Miscellaneous','MAA VAISHNAV TRADERS C0001250(BHADOHI)','Gauri','Cash']
        self.assertEqual(sc.rank_names('maa vaishnow',names)[0],names[1])
        self.assertEqual(len(sc.rank_names('maa vaishnow',names)),3)

    def test_combine_and_keep_price_variants(self):
        t=[voucher([line(qty=3,taxable=30),line(qty=7,taxable=70),line(rate=11,taxable=110)])]
        d=[voucher([line(),line(rate=11,taxable=110)])]
        self.assertTrue(all(r['matched'] for r in sc.compare_records(t,d)))
        self.assertEqual(len(sc.consolidate(t)),2)

    def test_no_averaging_or_arbitrary_pairing(self):
        t=[voucher([line(rate=9),line(rate=11)])]
        d=[voucher([line(rate=10)])]
        result=sc.compare_records(t,d)
        self.assertEqual(len(result),3)
        self.assertTrue(all(not r['matched'] for r in result))

    def test_zero_tolerance_two_decimals(self):
        self.assertTrue(sc.compare_records([voucher([line(rate=10.004)])],[voucher()])[0]['matched'])
        self.assertFalse(sc.compare_records([voucher([line(rate=10.005)])],[voucher()])[0]['matched'])

    def test_mapping_is_explicit_and_local(self):
        t=[voucher([line('Item New')])];d=[voucher()]
        self.assertEqual(len(sc.compare_records(t,d)),2)
        self.assertTrue(sc.compare_records(t,d,{'Item New':'Item'})[0]['matched'])
        self.assertEqual(t[0]['lines'][0]['name'],'Item New')

    def test_missing_item_and_discount_flag(self):
        self.assertIn('Only in Tally',sc.compare_records([voucher()],[])[0]['remarks'])
        self.assertIn('Discount',sc.compare_records([voucher([line(discount=6.25)])],[voucher()])[0]['remarks'])

    def test_selection_validation(self):
        self.assertEqual(sc.selection('2,1,2',['a','b'],True),['b','a'])
        with self.assertRaises(ValueError):sc.selection('3',['a','b'],True)
        with self.assertRaises(ValueError):sc.selection('1,2',['a','b'])

    def test_tally_free_items_and_date_boundary(self):
        def xml(date,discount,amount=''):
            return f'''<VOUCHER><DATE>{date}</DATE><PARTYLEDGERNAME>Party</PARTYLEDGERNAME>
            <VOUCHERNUMBER>1</VOUCHERNUMBER><VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
            <LEDGERENTRIES.LIST><LEDGERNAME>Party</LEDGERNAME><AMOUNT>-105</AMOUNT></LEDGERENTRIES.LIST>
            <ALLINVENTORYENTRIES.LIST><STOCKITEMNAME>Free</STOCKITEMNAME><BILLEDQTY>1 PCS</BILLEDQTY>
            <RATE>10/PCS</RATE><DISCOUNT>{discount}</DISCOUNT><AMOUNT>{amount}</AMOUNT>
            <ACCOUNTINGALLOCATIONS.LIST><LEDGERNAME>Sales</LEDGERNAME></ACCOUNTINGALLOCATIONS.LIST>
            <BATCHALLOCATIONS.LIST><BATCHNAME>7.00</BATCHNAME><MRP><UDF_AA_BATCH_MRP_RATE>20</UDF_AA_BATCH_MRP_RATE></MRP></BATCHALLOCATIONS.LIST>
            <RATEDETAILS.LIST><GSTRATEDUTYHEAD>CGST</GSTRATEDUTYHEAD><GSTRATE>0</GSTRATE></RATEDETAILS.LIST>
            <RATEDETAILS.LIST><GSTRATEDUTYHEAD>SGST/UTGST</GSTRATEDUTYHEAD><GSTRATE>0</GSTRATE></RATEDETAILS.LIST>
            <RATEDETAILS.LIST><GSTRATEDUTYHEAD>IGST</GSTRATEDUTYHEAD><GSTRATE>0</GSTRATE></RATEDETAILS.LIST>
            </ALLINVENTORYENTRIES.LIST></VOUCHER>'''
        root=ET.fromstring('<ENVELOPE>'+xml('20261114',100)+xml('20261107',100)+'</ENVELOPE>')
        with patch.object(sc,'_export',return_value=root):
            records,_=sc.fetch_tally({},'Party','2026-11-08','2026-11-14')
            self.assertEqual(len(records),1);self.assertEqual(records[0]['lines'][0]['taxable'],0)
        with patch.object(sc,'_export',return_value=ET.fromstring('<ENVELOPE>'+xml('20261114',0)+'</ENVELOPE>')):
            with self.assertRaises(ValueError):sc.fetch_tally({},'Party','2026-11-08','2026-11-14')

    def test_tally_blank_rate_is_zero_only_for_undiscounted_zero_value(self):
        def export(amount='', discount='', rate='', batch_rate='', batch_amount='', second=''):
            return ET.fromstring(f'''<ENVELOPE><VOUCHER><DATE>20261114</DATE>
            <PARTYLEDGERNAME>Party</PARTYLEDGERNAME><VOUCHERNUMBER>TEST-1</VOUCHERNUMBER>
            <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
            <LEDGERENTRIES.LIST><LEDGERNAME>Party</LEDGERNAME><AMOUNT>-105</AMOUNT></LEDGERENTRIES.LIST>
            <ALLINVENTORYENTRIES.LIST><STOCKITEMNAME>Scheme Item</STOCKITEMNAME>
            <BILLEDQTY>2 PCS</BILLEDQTY><RATE>{rate}</RATE><DISCOUNT>{discount}</DISCOUNT><AMOUNT>{amount}</AMOUNT>
            <ACCOUNTINGALLOCATIONS.LIST><LEDGERNAME>Sales</LEDGERNAME></ACCOUNTINGALLOCATIONS.LIST>
            <BATCHALLOCATIONS.LIST><BATCHNAME>7.00</BATCHNAME><BILLEDQTY>2 PCS</BILLEDQTY>
            <BATCHRATE>{batch_rate}</BATCHRATE><AMOUNT>{batch_amount}</AMOUNT>
            <MRP><UDF_AA_BATCH_MRP_RATE>20</UDF_AA_BATCH_MRP_RATE></MRP></BATCHALLOCATIONS.LIST>{second}
            <RATEDETAILS.LIST><GSTRATEDUTYHEAD>CGST</GSTRATEDUTYHEAD><GSTRATE>0</GSTRATE></RATEDETAILS.LIST>
            <RATEDETAILS.LIST><GSTRATEDUTYHEAD>SGST/UTGST</GSTRATEDUTYHEAD><GSTRATE>0</GSTRATE></RATEDETAILS.LIST>
            <RATEDETAILS.LIST><GSTRATEDUTYHEAD>IGST</GSTRATEDUTYHEAD><GSTRATE>0</GSTRATE></RATEDETAILS.LIST>
            </ALLINVENTORYENTRIES.LIST></VOUCHER></ENVELOPE>''')

        for amount in ('', '0', '0.00'):
            with self.subTest(amount=amount), patch.object(sc, '_export', return_value=export(amount=amount)):
                records, _ = sc.fetch_tally({}, 'Party', '2026-11-08', '2026-11-14')
                row = records[0]['lines'][0]
                self.assertEqual((row['rate'], row['discount'], row['taxable'], row['qty']), (0, 0, 0, 2))
                self.assertEqual((row['mrp'], row['ptr']), (20, '7.00'))

        for kwargs in ({'amount': '20'}, {'discount': '100'}, {'discount': '5'},
                       {'batch_amount': '20'}, {'amount': '20', 'second': '<BATCHALLOCATIONS.LIST/>'}):
            with self.subTest(kwargs=kwargs), patch.object(sc, '_export', return_value=export(**kwargs)):
                with self.assertRaisesRegex(ValueError, r'TEST-1: rate unavailable for Scheme Item .*PTR 7.00'):
                    sc.fetch_tally({}, 'Party', '2026-11-08', '2026-11-14')

        for kwargs in ({'rate': '10/PCS', 'amount': '20'},
                       {'batch_rate': '10/PCS', 'amount': '20'},
                       {'rate': '10/PCS', 'discount': '100'}):
            with self.subTest(kwargs=kwargs), patch.object(sc, '_export', return_value=export(**kwargs)):
                records, _ = sc.fetch_tally({}, 'Party', '2026-11-08', '2026-11-14')
                self.assertEqual(records[0]['lines'][0]['rate'], 10)

    def test_dms_cancel_dates_and_multiple_discounts(self):
        row={'Bill ID':'1','Bill Date':'2026-11-14','Buyer Name':'Party','Buyer ERP ID':'P',
             'Invoice Total Value':'105','Product Name':'Item','Invoiced Total Quantity':'10','MRP':'20','Unit Price':'10',
             'Gross Amount':'100','Total Discount':'19','Gross Amount Less Discount':'81',
             'Primary Discount Percent':'10','Secondary Discount Percent':'10','CGST Tax Rate':'2.5',
             'SGST Tax Rate':'2.5','IGST Tax Rate':'0','CGST Tax Amount':'2.025','SGST Tax Amount':'2.025',
             'IGST Tax Amount':'0','Net Value':'85.05'}
        records,excluded=sc.dms_invoices([row,dict(row,**{'Is Cancelled':'true'}),dict(row,**{'Bill Date':'2026-11-07'})],'2026-11-08','2026-11-14')
        self.assertEqual(len(records),1);self.assertEqual(excluded,1)
        self.assertEqual(records[0]['lines'][0]['discount'],19)

    def test_persisted_flow_and_company_guard(self):
        config={'tally':{'company_name':'Company'},'dms':{'company_id':1}}
        with tempfile.TemporaryDirectory() as tmp, patch.object(sc,'STATE_DIR',Path(tmp)),patch.object(sc,'_export',return_value=ET.fromstring('<ENVELOPE><LEDGER NAME="Party"/></ENVELOPE>')),patch.object(sc,'fetch_tally',return_value=([voucher()],0)),patch('src.dms_sales.fetch_sales',return_value=[]),patch.object(sc,'dms_invoices',return_value=([voucher()],0)):
            self.assertEqual(sc.advance('web:abc','compare party 14/11/26',config)['prompt']['stage'],'party')
            with self.assertRaises(ValueError):sc.advance('web:other','1',config)
            other=copy.deepcopy(config);other['tally']['company_name']='Other'
            with self.assertRaises(ValueError):sc.advance('web:abc','1',other)
            self.assertEqual(sc.advance('web:abc','1',config)['prompt']['stage'],'tally')
            self.assertEqual(sc.advance('web:abc','1',config)['prompt']['stage'],'dms')
            self.assertEqual(sc.advance('web:abc','1',config)['prompt']['stage'],'confirm')
            self.assertIsNone(sc.advance('web:abc','cancel',config)['prompt'])
            self.assertFalse(sc.has_pending('web:abc'))

    def test_workbook_15_rows_and_no_source_formulas(self):
        from openpyxl import load_workbook
        t=voucher();d=voucher();state={'party':'Party','company':'Company','start':'2026-11-08','end':'2026-11-14','tally':[t],'dms':[d]}
        with tempfile.TemporaryDirectory() as tmp:
            message,path=build_report(state,sc.compare_records([t],[d]),{'paths':{'processed':tmp},'dms':{'company_id':1}})
            wb=load_workbook(path)
            self.assertTrue(all(d.height==15 for s in wb for d in s.row_dimensions.values()))
            self.assertIn('Final invoice totals match',message)
            self.assertEqual(wb['Tally Detail']['D2'].value,'7.00')
            wb.close()

    def test_router_recognizes_compare(self):
        from src.command_router import recognized_command,run_command
        self.assertTrue(recognized_command('compare maa vaishnow 14/11/26'))
        with patch.object(sc,'advance',return_value={'message':'Choose party','files':[]}):
            r=asyncio.run(run_command('compare x 14/11/26',[],{},comparison_scope='whatsapp:one'))
            self.assertEqual(r.message,'Choose party')

    def test_web_job_uses_same_engine(self):
        from src.web_jobs import execute
        with patch.object(sc,'advance',return_value={'message':'Choose invoices','files':['C:/out/report.xlsx'],'prompt':{'stage':'tally','options':[],'multiple':True}}) as advance:
            result=execute({'kind':'compare','conversation_id':'abc','text':'1'})
            self.assertEqual(advance.call_args.args[0],'web:abc')
            self.assertEqual(result['files'],['report.xlsx'])
            self.assertEqual(result['comparison']['stage'],'tally')

    def test_expired_state_rejected(self):
        with tempfile.TemporaryDirectory() as tmp,patch.object(sc,'STATE_DIR',Path(tmp)):
            sc._path('one').write_text(json.dumps({'stage':'tally','updated':0}))
            self.assertFalse(sc.has_pending('one'))

    def test_non_numeric_inputs_fail_closed(self):
        for value in (None,'','n.a.'):
            with self.assertRaises(ValueError):sc.number(value)

    def test_formula_text_is_preserved_as_text(self):
        from openpyxl import load_workbook
        evil='=HYPERLINK("https://example.invalid","click")'
        t=voucher([line(evil)]);d=voucher([line(evil)])
        state=dict(party='Party',company='Company',start='2026-11-08',end='2026-11-14',tally=[t],dms=[d])
        with tempfile.TemporaryDirectory() as tmp:
            _,path=build_report(state,sc.compare_records([t],[d]),{'paths':{'processed':tmp},'dms':{'company_id':1}})
            wb=load_workbook(path);self.assertEqual(wb['Tally Detail']['B2'].data_type,'s');wb.close()


if __name__=='__main__':unittest.main()

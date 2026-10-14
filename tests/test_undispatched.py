import asyncio
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src.undispatched import parse_bills, apply_additions, fetch_additions
from src.command_router import run_command


class UndispatchedTests(unittest.TestCase):
    def test_exempt_sales_accepted_but_purchase_rejected(self):
        from src.tally_context import TallyConfig
        xml = '''<ENVELOPE><VOUCHER><VOUCHERNUMBER>EX-57</VOUCHERNUMBER>
        <ISCANCELLED>No</ISCANCELLED><ISOPTIONAL>No</ISOPTIONAL>
        <ALLINVENTORYENTRIES.LIST><STOCKITEMNAME>Roli</STOCKITEMNAME>
        <ACTUALQTY>10 PCS</ACTUALQTY><ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <ACCOUNTINGALLOCATIONS.LIST><LEDGERNAME>Exempt Sales</LEDGERNAME></ACCOUNTINGALLOCATIONS.LIST>
        <BATCHALLOCATIONS.LIST><BATCHNAME>43.00</BATCHNAME><ACTUALQTY>10 PCS</ACTUALQTY>
        <UDF:UDF_AA_BATCH_MRP_RATE>67</UDF:UDF_AA_BATCH_MRP_RATE></BATCHALLOCATIONS.LIST>
        </ALLINVENTORYENTRIES.LIST></VOUCHER></ENVELOPE>'''
        with patch('src.undispatched.get_tally_config', return_value=TallyConfig('Test')), patch('src.undispatched.ensure_configured_company'):
            with patch('src.undispatched.tally_request', return_value=xml):
                result = fetch_additions(['EX-57'])
                self.assertEqual(result[0]['qty'], 10)
            with patch('src.undispatched.tally_request', return_value=xml.replace('Exempt Sales','Purchase')), self.assertRaisesRegex(ValueError,'unsupported inventory accounting'):
                fetch_additions(['EX-57'])

    def test_list_and_none(self):
        self.assertEqual(parse_bills('26-27/A-1, 26-27/A-2\n26-27/A-1'), ['26-27/A-1', '26-27/A-2'])
        self.assertEqual(parse_bills('none'), [])
        with self.assertRaises(ValueError):
            parse_bills('yes please')

    def test_prompt_does_not_generate(self):
        with patch('src.command_router._unloading_slip') as build:
            result = asyncio.run(run_command('unloading slip 13-10-2026', [], {}))
            self.assertIn('not dispatched', result.message)
            build.assert_not_called()

    def test_exact_batch_and_new_batch(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d)/'master.xml'
            p.write_text('<ENVELOPE><STOCKITEM NAME="Tea"><DENOMINATOR>120</DENOMINATOR><CONVERSION>1</CONVERSION><BASEUNITS>PCS</BASEUNITS></STOCKITEM></ENVELOPE>')
            original = {'Tea': [dict(ptr='15.00', mrp='20.00', qty=120, case_qty=1, pcs_per_case=120)]}
            a = dict(bill='A-1',item='Tea',ptr='15.0',mrp=25,qty=60,unit='PCS')
            stock, audit = apply_additions(original, [a,a], p)
            self.assertEqual(original['Tea'][0]['qty'], 120)
            self.assertEqual(len(stock['Tea']), 2)
            self.assertEqual(stock['Tea'][1]['qty'], 120)
            self.assertEqual(audit[-1][4:7], [60,60,120])
            with self.assertRaises(ValueError):
                apply_additions({}, [dict(a,item='Missing')], p)

    def test_missing_duplicate_cancelled(self):
        from src.tally_context import TallyConfig
        with patch('src.undispatched.get_tally_config', return_value=TallyConfig('Test')), patch('src.undispatched.ensure_configured_company'):
            for xml in ['<ENVELOPE/>', '<ENVELOPE>'+('<VOUCHER><VOUCHERNUMBER>A-1</VOUCHERNUMBER></VOUCHER>'*2)+'</ENVELOPE>', '<ENVELOPE><VOUCHER><VOUCHERNUMBER>A-1</VOUCHERNUMBER><ISCANCELLED>Yes</ISCANCELLED></VOUCHER></ENVELOPE>']:
                with patch('src.undispatched.tally_request', return_value=xml), self.assertRaises(ValueError):
                    fetch_additions(['A-1'])

    def test_missing_mrp_uses_exact_stock_batch_only(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/'master.xml'
            p.write_text('<ENVELOPE><STOCKITEM NAME="Hing"><DENOMINATOR>720</DENOMINATOR><CONVERSION>1</CONVERSION><BASEUNITS>PCS</BASEUNITS></STOCKITEM></ENVELOPE>')
            a=dict(bill='A-1',item='Hing',ptr='43.00',mrp='',qty=200,unit='PCS')
            stock={'Hing':[dict(ptr='43.00',mrp='60.00',qty=100,case_qty=100/720,pcs_per_case=720)]}
            result,audit=apply_additions(stock,[a],p)
            self.assertEqual(result['Hing'][0]['qty'],300)
            self.assertEqual(audit[0][3],60)
            self.assertIn('Fresh stock',audit[0][-1])
            with self.assertRaises(ValueError):
                apply_additions(stock,[dict(a,ptr='43.0')],p)
            stock['Hing'].append(dict(stock['Hing'][0],mrp='65.00'))
            with self.assertRaisesRegex(ValueError,'Conflicting MRPs'):
                apply_additions(stock,[a],p)

    def test_history_fallback_and_ambiguity(self):
        with tempfile.TemporaryDirectory() as d:
            master=Path(d)/'master.xml';history=Path(d)/'history.xml'
            master.write_text('<ENVELOPE><STOCKITEM NAME="Hing"><DENOMINATOR>720</DENOMINATOR><CONVERSION>1</CONVERSION><BASEUNITS>PCS</BASEUNITS></STOCKITEM></ENVELOPE>')
            voucher='<VOUCHER><ALLINVENTORYENTRIES.LIST><STOCKITEMNAME>Hing</STOCKITEMNAME><BATCHALLOCATIONS.LIST><BATCHNAME>43.00</BATCHNAME><UDF:UDF_AA_BATCH_MRP_RATE>60</UDF:UDF_AA_BATCH_MRP_RATE></BATCHALLOCATIONS.LIST></ALLINVENTORYENTRIES.LIST></VOUCHER>'
            history.write_text('<ENVELOPE>'+voucher+'</ENVELOPE>')
            a=dict(bill='A-1',item='Hing',ptr='43.00',mrp='',qty=200,unit='PCS')
            result,audit=apply_additions({},[a],master,history)
            self.assertEqual(result['Hing'][0]['mrp'],'60.00')
            self.assertIn('history',audit[0][-1])
            history.write_text('<ENVELOPE>'+voucher+voucher.replace('>60<','>65<')+'</ENVELOPE>')
            with self.assertRaisesRegex(ValueError,'Conflicting MRPs'):
                apply_additions({},[a],master,history)


if __name__ == '__main__':
    unittest.main()

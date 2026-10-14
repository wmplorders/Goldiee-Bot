import asyncio
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src.undispatched import parse_bills, apply_additions, fetch_additions
from src.command_router import run_command


class UndispatchedTests(unittest.TestCase):
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


if __name__ == '__main__':
    unittest.main()

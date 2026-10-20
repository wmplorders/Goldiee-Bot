import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from src import web_app, web_jobs
from src import unloading_slip


class WebTests(unittest.TestCase):
    def test_download_boundary(self):
        for name in ['../config.json', '..\\config.json', 'config.json', 'file.exe']:
            with self.assertRaises(ValueError):
                web_app.file_path(name)

    def test_dispatch_choice_required(self):
        with patch('src.unloading_slip.build_unloading_slip') as build:
            with self.assertRaises(ValueError):
                web_jobs.execute({'kind':'unloading', 'date':'2026-10-13'})
            build.assert_not_called()

    def test_dispatch_bill_forwarding(self):
        with patch('src.unloading_slip.build_unloading_slip',return_value={'xlsx':'a.xlsx','pdf':'a.pdf','rows':3,'total_cases':25,'invoice_numbers':['SG-1']}) as build:
            result=web_jobs.execute({'kind':'unloading','date':'2026-10-13','dispatch':'bills','bills':'26-27/SG-TI-340, 26-27/SG-TI-341','invoices':'SG-1'})
            self.assertEqual(build.call_args.args[2], ['26-27/SG-TI-340','26-27/SG-TI-341'])
            self.assertEqual(result['files'],['a.xlsx','a.pdf'])

    def test_failed_job_is_recorded(self):
        with tempfile.TemporaryDirectory() as d, patch.object(web_app,'JOBS_DIR',Path(d)), patch.object(web_app.subprocess,'run',side_effect=RuntimeError('Tally offline')):
            job={'id':'test','company':'Test','created':'2026-10-14','request':{'kind':'stock'}}
            web_app.run_job(job)
            self.assertEqual(web_app.jobs()[0]['status'],'failed')
            self.assertIn('Tally offline',web_app.jobs()[0]['detail'])

    def test_conversation_round_trip(self):
        with tempfile.TemporaryDirectory() as d, patch.object(web_app,'CONVERSATIONS_DIR',Path(d)):
            conversation={
                'id':'a'*32,
                'title':'Unloading Slip',
                'created':'2026-10-15T09:00:00',
                'updated':'2026-10-15T09:00:00',
                'messages':[{'role':'user','type':'text','text':'Unloading Slip'}],
                'context':{'date':'2026-10-15'},
            }
            web_app.save_conversation(conversation)
            loaded=web_app.conversations()
            self.assertEqual(len(loaded),1)
            self.assertEqual(loaded[0]['context']['date'],'2026-10-15')

    def test_conversation_rejects_invalid_id(self):
        with tempfile.TemporaryDirectory() as d, patch.object(web_app,'CONVERSATIONS_DIR',Path(d)):
            with self.assertRaises(ValueError):
                web_app.save_conversation({'id':'../outside','messages':[]})

    def test_summary_search_prefilters_filename_date(self):
        with tempfile.TemporaryDirectory() as d:
            folder=Path(d)
            matching=folder/'SG-1 13-10-2026 Summary.xlsx'
            unrelated=folder/'Old Export.xlsx'
            matching.touch()
            unrelated.touch()
            with patch.object(unloading_slip,'SUMMARY_SEARCH_DIRS',[folder]), patch.object(
                unloading_slip,'workbook_has_purchase_date',return_value=True
            ) as inspect:
                result=unloading_slip.find_purchase_summaries(datetime(2026,10,13))
            self.assertEqual(result,[matching])
            inspect.assert_called_once_with(matching,datetime(2026,10,13),None)

    def test_combined_rows_do_not_count_duplicate_invoice_files(self):
        with tempfile.TemporaryDirectory() as d:
            paths=[]
            for folder_name in ('processed','downloads'):
                folder=Path(d)/folder_name
                folder.mkdir()
                path=folder/'SG-1 19-10-2026 Summary.xlsx'
                wb=Workbook()
                ws=wb.active
                ws.title='Purchase Summary'
                ws.append(['ITEM','PTR','MRP','Case','QTY','Invoice No.','Date'])
                ws.append(['TEST ITEM',10,20,635,6350,'SG-1','19-10-2026'])
                wb.save(path)
                paths.append(path)
            rows=unloading_slip.load_combined_purchase_rows(paths,datetime(2026,10,19))
            self.assertEqual(len(rows),1)
            self.assertEqual(rows[0]['cases'],635)
            self.assertEqual(rows[0]['pcs'],6350)


if __name__=='__main__':
    unittest.main()

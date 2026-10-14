import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src import web_app, web_jobs


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


if __name__=='__main__':
    unittest.main()

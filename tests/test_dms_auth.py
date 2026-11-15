import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from src import dms_client as dms


def token(exp, company=10456):
    payload=base64.urlsafe_b64encode(json.dumps({'exp':exp,'companyId':company}).encode()).decode().rstrip('=')
    return 'header.'+payload+'.signature'


class DmsAuthTests(unittest.TestCase):
    def test_new_cookie_token_beats_stale_local_storage(self):
        old,new=token(100),token(200)
        state={'cookies':[{'name':'accessToken','value':new,'domain':'dms.fieldassist.io','path':'/'}],
               'origins':[{'origin':dms.DMS_ORIGIN,'localStorage':[{'name':'accessToken','value':old}]}]}
        session=dms._session_from_storage(state)
        self.assertEqual(session.headers['Authorization'],'Bearer '+new)

    def test_new_local_storage_token_beats_old_cookie(self):
        old,new=token(100),token(200)
        state={'cookies':[{'name':'accessToken','value':old,'domain':'dms.fieldassist.io','path':'/'}],
               'origins':[{'origin':dms.DMS_ORIGIN,'localStorage':[{'name':'accessToken','value':new}]}]}
        session=dms._session_from_storage(state)
        dms._switch_company(session,10456)
        self.assertEqual(session.headers['Authorization'],'Bearer '+new)

    def test_public_page_or_unauthorized_state_not_accepted(self):
        context=MagicMock();context.storage_state.return_value={'cookies':[],'origins':[]}
        session=MagicMock();session.get.return_value.status_code=401
        with patch.object(dms,'_session_from_storage',return_value=session):
            self.assertIsNone(dms._verified_browser_state(context,{'dms':{'company_id':10456}}))

    def test_valid_state_requires_company_access(self):
        context=MagicMock();context.storage_state.return_value={'cookies':[]}
        session=MagicMock();response=session.get.return_value;response.status_code=200
        response.json.return_value=[{'companyId':10456}]
        with patch.object(dms,'_session_from_storage',return_value=session):
            self.assertEqual(dms._verified_browser_state(context,{'dms':{'company_id':10456}}),{'cookies':[]})
            with self.assertRaises(dms.DmsAuthenticationRequired):
                dms._verified_browser_state(context,{'dms':{'company_id':999}})

    def test_repeated_401_has_actionable_message(self):
        with tempfile.TemporaryDirectory() as tmp,patch.object(dms,'SESSION_PATH',Path(tmp)/'state.json'):
            dms.SESSION_PATH.write_text('{}')
            session=MagicMock();session.get.return_value.status_code=401
            with patch.object(dms,'_session_from_storage',return_value=session),patch.object(dms,'_renew_session_headlessly'):
                with self.assertRaisesRegex(dms.DmsAuthenticationRequired,'authorize_dms.bat'):
                    dms._authenticated_session({'dms':{'company_id':10456}})

    def test_failed_headless_renewal_does_not_overwrite_session(self):
        config={'dms':{'company_id':10456},'whatsapp':{'edge_executable':'edge.exe'}}
        runtime=MagicMock();page=runtime.chromium.launch.return_value.new_context.return_value.new_page.return_value
        page.url='https://login-v2.fieldassist.io/Account/Login'
        with tempfile.TemporaryDirectory() as tmp,patch.object(dms,'SESSION_PATH',Path(tmp)/'state.json'):
            dms.SESSION_PATH.write_text('{"original":true}')
            with patch.object(dms,'sync_playwright') as factory,patch.object(dms,'_verified_browser_state',return_value=None):
                factory.return_value.__enter__.return_value=runtime
                with self.assertRaises(dms.DmsAuthenticationRequired):dms._renew_session_headlessly(config)
            self.assertEqual(json.loads(dms.SESSION_PATH.read_text()),{'original':True})
            runtime.chromium.launch.return_value.close.assert_called_once()


if __name__=='__main__':unittest.main()

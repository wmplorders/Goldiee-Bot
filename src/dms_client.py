from __future__ import annotations

import json
import base64
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from playwright.sync_api import sync_playwright

from .settings import PROJECT_ROOT


DMS_ORIGIN = "https://dms.fieldassist.io"
SESSION_PATH = PROJECT_ROOT / "data" / "dms_storage_state.json"
PROFILE_PATH = PROJECT_ROOT / "data" / "dms-browser-profile"


class DmsAuthenticationRequired(RuntimeError):
    pass


AUTH_HELP = (
    'DMS authorization has expired on this laptop. Run authorize_dms.bat from the bot '
    'folder (or .\\.venv\\Scripts\\python.exe -m src.main dms-authorize), sign in in '
    'the dedicated Edge window, and wait for "DMS authorization verified and saved". '
    'Then click Retry on the failed comparison step. No invoice selection is lost.'
)


def authorize(config: dict) -> None:
    """One-time interactive authorization required by FieldAssist reCAPTCHA."""
    dms = config["dms"]
    PROFILE_PATH.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            str(PROFILE_PATH),
            executable_path=config["whatsapp"]["edge_executable"],
            headless=False,
        )
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(f"{DMS_ORIGIN}/signin-oidc?company-id={dms['company_id']}",
                      wait_until='domcontentloaded', timeout=120_000)
            print('Complete DMS login in the opened Edge window. Waiting for API verification...', flush=True)
            deadline = time.monotonic() + 600
            logged_out = False
            while time.monotonic() < deadline:
                try:
                    body = page.locator('body').inner_text(timeout=2_000)
                    if not logged_out and 'Error: Unauthorized' in body:
                        page.get_by_text('Logout', exact=True).click(timeout=5_000)
                        logged_out = True
                except Exception:
                    pass
                state = _verified_browser_state(context, config)
                if state is not None:
                    _write_storage_state(state)
                    # Verify the configured company switch as well, before closing Edge.
                    _authenticated_session(config)
                    print('DMS authorization verified and saved.', flush=True)
                    return
                page.wait_for_timeout(2_000)
            raise DmsAuthenticationRequired('DMS login was not verified within ten minutes. ' + AUTH_HELP)
        finally:
            context.close()


def fetch_inventory(config: dict) -> list[dict[str, Any]]:
    session = _authenticated_session(config)
    response = session.get(
        f"{DMS_ORIGIN}/api/v2/inventory/stocks",
        params={"category-info": "true"},
        timeout=120,
    )
    if response.status_code == 401:
        _renew_session_headlessly(config)
        session = _authenticated_session(config)
        response = session.get(
            f"{DMS_ORIGIN}/api/v2/inventory/stocks",
            params={"category-info": "true"},
            timeout=120,
        )
    if response.status_code in {401, 403}:
        raise DmsAuthenticationRequired(AUTH_HELP)
    response.raise_for_status()
    payload = response.json()
    if isinstance(payload, list):
        return payload
    for key in ("data", "stocks", "inventory", "items"):
        if isinstance(payload.get(key), list):
            return payload[key]
    raise RuntimeError("DMS inventory API returned an unsupported response format.")


def build_inventory_workbook(config: dict, output: Path | None = None) -> dict[str, Any]:
    rows = fetch_inventory(config)
    output = output or _default_output_path()
    output.parent.mkdir(parents=True, exist_ok=True)
    keys = _ordered_keys(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "DMS Current Stock"
    ws.append(keys)
    for row in rows:
        ws.append([_excel_value(row.get(key)) for key in keys])

    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    for column in ws.columns:
        letter = column[0].column_letter
        longest = max(len(str(cell.value or "")) for cell in column[:200])
        ws.column_dimensions[letter].width = min(max(longest + 2, 12), 45)
    if rows and keys:
        end = ws.cell(row=len(rows) + 1, column=len(keys)).coordinate
        table = Table(displayName="DmsCurrentStock", ref=f"A1:{end}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
    wb.save(output)
    return {"output": str(output), "rows": len(rows), "columns": len(keys)}


def _authenticated_session(config: dict) -> requests.Session:
    if not SESSION_PATH.exists():
        raise DmsAuthenticationRequired(
            "DMS has not been authorized. Run `python -m src.main dms-authorize` once."
        )
    session = _session_from_storage()
    response = session.get(f"{DMS_ORIGIN}/api/v1/user/companies", timeout=30)
    if response.status_code == 401:
        _renew_session_headlessly(config)
        session = _session_from_storage()
        response = session.get(f"{DMS_ORIGIN}/api/v1/user/companies", timeout=30)
    if response.status_code in {401, 403}:
        raise DmsAuthenticationRequired(AUTH_HELP)
    response.raise_for_status()
    expected_id = int(config["dms"]["company_id"])
    company_ids = {
        int(company.get("companyId") or company.get("id"))
        for company in response.json()
        if company.get("companyId") or company.get("id")
    }
    if expected_id not in company_ids:
        raise RuntimeError(f"DMS company ID {expected_id} is not accessible to this account.")
    _switch_company(session, expected_id)
    return session


def _switch_company(session: requests.Session, company_id: int) -> None:
    access_token = session.headers.get('Authorization', '').removeprefix('Bearer ').strip()
    if not access_token:
        access_token = session.cookies.get("accessToken", domain="dms.fieldassist.io")
    if access_token and _token_company_id(access_token) == company_id:
        session.headers["Authorization"] = f"Bearer {access_token}"
        return

    response = session.post(
        f"{DMS_ORIGIN}/api/v2/login/change-company/{company_id}",
        json={},
        timeout=60,
    )
    response.raise_for_status()
    tokens = response.json()
    access_token = tokens.get("accessToken")
    if not access_token or _token_company_id(access_token) != company_id:
        raise RuntimeError(f"DMS did not issue an access token for company ID {company_id}.")

    _store_token(session, "accessToken", access_token)
    if tokens.get("refreshToken"):
        _store_token(session, "refreshToken", tokens["refreshToken"])
    session.headers["Authorization"] = f"Bearer {access_token}"
    _persist_session_cookies(session)


def _token_company_id(token: str) -> int | None:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        return int(claims["companyId"])
    except (IndexError, KeyError, TypeError, ValueError, json.JSONDecodeError):
        return None


def _store_token(session: requests.Session, name: str, value: str) -> None:
    session.cookies.set(name, value, domain="dms.fieldassist.io", path="/")


def _persist_session_cookies(session: requests.Session) -> None:
    state = json.loads(SESSION_PATH.read_text(encoding="utf-8"))
    cookies = {
        (cookie["domain"], cookie["path"], cookie["name"]): cookie
        for cookie in state.get("cookies", [])
    }
    for cookie in session.cookies:
        key = (cookie.domain, cookie.path, cookie.name)
        saved = cookies.get(key, {
            "name": cookie.name,
            "domain": cookie.domain,
            "path": cookie.path,
            "httpOnly": False,
            "secure": True,
            "sameSite": "Lax",
        })
        saved["value"] = cookie.value
        token_expiry = _token_expiry(cookie.value)
        if token_expiry:
            saved["expires"] = token_expiry
        cookies[key] = saved
    state["cookies"] = list(cookies.values())
    SESSION_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def _token_expiry(token: str) -> int | None:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return int(json.loads(base64.urlsafe_b64decode(payload))["exp"])
    except (IndexError, KeyError, TypeError, ValueError, json.JSONDecodeError):
        return None


def _renew_session_headlessly(config: dict) -> None:
    dms = config["dms"]
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            executable_path=config["whatsapp"]["edge_executable"],
            headless=True,
        )
        try:
            context = browser.new_context(storage_state=str(SESSION_PATH))
            page = context.new_page()
            page.goto(
                f"{DMS_ORIGIN}/signin-oidc?company-id={dms['company_id']}",
                wait_until="domcontentloaded", timeout=120_000,
            )
            for _ in range(6):
                page.wait_for_timeout(2_000)
                state = _verified_browser_state(context, config)
                if state is not None:
                    _write_storage_state(state)
                    return
                if 'login-v2.fieldassist.io' in page.url:
                    break
            raise DmsAuthenticationRequired(AUTH_HELP)
        finally:
            browser.close()


def _verified_browser_state(context, config) -> dict | None:
    state = context.storage_state()
    session = _session_from_storage(state)
    try:
        response = session.get(f'{DMS_ORIGIN}/api/v1/user/companies', timeout=15)
        if response.status_code in {401, 403}:
            return None
        response.raise_for_status()
        companies = response.json()
    except (requests.RequestException, ValueError):
        return None
    expected = int(config['dms']['company_id'])
    ids = {int(c.get('companyId') or c.get('id')) for c in companies
           if c.get('companyId') or c.get('id')}
    if expected not in ids:
        raise DmsAuthenticationRequired(f'This DMS account cannot access configured company {expected}. Sign in using the correct account.')
    return state


def _write_storage_state(state) -> None:
    SESSION_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary = SESSION_PATH.with_suffix('.tmp')
    temporary.write_text(json.dumps(state, indent=2), encoding='utf-8')
    temporary.replace(SESSION_PATH)


def _save_storage_state(context) -> None:
    SESSION_PATH.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(SESSION_PATH))


def _session_from_storage(state=None) -> requests.Session:
    if state is None:
        state = json.loads(SESSION_PATH.read_text(encoding="utf-8"))
    session = requests.Session()
    session.headers.update({"Accept": "application/json", "User-Agent": "Goldiee-Purchase-Bot/1.0"})
    for cookie in state.get("cookies", []):
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain"),
            path=cookie.get("path", "/"),
        )
    tokens = []
    for cookie in state.get('cookies', []):
        if cookie.get('name') == 'accessToken' and cookie.get('domain','').lstrip('.') == 'dms.fieldassist.io':
            tokens.append(cookie['value'])
    for origin in state.get("origins", []):
        if origin.get("origin") != DMS_ORIGIN:
            continue
        storage = {entry["name"]: entry["value"] for entry in origin.get("localStorage", [])}
        token = storage.get("accessToken")
        if token:
            tokens.append(token)
    if tokens:
        token = max(tokens, key=lambda t: _token_expiry(t) or 0)
        session.headers['Authorization'] = f'Bearer {token}'
    return session


def _ordered_keys(rows: list[dict[str, Any]]) -> list[str]:
    preferred = [
        "productErpId", "productName", "variant", "productDivision", "primaryCategory",
        "secondaryCategory", "mrpBatchNo", "currentStockCases", "currentStockPieces",
        "currentStock", "mrp", "purchasePrice", "netValue", "warehouseErpId",
    ]
    found = {key for row in rows for key in row}
    return [key for key in preferred if key in found] + sorted(found.difference(preferred))


def _excel_value(value: Any) -> Any:
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def _default_output_path() -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    return PROJECT_ROOT / "downloads" / f"DMS Stock {stamp}.xlsx"

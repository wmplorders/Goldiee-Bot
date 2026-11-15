"""Fresh, browserless DMS sales dumps for invoice reconciliation."""
import csv
import io
import json
import time
from datetime import datetime, timezone, timedelta
from urllib.parse import urlparse

import requests

from .dms_client import DMS_ORIGIN, _authenticated_session


def fetch_sales(config, start, end):
    session = _authenticated_session(config)
    started = datetime.now(timezone.utc) - timedelta(seconds=2)
    filters = {k: '' for k in ('productDivisionIds', 'primaryCategoryIds', 'secondaryCategoryIds',
               'displayCategoryIds', 'alternateCategoryNames', 'variantNames', 'productIds',
               'mrp', 'productAttributeTexts1', 'productAttributeTexts2')}
    response = session.get(DMS_ORIGIN + '/api/v3/report/getReports', params={
        'report-type': 'sales_invoice_dump', 'from-date': start, 'to-date': end,
        'customer-report': 'false', 'show-edited-and-cancelled': 'false',
        'show-only-idt-invoices': 'false', 'show-proforma-invoices': 'false',
        'product-filters': json.dumps(filters), 'offset': '330'}, timeout=120)
    response.raise_for_status()
    for _ in range(80):
        response = session.get(DMS_ORIGIN + '/api/v3/report/report-generation', timeout=60)
        response.raise_for_status()
        candidates = []
        for row in response.json():
            params = row.get('reportParams') or {}
            try:
                created = datetime.fromisoformat(str(row.get('createdAt', '')).replace('Z', '+00:00'))
            except ValueError:
                continue
            if (row.get('reportType') == 'sales_invoice_dump' and created >= started
                    and str(params.get('fromDate', ''))[:10] == start
                    and str(params.get('toDate', ''))[:10] == end):
                candidates.append(row)
        if candidates:
            report = max(candidates, key=lambda r: r['createdAt'])
            status = report.get('reportGenerationStatus')
            if status in {'failed', 'cancelled'}:
                raise ValueError('DMS sales dump failed: ' + str(report.get('failureReason') or status))
            if status == 'completed':
                url = report.get('downloadUrl') or ''
                if url.startswith('/'):
                    url = DMS_ORIGIN + url
                parsed = urlparse(url)
                host = parsed.hostname or ''
                if parsed.scheme != 'https' or not (host == 'dms.fieldassist.io' or host.endswith('.blob.core.windows.net')):
                    raise ValueError('DMS returned an unsupported report download location.')
                download = requests.get(url, timeout=180)
                download.raise_for_status()
                if download.content.startswith(b'PK'):
                    from openpyxl import load_workbook
                    wb = load_workbook(io.BytesIO(download.content), read_only=True, data_only=True)
                    values = iter(wb.active.values)
                    headers = next(values)
                    rows = [dict(zip(headers, r)) for r in values]
                    wb.close()
                    return rows
                return list(csv.DictReader(io.StringIO(download.content.decode('utf-8-sig'))))
        time.sleep(15)
    raise TimeoutError('DMS sales dump is not ready. Please retry later.')

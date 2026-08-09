# CODEX_HANDOFF.md

## Project

Goldiee Purchase Bot is a local Windows automation project for Goldiee purchase workflows. It runs on the user's laptop, monitors Gmail and WhatsApp, processes Goldiee purchase invoice Excel files, interacts with Tally Prime over HTTP, fetches DMS stock from FieldAssist, and creates operational Excel/PDF outputs for purchase entry and warehouse unloading.

Repository:

```text
https://github.com/wmplorders/Goldiee-Bot.git
```

Primary local project path on the original laptop:

```text
C:\Users\goenk\Documents\Codex\2026-05-12\Goldiee Purchase Bot
```

## Objectives

The bot is intended to:

1. Create purchase summary Excel files from Goldiee invoice Excel attachments.
2. Preserve raw invoice data in one sheet and generate a clean `Purchase Summary` sheet.
3. Detect new Tally stock items while creating purchase summaries.
4. Monitor Gmail for Goldiee invoice mails and process invoice attachments.
5. Send purchase PDFs and processed summaries to WhatsApp.
6. Fetch detailed stock status from Tally, including item, batch/PTR, MRP, sell rate where available, quantity, and conversion data.
7. Create unloading slips from one or more purchase summaries and current Tally stock.
8. Fetch DMS stock from FieldAssist using browserless APIs after one-time authorization.
9. Compare Tally stock and DMS stock.
10. Support future direct Tally purchase voucher import from purchase summaries.

## Current Architecture

The project is a Python application with a small amount of Node.js tooling for spreadsheet/report generation.

Main Python entrypoint:

```text
src/main.py
```

Runtime modes:

- CLI commands through `python -m src.main ...`
- Batch wrappers such as `start_bot.bat`, `stock_status.bat`, `unloading_slip.bat`
- Continuous monitor mode for WhatsApp and Gmail

Core modules:

```text
src/purchase_processor.py   Purchase invoice parsing and summary generation
src/gmail_monitor.py        Gmail OAuth, mail scanning, attachment download
src/whatsapp_edge.py        WhatsApp Web automation in Microsoft Edge
src/command_router.py       WhatsApp command parsing and dispatch
src/tally_context.py        Tally company safety lock
src/tally_item_master.py    Tally item master cache and item existence checks
src/tally_batch_stock.py    Tally stock status extraction
src/unloading_slip.py       Unloading slip Excel/PDF generation
src/dms_client.py           FieldAssist/DMS session and stock download
src/stock_aging_report.py   Goldiee stock aging mail-format report
src/settings.py             Config loading from config.json and .env
src/state.py                Local state/log helpers
```

Supporting scripts and wrappers:

```text
setup.bat
start_bot.bat
start_whatsapp_login.bat
authorize_gmail.bat
authorize_dms.bat
dms_stock.bat
stock_status.bat
tally_check.bat
warm_item_master.bat
process_file.bat
unloading_slip.bat
```

Node/report tooling:

```text
build_stock_comparison.mjs
build_stock_comparison_current.mjs
tools/build_dms_july_sales_report.mjs
tools/build_physical_stock_sheet.mjs
tools/build_physical_stock_sheet_corrected.mjs
tools/build_stock_aging_analysis.py
```

## Important Decisions

### Purchase Summary Format

Every processed purchase invoice creates an Excel workbook with:

- `Raw Data`: original invoice sheet copied as-is
- `Purchase Summary`: clean output for Tally purchase entry

Required purchase summary columns:

```text
No.
PTR
MRP
Sell Rate
ITEM
Case
QTY
Billing Rate
Discount %
Invoice No.
Date
Status
```

Rules:

- `Sell Rate` is calculated from `Billing Rate`.
- Default sell rate: `Billing Rate * 1.03`, rounded to 2 decimals.
- Items starting with `Tea` use `Billing Rate * 1.025`, rounded to 2 decimals.
- `Discount %` is calculated from `DISCOUNT / AMOUNT`, rounded/formatted to 2 decimals.
- Items are sorted alphabetically.
- Duplicate rows are consolidated only when item name and billing rate are identical.
- If identical item names have different billing rates, they remain separate rows.
- Serial number is reapplied from 1 after sorting/consolidation.
- New items not present in Tally item master are marked in the `Status` column.
- Item matching against Tally is strict and case-sensitive.

### Tally Safety

The bot should never silently operate on the wrong Tally company.

Configured company information is stored in `config.json` or `.env`:

```text
TALLY_DATA_PATH
TALLY_COMPANY_FOLDER
TALLY_COMPANY_NAME
TALLY_HTTP_URL
TALLY_REQUIRE_EXACT_COMPANY
TALLY_CLOSE_OTHER_COMPANIES
```

Current implementation checks loaded Tally companies using Tally HTTP. If the configured company is not loaded, it stops instead of using another company.

Important current limitation:

- UI automation for Tally Prime `Select Company > Specify Path` has not been fully implemented.
- For now, open/select the configured Tally company manually before Tally-dependent commands.

### Tally Batch/PTR Rules

The Goldiee Tally company uses a custom TDL that adds batch/PTR, MRP, and sell-rate behavior.

For purchase import planning:

- The bot must match the unique combination of item name + PTR + MRP + sell rate.
- Discount can vary and should not be part of the batch identity.
- Numeric formatting differences in PTR can be deliberate.
- If Tally already has the same item/PTR but MRP or sell rate differs, the bot must stop and ask the user. It must never overwrite automatically.
- If the exact item/PTR/MRP/sell-rate combination exists, use the Tally value as-is.

### Ledger Decisions For Future Tally Purchase Import

Known decisions:

- Default purchase ledger: `SHUBHAM GOLDIEE MASALE PRIVATE LIMITED`
- Use `Round Off`, not `Round off - claims`
- Use input tax ledgers, for example `INPUT CGST 2.5%`, not generic `CGST 2.5%`
- 6% tax bracket support is not required currently
- Do not import if any item or ledger is missing. Ask the user to create it, recheck, then import.
- Never overwrite Tally masters automatically.

### Gmail Rules

Configured Gmail account:

```text
goenkatradeventures@gmail.com
```

Watched sender:

```text
information@goldiee.com
```

Processing rules:

- Process new Goldiee mails from configured start date.
- Ignore attachments with `pending` in the filename.
- If any Excel attachment has `invoice` in the name, process those Excel files.
- Otherwise process non-pending Excel files.
- Forward all PDF attachments as-is to WhatsApp in monitor mode.
- Emails 3 days or older are skipped and logged.
- Processed emails are labelled with the configured Gmail label.

### WhatsApp Rules

WhatsApp Web is automated through Microsoft Edge, not Chrome.

Reasons:

- The target laptop may already run other bots.
- Another Python bot may kill Chrome instances.
- Edge is expected not to be killed by the other automation.

Rules:

- WhatsApp command chat/group can be a personal chat or group.
- Configured group is usually `Goldiee Automation`.
- Bot replies in the same chat for WhatsApp commands.
- If WhatsApp is logged out, scan the QR code in Edge.

### DMS Rules

DMS is FieldAssist.

Configured company:

```text
Goldiee Group
```

DMS stock fetching was confirmed to be possible browserlessly after a one-time authorization flow. If the identity session expires or FieldAssist requires reCAPTCHA again, run DMS authorization again.

Do not store DMS passwords in this repository or handoff document.

## Completed Work

### Purchase Summary

Completed:

- Multiple invoice formats handled.
- Raw Data + Purchase Summary workbook format implemented.
- Sell rate calculation implemented.
- Tea-specific sell-rate calculation implemented.
- Discount percentage calculation implemented.
- Alphabetical sorting implemented.
- Duplicate consolidation implemented.
- Output filename format implemented:

```text
<Invoice No> <Invoice Date> Summary.xlsx
```

- New item detection against Tally item master implemented.
- Strict case-sensitive item matching implemented.

Recent created examples:

```text
processed\SG-ST-26-07111 07-08-2026 Summary.xlsx
processed\SG-SE-26-00852 07-08-2026 Summary.xlsx
```

### Gmail Monitoring

Completed:

- Gmail OAuth flow implemented.
- One-shot mail scan command implemented.
- Continuous monitor mode implemented.
- Goldiee sender filtering implemented.
- Attachment filtering implemented.
- Processed-label handling implemented.
- Old email skip logic implemented.

Command:

```powershell
.\.venv\Scripts\python.exe -m src.main gmail-once
```

### WhatsApp Automation

Completed:

- Edge-based WhatsApp Web automation implemented.
- QR login helper implemented.
- Same-chat reply logic implemented for monitor commands.
- Attachment sending implemented.

Commands:

```powershell
.\.venv\Scripts\python.exe -m src.main whatsapp-login
.\.venv\Scripts\python.exe -m src.main monitor
```

### Tally Stock Status

Completed:

- Tally HTTP stock extraction implemented.
- TDL-backed batch/PTR/MRP stock status extraction implemented.
- Item master cache implemented.
- Item conversion data extraction implemented.

Commands:

```powershell
.\.venv\Scripts\python.exe -m src.main tally-check
.\.venv\Scripts\python.exe -m src.main item-master
.\.venv\Scripts\python.exe -m src.main stock-status
```

### Unloading Slip

Completed:

- Combines purchase summaries for a date.
- Can restrict to specific invoice numbers from CLI.
- Uses current Tally stock.
- Matches item + MRP + PTR logic.
- Highlights rows:
  - Green: item not currently present in godown/current stock
  - Red: item present but MRP/PTR differs or multiple MRP/PTR stock exists
  - No highlight: exact MRP/PTR match with no conflicting stock
- Shows current stock as human-readable MRP quantity text.
- Supports full case + loose pcs display where conversion exists.
- Adds invoice numbers in first row.
- Adds total cases at bottom.
- Creates both XLSX and single-page PDF.
- PDF layout has been tightened vertically and horizontally for print.

Command:

```powershell
.\.venv\Scripts\python.exe -m src.main unloading-slip 07-08-2026 SG-ST-26-07111 SG-SE-26-00852
```

Recent created examples:

```text
processed\Unloading Slip 07-08-2026 SG-SE-26-00852 SG-ST-26-07111.xlsx
processed\Unloading Slip 07-08-2026 SG-SE-26-00852 SG-ST-26-07111.pdf
```

### DMS Stock

Completed:

- One-time browser/Edge authorization flow implemented.
- Browserless DMS stock fetch implemented.
- Current-stock report generation/download implemented.

Command:

```powershell
.\.venv\Scripts\python.exe -m src.main dms-authorize
.\.venv\Scripts\python.exe -m src.main dms-stock
```

### Tally vs DMS Stock Comparison

Completed manually/prototype:

- Fresh Tally stock and DMS stock can be compared.
- Comparison by exact item name + MRP.
- Tally quantity is net across nonzero PTR/batch rows, including negative batch rows.
- DMS quantity uses `Total Current Stock in Pcs`.
- Exact matches highlighted green.
- Mismatches/review rows highlighted red.

Recent output:

```text
downloads\Tally vs DMS Stock Comparison 2026-08-06.xlsx
```

Current status:

- Comparison logic exists as a manual/prototype builder, not yet fully integrated into `src.main` as a stable `compare stock` command.

### Stock Aging / Physical Stock

Completed/prototyped:

- Goldiee stock aging mail report command:

```powershell
.\.venv\Scripts\python.exe -m src.main stock-goldiee-mail 01-07-2026 08-07-2026
```

- Physical stock sheet prototypes created under `tools`.
- Conversion-based full cases and loose pcs reports have been generated from current Tally data.

## Current Status

As of this handoff:

- Purchase summary workflow is working.
- Gmail invoice fetch is working after OAuth authorization.
- WhatsApp sending works when Edge WhatsApp is logged in.
- Tally stock status works when the correct Tally company is open and reachable at the configured Tally HTTP URL.
- Unloading slip workflow is working and creates both Excel and PDF.
- DMS stock fetch works after DMS authorization, unless FieldAssist requires fresh login/reCAPTCHA.
- Tally vs DMS comparison has working manual output but should be made into a first-class bot command.
- Direct Tally purchase voucher import has been analyzed but should be treated as unfinished/experimental unless verified in current code.

## Relevant Files And Folders

### Source

```text
src/
```

### Config

```text
config.example.json
config.json
.env.example
```

`config.json` and `.env` can contain machine-specific paths. Review carefully on the new laptop.

### Local State And Secrets

```text
data/
edge-profile/
```

Do not commit these folders if they contain local authorization state.

### Inputs/Outputs

```text
downloads/
processed/
logs/
```

`downloads` stores downloaded Gmail/DMS/Tally source files.

`processed` stores generated summaries, unloading slips, PDFs, and reports.

`logs/processed_purchases.csv` stores processing history.

### Tally HTTP Logs

```text
tallyhttp.log
tallyerr.log
```

These may be local diagnostics only.

## Secrets To Transfer Separately

Do not put secret values into this file or commit them to GitHub.

On the new laptop, either re-authorize services or securely transfer the following local files if needed:

```text
data\gmail_credentials.json
data\token.json
data\dms_session.json
edge-profile\
```

Notes:

- `gmail_credentials.json` is the Google OAuth desktop client JSON.
- `token.json` is generated after Gmail authorization.
- DMS session files are generated by `dms-authorize`.
- WhatsApp login can be recreated by scanning the QR code in Edge.
- If copying `edge-profile`, close Edge first and use secure local transfer only.

Recommended safer approach on a new laptop:

1. Put fresh Gmail OAuth desktop credentials in `data\gmail_credentials.json`.
2. Run Gmail authorization.
3. Run DMS authorization.
4. Scan WhatsApp QR again.

## New Laptop Setup

### 1. Clone Repo

```powershell
cd "C:\Users\<user>\Documents"
git clone https://github.com/wmplorders/Goldiee-Bot.git
cd Goldiee-Bot
```

### 2. Create Python Virtual Environment

Use a project-local venv so this bot does not interfere with other bots.

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
```

Activation is optional. Prefer explicit venv Python:

```powershell
.\.venv\Scripts\python.exe -m src.main tally-check
```

If you want to activate in PowerShell:

```powershell
& .\.venv\Scripts\Activate.ps1
```

If blocked:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### 3. Install Node Dependencies If Needed

Some report-builder scripts use Node.js. If `package.json` exists in the repository:

```powershell
npm install
```

The main bot runtime is Python.

### 4. Configure `config.json`

Copy example if needed:

```powershell
copy config.example.json config.json
```

Review/update:

```json
{
  "gmail": {
    "account": "goenkatradeventures@gmail.com",
    "sender": "information@goldiee.com",
    "start_date": "2026-06-08",
    "max_email_age_days": 3,
    "processed_label": "Processed by Purchase Bot"
  },
  "whatsapp": {
    "command": "purchase summary",
    "group_name": "Goldiee Automation",
    "edge_executable": "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "edge_profile_dir": "edge-profile"
  },
  "tally": {
    "company_name": "Goenka Tradeventures 2025-2026 (Goldiee)",
    "data_path": "\\\\server\\Tally Server Data",
    "company_folder": "100011",
    "http_url": "http://localhost:9000",
    "require_exact_company": true,
    "close_other_companies": true
  },
  "dms": {
    "company_id": 10456,
    "company_name": "Goldiee Group"
  }
}
```

Important:

- Confirm server share path is accessible from the new laptop.
- Confirm Tally Prime is configured to serve HTTP on `http://localhost:9000`.
- Confirm the exact Tally company name matches what Tally reports.
- Confirm Edge executable path on the new laptop.

### 5. Create Required Local Folders

Usually these are created automatically, but they can be created manually:

```powershell
mkdir data, downloads, processed, logs
```

### 6. Gmail Authorization

Place the OAuth desktop-client JSON here:

```text
data\gmail_credentials.json
```

Then run:

```powershell
.\.venv\Scripts\python.exe -m src.main gmail-once
```

or:

```powershell
authorize_gmail.bat
```

Complete the browser authorization. A local token file should be created under `data`.

### 7. DMS Authorization

Run:

```powershell
.\.venv\Scripts\python.exe -m src.main dms-authorize
```

or:

```powershell
authorize_dms.bat
```

Complete the FieldAssist login/reCAPTCHA flow. Then test:

```powershell
.\.venv\Scripts\python.exe -m src.main dms-stock
```

### 8. WhatsApp Login

Run:

```powershell
.\.venv\Scripts\python.exe -m src.main whatsapp-login
```

or:

```powershell
start_whatsapp_login.bat
```

Scan QR in Edge if needed.

### 9. Tally Setup

In Tally Prime:

1. Ensure the correct Goldiee company is open.
2. Ensure Tally HTTP is enabled and reachable at configured URL, usually:

```text
http://localhost:9000
```

3. Ensure the required Goldiee TDL is loaded for batch/PTR/MRP/sell-rate stock reports.

Then test:

```powershell
.\.venv\Scripts\python.exe -m src.main tally-check
.\.venv\Scripts\python.exe -m src.main item-master
.\.venv\Scripts\python.exe -m src.main stock-status
```

## Main Commands

### Manual Invoice Processing

```powershell
.\.venv\Scripts\python.exe -m src.main process "C:\path\to\invoice.xlsx"
```

### Gmail One-Time Scan

```powershell
.\.venv\Scripts\python.exe -m src.main gmail-once
```

### Full Bot Monitor

```powershell
.\.venv\Scripts\python.exe -m src.main monitor
```

### Tally Check

```powershell
.\.venv\Scripts\python.exe -m src.main tally-check
```

### Refresh Tally Item Master

```powershell
.\.venv\Scripts\python.exe -m src.main item-master
```

### Tally Stock Status

```powershell
.\.venv\Scripts\python.exe -m src.main stock-status
```

### DMS Stock

```powershell
.\.venv\Scripts\python.exe -m src.main dms-stock
```

### Unloading Slip For Date

```powershell
.\.venv\Scripts\python.exe -m src.main unloading-slip 07-08-2026
```

### Unloading Slip For Specific Invoices

```powershell
.\.venv\Scripts\python.exe -m src.main unloading-slip 07-08-2026 SG-ST-26-07111 SG-SE-26-00852
```

### Stock Aging Mail Report

```powershell
.\.venv\Scripts\python.exe -m src.main stock-goldiee-mail 01-07-2026 08-07-2026
```

## WhatsApp Commands

Recognized commands:

```text
purchase summary
tally check
stock status
stock goldiee mail <fromdate> to <todate>
dms stock
unloading slip <date>
```

The CLI supports invoice-specific unloading slips, but the WhatsApp command router currently only parses the date text and does not expose invoice filters cleanly. Add this if needed.

## Tests And Verification Checklist

There is no formal automated test suite yet. Use the following smoke tests after setup.

### Python Import Test

```powershell
.\.venv\Scripts\python.exe -c "import src.main; print('ok')"
```

### Tally Company Lock Test

```powershell
.\.venv\Scripts\python.exe -m src.main tally-check
```

Expected:

- Configured company shown.
- Loaded companies shown.
- `Tally company lock OK.`

If it fails:

- Open/select correct Tally company.
- Check Tally HTTP URL.
- Check company name and data path in config.

### Item Master Cache Test

```powershell
.\.venv\Scripts\python.exe -m src.main item-master
```

Expected:

- Prints number of stock item masters cached.

### Purchase Summary Test

Use a known invoice Excel:

```powershell
.\.venv\Scripts\python.exe -m src.main process "C:\path\to\invoice.xlsx"
```

Verify:

- Output file is in `processed`.
- Has `Raw Data` and `Purchase Summary`.
- Summary has expected required columns.
- Discounts are two decimals.
- Sell rate is two decimals.
- Duplicate items are consolidated correctly.
- New items are marked in `Status`.

### Gmail Test

```powershell
.\.venv\Scripts\python.exe -m src.main gmail-once
```

Verify:

- New Goldiee invoice mails are processed.
- Pending-order Excel is ignored.
- Invoice PDF is downloaded.
- Purchase summary is created.
- Gmail label is applied.

If WhatsApp is logged out, the file may still be created but sending may fail.

### DMS Test

```powershell
.\.venv\Scripts\python.exe -m src.main dms-stock
```

Expected:

- Creates current DMS stock workbook in `downloads`.

If authorization fails:

```powershell
.\.venv\Scripts\python.exe -m src.main dms-authorize
```

### Stock Status Test

```powershell
.\.venv\Scripts\python.exe -m src.main stock-status
```

Expected:

- Creates timestamped stock status workbook in `downloads`.
- Includes batch/PTR/MRP details where available.

### Unloading Slip Test

```powershell
.\.venv\Scripts\python.exe -m src.main unloading-slip 07-08-2026 SG-ST-26-07111 SG-SE-26-00852
```

Verify:

- Creates XLSX and PDF in `processed`.
- First row lists combined invoice numbers.
- PDF is one page.
- Total cases appears at bottom.
- Highlight colors match rules.

## Known Issues / Unresolved Problems

### Tally UI Company Selection

Desired behavior:

- Bot should open configured Tally data location and company regardless of what company is currently open.
- Bot should close other open Tally companies if configured.
- It should use `Select Company > Specify Path` without changing Tally's default data path.

Current state:

- Safety checks exist through Tally HTTP.
- Fully independent Tally UI selection is not implemented.
- User must ensure correct company is open before Tally-dependent work.

### Tally Purchase Voucher Import

Desired behavior:

- Bot should create/import Tally purchase vouchers from purchase summaries or raw invoice Excel.
- It must fill supplier invoice number/date.
- It must fill supplier details if needed.
- It must enter item batch/PTR/MRP/sell-rate/purchase-rate/discount.
- It must validate tax ledgers and round-off.
- It must stop if items/ledgers/batches are missing or conflicting.

Current state:

- Business rules and pitfalls have been analyzed.
- Do not assume production-ready import unless code has been reviewed and tested in dummy Tally data.

### DMS Authorization Expiry

DMS browserless fetch works while FieldAssist identity/session remains valid. If FieldAssist requires reCAPTCHA or re-login, run `dms-authorize` again.

### WhatsApp Login

WhatsApp Web login can expire. If `gmail-once` or `monitor` fails during send with QR/login message, run WhatsApp login and scan QR.

### Compare Stock Command

Desired WhatsApp command:

```text
compare stock
```

Current state:

- Manual/prototype comparison report exists.
- It should be integrated into `src.main` and `command_router.py`.

### Formal Tests

There is no formal automated test suite yet. Smoke tests are manual commands.

## Recommended Next Steps

1. Set up the new laptop using a project-local `.venv`.
2. Configure `config.json` for the new laptop paths, Tally company, Edge path, and stock aging template path.
3. Re-authorize Gmail, DMS, and WhatsApp rather than copying secret/session files where possible.
4. Run smoke tests in this order:
   - Python import
   - Gmail auth
   - WhatsApp login
   - Tally check
   - Item master
   - Stock status
   - DMS stock
   - Purchase summary from known Excel
   - Unloading slip from known summaries
5. Implement `compare stock` as a first-class CLI and WhatsApp command.
6. Implement Tally company UI selection by configured data path/company folder if still required.
7. Add a proper `tests/` folder with fixture invoices and expected purchase-summary outputs.
8. Add validation for direct Tally purchase import only after repeated dummy-company testing.
9. Add a safe secrets setup guide:
   - what files are required
   - where to place them
   - how to re-authorize
   - what must never be committed

## Git Hygiene

Before pushing:

```powershell
git status --short
```

Do not commit:

```text
.venv/
data/token.json
data/dms_session.json
edge-profile/
downloads/
processed/
logs/
*.log
```

Review `.gitignore` and ensure secret/session files are excluded.

## Practical Recovery Notes

If PowerShell opens Notepad for `Activate.ps1`, do not rely on activation. Use explicit venv Python:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m src.main tally-check
```

If Tally commands fail with configured company not loaded:

1. Open Tally Prime.
2. Load the configured Goldiee company.
3. Confirm Tally HTTP is active.
4. Run:

```powershell
.\.venv\Scripts\python.exe -m src.main tally-check
```

If Gmail processing creates files but WhatsApp send fails:

1. Check `processed` for created files.
2. Run WhatsApp login.
3. Scan QR.
4. Restart monitor.

If DMS stock fails authorization:

```powershell
.\.venv\Scripts\python.exe -m src.main dms-authorize
```

Then retry:

```powershell
.\.venv\Scripts\python.exe -m src.main dms-stock
```

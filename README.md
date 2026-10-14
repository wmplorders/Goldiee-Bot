# Goldiee Purchase Bot

Local bot for WhatsApp purchase commands, preparing purchase-summary Excel files, fetching Tally stock, creating unloading slips, and monitoring Goldiee invoice emails.

## Current Rules

- WhatsApp commands:
  - `purchase summary`
  - `tally check`
  - `stock status`
  - `stock goldiee mail <fromdate> to <todate>`
  - `dms stock`
  - `unloading slip <date>`
  - `delivery vouchers`
- Gmail account: `goenkatradeventures@gmail.com`
- Sender watched: `information@goldiee.com`
- Start date: `2026-06-08`
- Emails 3 days or older are skipped and logged.
- Excel attachments with `pending` in the filename are ignored.
- If any Excel filename contains `invoice`, those are processed. Otherwise all non-pending Excel files are processed.
- All PDF attachments are forwarded as-is.
- Output workbook has:
  - `Raw Data`: original purchase Excel sheet
  - `Purchase Summary`: processed summary
- Purchase items are matched strictly and case-sensitively against Tally's complete Stock Item Master.
- Items not present in Tally are highlighted yellow and marked `NEW ITEM` in the Status column.
- Item-master and stock matching use the explicit Tally company configured in `config.json`
  or `.env`; they never silently use another loaded company.
- Output filename format: `<Invoice No> <Invoice Date> Summary.xlsx`

## First-Time Setup

1. Run `setup.bat`.
2. Create Gmail OAuth desktop credentials in Google Cloud.
3. Save that JSON file as `data\gmail_credentials.json`.
4. Run `start_whatsapp_login.bat` and scan the WhatsApp Web QR code in Edge.
5. Configure Tally company lock in `config.json` or copy `.env.example` to `.env` and edit:
   - `TALLY_DATA_PATH`
   - `TALLY_COMPANY_FOLDER`
   - `TALLY_COMPANY_NAME`
   - `TALLY_HTTP_URL`
6. Open Tally Prime with the configured company and run:

```bat
.\.venv\Scripts\python.exe -m src.main tally-check
```

7. Run `start_bot.bat`.
8. To start automatically with Windows, run `install_startup.bat`.

`start_bot.bat` keeps one Edge window open, listens for commands in unread WhatsApp chats,
replies in the same chat, and runs Gmail monitoring in the background. Gmail results are
sent to the configured `Goldiee Automation` group.

For `purchase summary`, attach the Excel to the command message or send the Excel shortly
before the command in the same chat.

For `delivery vouchers`, attach the transport-details Excel to the command message, send it
shortly before the command in the same chat, or put it in the bot's `uploads` folder. If no
WhatsApp Excel is available, the bot uses the newest Excel in `uploads`. The bot reads only
`Sheet1`, uses headers on row 10, uses the first `Despatch Date` in column B, skips invalid
rows while reporting them, and replies with one combined PDF containing one A4 page per
valid row.

## Delivery Vouchers

WhatsApp command:

```text
delivery vouchers
```

Manual command:

```bat
.\.venv\Scripts\python.exe -m src.main delivery-vouchers "C:\path\to\Transport Details.xlsx"
```

The generated PDF is saved in `processed`. `Eway No.` and `Remarks` may be blank; all other
mapped fields are mandatory. The final compact voucher template is isolated in
`src\delivery_voucher_template.py` so its layout can be changed independently of Excel and
WhatsApp processing.

## Manual Processing

Use:

```bat
process_file.bat "C:\path\to\invoice.xlsx"
```

Processed files are saved in `processed`.

To manually refresh the complete Tally item-master cache, run:

```bat
warm_item_master.bat
```

The bot also refreshes this cache at startup and checks the current Tally company context
when creating each purchase summary. Cache files use an exact-name company fingerprint.

## Tally Company Lock

Before Tally-dependent work, the bot checks that the exact configured company is loaded in
Tally. If it is missing, or if other companies are open while `close_other_companies` is
enabled, the bot stops and reports the problem instead of using the wrong company.

Current implementation enforces the safety check through Tally HTTP. Tally Prime UI
navigation for `Select Company > Specify Path` is intentionally kept separate; until that
UI automation is added, open/select the configured company manually before running imports
or stock commands.

## Stock Status

Use:

```bat
stock_status.bat
```

This fetches Tally's batch/PTR stock status using the loaded TDL-backed `Item Batch Summary` report and saves a timestamped Excel file in `downloads`.

## Goldiee Stock Aging Mail Report

Use the WhatsApp command:

```text
stock goldiee mail 01-07-2026 to 08-07-2026
```

Or run manually:

```bat
.\.venv\Scripts\python.exe -m src.main stock-goldiee-mail 01-07-2026 08-07-2026
```

This fills the configured `stock_aging.template` workbook from current Tally stock. It uses
Tally stock item aliases as item codes, stock groups as item categories, parsed item-name
weights where possible, base quantity in `In PCS`, and Tally alternate quantity in
`In CASES`. Manufacturing date, expiry date, and damage stock stay blank unless those
fields become available from Tally/DMS later.

## DMS Stock

Run `authorize_dms.bat` once and complete FieldAssist's required reCAPTCHA login. The bot
stores the resulting DMS session locally and uses direct DMS APIs for routine stock fetches.

Use the WhatsApp command `dms stock` or run `dms_stock.bat`. Routine DMS stock fetching is
browserless and does not open a visible window. The bot silently renews the portal session
while FieldAssist's identity session remains valid. If FieldAssist expires the identity
session and requires reCAPTCHA again, run `authorize_dms.bat` once to renew authorization.

## Unloading Slip

Use:

```bat
unloading_slip.bat 30-05-2026
```

This combines all purchase summaries for the date, fetches current stock status from Tally, matches item + MRP + PTR, and saves both `.xlsx` and `.pdf` unloading slips in `processed`.

## Logs

Processing log:

```text
logs\processed_purchases.csv
```

Bot state:

```text
data\state.json
```

## Notes

WhatsApp Web is intentionally configured to use Microsoft Edge with its own profile folder: `edge-profile`.
If WhatsApp is logged out, Edge will stay open for QR scanning.
# Undispatched Bills in Unloading Slips

Send `unloading slip 13-10-2026`. The bot asks whether any billed goods remain
undispatched. Reply in the same chat with full sales bill numbers separated by
commas or new lines, `none`, or `cancel`. A new unloading request replaces the
pending request in that chat. Pending requests are cleared when the bot restarts;
send the command again after a restart. On a validation failure, reply with the
corrected full list to retry.

The bot validates the bills, exports fresh stock, and adds their actual batch
quantities to the report only. Tally is not changed. Batch-level MRP and item-master
case conversions are used. The workbook includes an Undispatched Add-back audit
sheet. Missing or ambiguous bills, cancelled/optional vouchers, unsupported sales
ledgers, missing MRP/conversions, and stock-export errors stop generation.

The CLI asks the same question. For unattended use, supply
`--undispatched 26-27/SG-TI-340 26-27/SG-TI-341`, or `--undispatched` without values
to explicitly select none. Add-backs are never carried over to the next request.

## Local Web Dashboard

Run `start_web.bat` and open `http://localhost:9001`. No Node installation or
additional Python packages are needed beyond the existing bot requirements.
The dashboard runs independently of the WhatsApp listener.

Available screens: Purchase (Gmail by received date or uploaded invoice Excel),
Unloading (purchase invoice selection and undispatched sales bills), Tally/DMS
stock exports, Delivery Challans, Files & History, and Settings.
Web company selection is stored in `data/web_settings.json`; it does not alter
the WhatsApp configuration or Tally's default data path. Select the exact loaded
Goldiee company in Settings before processing. Web summaries require a fresh
item-master read rather than silently using an old cache.

Web jobs run serially, and job history is stored in `data/web_jobs`. Tally stock
exports share a process lock with the updated WhatsApp code. Restart any running
older bot after updating so it uses the same lock. Interrupted jobs are marked
failed on dashboard restart. Files are available through the dashboard; web jobs
do not send WhatsApp messages or mark Gmail messages processed.

For access from another device on your trusted LAN, stop the local web server,
set `GOLDIEE_WEB_PASSWORD` in the launching environment, then run:

```powershell
.\.venv\Scripts\python.exe -m src.web_app --host 0.0.0.0 --port 9001
```

Open `http://<bot-laptop-IP>:9001` from the other device. Permit the port only on
the Windows private-network firewall profile. LAN mode requires the password;
the password is not saved in the repository. This is a local HTTP service, not
an internet-facing deployment. Use localhost mode when network access is not
needed. Gmail authorization is performed using the existing `authorize_gmail.bat`
on the bot laptop, not an unattended web job.

The dashboard currently exposes Tally and DMS stock exports separately; stock
comparison, aging, and direct purchase import are not part of this first web version.

The Chat home screen provides action buttons, request/result messages, direct
file downloads, and retry buttons for failed jobs. Typing a workflow name opens
its form; the text box is navigation, not an AI assistant. Technical tracebacks
are available separately from the readable error message.

If an undispatched sales voucher omits batch MRP, the report checks the exact
item and exact PTR spelling in fresh stock, then freshly exported batch history.
It uses only a unique positive MRP, never the item-level MRP or a numerically
similar PTR. Conflicts and missing values block generation. The audit sheet
records the MRP source for each adjustment.

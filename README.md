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

The web interface is a single persistent chat. Its action buttons remain above the
message box: Purchase Summary, Unloading Slip, Stock Status, DMS Stock, Delivery
Challans, and Files. Settings stays in the top-right corner. Conversations are
saved by day in `data/web_conversations` and remain available after a restart.

Each action is a guided conversation. Unloading first previews the purchase
invoices found for the date, lets the operator select them, asks for any
undispatched sales bills, and shows a final review card. Tally is not contacted
until the operator presses `Create Slip` on that card. Long-running jobs keep
their status in the same conversation and recover after a browser refresh.
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

The chat provides request/result messages, direct file downloads, and retry
buttons for failed jobs. Typing a workflow name starts the matching guided flow;
the text box is command navigation, not an AI assistant. Technical tracebacks
are available separately from the readable error message.

If an undispatched sales voucher omits batch MRP, the report checks the exact
item and exact PTR spelling in fresh stock, then freshly exported batch history.
It uses only a unique positive MRP, never the item-level MRP or a numerically
similar PTR. Conflicts and missing values block generation. The audit sheet
records the MRP source for each adjustment.
# Sales Invoice Comparison

Use `compare <customer name> <date>`, for example `compare maa vaishnow 14/11/26`.
The date window is 8-14 November inclusive (the specified date and preceding six days).

Available through WhatsApp and the web chat's **Compare Sales** action. No WhatsApp
login is required when using the web interface.

1. Select one of the top three fuzzy matches from the configured Tally company's ledger list.
2. Select one or more active Tally sales invoices by the displayed numbers (e.g. `1,2` or `all`).
3. The bot requests a fresh DMS Sales Invoice Dump. Confirm the DMS party if its name differs,
   then select one or more DMS invoices.
4. Accept or reject proposed item-name pairs for this comparison only.
5. Confirm **Create Comparison** (reply `yes` in WhatsApp).

The reply shows exceptions first and attaches an Excel with the snapshot, full comparison,
original Tally item/ledger details and DMS raw rows. All rows are height 15. Quantities,
MRP, rate, effective discount, GST rates, taxable values, tax ledgers and final totals are
checked to two decimal places with zero tolerance. Identical item/price combinations are
consolidated; ambiguous multiple-price combinations remain separate rather than averaged.
Cancelled and optional Tally invoices are excluded. DMS requests exclude cancelled and
proforma invoices. Send `cancel` to stop a comparison.

This is read-only. It never alters vouchers, batches, ledgers or DMS records. It uses the
configured Tally company and the existing DMS authorization/company ID. If DMS authorization
expires, authorize it on the bot laptop and retry. Conversation snapshots are private files
under `data/sales_compare/`, expire after 24 hours, and are excluded from Git. A failed step
can be retried without advancing its saved state. The saved source snapshot is compared;
start a new comparison if either source has changed since selection.

Test: `.\.venv\Scripts\python.exe -m unittest discover -s tests`

If DMS reports expired authorization, run `authorize_dms.bat` on the bot laptop,
complete the sign-in in its dedicated Edge window, and wait for
`DMS authorization verified and saved`. The helper verifies API access and the
configured DMS company before it closes Edge. Then retry the failed web step.
Sessions are local to each laptop; Git does not distribute login tokens.

## Sales and Payment Report

WhatsApp / web chat: `sales report 1/11/26 to 18/11/26`.
The shorthand `sales report 18/11/26` starts on the first day of that month.
The web chat also has a **Sales Report** button with date selection and confirmation.
Use a date range within one month. Dates outside the requested range remain blank.

The bot returns an alphabetical monthly Excel and a shareable PNG containing party,
post code, the last date and preceding calendar date, and four totals/balance columns.
If the preceding date is outside the requested period, its snapshot cells stay blank.
All Excel rows are height 15. Dates have bold outlines and amounts use Indian comma
grouping, displayed as whole rupees while retaining paise internally.
PNG snapshots use the compact WhatsApp layout: full names wrap in a narrower party
column, totals have two-line headings, and each date block retains its bold outline.
Text is rendered directly at 3x resolution into a lossless, optimized 300-DPI PNG
(about 3989 pixels wide for normal figures). Long names/towns wrap without truncation,
and exceptionally large figures expand only their numeric column. Excel formatting
and report calculations are unchanged.

- Spices: final customer amounts of **Tax Invoice** and **Bill of Supply** vouchers.
- Detergent: final customer amounts of **Detergent Invoice** vouchers.
- Payment: credits to a customer against **Main Branch**, regardless of voucher type.
- Credit notes never reduce sales, but actual Tally closing balances include them and
  other ledger adjustments. Debit/credit balances are directly exported as of the end date.
- Ledgers under **Closed Customers**, including its descendants, are excluded entirely.
  TDS/TCS numbered tax-adjustment ledgers under Sundry Debtors are not customer rows.
- Cancelled/optional vouchers are excluded. Unexpected MAIN BRANCH reversals, missing
  parties, incomplete voucher batches or failed balance reconciliation stop the report.

New active customers prompt for their town. On a new laptop, upload a previous sales
report when prompted (**Import Towns from Previous Report** in the web chat). The first
two columns must be PARTY NAME and POST CODE; the bot matches only current active
customers, never imports its sales values, and blocks conflicting saved towns.
You can also reply with one town at a time. Towns are stored privately per configured
company in `data/sales_report_towns.json` and are not pushed to GitHub. Send `refresh`
or use **Recheck Customers** if customers have been moved into Closed Customers while
the bot was asking for towns. Send `cancel` to stop.

All access is read-only. The bot fetches lightweight voucher headers, then accounting
entries in batches of ten IDs, rather than a broad nested voucher export. It verifies
each customer's month-opening balance plus signed movements equals the closing balance,
and checks for group/balance changes while fetching. Pending town conversations expire
after 24 hours. PNG generation uses Pillow, not an external browser or Node installation.

After updating, install dependencies and restart the portal / WhatsApp listener:
```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

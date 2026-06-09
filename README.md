# Goldiee Purchase Bot

Local bot for WhatsApp purchase commands, preparing purchase-summary Excel files, fetching Tally stock, creating unloading slips, and monitoring Goldiee invoice emails.

## Current Rules

- WhatsApp commands:
  - `purchase summary`
  - `stock status`
  - `dms stock`
  - `unloading slip <date>`
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
- Item-master matching uses the explicit Tally company configured in `config.json`; it never
  silently uses another loaded company.
- Output filename format: `<Invoice No> <Invoice Date> Summary.xlsx`

## First-Time Setup

1. Run `setup.bat`.
2. Create Gmail OAuth desktop credentials in Google Cloud.
3. Save that JSON file as `data\gmail_credentials.json`.
4. Run `start_whatsapp_login.bat` and scan the WhatsApp Web QR code in Edge.
5. Run `start_bot.bat`.
6. To start automatically with Windows, run `install_startup.bat`.

`start_bot.bat` keeps one Edge window open, listens for commands in unread WhatsApp chats,
replies in the same chat, and runs Gmail monitoring in the background. Gmail results are
sent to the configured `Goldiee Automation` group.

For `purchase summary`, attach the Excel to the command message or send the Excel shortly
before the command in the same chat.

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

## Stock Status

Use:

```bat
stock_status.bat
```

This fetches Tally's batch/PTR stock status using the loaded TDL-backed `Item Batch Summary` report and saves a timestamped Excel file in `downloads`.

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

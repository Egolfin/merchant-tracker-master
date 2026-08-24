# Merchant Tracker Master

A single-page merchant operations command center designed around one Google Sheets workbook as the source of truth.

## Architecture

- `Leads` - one row per merchant / Store ID
- `Activity_Log` - calls, emails, meetings, follow-ups
- `Photos_Added` - photo changes with automatic subtraction
- `Videos_Uploaded` - video activity
- `Approvals` - photo/video approval history
- `Photoshoot_Tracker` - photoshoot history
- `Case_Tracker` - support cases
- `Opportunities` - SL, Promotions, Co-Funding and other opportunities
- `Assets` - merchant assets
- `Audit_Log` - optional technical/audit history

## Key principle

The connected Google Sheets workbook is the single source of truth. The tracker uses `Store ID` as the primary merchant key and writes operational activity into structured tabs in the same workbook.

## Front-end

The app is intentionally kept as a single `index.html` so it is easy to deploy as a static web app.

## Google Sheets integration

The production write/read layer should be provided by a Google Apps Script Web App bound to the target workbook. The front-end accepts the workbook URL and uses the Apps Script endpoint to read/write the workbook.

Do not place credentials, API keys, or service-account secrets in this repository.

## Deployment

1. Create the Google Sheets workbook.
2. Add the required tabs above.
3. Deploy the companion Apps Script Web App with access appropriate to your Google Workspace.
4. Enter the workbook URL and Apps Script Web App URL in the tracker configuration.
5. Test read/write with a sandbox merchant before production use.

## Current master

The included `index.html` contains the current master tracker UI and functionality developed in this conversation, including merchant search, dashboard analytics, quick filters, activity logging, Photos Added calculations, Merchant 360, and lead-list connection controls.

# Merchant Tracker Master

## Current source of truth

This build is locked to the Google Sheet workbook:

Copy of CA Lead List Q3 - 2026- Esteban Golfin
Spreadsheet ID: 19dtcq3g291NpwvgUBvIIzaC4pt6eQqpQrpHN5Jyu5Xs

The `Leads` tab is the merchant master source. Store ID is the primary merchant key.

## Why the lead list was returning 0

The front end was expecting a JSONP response shape that was not guaranteed by the previous Apps Script implementation. The new `Code.gs` returns both:

- `rows`
- `data.rows`
- `headers`
- `data.headers`
- `count`

The front end now accepts all of these compatible shapes and displays the number of loaded merchants.

## Install / deploy backend

1. Open the master Google Sheet.
2. Extensions -> Apps Script.
3. Replace the script with `Code.gs` from this repository.
4. Save.
5. Deploy -> New deployment -> Web app.
6. Execute as: Me.
7. Use the access level appropriate for your Workspace environment.
8. Copy the `/exec` URL.
9. Put that URL into `APPS_SCRIPT_URL` inside `index.html`.
10. Redeploy the web app hosting `index.html`.

The backend is locked to the master workbook and does not use the legacy tracker database.

## Automatic tracker tabs

The first backend call creates these tabs in the SAME workbook if they do not already exist:

- Tracker_Activity_Log
- Tracker_Photos_Added
- Tracker_Videos
- Tracker_Approvals
- Tracker_Photoshoots
- Tracker_Cases
- Tracker_Opportunities
- Tracker_Assets
- Tracker_Audit

The existing `Leads` and `_Logs` tabs are not replaced.

## Lead-list read contract

GET action:

`getLeadList`

Required params:
- `spreadsheetId`
- `sheetName=Leads`
- `callback`

Response contains `success`, `spreadsheetId`, `spreadsheetName`, `sheetName`, `headers`, `rows`, `data`, and `count`.

## Production note

Do not place OAuth credentials or secrets in the HTML repository.

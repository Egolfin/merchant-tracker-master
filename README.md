# Merchant Tracker Master - Single Workbook Edition

This build is intentionally locked to the new master workbook:
**Copy of CA Lead List Q3 - 2026- Esteban Golfin**

Workbook ID:
`19dtcq3g291NpwvgUBvIIzaC4pt6eQqpQrpHN5Jyu5Xs`

## Source of truth

The `Leads` tab is the only merchant master source. Existing lead-list URLs and browser caches from previous tracker versions are explicitly cleared on initialization.

## Same-workbook storage

The tracker creates and uses these tabs in the same workbook:

- `Leads`
- `Tracker_Activity_Log`
- `Tracker_Photos_Added`
- `Tracker_Videos`
- `Tracker_Approvals`
- `Tracker_Photoshoots`
- `Tracker_Cases`
- `Tracker_Opportunities`
- `Tracker_Assets`
- `Tracker_Audit`

The existing `_Logs` tab is preserved.

## Required deployment step

Deploy `Code.gs` as a Google Apps Script Web App bound to the master workbook.

Then paste the resulting `/exec` URL into `APPS_SCRIPT_URL` in `index.html`.

The front-end uses:
- `getLeadList` to load only the `Leads` tab
- `getWorkbook` to load the tracker tabs from the same workbook
- POST writes for new records and edits/deletes

## Merchant key

`Store ID` is the primary merchant key. Lead fields are pulled from the `Leads` tab and operational records reference the same Store ID.

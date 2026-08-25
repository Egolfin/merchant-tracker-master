/**
 * Merchant Tracker Master - Google Apps Script backend
 *
 * IMPORTANT:
 * Bind this script to the master Google Sheet or use it as a Web App.
 * The backend is locked to the approved master workbook.
 */

const CONFIG = {
  SPREADSHEET_ID: '19dtcq3g291NpwvgUBvIIzaC4pt6eQqpQrpHN5Jyu5Xs',
  LEAD_SHEET_NAME: 'Leads',
  TRACKER_TABS: {
    calls: 'Tracker_Activity_Log',
    photosAdded: 'Tracker_Photos_Added',
    videosUploaded: 'Tracker_Videos',
    contentApprovals: 'Tracker_Approvals',
    photoshootTracker: 'Tracker_Photoshoots',
    caseTracker: 'Tracker_Cases',
    opportunities: 'Tracker_Opportunities',
    assets: 'Tracker_Assets',
    audit: 'Tracker_Audit'
  }
};

function getWorkbook_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = String(params.action || 'getSetupStatus').trim();
    const result = routeGet_(action, params);
    return output_(result, params.callback);
  } catch (err) {
    return output_({
      success: false,
      message: err && err.message ? err.message : String(err)
    }, e && e.parameter ? e.parameter.callback : '');
  }
}

function doPost(e) {
  try {
    const raw = e && e.parameter ? e.parameter.payload : '';
    const payload = raw ? JSON.parse(raw) : {};
    const result = handleWrite_(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: err && err.message ? err.message : String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function routeGet_(action, params) {
  switch (action) {
    case 'ping':
      return {
        success: true,
        action: 'ping',
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        spreadsheetName: getWorkbook_().getName(),
        leadSheet: CONFIG.LEAD_SHEET_NAME,
        timestamp: new Date().toISOString()
      };
    case 'getLeadList':
      return getLeadList_(params);
    case 'getWorkbook':
      return getWorkbookSnapshot_();
    case 'getSetupStatus':
      return getSetupStatus_();
    case 'initializeWorkbook':
      return initializeWorkbook_();
    default:
      throw new Error('Unsupported GET action: ' + action);
  }
}

function getLeadList_(params) {
  const ss = getWorkbook_();
  const requestedSpreadsheetId = String(params.spreadsheetId || '').trim();
  if (requestedSpreadsheetId && requestedSpreadsheetId !== CONFIG.SPREADSHEET_ID) {
    throw new Error('Spreadsheet ID does not match the approved master workbook.');
  }

  const sheetName = String(params.sheetName || CONFIG.LEAD_SHEET_NAME).trim();
  if (sheetName !== CONFIG.LEAD_SHEET_NAME) {
    throw new Error('Only the Leads tab is allowed for merchant master data.');
  }

  const sheet = ss.getSheetByName(CONFIG.LEAD_SHEET_NAME);
  if (!sheet) throw new Error('Leads tab not found.');

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetName: CONFIG.LEAD_SHEET_NAME,
      headers: [],
      rows: [],
      data: { headers: [], rows: [] },
      count: 0
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = dedupeHeaders_(values[0]);
  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => String(v).trim() === '')) continue;
    rows.push(rowToObject_(headers, row));
  }

  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetName: CONFIG.LEAD_SHEET_NAME,
    headers,
    rows,
    data: { headers, rows },
    count: rows.length
  };
}

function getWorkbookSnapshot_() {
  const ss = getWorkbook_();
  ensureTrackerSheets_(ss);
  const database = {};
  Object.keys(CONFIG.TRACKER_TABS).forEach(key => {
    const tab = CONFIG.TRACKER_TABS[key];
    database[tab] = readSheetObjects_(ss.getSheetByName(tab));
  });

  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    leadSheet: CONFIG.LEAD_SHEET_NAME,
    data: database,
    rows: database
  };
}

function getSetupStatus_() {
  const ss = getWorkbook_();
  const missing = [];
  Object.keys(CONFIG.TRACKER_TABS).forEach(key => {
    if (!ss.getSheetByName(CONFIG.TRACKER_TABS[key])) missing.push(CONFIG.TRACKER_TABS[key]);
  });
  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    leadSheetExists: Boolean(ss.getSheetByName(CONFIG.LEAD_SHEET_NAME)),
    leadRowCount: Math.max(0, ss.getSheetByName(CONFIG.LEAD_SHEET_NAME)?.getLastRow() - 1 || 0),
    missingTrackerTabs: missing,
    ready: Boolean(ss.getSheetByName(CONFIG.LEAD_SHEET_NAME)) && missing.length === 0
  };
}

function initializeWorkbook_() {
  const ss = getWorkbook_();
  ensureTrackerSheets_(ss);
  return getSetupStatus_();
}

function ensureTrackerSheets_(ss) {
  Object.keys(CONFIG.TRACKER_TABS).forEach(key => {
    const name = CONFIG.TRACKER_TABS[key];
    if (ss.getSheetByName(name)) return;
    const sheet = ss.insertSheet(name);
    const headers = defaultHeadersFor_(key);
    if (headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  });
}

function defaultHeadersFor_(key) {
  const common = ['ID', 'Timestamp', 'Store ID', 'Business ID', 'Business Name', 'Owner', 'Notes'];
  if (key === 'calls') return common.concat(['Activity Type', 'Call Purpose', 'Call Disposition', 'Follow-Up Date', 'Follow-Up Time']);
  if (key === 'photosAdded') return common.concat(['Previous Missing Photos', 'Current Missing Photos', 'Photos Added']);
  if (key === 'videosUploaded') return common.concat(['Video Type']);
  if (key === 'contentApprovals') return common.concat(['Approval Type']);
  if (key === 'photoshootTracker') return common.concat(['Shoot Date', 'Shoot Time', 'Status']);
  if (key === 'caseTracker') return common.concat(['Case #', 'Case Type', 'Status', 'Priority']);
  if (key === 'opportunities') return common.concat(['Opportunity Type', 'Co-Funded', 'Co-Funded Split', 'Budget', 'Salesforce Link']);
  if (key === 'assets') return common.concat(['Asset Type', 'Asset URL', 'Status']);
  if (key === 'audit') return ['ID', 'Timestamp', 'Operation', 'Target Tab', 'Store ID', 'Record ID', 'Status', 'Message'];
  return common;
}

function handleWrite_(payload) {
  const ss = getWorkbook_();
  ensureTrackerSheets_(ss);

  if (payload && payload.operationType === 'edit') {
    return editRecord_(ss, payload.targetTab, payload.recordId, payload.updatedFields || {});
  }

  if (payload && payload.operationType === 'delete') {
    return deleteRecord_(ss, payload.targetTab, payload.recordId);
  }

  const formSource = String(payload.formSource || '').trim();
  const targetTab = CONFIG.TRACKER_TABS[formSource] || formSource;
  if (!targetTab) throw new Error('No target tracker tab supplied.');

  const record = normalizePayloadForWrite_(payload);
  appendRecord_(ss.getSheetByName(targetTab), record);
  maybeUpdateLead_(ss, record, formSource);
  writeAudit_(ss, 'create', targetTab, record, 'OK', 'Record created');

  return { success: true, targetTab, record };
}

function normalizePayloadForWrite_(payload) {
  const out = Object.assign({}, payload);
  delete out.operationType;

  const now = new Date();
  out.Timestamp = out.Timestamp || out.timestamp || now.toISOString();
  if (!out.ID) out.ID = makeId_(out.formSource || 'REC');

  if (out.formSource === 'photosAdded' || out['Photos Added'] !== undefined) {
    const prev = number_(out['Previous Missing Photos'] ?? out.prevMissing ?? out['Prev Missing']);
    const curr = number_(out['Current Missing Photos'] ?? out.currMissing ?? out['Curr Missing']);
    if (curr > prev) throw new Error('Current missing photos cannot be greater than previous missing photos.');
    out['Previous Missing Photos'] = prev;
    out['Current Missing Photos'] = curr;
    out['Photos Added'] = Math.max(0, prev - curr);
  }

  return out;
}

function appendRecord_(sheet, record) {
  if (!sheet) throw new Error('Target tracker sheet not found.');
  const headers = getOrCreateHeaders_(sheet, record);
  const row = headers.map(h => record[h] !== undefined && record[h] !== null ? record[h] : '');
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
}

function getOrCreateHeaders_(sheet, record) {
  let lastColumn = sheet.getLastColumn();
  let headers = lastColumn ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String) : [];
  headers = headers.filter((_, i) => i < lastColumn);
  const incoming = Object.keys(record).filter(k => k !== 'operationType');
  const missing = incoming.filter(k => headers.indexOf(k) === -1);
  if (missing.length) {
    const start = headers.length + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return headers;
}

function editRecord_(ss, targetTab, recordId, updatedFields) {
  const sheet = ss.getSheetByName(targetTab);
  if (!sheet) throw new Error('Tracker tab not found: ' + targetTab);
  const found = findRecord_(sheet, recordId);
  if (!found) throw new Error('Record not found: ' + recordId);

  const headers = getOrCreateHeaders_(sheet, updatedFields);
  headers.forEach((h, i) => {
    if (updatedFields[h] !== undefined) sheet.getRange(found.row, i + 1).setValue(updatedFields[h]);
  });

  if (targetTab === CONFIG.TRACKER_TABS.photosAdded) {
    const map = headerMap_(sheet);
    const prevCol = map['Previous Missing Photos'] || map['Prev Missing'];
    const currCol = map['Current Missing Photos'] || map['Curr Missing'];
    const photoCol = map['Photos Added'];
    if (prevCol && currCol && photoCol) {
      const prev = number_(sheet.getRange(found.row, prevCol).getDisplayValue());
      const curr = number_(sheet.getRange(found.row, currCol).getDisplayValue());
      sheet.getRange(found.row, photoCol).setValue(Math.max(0, prev - curr));
    }
  }

  const updated = readRowObject_(sheet, found.row);
  maybeUpdateLead_(ss, updated, keyForTab_(targetTab));
  writeAudit_(ss, 'edit', targetTab, updated, 'OK', 'Record updated');
  return { success: true, targetTab, record: updated };
}

function deleteRecord_(ss, targetTab, recordId) {
  const sheet = ss.getSheetByName(targetTab);
  if (!sheet) throw new Error('Tracker tab not found: ' + targetTab);
  const found = findRecord_(sheet, recordId);
  if (!found) throw new Error('Record not found: ' + recordId);
  const record = readRowObject_(sheet, found.row);
  sheet.deleteRow(found.row);
  writeAudit_(ss, 'delete', targetTab, record, 'OK', 'Record deleted');
  return { success: true, targetTab, record };
}

function findRecord_(sheet, recordId) {
  const map = headerMap_(sheet);
  const idCol = map['ID'];
  if (!idCol) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, idCol, lastRow - 1, 1).getDisplayValues();
  const wanted = String(recordId || '').trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === wanted) return { row: i + 2 };
  }
  return null;
}

function readSheetObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  const headers = dedupeHeaders_(values[0]);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(v => String(v).trim() === '')) continue;
    rows.push(rowToObject_(headers, values[i]));
  }
  return rows;
}

function readRowObject_(sheet, rowNumber) {
  const headers = dedupeHeaders_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
  return rowToObject_(headers, values);
}

function maybeUpdateLead_(ss, record, formSource) {
  const storeId = String(record['Store ID'] || record['Store Id'] || record.storeId || '').trim();
  if (!storeId) return;
  const sheet = ss.getSheetByName(CONFIG.LEAD_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const map = headerMap_(sheet);
  const storeCol = map['Store Id'] || map['Store ID'];
  if (!storeCol) return;

  const values = sheet.getRange(2, storeCol, sheet.getLastRow() - 1, 1).getDisplayValues();
  let rowNumber = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === storeId) {
      rowNumber = i + 2;
      break;
    }
  }
  if (rowNumber < 0) return;

  const updates = {};
  const now = record.Timestamp || new Date().toISOString();
  if (map['Last Contacted'] && (formSource === 'calls' || formSource === 'activityLog')) updates['Last Contacted'] = now;
  if (map['Last Activity Type']) updates['Last Activity Type'] = record['Activity Type'] || record['Call Purpose'] || formSource;
  if (map['Last Activity At']) updates['Last Activity At'] = now;
  if (map['Notes Summary']) {
    const noteValue = record['Notes'] || record['Call Notes'] || record['caseNotes'] || record['approvalNotes'] || record['shootNotes'];
    if (noteValue) updates['Notes Summary'] = String(noteValue).slice(0, 500);
  }
  if (map['Next Follow-Up'] && (record['Follow-Up Date'] || record.followUpDate || record['Next Follow-Up'])) {
    updates['Next Follow-Up'] = record['Follow-Up Date'] || record.followUpDate || record['Next Follow-Up'];
  }

  Object.keys(updates).forEach(header => sheet.getRange(rowNumber, map[header]).setValue(updates[header]));
}

function writeAudit_(ss, operation, targetTab, record, status, message) {
  const sheet = ss.getSheetByName(CONFIG.TRACKER_TABS.audit);
  if (!sheet) return;
  const row = [
    makeId_('AUDIT'), new Date().toISOString(), operation, targetTab,
    record['Store ID'] || record['Store Id'] || '', record.ID || '', status, message
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function keyForTab_(tab) {
  for (const key in CONFIG.TRACKER_TABS) if (CONFIG.TRACKER_TABS[key] === tab) return key;
  return tab;
}

function headerMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return {};
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (String(h).trim()) map[String(h).trim()] = i + 1;
  });
  return map;
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) obj[key] = row[i] === undefined ? '' : row[i];
  });
  return obj;
}

function dedupeHeaders_(headers) {
  const seen = {};
  return headers.map((header, index) => {
    const base = String(header || '').trim() || ('Column ' + (index + 1));
    if (!seen[base]) {
      seen[base] = 1;
      return base;
    }
    seen[base] += 1;
    return base + ' ' + seen[base];
  });
}

function number_(value) {
  const n = Number(String(value == null ? '' : value).replace(/[$,%\s,]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function makeId_(prefix) {
  const tz = Session.getScriptTimeZone() || 'America/Guatemala';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  const rand = Math.floor(Math.random() * 1e6).toString().padStart(6, '0');
  return String(prefix || 'REC').toUpperCase() + '-' + stamp + '-' + rand;
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

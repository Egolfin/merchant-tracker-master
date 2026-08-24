/**
 * Merchant Tracker Master - Google Apps Script backend
 * Bind this project to the master Google Sheet:
 * Copy of CA Lead List Q3 - 2026- Esteban Golfin
 *
 * Deploy as a Web App:
 * Execute as: Me
 * Who has access: the intended Google Workspace users
 */

const MASTER_SHEET_ID = '19dtcq3g291NpwvgUBvIIzaC4pt6eQqpQrpHN5Jyu5Xs';
const LEAD_SHEET_NAME = 'Leads';
const PROTECTED_SHEETS = new Set(['Leads', '_Logs']);

const TRACKER_TABS = {
  calls: 'Tracker_Activity_Log',
  photosAdded: 'Tracker_Photos_Added',
  videosUploaded: 'Tracker_Videos',
  contentApprovals: 'Tracker_Approvals',
  photoshootTracker: 'Tracker_Photoshoots',
  caseTracker: 'Tracker_Cases',
  opportunities: 'Tracker_Opportunities',
  assets: 'Tracker_Assets',
  audit: 'Tracker_Audit'
};

const HEADERS = {
  Tracker_Activity_Log: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Rx Name','Activity Type',
    'Call Purpose','Call Disposition','Notes','Outcome','Owner','Next Follow-Up','Form Source'
  ],
  Tracker_Photos_Added: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Previous Missing Photos',
    'Current Missing Photos','Photos Added','Notes','Owner','Form Source'
  ],
  Tracker_Videos: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Notes','Owner','Form Source'
  ],
  Tracker_Approvals: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Approval Type','Notes','Owner','Form Source'
  ],
  Tracker_Photoshoots: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Shoot Date','Shoot Time','Status','Notes','Owner','Form Source'
  ],
  Tracker_Cases: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Case Number','Case Type','Status','Notes','Owner','Form Source'
  ],
  Tracker_Opportunities: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Opportunity Type','Co-Funded','Split','Budget','Salesforce Link','Notes','Owner','Status','Pipeline Stage','Form Source'
  ],
  Tracker_Assets: [
    'ID','Timestamp','Store ID','Business ID','Business Name','Asset Type','Asset Status','Asset URL','Notes','Owner','Form Source'
  ],
  Tracker_Audit: [
    'ID','Timestamp','Action','Target Tab','Record ID','Store ID','Details','Actor'
  ]
};

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'getWorkbook').trim();
  try {
    const data = routeGet_(action, e || {});
    return respond_(data, e && e.parameter ? e.parameter.callback : '');
  } catch (err) {
    return respond_({ status: 'error', message: err.message, error: String(err) }, e && e.parameter ? e.parameter.callback : '');
  }
}

function doPost(e) {
  try {
    const raw = e && e.parameter ? e.parameter.payload : '';
    const payload = raw ? JSON.parse(raw) : {};
    const result = routeWrite_(payload || {});
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function routeGet_(action, e) {
  switch (action) {
    case 'getLeadList':
      return { rows: readSheetObjects_(getMaster_().getSheetByName(LEAD_SHEET_NAME)) };
    case 'getWorkbook':
      return getTrackerWorkbook_();
    case 'health':
      return { ok: true, workbookId: getMaster_().getId(), title: getMaster_().getName(), timestamp: new Date().toISOString() };
    default:
      throw new Error('Unsupported GET action: ' + action);
  }
}

function routeWrite_(payload) {
  if (payload.operationType === 'edit') return editRecord_(payload);
  if (payload.operationType === 'delete') return deleteRecord_(payload);
  return appendActivity_(payload);
}

function getMaster_() {
  return SpreadsheetApp.openById(MASTER_SHEET_ID);
}

function ensureTrackerTabs_() {
  const ss = getMaster_();
  Object.values(TRACKER_TABS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    ensureHeaders_(sheet, HEADERS[name] || []);
  });
  return ss;
}

function ensureHeaders_(sheet, headers) {
  if (!headers.length) return;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    if (existing.every(v => String(v || '').trim() === '')) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
}

function getTrackerWorkbook_() {
  const ss = ensureTrackerTabs_();
  const output = {};
  Object.values(TRACKER_TABS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    output[name] = readSheetObjects_(sheet, 20000);
  });
  return output;
}

function readSheetObjects_(sheet, maxRows) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  const rowCount = Math.min(Math.max(sheet.getLastRow() - 1, 0), maxRows || 20000);
  if (!rowCount) return [];
  const colCount = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, rowCount + 1, colCount).getDisplayValues();
  const headers = values[0].map(v => String(v || '').trim());
  return values.slice(1).filter(row => row.some(v => String(v || '').trim() !== '')).map(row => {
    const out = {};
    headers.forEach((header, i) => { if (header) out[header] = row[i] || ''; });
    return out;
  });
}

function appendActivity_(payload) {
  const ss = ensureTrackerTabs_();
  const category = String(payload.formSource || '').trim();
  const targetTab = TRACKER_TABS[category];
  if (!targetTab) throw new Error('Unknown form source: ' + category);

  const storeId = String(payload['Store ID'] || payload.storeId || '').trim();
  const businessId = String(payload['Business ID'] || payload.businessId || '').trim();
  const businessName = String(payload['Business Name'] || payload.businessName || payload.storeName || '').trim();
  const now = new Date();
  const id = 'TRK-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 10000);
  const owner = String(payload.Owner || payload.owner || 'Esteban Golfin').trim();

  const row = buildRowForCategory_(category, payload, { id, now, storeId, businessId, businessName, owner });
  const sheet = ss.getSheetByName(targetTab);
  const headers = HEADERS[targetTab];
  const values = headers.map(h => row[h] !== undefined ? row[h] : '');
  sheet.appendRow(values);

  updateLeadSummary_(storeId, category, row);
  audit_('CREATE', targetTab, id, storeId, JSON.stringify({ businessName, category }), owner);
  return { id, targetTab, row };
}

function buildRowForCategory_(category, payload, ctx) {
  const p = payload || {};
  const out = { ID: ctx.id, Timestamp: ctx.now, 'Store ID': ctx.storeId, 'Business ID': ctx.businessId, 'Business Name': ctx.businessName, Owner: ctx.owner, 'Form Source': category };
  if (category === 'calls') {
    out['Rx Name'] = String(p.rxName || '').trim();
    out['Activity Type'] = 'Call';
    out['Call Purpose'] = String(p.callGoal || p['Call Purpose'] || '').trim();
    out['Call Disposition'] = String(p.callResult || p['Call Disposition'] || '').trim();
    out.Notes = String(p.callNotes || p.Notes || '').trim();
    out.Outcome = String(p.callOutcome || p.Outcome || '').trim();
    out['Next Follow-Up'] = combineFollowUp_(p.followUpDate, p.followUpTime);
  } else if (category === 'photosAdded') {
    const prev = toNumber_(p.prevMissing ?? p['Previous Missing Photos']);
    const curr = toNumber_(p.currMissing ?? p['Current Missing Photos']);
    out['Previous Missing Photos'] = prev;
    out['Current Missing Photos'] = curr;
    out['Photos Added'] = Math.max(0, prev - curr);
    out.Notes = String(p.photosNotes || p.Notes || '').trim();
  } else if (category === 'videosUploaded') {
    out.Notes = String(p.videoNotes || p.Notes || '').trim();
  } else if (category === 'contentApprovals') {
    out['Approval Type'] = String(p.approvalType || '').trim();
    out.Notes = String(p.approvalNotes || '').trim();
  } else if (category === 'photoshootTracker') {
    out['Shoot Date'] = String(p.shootDate || '').trim();
    out['Shoot Time'] = String(p.shootTime || '').trim();
    out.Status = String(p.shootStatus || '').trim();
    out.Notes = String(p.shootNotes || '').trim();
  } else if (category === 'caseTracker') {
    out['Case Number'] = String(p.caseNumber || '').trim();
    out['Case Type'] = String(p.caseType || '').trim();
    out.Status = String(p.caseResolved || '').trim();
    out.Notes = String(p.caseNotes || '').trim();
  } else if (category === 'opportunities') {
    out['Opportunity Type'] = String(p.oppType || '').trim();
    out['Co-Funded'] = String(p.isCoFunded || 'No').trim();
    out.Split = String(p.coFundedSplit || '').trim();
    out.Budget = String(p.oppBudget || '').trim();
    out['Salesforce Link'] = String(p.salesforceLink || '').trim();
    out.Notes = String(p.oppNotes || p.Notes || '').trim();
    out.Status = String(p.oppStatus || '').trim();
    out['Pipeline Stage'] = String(p.pipelineStage || '').trim();
  }
  return out;
}

function combineFollowUp_(dateValue, timeValue) {
  const d = String(dateValue || '').trim();
  const t = String(timeValue || '').trim();
  if (!d) return '';
  return t ? d + ' ' + t : d;
}

function editRecord_(payload) {
  const targetTab = String(payload.targetTab || '').trim();
  const recordId = String(payload.recordId || '').trim();
  if (!targetTab || !recordId) throw new Error('Missing target tab or record ID.');
  if (!Object.values(TRACKER_TABS).includes(targetTab)) throw new Error('Invalid tracker tab.');

  const ss = ensureTrackerTabs_();
  const sheet = ss.getSheetByName(targetTab);
  const headers = HEADERS[targetTab];
  const rows = sheet.getDataRange().getValues();
  const idCol = headers.indexOf('ID');
  if (idCol < 0) throw new Error('ID column missing from ' + targetTab);
  let rowNumber = -1;
  for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][idCol]) === recordId) { rowNumber = r + 1; break; }
  }
  if (rowNumber < 0) throw new Error('Record not found: ' + recordId);

  const updates = payload.updatedFields || {};
  const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  Object.entries(updates).forEach(([key, value]) => {
    const idx = headers.indexOf(key);
    if (idx >= 0 && key !== 'ID' && key !== 'Timestamp') rowValues[idx] = value;
  });

  if (targetTab === TRACKER_TABS.photosAdded) {
    const prev = toNumber_(rowValues[headers.indexOf('Previous Missing Photos')]);
    const curr = toNumber_(rowValues[headers.indexOf('Current Missing Photos')]);
    rowValues[headers.indexOf('Photos Added')] = Math.max(0, prev - curr);
  }

  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([rowValues]);
  const storeId = String(rowValues[headers.indexOf('Store ID')] || '').trim();
  const owner = String(rowValues[headers.indexOf('Owner')] || 'Esteban Golfin').trim();
  updateLeadSummary_(storeId, categoryFromTab_(targetTab), objectFromRow_(headers, rowValues));
  audit_('EDIT', targetTab, recordId, storeId, JSON.stringify(updates), owner);
  return { id: recordId, targetTab };
}

function deleteRecord_(payload) {
  const targetTab = String(payload.targetTab || '').trim();
  const recordId = String(payload.recordId || '').trim();
  if (!targetTab || !recordId) throw new Error('Missing target tab or record ID.');
  if (!Object.values(TRACKER_TABS).includes(targetTab)) throw new Error('Invalid tracker tab.');

  const ss = ensureTrackerTabs_();
  const sheet = ss.getSheetByName(targetTab);
  const headers = HEADERS[targetTab];
  const idCol = headers.indexOf('ID');
  const values = sheet.getDataRange().getValues();
  let rowNumber = -1;
  let storeId = '';
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === recordId) {
      rowNumber = r + 1;
      storeId = String(values[r][headers.indexOf('Store ID')] || '').trim();
      break;
    }
  }
  if (rowNumber < 0) throw new Error('Record not found: ' + recordId);
  sheet.deleteRow(rowNumber);
  audit_('DELETE', targetTab, recordId, storeId, 'Record deleted', Session.getActiveUser().getEmail() || 'Tracker');
  if (targetTab === TRACKER_TABS.caseTracker && storeId) refreshOpenCaseCount_(storeId);
  return { id: recordId, targetTab };
}

function categoryFromTab_(tabName) {
  return Object.keys(TRACKER_TABS).find(k => TRACKER_TABS[k] === tabName) || '';
}

function objectFromRow_(headers, values) {
  const out = {};
  headers.forEach((h, i) => out[h] = values[i]);
  return out;
}

function findLeadRow_(storeId) {
  const sheet = getMaster_().getSheetByName(LEAD_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(v => String(v || '').trim());
  const storeIdx = headers.findIndex(h => ['Store Id','Store ID','StoreId'].includes(h));
  if (storeIdx < 0) return null;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][storeIdx]).trim() === String(storeId).trim()) return { sheet, headers, values, rowNumber: r + 1, row: values[r] };
  }
  return null;
}

function updateLeadSummary_(storeId, category, row) {
  if (!storeId) return;
  const lead = findLeadRow_(storeId);
  if (!lead) return;
  const index = {};
  lead.headers.forEach((h, i) => index[h] = i);
  const setByNames = (names, value) => {
    const idx = names.map(n => index[n]).find(i => i !== undefined);
    if (idx !== undefined && idx >= 0) lead.row[idx] = value;
  };

  setByNames(['Last Contacted'], category === 'calls' ? row.Timestamp : lead.row[index['Last Contacted']]);
  setByNames(['Last Activity Type'], activityLabel_(category, row));
  setByNames(['Last Activity At'], row.Timestamp || new Date());
  if (row['Next Follow-Up']) setByNames(['Next Follow-Up'], row['Next Follow-Up']);

  const touchIdx = ['Touchpoint Count','Touchpoints'].map(n => index[n]).find(i => i !== undefined);
  if (touchIdx !== undefined) lead.row[touchIdx] = toNumber_(lead.row[touchIdx]) + 1;

  if (category === 'calls' && row['Call Disposition']) {
    if (String(row['Call Disposition']).toLowerCase().includes('not interested')) setByNames(['Lead Status'], 'Not Interested');
  }
  if (category === 'opportunities' && row['Pipeline Stage']) setByNames(['Pipeline Stage'], row['Pipeline Stage']);
  if (category === 'caseTracker') refreshOpenCaseCount_(storeId);

  lead.sheet.getRange(lead.rowNumber, 1, 1, lead.headers.length).setValues([lead.row]);
}

function refreshOpenCaseCount_(storeId) {
  const ss = ensureTrackerTabs_();
  const sheet = ss.getSheetByName(TRACKER_TABS.caseTracker);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(v => String(v || '').trim());
  const storeIdx = headers.indexOf('Store ID');
  const statusIdx = headers.indexOf('Status');
  let count = 0;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][storeIdx]).trim() !== String(storeId).trim()) continue;
    const status = String(values[r][statusIdx] || '').toLowerCase();
    if (!['resolved','closed','solved w/ no resolution'].includes(status)) count++;
  }
  const lead = findLeadRow_(storeId);
  if (!lead) return;
  const idx = lead.headers.indexOf('Open Case Count');
  if (idx >= 0) {
    lead.sheet.getRange(lead.rowNumber, idx + 1).setValue(count);
  }
}

function activityLabel_(category, row) {
  if (category === 'calls') return 'Call - ' + String(row['Call Disposition'] || row['Call Purpose'] || 'Logged');
  if (category === 'photosAdded') return 'Photos Added - ' + String(row['Photos Added'] || 0);
  if (category === 'videosUploaded') return 'Video Logged';
  if (category === 'contentApprovals') return 'Approval - ' + String(row['Approval Type'] || 'Logged');
  if (category === 'photoshootTracker') return 'Photoshoot - ' + String(row.Status || 'Logged');
  if (category === 'caseTracker') return 'Case - ' + String(row.Status || 'Logged');
  if (category === 'opportunities') return 'Opportunity - ' + String(row['Opportunity Type'] || 'Logged');
  return category;
}

function audit_(action, targetTab, recordId, storeId, details, actor) {
  const ss = ensureTrackerTabs_();
  const sheet = ss.getSheetByName(TRACKER_TABS.audit);
  sheet.appendRow([Utilities.getUuid(), new Date(), action, targetTab, recordId, storeId, details || '', actor || 'Tracker']);
}

function toNumber_(value) {
  const n = Number(String(value ?? '').replace(/[$,%\s,]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

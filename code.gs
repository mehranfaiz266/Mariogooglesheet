// ─── CONFIGURATION ─────────────────────────────────────────────────────────────
const CONFIG = {
  API_BASE: 'https://services.leadconnectorhq.com',
  TOKEN_PROPERTY: 'GHL_PRIVATE_TOKEN',
  LOCATION_PROPERTY: 'GHL_LOCATION_ID',
  PIPELINE_ID: '9ZudIayf0HLmHWoForra',
  INITIAL_STAGE_ID: 'd0a21e32-d181-4c05-ba8a-7d501b1e1bba',
  STAGE_MAP: {
    'New Inquiry': 'd0a21e32-d181-4c05-ba8a-7d501b1e1bba',
    'Contacted': 'e13adfb8-768b-4fa2-8657-8c4c517742e3',
    'Engaged': '37c047ed-f269-4314-aa85-d313e8539557',
    'Site Tour Scheduled': '04bf3db6-c7a1-4b17-85c5-2579442c7c1d',
    'Proposal Sent': '11ffff94-4302-4731-b23f-dc513d9be517',
    'Follow-Up': '6b189b33-9586-4dda-a071-8dd900774f99',
    'Booked Event': '59e4a8f9-95df-4857-a663-dac6535bcfe1',
    'Lost / Not Interested': '44d2cd02-10be-456b-a59c-82136c6c6f5a',
    'Unable to Reach': 'ee8c17cd-a01b-4cea-97f0-a04c233cde29'
  },
  CUSTOM_OPPORTUNITY_FIELDS: {
    'Event Date': 'event_date',
    'Guest Count': 'guest_count',
    'Proposal Amount': 'proposal_amount',
    'Estimated Budget': 'estimated_budget',
    'Follow-Up Date': 'followup_date',
    'Status - EMRG': 'status__emrg',
    'Probability': 'probability',
    'Event Type': 'event_type'
  }
};

const SHEET_NAMES = {
  leads: 'Leads',
  errorLog: 'Error Log'
};

// ─── UTILITIES ─────────────────────────────────────────────────────────────────
String.prototype.camelize = function () {
  return this.toLowerCase().replace(/[^a-z0-9]+(.)/g, (_, c) => c.toUpperCase());
};

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function colIndex(headers, name) {
  const idx = headers.indexOf(name);
  if (idx < 0) throw new Error(`Header "${name}" not found`);
  return idx + 1;
}

function getToken() {
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('GHL API Token not set');
  return token;
}

function getLocation() {
  const loc = PropertiesService.getScriptProperties().getProperty(CONFIG.LOCATION_PROPERTY);
  if (!loc) throw new Error('GHL Location ID not set');
  return loc;
}

function apiFetch(path, method, payload) {
  const options = {
    method: method.toUpperCase(),
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'Authorization': `Bearer ${getToken()}`,
      'Accept': 'application/json',
      'Version': '2021-07-28'
    }
  };
  if (payload) options.payload = JSON.stringify(payload);
  const response = UrlFetchApp.fetch(CONFIG.API_BASE + path, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  let json;
  try {
    json = JSON.parse(body);
  } catch {
    throw new Error(`Invalid JSON: ${body}`);
  }
  if (code < 200 || code >= 300) {
    throw new Error(`GHL API Error (${code}): ${json.message || body}`);
  }
  return json;
}

// ─── MENU ─────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚖️ EMRG Tools')
    .addItem('📊 Show Dashboard', 'showSidebar')
    .addItem('➕ Add New Lead', 'showLeadForm')
    .addItem('🔄 Re-sync All Rows', 'resyncAllRows')
    .addItem('🛠️ Initialize Leads Sheet', 'setupSheet')
    .addToUi();
}

function showLeadForm() {
  const html = HtmlService.createHtmlOutputFromFile('LeadForm')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Lead');
}

function showLeadForm() {
  const html = HtmlService.createHtmlOutput('<p>Lead form goes here.</p>')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Lead');
}

// ─── INSTALLABLE ONEDIT TRIGGER ───────────────────────────────────────────────
function onEditTrigger(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.leads || e.range.getRow() === 1) return;

  const headers = getHeaders(sheet);
  const oppIdCol = colIndex(headers, 'Opportunity ID');
  const oppId = sheet.getRange(e.range.getRow(), oppIdCol).getValue();
  if (!oppId) return;

  const syncCol = colIndex(headers, 'Sync?');
  const syncCell = sheet.getRange(e.range.getRow(), syncCol);
  syncCell.clearContent();
  syncCell.setBackground(null);

  syncRow(sheet, headers, e.range.getRow(), false);
}

// ─── SYNC ─────────────────────────────────────────────────────────────────────
function syncRow(sheet, headers, row, silent = false) {
  const data = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const formData = {};
  headers.forEach((h, i) => formData[h.camelize()] = data[i]);

  const oppId = formData.opportunityId;
  const syncCol = colIndex(headers, 'Sync?');
  const syncCell = sheet.getRange(row, syncCol);
  const reasonCol = colIndex(headers, 'Sync Reason');
  const reasonCell = sheet.getRange(row, reasonCol);

  const resetRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['True'], true).setAllowInvalid(false).build();

  if (!oppId) {
    syncCell.setValue('❌ Failed').clearDataValidations();
    syncCell.setBackground('#f8d7da');
    reasonCell.setValue('Missing Opportunity ID');
    return;
  }

  try {
    updateOpportunity(formData, oppId);
    syncCell.setValue('✅ Success').setDataValidation(resetRule);
    syncCell.setBackground('#d4edda');
    reasonCell.clearContent();
    if (!silent) SpreadsheetApp.getUi().toast(`✅ Row ${row} synced.`);
  } catch (err) {
    logErrorToSheet('syncRow', err);
    syncCell.setValue('❌ Failed').clearDataValidations();
    syncCell.setBackground('#f8d7da');
    reasonCell.setValue(err.message);
    if (!silent) SpreadsheetApp.getUi().toast(`❌ Sync failed: ${err.message}`, 'Error');
  }
}

function resyncAllRows() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.leads);
  const headers = getHeaders(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  let success = 0, failed = 0;

  data.forEach((_, i) => {
    const result = syncRow(sheet, headers, i + 2, true);
    if (result === 'ok') success++;
    else failed++;
  });

  SpreadsheetApp.getUi().alert(`🔄 Resync completed:\n✅ ${success} successful\n❌ ${failed} failed`);
}

// ─── GHL OPPORTUNITY ──────────────────────────────────────────────────────────
function updateOpportunity(formData, oppId) {
  const payload = {};
  if (formData.stage) {
    const stageId = CONFIG.STAGE_MAP[formData.stage];
    if (!stageId) throw new Error(`Unknown stage: ${formData.stage}`);
    payload.pipelineStageId = stageId;
  }
  if (formData.status) payload.status = formData.status;
  if (formData.proposalAmount) payload.monetaryValue = Number(formData.proposalAmount);

  const customFields = [];
  Object.entries(CONFIG.CUSTOM_OPPORTUNITY_FIELDS).forEach(([label, id]) => {
    const val = formData[label.camelize()];
    if (val !== undefined && val !== '') {
      customFields.push({ id, value: val });
    }
  });
  if (customFields.length) payload.customFields = customFields;
  if (!Object.keys(payload).length) throw new Error('No fields to update');

  apiFetch(`/opportunities/${oppId}`, 'put', payload);
}

// ─── SHEET SETUP ──────────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  const leads = ss.getSheetByName(SHEET_NAMES.leads) || ss.insertSheet(SHEET_NAMES.leads);
  leads.clear();

  const headers = [
    'Opportunity ID', 'Date Received', 'Stage', 'Status', 'Opportunity Name', 'Opportunity Source',
    'Event Date', 'Guest Count', 'Proposal Amount', 'Estimated Budget',
    'Follow-Up Date', 'Probability', 'Event Type',
    'First Name', 'Last Name', 'Email', 'Phone',
    'Sync?', 'Sync Reason'
  ];
  leads.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  leads.setFrozenRows(1);
  SpreadsheetApp.flush();
  leads.autoResizeColumns(1, headers.length);

  const dropdowns = {
    'Stage': Object.keys(CONFIG.STAGE_MAP),
    'Status': ['open', 'won', 'abandoned', 'lost'],
    'Opportunity Source': ['PartySlate', 'Website', 'Instagram', 'Referral', 'Other'],
    'Probability': ['Hot', 'Warm', 'Cold'],
    'Event Type': ['Corporate Event', 'Social Event', 'Mitzvah', 'Fundraiser', 'Other'],
    'Sync?': ['True']
  };

  Object.entries(dropdowns).forEach(([col, list]) => {
    const colNum = headers.indexOf(col) + 1;
    if (colNum > 0) {
      leads.getRange(2, colNum, leads.getMaxRows() - 1)
        .setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInList(list).setAllowInvalid(false).build());
    }
  });

  leads.getRange(2, 1, leads.getMaxRows() - 1).protect().setDescription('Auto-generated');
  leads.hideColumn(leads.getRange(1, 1));

  const log = ss.getSheetByName(SHEET_NAMES.errorLog) || ss.insertSheet(SHEET_NAMES.errorLog);
  log.clear();
  log.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Function', 'Details']]);
}

// ─── ERROR LOGGING ────────────────────────────────────────────────────────────
function logErrorToSheet(func, err) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.errorLog);
  if (!sheet) return;
  sheet.appendRow([new Date(), func, err.stack || err.message || err.toString()]);
}

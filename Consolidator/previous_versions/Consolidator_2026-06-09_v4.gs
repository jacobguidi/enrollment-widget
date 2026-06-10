// ============================================================
// CHANGE LOG
// ============================================================
// 2026-06-09 v4 — Add onOpen() menu: Run Consolidation Now, Generate Recuro File, Setup Triggers.
// 2026-06-09 v3 — Split Master Eligibility into its own private workbook (ELIGIBILITY_WORKBOOK_ID)
//                 so PII is never in the same sheet as the published Clients tab.
//                 Clients tab still read from MASTER_WORKBOOK_ID. CSV saved to eligibility
//                 workbook's Drive folder.
// 2026-06-09 v2 — Fix Clients tab header mismatch: 'Workbook ID' → 'PPS Workbook ID'
//                 (actual header in Master Workbook). Falls back to 'Workbook ID' for
//                 compatibility. Without this fix every client was silently skipped.
// 2026-06-09 v1 — Initial extraction from Consolidator.rtf into version control.
//                 Fixed RTF artifacts: \n literals in _csvEscape, CSV join, and
//                 summary/skipped string joins. No logic changes.
// ============================================================

// ============================================================
// RECURO PPS CONSOLIDATOR
// Standalone Apps Script — Master Runner
// Reads PPS workbooks (headers row 6, data row 7+)
// Writes to Master Eligibility sheet (headers row 1)
// ============================================================

// ============================================================
// CONFIGURATION — set these once
// ============================================================

// Contains only the Clients registry tab — safe to share/publish
const MASTER_WORKBOOK_ID = '1jqzmXQPbI0jlIvVN7dMgVL6YNHz8oxTEwboh8ROVbtU';

// Private workbook — contains Master Eligibility tab with employee PII
// Create a new blank Google Sheet, paste its ID here, share with NO ONE
const ELIGIBILITY_WORKBOOK_ID = 'PASTE_ELIGIBILITY_SHEET_ID_HERE';

const NOTIFY_EMAIL = 'jacob@thrivebg.com';

const PPS_HEADER_ROW = 6;
const PPS_DATA_START = 7;

// ============================================================
// PPS HEADER → RECURO FIELD MAPPING
// Keys are normalized (lowercase, stripped) PPS header names
// Values are Recuro column names
// ============================================================

const PPS_TO_RECURO = {
  'lastname':        'LastName',
  'firstname':       'FirstName',
  'gender':          'Gender',
  'dateofbirth':     'DateOfBirth',
  'dob':             'DateOfBirth',
  'address':         'AddressLine1',
  'addressline1':    'AddressLine1',
  'streetaddress':   'AddressLine1',
  'mailingaddress1': 'AddressLine1',
  'city':            'City',
  'state':           'State',
  'zip':             'ZipCode',
  'zipcode':         'ZipCode',
  'postalcode':      'ZipCode',
  'phone':           'MobilePhone',
  'phonenumber':     'MobilePhone',
  'mobilephone':     'MobilePhone',
  'emailaddress':    'EmailAddress',
  'email':           'EmailAddress',
  'effectivedate':   'EffectiveStart',
  'effectivestart':  'EffectiveStart',
  'terminationdate': 'EffectiveEnd',
  'effectiveend':    'EffectiveEnd',
};

// SSN Last 4 — checked in priority order
const SSN_LAST4_CANDIDATES = [
  'Employee SSNLast 4 digits',
  'Employee SSN Last 4 digits',
  'Employee SSN Last 4 Digits',
  'SSN Last 4 Digits',
  'SSN Last 4',
];

// Recuro output columns — fixed order, matches spec exactly
const RECURO_HEADERS = [
  'LastName', 'FirstName', 'Gender', 'DateOfBirth',
  'AddressLine1', 'AddressLine2', 'City', 'State', 'ZipCode', 'CountryCode',
  'MobilePhone', 'EmailAddress',
  'EffectiveStart', 'EffectiveEnd',
  'MemberType', 'ClientMemberID', 'SecondaryClientMemberID', 'ClientPrimaryMemberID',
  'ServiceOffering', 'GroupID', 'GroupName',
  'MetaTag1', 'MetaTag2', 'MetaTag3', 'MetaTag4', 'MetaTag5',
  // TBG operational columns (appended after Recuro spec)
  '_TBG_ClientID', '_TBG_SourceWorkbookID', '_TBG_LastSync', '_TBG_Status',
];

// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TBG Consolidator')
    .addItem('Run Consolidation Now',  'runMasterConsolidation')
    .addItem('Generate Recuro File',   'generateRecuroFile')
    .addSeparator()
    .addItem('Setup Triggers',         'createTriggers')
    .addToUi();
}

// ============================================================
// UTILITIES
// ============================================================

function buildColumnMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    if (h !== null && h !== undefined) map[String(h).trim()] = i;
  });
  return map;
}

function _normalizeHeader(h) {
  return String(h || '')
    .replace(/\s+/g, ' ')
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .replace(/ /g, '')
    .toLowerCase()
    .trim();
}

function _formatDate(val) {
  if (!val) return '';
  const d = val instanceof Date ? val : new Date(val);
  if (isNaN(d.getTime())) return String(val);
  const m  = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return m + '/' + dd + '/' + d.getFullYear();
}

function _formatPhone(val) {
  if (!val) return '';
  const digits = String(val).replace(/[^0-9]/g, '');
  if (digits.length === 10)
    return digits.slice(0,3) + '-' + digits.slice(3,6) + '-' + digits.slice(6);
  if (digits.length === 11 && digits[0] === '1')
    return digits.slice(1,4) + '-' + digits.slice(4,7) + '-' + digits.slice(7);
  return String(val);
}

function _isOlderThanDays(dateVal, days) {
  if (!dateVal) return false;
  const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return false;
  return (new Date() - d) > days * 86400000;
}

function _csvEscape(v) {
  const s = String(v == null ? '' : v);
  return (s.includes(',') || s.includes('"') || s.includes('\n'))
    ? '"' + s.replace(/"/g, '""') + '"'
    : s;
}

// ============================================================
// CORE: PULL ONE PPS WORKBOOK → ARRAY OF NORMALIZED ROWS
// ============================================================

function _pullFromPPS(workbookId, clientMeta, ts) {
  const ss    = SpreadsheetApp.openById(workbookId);
  const wsPPS = ss.getSheetByName('PPS');
  if (!wsPPS) {
    Logger.log('No PPS tab in workbook: ' + workbookId);
    return [];
  }

  const lastRow = wsPPS.getLastRow();
  if (lastRow < PPS_DATA_START) {
    Logger.log('No data rows in PPS: ' + workbookId);
    return [];
  }

  const lastCol    = wsPPS.getLastColumn();
  const rawHeaders = wsPPS.getRange(PPS_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const data       = wsPPS.getRange(PPS_DATA_START, 1, lastRow - PPS_DATA_START + 1, lastCol).getValues();

  const exactColMap = buildColumnMap(rawHeaders);
  const normColMap  = {};
  rawHeaders.forEach((h, i) => {
    const norm = _normalizeHeader(h);
    if (norm) normColMap[norm] = i;
  });

  let ssnCol = -1;
  for (const candidate of SSN_LAST4_CANDIDATES) {
    if (exactColMap[candidate] !== undefined) {
      ssnCol = exactColMap[candidate];
      break;
    }
  }
  if (ssnCol === -1) {
    for (const key of Object.keys(normColMap)) {
      if (key.includes('ssnlast4') || key.includes('ssn4')) {
        ssnCol = normColMap[key];
        break;
      }
    }
  }

  const _get = (row, recuroField) => {
    for (const [normKey, rField] of Object.entries(PPS_TO_RECURO)) {
      if (rField === recuroField && normColMap[normKey] !== undefined) {
        return row[normColMap[normKey]];
      }
    }
    return '';
  };

  // Stop reading columns at DeductionFrequency or PayFrequency to avoid trailing junk
  const stopCol = (() => {
    for (const key of Object.keys(normColMap)) {
      if (key === 'deductionfrequency') return normColMap[key];
    }
    for (const key of Object.keys(normColMap)) {
      if (key === 'payfrequency') return normColMap[key];
    }
    return lastCol;
  })();

  const rows = [];

  for (const row of data) {
    const rowSlice = row.slice(0, stopCol + 1);
    if (rowSlice.every(v => v === '' || v === null || v === undefined)) continue;

    const fn = String(_get(row, 'FirstName') || '').trim();
    const ln = String(_get(row, 'LastName')  || '').trim();
    if (!fn && !ln) continue;

    const termDateRaw = _get(row, 'EffectiveEnd');
    const effDateRaw  = _get(row, 'EffectiveStart');
    const termDate = termDateRaw ? (termDateRaw instanceof Date ? termDateRaw : new Date(termDateRaw)) : null;
    const effDate  = effDateRaw  ? (effDateRaw  instanceof Date ? effDateRaw  : new Date(effDateRaw))  : null;

    if (!effDate || isNaN(effDate.getTime())) continue;

    const isTerminated = termDate && !isNaN(termDate.getTime());
    if (isTerminated && _isOlderThanDays(termDate, 60)) continue;

    const status      = isTerminated ? 'Terminated' : 'Active';
    const ssnLast4    = ssnCol !== -1 ? String(row[ssnCol] || '').replace(/[^0-9]/g, '').slice(-4) : '';
    const clientMemberID = ssnLast4 ? (clientMeta.clientId + ssnLast4).slice(0, 15) : '';

    rows.push([
      ln,
      fn,
      String(_get(row, 'Gender')       || '').trim(),
      _formatDate(_get(row, 'DateOfBirth')),
      String(_get(row, 'AddressLine1') || '').trim(),
      '',
      String(_get(row, 'City')         || '').trim(),
      String(_get(row, 'State')        || '').trim(),
      String(_get(row, 'ZipCode')      || '').trim(),
      'US',
      _formatPhone(_get(row, 'MobilePhone')),
      String(_get(row, 'EmailAddress') || '').trim(),
      _formatDate(effDate),
      isTerminated ? _formatDate(termDate) : '',
      'Primary',
      clientMemberID,
      '',
      clientMemberID,
      clientMeta.serviceOffering || '',
      clientMeta.groupId         || '',
      clientMeta.groupName       || '',
      '', '', '', '', '',
      clientMeta.clientId,
      clientMeta.workbookId,
      ts,
      status,
    ]);
  }

  return rows;
}

// ============================================================
// MAIN: CONSOLIDATION — reads all client PPS, writes master
// ============================================================

function runMasterConsolidation() {
  const masterSS  = SpreadsheetApp.openById(MASTER_WORKBOOK_ID);
  const eligSS    = SpreadsheetApp.openById(ELIGIBILITY_WORKBOOK_ID);

  const clientSheet = masterSS.getSheetByName('Clients');
  const eligSheet   = eligSS.getSheetByName('Master Eligibility')
                   || eligSS.insertSheet('Master Eligibility');

  if (!clientSheet) throw new Error('No "Clients" tab in Master Workbook');

  const existingHeader = eligSheet.getRange(1, 1, 1, 1).getValue();
  if (!existingHeader || existingHeader !== 'LastName') {
    eligSheet.getRange(1, 1, 1, RECURO_HEADERS.length).setValues([RECURO_HEADERS]);
  }

  const clientData    = clientSheet.getDataRange().getValues();
  const clientHeaders = clientData[0];
  const cm            = buildColumnMap(clientHeaders);

  const ts      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const allRows = [];
  const log     = [];

  for (let i = 1; i < clientData.length; i++) {
    const client = clientData[i];
    const status = String(client[cm['Status']] || '').toLowerCase().trim();
    if (status !== 'active') continue;

    const clientMeta = {
      clientId:        String(client[cm['Plan Builder ID']]  || '').trim(),
      // 'PPS Workbook ID' is the actual header; fall back to 'Workbook ID' for older sheets
      workbookId:      String(client[cm['PPS Workbook ID']] !== undefined
                         ? client[cm['PPS Workbook ID']]
                         : client[cm['Workbook ID']] || '').trim(),
      groupId:         String(client[cm['Group ID']]         || '').trim(),
      groupName:       String(client[cm['Company Name']]     || '').trim(),
      serviceOffering: String(client[cm['Service Offering']] || '').trim(),
    };

    if (!clientMeta.workbookId) {
      log.push('SKIP no Workbook ID: ' + clientMeta.clientId);
      continue;
    }

    try {
      const rows = _pullFromPPS(clientMeta.workbookId, clientMeta, ts);
      allRows.push(...rows);
      log.push('OK ' + clientMeta.clientId + ': ' + rows.length + ' rows');
    } catch (e) {
      log.push('ERROR ' + clientMeta.clientId + ': ' + e.message);
    }
  }

  const existingLastRow = eligSheet.getLastRow();
  if (existingLastRow > 1) {
    eligSheet.getRange(2, 1, existingLastRow - 1, RECURO_HEADERS.length).clearContent();
  }
  if (allRows.length > 0) {
    eligSheet.getRange(2, 1, allRows.length, RECURO_HEADERS.length).setValues(allRows);
  }

  Logger.log('=== Consolidation Complete ===');
  Logger.log('Total rows written: ' + allRows.length);
  log.forEach(l => Logger.log(l));
}

// ============================================================
// RECURO FILE GENERATOR
// Reads Master Eligibility → outputs CSV to Drive + email
// ============================================================

function generateRecuroFile() {
  const eligSS    = SpreadsheetApp.openById(ELIGIBILITY_WORKBOOK_ID);
  const eligSheet = eligSS.getSheetByName('Master Eligibility');
  const data      = eligSheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('Master Eligibility is empty — run consolidation first.');
    return;
  }

  const headers       = data[0];
  const cm            = buildColumnMap(headers);
  const recuroSpecCols = RECURO_HEADERS.slice(0, 26);
  const csvRows       = [recuroSpecCols];
  const skipped       = [];

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const status = String(row[cm['_TBG_Status']] || '').toLowerCase();

    if (status === 'pending') {
      skipped.push('PENDING: ' + row[cm['FirstName']] + ' ' + row[cm['LastName']]);
      continue;
    }

    const missing = [];
    if (!row[cm['LastName']])       missing.push('LastName');
    if (!row[cm['FirstName']])      missing.push('FirstName');
    if (!row[cm['DateOfBirth']])    missing.push('DateOfBirth');
    if (!row[cm['EmailAddress']])   missing.push('EmailAddress');
    if (!row[cm['EffectiveStart']]) missing.push('EffectiveStart');
    if (!row[cm['MemberType']])     missing.push('MemberType');
    if (!row[cm['ClientMemberID']]) missing.push('ClientMemberID');
    if (!row[cm['ServiceOffering']]) missing.push('ServiceOffering');
    if (!row[cm['GroupID']])        missing.push('GroupID');
    if (!row[cm['GroupName']])      missing.push('GroupName');

    if (missing.length > 0) {
      skipped.push('MISSING [' + missing.join(', ') + ']: ' + row[cm['FirstName']] + ' ' + row[cm['LastName']]);
      continue;
    }

    csvRows.push(recuroSpecCols.map(col => cm[col] !== undefined ? row[cm[col]] : ''));
  }

  const csv      = csvRows.map(r => r.map(_csvEscape).join(',')).join('\n');
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fileName = 'Recuro_Eligibility_' + today + '.csv';

  // Save CSV alongside the private eligibility workbook, not the published master
  const masterFile = DriveApp.getFileById(ELIGIBILITY_WORKBOOK_ID);
  const folder     = masterFile.getParents().next();
  const existing   = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  const savedFile  = folder.createFile(fileName, csv, MimeType.CSV);

  const activeCount = csvRows.length - 1;
  const termCount   = csvRows.filter((r, i) => i > 0 && r[13] !== '').length;

  const summary = [
    'Recuro Eligibility File — ' + today,
    '',
    'Total rows: '  + activeCount,
    'Active: '      + (activeCount - termCount),
    'Terminations: '+ termCount,
    'Skipped: '     + skipped.length,
    '',
    skipped.length > 0 ? 'Skipped:\n' + skipped.join('\n') : 'No records skipped.',
    '',
    'Drive copy: ' + savedFile.getUrl(),
    '',
    'CSV attached — ready to upload to Recuro.',
  ].join('\n');

  GmailApp.sendEmail(
    NOTIFY_EMAIL,
    'Recuro Eligibility File — ' + today,
    summary,
    { attachments: [Utilities.newBlob(csv, MimeType.CSV, fileName)] }
  );

  Logger.log(summary);
}

// ============================================================
// TRIGGERS — run createTriggers() once manually to schedule
// ============================================================

function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['runMasterConsolidation', 'generateRecuroFile'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runMasterConsolidation')
    .timeBased().everyDays(1).atHour(2).create();

  ScriptApp.newTrigger('generateRecuroFile')
    .timeBased().onMonthDay(5).atHour(6).create();

  Logger.log('Triggers created: nightly consolidation + monthly Recuro file.');
}

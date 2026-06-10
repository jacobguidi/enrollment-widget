// ============================================================
// TBG TEMPLATE WORKBOOK — CENTRAL PRODUCTION ENGINE / LIBRARY
// PATCHED §1 → §4  (replace everything from the top of the file
// down to but NOT including "// §5  STEP 3 — SYNC OUT")
// ============================================================
//
// What this patch fixes:
//   1. registerWithMaster()  →  TDZ bug:  const ss = ss || ...
//   2. registerWithMaster()  →  tolerates either "Workbook ID" or
//                                "PPS Workbook ID" as the Master header
//   3. loadCurrentConfig()   →  leaveInPaycheck no longer writes "0"
//                                into a blank input on first load
//   4. RE-ADDED loadSTDOptionsForWizard()  (was dropped — caused
//                                "MasterEngine.loadSTDOptionsForWizard
//                                is not a function")
//   5. RE-ADDED loadPPSColumnsForWizard()  (also dropped)
//
// Intentionally kept from Gemini's pass:
//   • runBuildPlans() builds for every row with a DOB + Allotments > 0
//     (no Relation === 'EE' guard) — confirmed desired behavior.
//
// IMPORTANT — also do this manually:
//   Scroll to §10 and DELETE the second _buildWizardHTML() function
//   that lives just above "MISSING CONSTANTS FOR NAME CASING & SYNC
//   ENGINE". It's the older 4-tab wizard, and because it's declared
//   later in the file it overrides the new 5-tab wizard below.
// ============================================================


// ============================================================
// §1  CONSTANTS + MENU
// ============================================================
const CONFIG_SHEET_NAME  = '__config';
const GHL_BASE           = 'https://services.leadconnectorhq.com';
const GHL_API_VERSION    = '2021-07-28';
const MASTER_WORKBOOK_ID = '1MWenQPb4CyASD-SOJIkmQIHNvgkhTCBrS9wjYcLLlUM'; // ← paste once

// ============================================================
// PPS ENGINE MAPPING TABLES (required by §11)
// ============================================================
const SSN_LAST4_CANDIDATES = [
  'Employee SSNLast 4 digits', 'Employee SSN Last 4 digits',
  'Employee SSN Last 4 Digits', 'SSN Last 4 Digits', 'SSN Last 4'
];
const PPS_DEMO_FIELD_MAP = [
  { src: 'First Name',       dst: 'First Name'   },
  { src: 'Last Name',        dst: 'Last Name'    },
  { src: 'Gender',           dst: 'Gender'       },
  { src: 'DOB',              dst: 'Date of Birth'},
  { src: 'Mailing Address1', dst: 'Address'      },
  { src: 'City',             dst: 'City'         },
  { src: 'State',            dst: 'State'        },
  { src: 'Zip',              dst: 'Zip'          },
  { src: 'Phone Number',     dst: 'Phone'        },
  { src: 'Email Address',    dst: 'Email'        }
];
const PPS_PREMIUM_FIELD_MAP = [
  { src: 'AE Premium Amount',   label: 'Accident',           dsts: ['Accident After Tax']                                       },
  { src: 'CI Premium Amount',   label: 'Critical Illness',   dsts: ['Critical Illness After Tax']                               },
  { src: 'HI Premium Amount',   label: 'Hospital Indemnity', dsts: ['Hospital Indemnity After Tax']                             },
  { src: 'Life Premium Amount', label: 'Whole Life',         dsts: ['Whole Life After Tax']                                     },
  { src: 'DI Premium Amount',   label: 'Disability / STD',   dsts: ['Short Term Disability After Tax', 'Disability After Tax'] }
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TBG Plan Builder')
    .addItem('1. Setup & Import',          'showSetupWizard')
    .addItem('2. Build Plans',             'runBuildPlans')
    .addItem('3. Sync Out',                'runSyncOut')
    .addSeparator()
    .addSubMenu(ui.createMenu('Tools')
      .addItem('Edit Setup',                    'showSetupWizard')
      .addItem('Stack All Plans',               'stackAllPlans')
      .addItem('Sync Demographics',             'syncDemographics')
      .addItem('Sync to PPS (sidebar)',         'showPPSSidebar')
      .addItem('Push to GHL only',              'pushDemographicToGHL')
      .addSeparator()
      .addItem('Demo Import (legacy)',          'showDemoSidebar')
      .addItem('Auto Plan Builder (legacy)',    'runAutoPlanBuilder')
      .addItem('Build Benefit Package (legacy)','runBenefitPackageBuilder')
      .addSeparator()
      .addItem('Test GHL Connection',           'testGHLConnection')
      .addItem('Test Widget Sync',              'testSyncCustomizedEmployee')
    )
    .addToUi();
}

// ============================================================
// §2  __config TAB HELPERS
// ============================================================
function _getConfigSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET_NAME);
    sh.hideSheet();
    sh.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    sh.setColumnWidth(1, 220);
    sh.setColumnWidth(2, 600);
  }
  return sh;
}

function getConfig(ss) {
  const sh   = _getConfigSheet(ss);
  const last = sh.getLastRow();
  if (last < 2) return {};
  const rows = sh.getRange(2, 1, last - 1, 2).getValues();
  const out  = {};
  for (const [k, v] of rows) {
    if (!k) continue;
    const key = String(k).trim();
    if (typeof v === 'string' && v.startsWith('{')) {
      try { out[key] = JSON.parse(v); continue; } catch (e) {}
    }
    out[key] = v;
  }
  return out;
}

function saveConfig(obj, ss) {
  const sh   = _getConfigSheet(ss);
  const last = sh.getLastRow();
  const existing = last < 2 ? {} : (() => {
    const m = {};
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(([k, v]) => {
      if (k) m[String(k).trim()] = v;
    });
    return m;
  })();
  Object.keys(obj).forEach(k => {
    let v = obj[k];
    if (v !== null && typeof v === 'object') v = JSON.stringify(v);
    existing[k] = v;
  });
  const merged = Object.keys(existing).sort().map(k => [k, existing[k]]);
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  if (merged.length) sh.getRange(2, 1, merged.length, 2).setValues(merged);
  sh.getRange(merged.length + 2, 1, 1, 2)
    .setValues([['Setup Last Saved', new Date().toISOString()]]);
}

function configValue(key, fallback, ss) {
  const c = getConfig(ss);
  return c[key] !== undefined ? c[key] : (fallback === undefined ? null : fallback);
}

// ============================================================
// §3  SETUP & IMPORT WIZARD — public entry points
// ============================================================
function showSetupWizard() {
  const html = HtmlService.createHtmlOutput(_buildWizardHTML())
    .setTitle('Setup & Import')
    .setWidth(460);
  SpreadsheetApp.getUi().showSidebar(html);
}

function loadCurrentConfig() {
  const cfg = getConfig();
  return {
    company:          cfg['Company Name']        || '',
    groupId:          cfg['Group ID']            || '',
    enrollmentCutoff: cfg['Enrollment Cutoff']   || '',
    effectiveDate:    cfg['Effective Date']       || '',
    signDate:         cfg['Sign Date']            || '',
    rateSheetId:      cfg['Rate Tables Sheet ID'] || '',
    stdLabel:         cfg['STD Benefit Period']   || '',
    payFreq:          cfg['Pay Frequency']        || 26,
    ppsWorkbookId:    cfg['PPS Workbook ID']      || '',
    ppsMapping:       cfg['PPS Mapping']          || null,
    // PATCH: use != null so a saved 0 still loads, but an unsaved
    // value returns '' instead of writing "0" into a blank field.
    leaveInPaycheck:  cfg['Leave in Paycheck'] != null ? cfg['Leave in Paycheck'] : '',
  };
}

// PATCH: re-added — was dropped, caused
// "MasterEngine.loadSTDOptionsForWizard is not a function"
function loadSTDOptionsForWizard(rateSheetId) {
  if (!rateSheetId) return { error: 'No sheet ID' };
  let rateSS;
  try { rateSS = SpreadsheetApp.openById(rateSheetId.trim()); }
  catch (e) { return { error: 'Cannot open rate sheet (check ID + permissions).' }; }
  const std = rateSS.getSheetByName('Short Term Disability');
  if (!std) return { error: 'Rate sheet has no "Short Term Disability" tab.' };
  const opts = detectSTDOptions(std).map(o => o.label);
  return { options: opts };
}

// PATCH: re-added — was dropped along with loadSTDOptionsForWizard
function loadPPSColumnsForWizard() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wsP = ss.getSheetByName('PPS');
  if (!wsP) return { error: 'PPS tab not found.' };
  const cols = wsP.getRange(6, 1, 1, wsP.getLastColumn()).getValues()[0]
    .map(c => String(c).trim()).filter(c => c !== '');
  let detected = null;
  try { detected = getPPSMappingData(); } catch (e) {}
  return { columns: cols, detected: detected };
}

function saveSetupAndImport(formJSON) {
  const form = JSON.parse(formJSON);
  const cfg = {
    'Company Name':         form.company,
    'Group ID':             form.groupId          || '',
    'Enrollment Cutoff':    form.enrollmentCutoff,
    'Effective Date':       form.effectiveDate,
    'Sign Date':            form.signDate,
    'Rate Tables Sheet ID': form.rateSheetId,
    'STD Benefit Period':   form.stdLabel,
    'Pay Frequency':        Number(form.payFreq)  || 26,
    'PPS Workbook ID':      form.ppsWorkbookId    || '',
    'PPS Mapping':          form.ppsMapping       || null,
    'Leave in Paycheck':    Number(form.leaveInPaycheck) || 0,
  };
  saveConfig(cfg);

  let regMsg = '';
  try {
    registerWithMaster();
    regMsg = '<br>✓ Registered with Master Workbook.';
  } catch (e) {
    regMsg = '<br>⚠ Master registration skipped: ' + e.message;
  }

  let importMsg = 'Import skipped — no savings/salary data pasted.';
  if ((form.savingsRaw && form.savingsRaw.trim()) ||
      (form.salaryRaw  && form.salaryRaw.trim())) {
    try {
      importMsg = executeDemoImport(form.savingsRaw || '', form.salaryRaw || '');
    } catch (err) {
      importMsg = '<span style="color:#c5221f">Import error: ' + err.message + '</span>';
    }
  }
  return '<b>Configuration saved.</b>' + regMsg + '<br>' + importMsg;
}

// ============================================================
// §3a  MASTER WORKBOOK REGISTRATION  (PATCHED)
// ============================================================
function registerWithMaster() {
  // PATCH: was  const ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  // which threw a ReferenceError every call (TDZ on `ss`).
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig(ss);
  const ppsId = (cfg['PPS Workbook ID'] || '').trim() || ss.getId();

  const masterSS    = SpreadsheetApp.openById(MASTER_WORKBOOK_ID);
  const clientSheet = masterSS.getSheetByName('Clients');
  if (!clientSheet) throw new Error('No "Clients" tab found in Master Workbook.');

  const data          = clientSheet.getDataRange().getValues();
  const clientHeaders = data[0];
  const cm            = buildColumnMap(clientHeaders);
  const row           = new Array(clientHeaders.length).fill('');

  // Helper: write the first header name that actually exists in the Master.
  // This is the fix for the "Workbook ID" vs "PPS Workbook ID" mismatch.
  const setIfHeader = (candidates, value) => {
    for (const name of candidates) {
      if (cm[name] !== undefined) { row[cm[name]] = value; return; }
    }
  };

  setIfHeader(['Company Name'],                            cfg['Company Name'] || '');
  setIfHeader(['Plan Builder ID'],                         ss.getId());
  setIfHeader(['PPS Workbook ID', 'Workbook ID'],          ppsId);
  setIfHeader(['Rate Table ID', 'Rate Tables Sheet ID'],   cfg['Rate Tables Sheet ID'] || '');
  setIfHeader(['Group ID'],                                cfg['Group ID'] || '');
  setIfHeader(['Service Offering'],                        '');
  setIfHeader(['Enrollment Cutoff'],                       cfg['Enrollment Cutoff'] || '');
  setIfHeader(['Effective Date'],                          cfg['Effective Date'] || '');
  setIfHeader(['Pay Frequency'],                           cfg['Pay Frequency'] || '');
  setIfHeader(['Status'],                                  'Active');

  const planBuilderCol = cm['Plan Builder ID'];
  if (planBuilderCol !== undefined) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][planBuilderCol] || '').trim() === ss.getId()) {
        clientSheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        Logger.log('Master registration updated: ' + cfg['Company Name']);
        return;
      }
    }
  }

  clientSheet.appendRow(row);
  Logger.log('Master registration added: ' + cfg['Company Name']);
}

// ============================================================
// §3b  WIZARD HTML  (5-tab version — group type radio + cushion)
// ============================================================
function _buildWizardHTML() {
  return `<!DOCTYPE html>
<html><head><style>
  body        { font-family:'Google Sans',Roboto,Arial,sans-serif; padding:14px; font-size:13px; color:#202124; }
  h3          { margin:0 0 4px 0; color:#1a73e8; font-size:15px; }
  .sub        { color:#5f6368; font-size:11px; margin-bottom:14px; }
  .steps      { display:flex; gap:4px; margin-bottom:16px; }
  .step       { flex:1; height:4px; background:#e0e0e0; border-radius:2px; }
  .step.active,.step.done { background:#1a73e8; }
  .tab        { display:none; }
  .tab.active { display:block; animation:fadeIn .2s; }
  @keyframes fadeIn { from{opacity:0} to{opacity:1} }
  .field      { margin-bottom:12px; }
  .lbl        { font-size:11px; font-weight:600; color:#3c4043; margin-bottom:4px; display:block; text-transform:uppercase; letter-spacing:.04em; }
  .hint       { font-size:10.5px; color:#5f6368; margin-top:3px; line-height:1.4; }
  input,select,textarea { width:100%; box-sizing:border-box; padding:7px 9px; border:1px solid #dadce0; border-radius:4px; font-size:12.5px; font-family:inherit; color:#202124; }
  textarea { font-family:monospace; font-size:10.5px; min-height:110px; resize:vertical; }
  input:focus,select:focus,textarea:focus { outline:none; border-color:#1a73e8; box-shadow:0 0 0 2px rgba(26,115,232,.12); }
  .row2       { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
  .btn-row    { display:flex; gap:8px; margin-top:14px; }
  .btn        { flex:1; padding:9px 12px; border-radius:4px; border:none; font-size:13px; font-weight:600; cursor:pointer; font-family:inherit; }
  .btn-primary              { background:#1a73e8; color:#fff; }
  .btn-primary:hover        { background:#1557b0; }
  .btn-primary:disabled     { background:#dadce0; color:#5f6368; cursor:not-allowed; }
  .btn-ghost                { background:#fff; color:#1a73e8; border:1px solid #dadce0; }
  .btn-ghost:hover          { background:#f1f3f4; }
  .map-row    { display:grid; grid-template-columns:1fr 1fr; gap:6px; align-items:center; padding:5px 0; border-bottom:1px solid #f1f3f4; }
  .map-row:last-child       { border-bottom:none; }
  .map-lbl    { font-size:11px; color:#3c4043; font-weight:500; }
  .map-row select           { font-size:11px; padding:5px 7px; }
  #status     { margin-top:12px; padding:10px 12px; border-radius:4px; display:none; font-size:12px; line-height:1.5; }
  .ok-bg      { background:#e6f4ea; color:#137333; }
  .err-bg     { background:#fce8e6; color:#c5221f; }
  .info-card  { background:#f1f3f4; padding:8px 10px; border-radius:4px; font-size:11px; line-height:1.5; margin-bottom:12px; }
  .radio-group        { display:flex; flex-direction:column; gap:8px; margin-top:6px; }
  .radio-group label  { display:flex; align-items:center; gap:8px; font-size:12.5px; cursor:pointer; font-weight:500; }
  .pps-card   { background:#e8f0fe; border:1px solid #c5d6f5; border-radius:4px; padding:10px 12px; font-size:11.5px; line-height:1.6; margin-bottom:10px; }
  .pps-card b { color:#1a73e8; }
  .pps-card ol{ margin:6px 0 0 16px; padding:0; }
</style></head><body>

  <h3>Setup &amp; Import</h3>
  <p class="sub">Lock template constraints once — Steps 2 and 3 read these automatically.</p>

  <div class="steps">
    <div class="step active" id="s1"></div>
    <div class="step" id="s2"></div>
    <div class="step" id="s3"></div>
    <div class="step" id="s4"></div>
    <div class="step" id="s5"></div>
  </div>

  <div class="tab active" id="tab1">
    <div class="field"><label class="lbl">Company Name</label>
      <input id="company" type="text" placeholder="City of Refuge" /></div>
    <div class="field">
      <label class="lbl">Group ID <span style="font-weight:400; text-transform:none; letter-spacing:0">— assigned by Recuro</span></label>
      <input id="groupId" type="text" placeholder="Leave blank until Recuro provides" />
      <p class="hint">Recuro assigns this when they onboard the group. Save now and re-run Setup to fill in later — it will update, not duplicate.</p>
    </div>
    <div class="field"><label class="lbl">Enrollment Cutoff</label>
      <input id="enrollmentCutoff" type="text" placeholder="MM/DD/YYYY" /></div>
    <div class="row2">
      <div class="field"><label class="lbl">Effective Date</label>
        <input id="effectiveDate" type="text" placeholder="MM/DD/YYYY" /></div>
      <div class="field"><label class="lbl">Sign Date</label>
        <input id="signDate" type="text" placeholder="MM/DD/YYYY" /></div>
    </div>
    <div class="btn-row"><button class="btn btn-primary" onclick="next(1)">Next →</button></div>
  </div>

  <div class="tab" id="tab2">
    <div class="field"><label class="lbl">Rate Tables Sheet ID</label>
      <input id="rateSheetId" type="text" placeholder="1ABCdef…" onblur="onRateIdChange()" /></div>
    <div class="field"><label class="lbl">STD Benefit Period</label>
      <select id="stdLabel"><option value="">— pick rate sheet first —</option></select>
      <p class="hint" id="stdHint">Auto-loads from the rate sheet's Short Term Disability tab.</p>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost"   onclick="prev(2)">← Back</button>
      <button class="btn btn-primary" onclick="next(2)">Next →</button>
    </div>
  </div>

  <div class="tab" id="tab3">
    <div class="field">
      <label class="lbl">Group Onboarding Classification</label>
      <div class="radio-group">
        <label><input type="radio" name="groupType" value="existing" onchange="onGroupTypeChange('existing')" /> Existing Group with an Active PPS Tracker</label>
        <label><input type="radio" name="groupType" value="new" onchange="onGroupTypeChange('new')" /> New Group without a PPS Template Workbook</label>
      </div>
    </div>

    <div id="newGroupCard" style="display:none">
      <div class="pps-card">
        <b>Action Required — Initialize New PPS File:</b>
        <ol>
          <li>Open the master PPS template:<br>
              <a href="https://docs.google.com/spreadsheets/d/1G36w4u6A5YRw-JPSOLKdU93PXsjPtkRUa8Yyp-O47xw/edit?gid=0#gid=0" target="_blank" style="color:#1a73e8; font-weight:bold;">Open Master PPS Template</a>
          </li>
          <li>Clone it into your tracking drive: <b>File → Make a copy</b></li>
          <li>Rename the copy to this client's company name</li>
          <li style="color:#c5221f; font-weight:bold;">Configure benefit headers in Row 6 only. Do not touch default columns.</li>
          <li>Copy the Sheet ID from the new file's URL.</li>
        </ol>
      </div>
    </div>

    <div id="existingGroupCard" style="display:none">
      <div class="pps-card">
        <b>Connect Existing Group Tracker:</b>
        <ol>
          <li>Open the group's existing PPS workbook</li>
          <li>Copy the Sheet ID from the URL (the long string between /d/ and /edit)</li>
        </ol>
      </div>
    </div>

    <div class="field" id="ppsIdField" style="display:none">
      <label class="lbl">Target PPS Workbook Sheet ID</label>
      <input id="ppsWorkbookId" type="text" placeholder="Paste Sheet ID here" />
      <p class="hint">Leave blank if PPS lives in this same workbook.</p>
    </div>

    <div class="btn-row">
      <button class="btn btn-ghost"   onclick="prev(3)">← Back</button>
      <button class="btn btn-primary" onclick="next(3)">Next →</button>
    </div>
  </div>

  <div class="tab" id="tab4">
    <div class="field">
      <label class="lbl">Pay Frequency</label>
      <select id="payFreq">
        <option value="52">Weekly (52/year)</option>
        <option value="26" selected>Bi-Weekly (26/year)</option>
        <option value="24">Semi-Monthly (24/year)</option>
        <option value="12">Monthly (12/year)</option>
      </select>
    </div>
    <div class="field" style="margin-top:14px;">
      <label class="lbl" style="color:#1a73e8;">Money to Leave in Paycheck ($)</label>
      <input id="leaveInPaycheck" type="number" min="0" placeholder="0" />
      <p class="hint">Reduces the per-employee build budget by this amount during auto-enroll. The Remaining Formula column still shows the full allotment minus all premiums.</p>
    </div>
    <p class="lbl" style="margin-top:14px">PPS Column Mapping</p>
    <div id="ppsMapBox" class="info-card">Loading PPS columns…</div>
    <div class="btn-row">
      <button class="btn btn-ghost"   onclick="prev(4)">← Back</button>
      <button class="btn btn-primary" onclick="next(4)">Next →</button>
    </div>
  </div>

  <div class="tab" id="tab5">
    <p class="lbl">Savings Data (optional)</p>
    <textarea id="savingsRaw" placeholder="Jane&#9;Smith&#9;$54.50"></textarea>
    <p class="lbl" style="margin-top:14px">Salary Data (optional)</p>
    <textarea id="salaryRaw" placeholder="Smith&#9;Jane&#9;$48,000"></textarea>
    <div class="btn-row">
      <button class="btn btn-ghost"   onclick="prev(5)">← Back</button>
      <button class="btn btn-primary" id="saveBtn" onclick="doSave()">Save &amp; Import →</button>
    </div>
  </div>

  <div id="status"></div>

<script>
  var current    = 1;
  var TOTAL_TABS = 5;
  var ppsCols    = [];
  var ppsDetected = null;

  google.script.run.withSuccessHandler(function(c) {
    var setIf = function(id, v) {
      var el = document.getElementById(id);
      if (el && !el.value && v !== '' && v != null) el.value = v;
    };
    setIf('company',          c.company);
    setIf('groupId',          c.groupId);
    setIf('enrollmentCutoff', c.enrollmentCutoff);
    setIf('effectiveDate',    c.effectiveDate);
    setIf('signDate',         c.signDate);
    setIf('rateSheetId',      c.rateSheetId);
    setIf('ppsWorkbookId',    c.ppsWorkbookId);
    setIf('leaveInPaycheck',  c.leaveInPaycheck);

    if (c.payFreq) document.getElementById('payFreq').value = String(c.payFreq);
    if (document.getElementById('rateSheetId').value) onRateIdChange();
    if (c.stdLabel) setTimeout(function() {
      var sel = document.getElementById('stdLabel');
      for (var i = 0; i < sel.options.length; i++) {
        if (sel.options[i].value === c.stdLabel) { sel.selectedIndex = i; break; }
      }
    }, 800);

    if (c.ppsWorkbookId) {
      document.querySelectorAll('input[name="groupType"]').forEach(function(r) {
        if (r.value === 'existing') r.checked = true;
      });
      onGroupTypeChange('existing');
    }
    window.__savedPPSMapping = c.ppsMapping;
  }).loadCurrentConfig();

  function onGroupTypeChange(type) {
    document.getElementById('newGroupCard').style.display      = type === 'new'      ? 'block' : 'none';
    document.getElementById('existingGroupCard').style.display = type === 'existing' ? 'block' : 'none';
    document.getElementById('ppsIdField').style.display        = 'block';
  }

  function showTab(n) {
    for (var i = 1; i <= TOTAL_TABS; i++) {
      document.getElementById('tab' + i).classList.toggle('active', i === n);
      var s = document.getElementById('s' + i);
      if (s) {
        s.className = 'step';
        if (i < n) s.classList.add('done');
        if (i === n) s.classList.add('active');
      }
    }
    current = n;
    if (n === 4 && ppsCols.length === 0) loadPPSCols();
  }

  function next(from) { if (!validateTab(from)) return; showTab(from + 1); }
  function prev(from) { showTab(from - 1); }

  function validateTab(n) {
    var err = '';
    if (n === 1) {
      if (!val('company'))                  err = 'Company name is required.';
      if (!err && !val('enrollmentCutoff')) err = 'Enrollment Cutoff is required.';
      if (!err && !val('effectiveDate'))    err = 'Effective Date is required.';
      if (!err && !val('signDate'))         err = 'Sign Date is required.';
    }
    if (n === 2) {
      if (!val('rateSheetId')) err = 'Rate Tables Sheet ID is required.';
      if (!err && !val('stdLabel')) err = 'Select an STD benefit period.';
    }
    if (n === 3) {
      if (!document.querySelector('input[name="groupType"]:checked'))
        err = 'Please pick New or Existing group.';
    }
    if (err) { setStatus(err, 'err-bg'); return false; }
    hideStatus();
    return true;
  }

  function val(id) {
    var el = document.getElementById(id);
    return el ? (el.value || '').trim() : '';
  }

  function onRateIdChange() {
    var id = val('rateSheetId'); if (!id) return;
    var hint = document.getElementById('stdHint');
    hint.textContent = 'Loading STD options…';
    google.script.run
      .withSuccessHandler(function(r) {
        if (r.error) { hint.textContent = 'Error: ' + r.error; return; }
        var sel = document.getElementById('stdLabel');
        sel.innerHTML = '<option value="">— select —</option>' +
          r.options.map(function(o) { return '<option value="' + esc(o) + '">' + esc(o) + '</option>'; }).join('');
        hint.textContent = r.options.length + ' option(s) found.';
      })
      .withFailureHandler(function(e) { hint.textContent = 'Failed: ' + (e.message || e); })
      .loadSTDOptionsForWizard(id);
  }

  function loadPPSCols() {
    document.getElementById('ppsMapBox').textContent = 'Loading PPS columns…';
    google.script.run.withSuccessHandler(function(r) {
      if (r.error) { document.getElementById('ppsMapBox').textContent = 'Error: ' + r.error; return; }
      ppsCols     = r.columns  || [];
      ppsDetected = r.detected || null;
      renderPPSMap();
    }).loadPPSColumnsForWizard();
  }

  function ddOpts(selected) {
    var html = '<option value="">— not mapped —</option>';
    ppsCols.forEach(function(c) {
      html += '<option value="' + esc(c) + '"' + (c === selected ? ' selected' : '') + '>' + esc(c) + '</option>';
    });
    return html;
  }

  function renderPPSMap() {
    var box  = document.getElementById('ppsMapBox');
    box.style.cssText = 'background:#fff;padding:6px 8px;border:1px solid #e0e0e0';
    var saved = window.__savedPPSMapping || null;
    var savedPrem = {}, savedTL = {}, savedSSN = '';
    if (saved) {
      (saved.premium || []).forEach(function(p) { savedPrem[p.src] = p.dst; });
      (saved.tl      || []).forEach(function(t) {
        if (t.slot === -1) savedTL.consolidated = t.dst;
        else savedTL['slot' + t.slot] = t.dst;
      });
      savedSSN = saved.ssnLast4Dst || '';
    }
    if (!saved && ppsDetected) {
      (ppsDetected.premiumMappings  || []).forEach(function(m) { if (m.matchedDst) savedPrem[m.src] = m.matchedDst; });
      (ppsDetected.tlMappings       || []).forEach(function(m, i) { if (m.matchedDst) savedTL['slot' + i] = m.matchedDst; });
      if (ppsDetected.consolidatedTLDst) savedTL.consolidated = ppsDetected.consolidatedTLDst;
      savedSSN = ppsDetected.ssnLast4Matched || '';
    }
    var rows = mapRow('SSN Last 4 → PPS', 'ssn_last4', savedSSN);
    [
      { src: 'AE Premium Amount',   label: 'Accident'          },
      { src: 'CI Premium Amount',   label: 'Critical Illness'  },
      { src: 'Life Premium Amount', label: 'Whole Life'        },
      { src: 'HI Premium Amount',   label: 'Hospital Indemnity'},
      { src: 'DI Premium Amount',   label: 'Disability / STD'  },
    ].forEach(function(p, i) {
      rows += mapRow(p.label, 'prem_' + i + '__' + esc(p.src), savedPrem[p.src] || '');
    });
    var hasConsolidated = ppsCols.indexOf('Term Life After Tax') !== -1;
    var hasSpecific     = ppsCols.indexOf('Term Life 10 Yr After Tax')    !== -1 ||
                          ppsCols.indexOf('Term Life 20 Yr After Tax')    !== -1 ||
                          ppsCols.indexOf('Term Life To Age 70 After Tax') !== -1;
    if (hasSpecific || !hasConsolidated) {
      ['10 Year','20 Year','To Age 70'].forEach(function(lbl, slot) {
        rows += mapRow('Term Life — ' + lbl, 'tl_' + slot, savedTL['slot' + slot] || '');
      });
    } else {
      rows += mapRow('Term Life (single col)', 'tl_consolidated', savedTL.consolidated || '');
    }
    box.innerHTML = rows;
  }

  function mapRow(label, id, selected) {
    return '<div class="map-row"><div class="map-lbl">' + esc(label) +
      '</div><select id="' + id + '">' + ddOpts(selected) + '</select></div>';
  }

  function collectPPSMapping() {
    var gv = function(id) { var el = document.getElementById(id); return el ? (el.value || '').trim() : ''; };
    var premium = [];
    ['AE Premium Amount','CI Premium Amount','Life Premium Amount','HI Premium Amount','DI Premium Amount']
      .forEach(function(src, i) {
        var dst = gv('prem_' + i + '__' + src);
        if (dst) premium.push({ src: src, dst: dst });
      });
    var tl = [], c = gv('tl_consolidated');
    if (c) {
      tl.push({ slot: -1, src: 'consolidated', dst: c });
    } else {
      ['tl_0','tl_1','tl_2'].forEach(function(id, i) {
        tl.push({ slot: i, src: ['TL Premium','TL Premium 2','TL Premium 3'][i], dst: gv(id) });
      });
    }
    return { ssnLast4Dst: gv('ssn_last4'), premium: premium, tl: tl };
  }

  function doSave() {
    hideStatus();
    var btn = document.getElementById('saveBtn');
    btn.disabled = true; btn.textContent = 'Saving…';
    google.script.run
      .withSuccessHandler(function(html) {
        setStatus(html, 'ok-bg');
        btn.disabled = false; btn.textContent = 'Save & Import →';
      })
      .withFailureHandler(function(err) {
        setStatus('Error: ' + (err.message || err), 'err-bg');
        btn.disabled = false; btn.textContent = 'Save & Import →';
      })
      .saveSetupAndImport(JSON.stringify({
        company:          val('company'),
        groupId:          val('groupId'),
        enrollmentCutoff: val('enrollmentCutoff'),
        effectiveDate:    val('effectiveDate'),
        signDate:         val('signDate'),
        rateSheetId:      val('rateSheetId'),
        stdLabel:         val('stdLabel'),
        payFreq:          val('payFreq'),
        ppsWorkbookId:    val('ppsWorkbookId'),
        ppsMapping:       collectPPSMapping(),
        savingsRaw:       val('savingsRaw'),
        salaryRaw:        val('salaryRaw'),
        leaveInPaycheck:  val('leaveInPaycheck'),
      }));
  }

  function setStatus(html, cls) {
    var s = document.getElementById('status');
    s.innerHTML = html; s.className = cls; s.style.display = 'block';
  }
  function hideStatus() { document.getElementById('status').style.display = 'none'; }
  function esc(s) {
    return String(s || '').replace(/[&<>"']/g, function(c) {
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
    });
  }
</script>
</body></html>`;
}

// ============================================================
// §4  STEP 2 — BUILD PLANS  (cushion + universal row build)
// ============================================================
function runBuildPlans() {
  const ui  = SpreadsheetApp.getUi();
  const cfg = getConfig();

  const missing = [];
  if (!cfg['Rate Tables Sheet ID']) missing.push('Rate Tables Sheet ID');
  if (!cfg['Effective Date'])       missing.push('Effective Date');
  if (!cfg['Sign Date'])            missing.push('Sign Date');
  if (!cfg['Company Name'])         missing.push('Company Name');
  if (!cfg['Enrollment Cutoff'])    missing.push('Enrollment Cutoff');
  if (!cfg['STD Benefit Period'])   missing.push('STD Benefit Period');
  if (missing.length) {
    ui.alert('Setup incomplete', 'Missing config:\n• ' + missing.join('\n• ') +
             '\n\nOpen Step 1: Setup & Import to lock these in.', ui.ButtonSet.OK);
    return;
  }

  let rateSS;
  try { rateSS = SpreadsheetApp.openById(cfg['Rate Tables Sheet ID']); }
  catch (e) { ui.alert('Cannot open Rate Tables sheet — check the ID.'); return; }

  const stdSheet = rateSS.getSheetByName('Short Term Disability');
  if (!stdSheet) { ui.alert('Rate sheet has no "Short Term Disability" tab.'); return; }
  const stdOpts = detectSTDOptions(stdSheet);
  const selectedSTD = stdOpts.find(o => o.label === cfg['STD Benefit Period']) || stdOpts[0];
  if (!selectedSTD) { ui.alert('No STD options found in rate sheet.'); return; }

  const effectiveDate = new Date(cfg['Effective Date']);
  const signDate      = new Date(cfg['Sign Date']);
  if (isNaN(effectiveDate.getTime())) { ui.alert('Bad Effective Date in config.'); return; }
  if (isNaN(signDate.getTime()))      { ui.alert('Bad Sign Date in config.'); return; }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsAE  = ss.getSheetByName('Auto Enroll Sheet');
  const wsDem = ss.getSheetByName('Demographic');
  if (!wsAE)  { ui.alert('Missing "Auto Enroll Sheet" tab.'); return; }
  if (!wsDem) { ui.alert('Missing "Demographic" tab.');       return; }

  let aeLastCol = wsAE.getLastColumn();
  let aeHeaders = wsAE.getRange(1, 1, 1, aeLastCol).getValues()[0];
  let colMap    = buildColumnMap(aeHeaders);
  if (colMap['Remaining Formula'] === undefined) {
    wsAE.getRange(1, aeHeaders.length + 1).setValue('Remaining Formula');
    aeHeaders = [...aeHeaders, 'Remaining Formula'];
    colMap    = buildColumnMap(aeHeaders);
  }

  const lastRow = wsAE.getLastRow();
  if (lastRow < 2) { ui.alert('No employee data on Auto Enroll Sheet.'); return; }
  const dataRange = wsAE.getRange(2, 1, lastRow - 1, aeHeaders.length);
  const data      = dataRange.getValues();
  const rates     = loadAllRateTables(rateSS, selectedSTD);

  const leaveInPaycheck = Number(cfg['Leave in Paycheck']) || 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // Calculate age for every row that carries a DOB — Auto Enroll
    // writes plans against any row with Allotments > 0 regardless of
    // Relation. (Intentional — no EE-only guard.)
    const dob = row[colMap['DOB']];
    row[colMap['Age']] = (dob instanceof Date) ? calcAge(dob, effectiveDate) : '';

    const age       = row[colMap['Age']];
    const allotment = parseFloat(row[colMap['Allotments']])    || 0;
    const salary    = parseFloat(row[colMap['Annual Salary']]) || 0;
    if (age === '' || allotment <= 0) continue;

    // Cushion applies only to the build budget; Remaining Formula still
    // reports against the full allotment so the cushion is visible.
    const buildBudget = Math.max(0, allotment - leaveInPaycheck);
    let used = 0;

    const tl = assignTermLife(age, buildBudget, rates);
    if (tl) {
      row[colMap['TL Plan Type']]      = tl.plan;
      row[colMap['TL Insured Option']] = 'EE';
      row[colMap['TL Benefit Amount']] = tl.benefit;
      row[colMap['TL Premium']]        = tl.rate;
    }
    normalizeTLRow(row, colMap);
    used = totalTLPremium(row, colMap);

    const wl = assignWholeLife(age, buildBudget - used, rates);
    if (wl) {
      row[colMap['Life Plan Type']]          = 'Whole Life';
      row[colMap['Life Insured Option']]     = 'EO';
      row[colMap['Life Certificate Amount']] = wl.benefit;
      row[colMap['Life Premium Amount']]     = wl.rate;
      used += wl.rate;
    }

    const ae = assignAccident(buildBudget - used, rates);
    if (ae) {
      row[colMap['AE Plan Type']]      = ae.plan;
      row[colMap['AE Insured Option']] = 'EE';
      row[colMap['AE Premium Amount']] = ae.rate;
      used += ae.rate;
    }

    const ci = assignCriticalIllness(age, buildBudget - used, rates);
    if (ci) {
      row[colMap['CI Plan Type']]      = 'Critical Tier 2';
      row[colMap['CI Insured Option']] = 'EE';
      row[colMap['CI Benefit Amount']] = ci.benefit;
      row[colMap['CI Premium Amount']] = ci.rate;
      used += ci.rate;
    }

    const hi = assignHospitalIndemnity(buildBudget - used, rates);
    if (hi) {
      row[colMap['HI Plan Type']]      = hi.plan;
      row[colMap['HI Insured Option']] = 'EE';
      row[colMap['HI Premium Amount']] = hi.rate;
      row[colMap['HI Benefit Amount']] = 1000;
      used += hi.rate;
    }

    const di = assignSTD(age, salary, buildBudget - used, rates);
    if (di) {
      row[colMap['DI Plan Type']]      = selectedSTD.label;
      row[colMap['DI Benefit Amount']] = di.benefit;
      row[colMap['DI Premium Amount']] = di.rate;
      used += di.rate;
    }

    const v = (c) => (colMap[c] !== undefined ? (parseFloat(row[colMap[c]]) || 0) : 0);
    const spouseWLPremium = colMap['Spouse Life Premium Amount'] !== undefined ? v('Spouse Life Premium Amount') : 0;

    row[colMap['Remaining Formula']] =
      v('Allotments') - v('TL Premium 3') - v('TL Premium 2') - v('TL Premium') -
      v('Life Premium Amount') - spouseWLPremium - v('HI Premium Amount') -
      v('CI Premium Amount') - v('DI Premium Amount') - v('AE Premium Amount');

    populatePolicyDates(row, colMap, effectiveDate, signDate);
  }

  dataRange.setValues(data);
  writeGroupMetadataToDemographic(wsDem, cfg['Company Name'], cfg['Enrollment Cutoff']);
  runBenefitPackageBuilder();

  ui.alert('Build Plans complete',
    'Company: ' + cfg['Company Name'] + '\n' +
    'Effective: ' + Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') + '\n' +
    'STD: ' + selectedSTD.label + '\n' +
    'Cushion: $' + leaveInPaycheck + '/employee',
    ui.ButtonSet.OK);
}

// ============================================================
// §5  STEP 3 — SYNC OUT
// ============================================================
function runSyncOut() {
    const ui  = SpreadsheetApp.getUi();
    const cfg = getConfig();
    if (!cfg['PPS Mapping']) {
      ui.alert('PPS mapping not configured',
        'Open Setup & Import and complete the "Pay & PPS Mapping" section first.',
        ui.ButtonSet.OK);
      return;
    }

    const log = [];
    try { stackAllPlans();         log.push('✓ Stack All Plans');     } catch (e) { log.push('✗ Stack: ' + e.message); }
    try { syncDemographics();      log.push('✓ Sync Employees');      } catch (e) { log.push('✗ Employees: ' + e.message); }
    try { _runPPSSyncFromConfig(); log.push('✓ Sync to PPS');         } catch (e) { log.push('✗ PPS: ' + e.message); }

    // GHL push is intentionally manual now — run Tools ▸ "Push to GHL only" when ready.
    log.push('⏭ GHL push skipped (manual — Tools ▸ Push to GHL only)');

    ui.alert('Sync Out complete', log.join('\n'), ui.ButtonSet.OK);
  }


function _runPPSSyncFromConfig() {
  const cfg = getConfig();
  const m   = cfg['PPS Mapping'];
  if (!m) throw new Error('PPS Mapping not in config — re-run Setup & Import.');
  const mappingsJson = JSON.stringify({
    premium:     m.premium     || [],
    tl:          m.tl          || [],
    ssnLast4Dst: m.ssnLast4Dst || '',
  });
  const payFreq = String(cfg['Pay Frequency'] || 26);
  return runPPSSyncFromSidebar(mappingsJson, payFreq);
}

// ============================================================
// §6  PUSH TO GHL (Demographic → GHL contacts)
// ============================================================
function pushDemographicToGHL() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GHL_PIT');
  const locId = props.getProperty('GHL_LOCATION_ID');

  if (!token || !locId) {
    ui.alert('Missing GHL credentials',
      'Add GHL_PIT and GHL_LOCATION_ID to Script Properties first.',
      ui.ButtonSet.OK);
    return;
  }

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ssId   = ss.getId();
  const wsDem  = ss.getSheetByName('Demographic');
  if (!wsDem) { ui.alert('Demographic tab not found.'); return; }

  const filterResp = ui.alert('Push to GHL',
    'YES  = only Yes/Customized employees\n' +
    'NO   = every Demographic row (full roster)\n' +
    'CANCEL = abort',
    ui.ButtonSet.YES_NO_CANCEL);
  if (filterResp === ui.Button.CANCEL) return;
  const onlyApproved = (filterResp === ui.Button.YES);

  const lastCol = wsDem.getLastColumn();
  const lastRow = wsDem.getLastRow();
  if (lastRow < 2) { ui.alert('Demographic has no data.'); return; }
  let headers   = wsDem.getRange(1, 1, 1, lastCol).getValues()[0];
  let colMap    = buildColumnMap(headers);

  let newCols = lastCol;
  const ensureCol = (name) => {
    if (colMap[name] === undefined) {
      newCols++;
      wsDem.getRange(1, newCols).setValue(name);
      colMap[name] = newCols - 1;
    }
  };
  ensureCol('Contact ID');
  ensureCol('GHL Sync Status');
  ensureCol('GHL Sync Last Run');
  if (newCols !== lastCol) headers = wsDem.getRange(1, 1, 1, newCols).getValues()[0];

  const data = wsDem.getRange(2, 1, lastRow - 1, newCols).getValues();
  for (let r = 0; r < data.length; r++) while (data[r].length < newCols) data[r].push('');

  let fieldMap;
  try { fieldMap = _ghlFetchCustomFields(locId, token); }
  catch (err) { ui.alert('Failed to load GHL custom fields', String(err.message || err), ui.ButtonSet.OK); return; }

  const CUSTOM_FIELD_BINDINGS = [
    { sheet: 'Net Tax Savings',    ghl: 'Allotment' },
    { sheet: 'Enrollment Cutoff',  ghl: 'Enrollment Cutoff' },
    { sheet: 'Benefits Package',   ghl: 'Benefits Package' },
    { sheet: 'Savings',            ghl: 'Savings' },
    { sheet: 'Benefit Package',    ghl: 'Benefit Package' },
    { sheet: 'Company Name',       ghl: 'Company Name' },
    { sheet: 'Gender',             ghl: 'Gender' },
    { sheet: 'Age',                ghl: 'Age' },
  ];

  const tsStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const ssnCol = colMap['Employee SSN'];

  let pushed = 0, updated = 0, skipped = 0, failed = 0;
  const failures = [];

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (onlyApproved) {
      const flag = String(row[0]).trim().toLowerCase();
      if (flag !== 'yes' && flag !== 'customized') { skipped++; continue; }
    }
    const email = String(row[colMap['Email Address']] || '').trim();
    const fn    = String(row[colMap['First Name']]    || '').trim();
    const ln    = String(row[colMap['Last Name']]     || '').trim();
    if (!email && !fn && !ln) { skipped++; continue; }

    const body = {
      locationId:  locId,
      firstName:   fn,
      lastName:    ln,
      email:       email,
      phone:       _normalizePhone(row[colMap['Phone Number']]),
      companyName: String(row[colMap['Company Name']] || '').trim(),
      address1:    String(row[colMap['Mailing Address1']] || '').trim(),
      city:        String(row[colMap['City']]  || '').trim(),
      state:       String(row[colMap['State']] || '').trim(),
      postalCode:  String(row[colMap['Zip']]   || '').trim(),
      dateOfBirth: _normalizeDOB(row[colMap['DOB']]),
      source:      'TBG Plan Builder',
    };

    const customFields = [];
    const pushCF = (ghlName, val) => {
      if (val === undefined || val === null || val === '') return;
      const id = _findFieldId(fieldMap, ghlName);
      if (!id) return;
      customFields.push({ id: id, value: String(val) });
    };

    const ssnDigits = String(row[ssnCol] || '').replace(/[^0-9]/g, '');
    const last4 = ssnDigits.length >= 4 ? ssnDigits.slice(-4) : '';
    pushCF('SSN Last 4', last4);
    pushCF('Employee SSN Last 4 digits', last4);

    CUSTOM_FIELD_BINDINGS.forEach(b => {
      if (colMap[b.sheet] === undefined) return;
      pushCF(b.ghl, row[colMap[b.sheet]]);
    });
    pushCF('Plan Builder Sheet ID', ssId);

    if (customFields.length) body.customFields = customFields;

    try {
      const result = _ghlUpsertContact(body, token);
      const contactId = result && result.contact && result.contact.id;
      if (contactId) {
        row[colMap['Contact ID']]        = contactId;
        row[colMap['GHL Sync Status']]   = result.new ? 'Created' : 'Updated';
        row[colMap['GHL Sync Last Run']] = tsStr;
        if (result.new) pushed++; else updated++;
      } else {
        row[colMap['GHL Sync Status']]   = 'No contact ID returned';
        row[colMap['GHL Sync Last Run']] = tsStr;
        failed++;
      }
    } catch (err) {
      row[colMap['GHL Sync Status']]   = 'ERROR: ' + String(err.message || err).slice(0, 200);
      row[colMap['GHL Sync Last Run']] = tsStr;
      failures.push((fn + ' ' + ln + ' — ' + err.message).slice(0, 160));
      failed++;
    }
    Utilities.sleep(120);
  }

  wsDem.getRange(2, 1, data.length, newCols).setValues(data);

  let msg = 'GHL Push complete.\n\n' +
            '✓ Created:  ' + pushed  + '\n' +
            '✓ Updated:  ' + updated + '\n' +
            '⏭ Skipped:  ' + skipped + '\n' +
            '✗ Failed:   ' + failed;
  if (failures.length) msg += '\n\nFirst few errors:\n• ' + failures.slice(0, 5).join('\n• ');
  ui.alert(msg);
}

// ============================================================
// §7  GHL API HELPERS
// ============================================================
function _ghlFetchCustomFields(locationId, token) {
  const url = GHL_BASE + '/locations/' + locationId + '/customFields';
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token, 'Version': GHL_API_VERSION, 'Accept': 'application/json' },
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode(), txt = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error('GHL customFields ' + code + ': ' + txt.slice(0, 200));
  const json = JSON.parse(txt);
  const list = json.customFields || json.fields || [];
  const map = {};
  list.forEach(f => { const name = String(f.name || '').toLowerCase().trim(); if (name) map[name] = f.id; });
  return map;
}

function _findFieldId(fieldMap, name) {
  if (!name) return null;
  return fieldMap[String(name).toLowerCase().trim()] || null;
}

function _ghlUpsertContact(body, token) {
  const resp = UrlFetchApp.fetch(GHL_BASE + '/contacts/upsert', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token, 'Version': GHL_API_VERSION, 'Accept': 'application/json' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode(), txt = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error('GHL upsert ' + code + ': ' + txt.slice(0, 200));
  return JSON.parse(txt);
}

function _ghlGetContact(contactId, token) {
  const resp = UrlFetchApp.fetch(GHL_BASE + '/contacts/' + encodeURIComponent(contactId), {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token, 'Version': GHL_API_VERSION, 'Accept': 'application/json' },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) return null;
  return JSON.parse(resp.getContentText() || '{}');
}

// ============================================================
// §8  FIELD NORMALIZERS
// ============================================================
function _normalizePhone(raw) {
  if (!raw) return '';
  const digits = String(raw).replace(/[^0-9]/g, '');
  if (!digits) return '';
  if (digits.length === 10) return '+1' + digits;
  if (digits.length === 11 && digits[0] === '1') return '+' + digits;
  return '+' + digits;
}

function _normalizeDOB(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  const d = new Date(val);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + dd;
  }
  return String(val);
}

function testGHLConnection() {
  const props = PropertiesService.getScriptProperties();
  const locId = props.getProperty('GHL_LOCATION_ID');
  const token = props.getProperty('GHL_PIT');
  if (!locId || !token) { Logger.log('Missing GHL_LOCATION_ID or GHL_PIT in Script Properties'); return; }
  const map = _ghlFetchCustomFields(locId, token);
  Logger.log('Custom fields found in GHL (' + Object.keys(map).length + '):');
  Object.keys(map).sort().forEach(k => Logger.log('  - ' + k));
}

// ============================================================
// §9  WEBHOOK — doPost (Widget Customization / Concurrency Protected)
// ============================================================
  function doPost(e) {
    const lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(30000)) {
        return _jsonOut({ ok:false, error:'System busy. Try again.' });
      }
      if (!e || !e.postData || !e.postData.contents) throw new Error('Empty POST body');
      const payload = JSON.parse(e.postData.contents);

      const expected = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
      if (expected && payload.secret !== expected) {
        return _jsonOut({ ok:false, error:'Unauthorized' });
      }

      // Route: GHL form submission with benefit_summary text vs structured widget payload
      const result = (payload.benefit_summary && !payload.benefits)
        ? syncFromBenefitSummary(payload)
        : syncCustomizedEmployee(payload);

      return _jsonOut({ ok:true, ...result });                                                                                                                                     
    } catch (err) {
      console.error('doPost error:', err);
      return _jsonOut({ ok:false, error:String(err.message || err) });
    } finally {
      lock.releaseLock();                                                                                                                                                          
    }
  }

function doGet() {
  return _jsonOut({ ok: true, service: 'TBG Widget Sync Engine', ts: new Date().toISOString() });
}

function _jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function parseBenefitSummary(text) {
    const benefits = {
      accident:    { on: false }, critical:    { on: false },
      hospital:    { on: false }, disability:  { on: false },
      term:        { on: false }, wholeEmp:    { on: false }, wholeSpouse: { on: false }
    };
    if (!text) return benefits;

    const tierMap = s => {
      const t = (s || '').toLowerCase();
      if (t.includes('family'))                            return 'family';                                                                                                        
      if (t.includes('spouse'))                            return 'employee_spouse';
      if (t.includes('children') || t.includes('child'))  return 'employee_children';
      return 'employee';
    };
    const parseAmt = s => parseFloat(String(s).replace(/[$,\/wk\s]/g, '')) || 0;
    const parsePlan = s => /plan\s*2/i.test(s || '') ? 'plan2' : 'plan1';

    for (const raw of text.split(/[\n\r]+/)) {
      const line  = raw.trim();
      const parts = line.split('|').map(p => p.trim());
      const head  = parts[0] || '';
      const premM = head.match(/\$([\d,]+\.?\d*)\/mo/);
      const prem  = premM ? parseFloat(premM[1].replace(/,/g,'')) : 0;

      if (/^Accident Expense:/i.test(head)) {
        benefits.accident = { on:true, tier:tierMap(parts[1]), plan:parsePlan(parts[2]), premium:prem };

      } else if (/^Critical Illness:/i.test(head)) {
        const benPart = parts.find(p => /benefit/i.test(p)) || '';
        benefits.critical = { on:true, tier:tierMap(parts[1]), benefit:parseAmt(benPart.replace(/benefit/i,'')), premium:prem };

      } else if (/^Hospital Indemnity:/i.test(head)) {
        benefits.hospital = { on:true, tier:tierMap(parts[1]), plan:parsePlan(parts[2]), premium:prem };

      } else if (/^Short.?Term Disability:/i.test(head)) {
        const wkPart = parts.find(p => /\/wk/i.test(p)) || '';
        benefits.disability = { on:true, weekly:parseAmt(wkPart), period:'26wk', premium:prem };

      } else if (/^Term Life:/i.test(head)) {
        const benPart = parts.find(p => /^\$[\d,]+\.?\d*$/.test(p)) || '';
        const typPart = parts.find(p => /10.year|20.year|age\s*70/i.test(p)) || '';
        const termType = /10.year/i.test(typPart) ? 'term_10' : /20.year/i.test(typPart) ? 'term_20' : 'term_to_70';
        benefits.term = { on:true, tier:tierMap(parts[1]), termType, benefit:parseAmt(benPart), premium:prem };

      } else if (/^Whole Life.*You:/i.test(head)) {
        const benPart = parts.find(p => /^\$[\d,]+\.?\d*$/.test(p)) || '';
        benefits.wholeEmp = { on:true, benefit:parseAmt(benPart), childRider:0, premium:prem };

      } else if (/^Whole Life.*Spouse:/i.test(head)) {
        const benPart = parts.find(p => /^\$[\d,]+\.?\d*$/.test(p)) || '';
        benefits.wholeSpouse = { on:true, benefit:parseAmt(benPart), premium:prem };
      }
    }
    return benefits;
  }

  function syncFromBenefitSummary(payload) {
    const raw = payload.benefit_summary || payload.benefitSummary || '';
    const summary = raw.replace(/\\n/g, '\n');   // ← fixes literal \n
    if (!summary.trim()) throw new Error('No benefit_summary in payload');
    const normalized = {
      spreadsheetId: payload.spreadsheetId || null,
      match: payload.match || {
        contactId: payload.contact_id || '',
        email:     payload.email      || '',
        firstName: payload.first_name || '',                                                                                                                                       
        lastName:  payload.last_name  || ''
      },
      benefits:      parseBenefitSummary(summary),
      benefitSummary: summary,
      totalMonthly:  payload.total_monthly || 0                                                                                                                                    
    };
    return syncCustomizedEmployee(normalized);
  }

// ============================================================
// §10  WEBHOOK HELPERS
// ============================================================
function syncCustomizedEmployee(payload) {
  const props = PropertiesService.getScriptProperties();
  const payFreqGlobal = parsePayFrequency(props.getProperty('PAY_FREQUENCY')) || 26;

  const ssId = _resolveSpreadsheetId(payload, props);
  if (!ssId) throw new Error('Could not resolve target spreadsheet matrix.');

  const ss    = SpreadsheetApp.openById(ssId);
  const wsAE  = ss.getSheetByName('Auto Enroll Sheet');
  const wsDem = ss.getSheetByName('Demographic');
  const wsPPS = ss.getSheetByName('PPS');
  if (!wsAE)  throw new Error("'Auto Enroll Sheet' tab not found in workbook " + ssId);
  if (!wsDem) throw new Error("'Demographic' tab not found in workbook " + ssId);
  if (!wsPPS) throw new Error("'PPS' tab not found in workbook " + ssId);

  const wbConfig = getConfig(ss);
  const payFreq  = Number(wbConfig['Pay Frequency']) || payFreqGlobal;

  const demHeaders = wsDem.getRange(1, 1, 1, wsDem.getLastColumn()).getValues()[0];
  const demColMap  = buildColumnMap(demHeaders);
  const demLastRow = wsDem.getLastRow();
  if (demLastRow < 2) throw new Error('Demographic has no data rows.');
  const demData = wsDem.getRange(2, 1, demLastRow - 1, demHeaders.length).getValues();

  const demContactCol = demColMap['Contact ID'];
  const demEmailCol   = demColMap['Email Address'];
  const demSSNCol     = demColMap['Employee SSN'];
  const demFNCol      = demColMap['First Name'];
  const demLNCol      = demColMap['Last Name'];

  const m = payload.match || {};
  const wantCID = String(m.contactId || '').trim();
  const wantEm  = String(m.email || '').toLowerCase().trim();
  const wantFN  = String(m.firstName || '').toLowerCase().trim();
  const wantLN  = String(m.lastName  || '').toLowerCase().trim();

  let demRowIdx = -1, matchedBy = '';
  if (wantCID && demContactCol !== undefined) {
    for (let i = 0; i < demData.length; i++) {
      if (String(demData[i][demContactCol] || '').trim() === wantCID) { demRowIdx = i; matchedBy = 'contact_id'; break; }
    }
  }
  if (demRowIdx === -1 && wantEm && demEmailCol !== undefined) {
    for (let i = 0; i < demData.length; i++) {
      if (String(demData[i][demEmailCol] || '').toLowerCase().trim() === wantEm) { demRowIdx = i; matchedBy = 'email'; break; }
    }
  }
  if (demRowIdx === -1 && wantFN && wantLN && demFNCol !== undefined && demLNCol !== undefined) {
    for (let i = 0; i < demData.length; i++) {
      const fn = String(demData[i][demFNCol] || '').toLowerCase().trim();
      const ln = String(demData[i][demLNCol] || '').toLowerCase().trim();
      if (fn === wantFN && ln === wantLN) { demRowIdx = i; matchedBy = 'name'; break; }
    }
  }
  if (demRowIdx === -1) {
    throw new Error('Employee record mismatch on database execution pass.');
  }
  const ssn = padSSN(String(demData[demRowIdx][demSSNCol] || ''));
  if (!ssn) throw new Error('Matched Demographic row has no valid SSN lookup token.');

  const isAutomaticOptIn = (payload.event === 'opt_in');
  demData[demRowIdx][0] = isAutomaticOptIn ? 'Yes' : 'Customized';

  if (demColMap['Benefits Package'] !== undefined && payload.benefitSummary) {
    demData[demRowIdx][demColMap['Benefits Package']] = payload.benefitSummary;
  }
  if (demContactCol !== undefined && wantCID && !demData[demRowIdx][demContactCol]) {
    demData[demRowIdx][demContactCol] = wantCID;
  }
  wsDem.getRange(demRowIdx + 2, 1, 1, demHeaders.length).setValues([demData[demRowIdx]]);

  const aeHeaders = wsAE.getRange(1, 1, 1, wsAE.getLastColumn()).getValues()[0];
  const aeColMap  = buildColumnMap(aeHeaders);
  const ensureAECol = (name) => {
    if (aeColMap[name] === undefined) {
      const hdrs = wsAE.getRange(1, 1, 1, wsAE.getLastColumn()).getValues()[0];
      wsAE.getRange(1, hdrs.length + 1).setValue(name);
      aeColMap[name] = hdrs.length;
    }
  };
  ensureAECol('Remaining Formula');
  ensureAECol('Benefit Package');

  const aeWidth   = wsAE.getLastColumn();
  const aeLastRow = wsAE.getLastRow();
  if (aeLastRow < 2) throw new Error('Auto Enroll Sheet has no data rows.');
  const aeData    = wsAE.getRange(2, 1, aeLastRow - 1, aeWidth).getValues();

  let aeRowIdx = -1;
  for (let i = 0; i < aeData.length; i++) {
    const rel = String(aeData[i][aeColMap['Relation']] || '').toUpperCase().trim();
    if (rel !== 'EE') continue;
    if (padSSN(String(aeData[i][aeColMap['Employee SSN']] || '')) === ssn) { aeRowIdx = i; break; }
  }
  if (aeRowIdx === -1) throw new Error('Employee not found on Auto Enroll Sheet by SSN ' + ssn);
  while (aeData[aeRowIdx].length < aeWidth) aeData[aeRowIdx].push('');

  const aeRow = aeData[aeRowIdx];

  if (!isAutomaticOptIn) {
    _clearPlanFields(aeRow, aeColMap);
    _writePlansFromPayload(aeRow, aeColMap, payload);
  }

  const num = (v) => (typeof v === 'number' ? v : parseFloat(v) || 0);
  const allot = num(aeRow[aeColMap['Allotments']]);
  const used =
    num(aeRow[aeColMap['TL Premium 3']])        + num(aeRow[aeColMap['TL Premium 2']]) +
    num(aeRow[aeColMap['TL Premium']])          + num(aeRow[aeColMap['Life Premium Amount']]) +
    (aeColMap['Spouse Life Premium Amount'] !== undefined ? num(aeRow[aeColMap['Spouse Life Premium Amount']]) : 0) +
    num(aeRow[aeColMap['HI Premium Amount']])   + num(aeRow[aeColMap['CI Premium Amount']]) +
    num(aeRow[aeColMap['DI Premium Amount']])   + num(aeRow[aeColMap['AE Premium Amount']]);
    
  if (aeColMap['Remaining Formula'] !== undefined) aeRow[aeColMap['Remaining Formula']] = allot - used;
  if (aeColMap['Benefit Package'] !== undefined && payload.benefitSummary) {
    aeRow[aeColMap['Benefit Package']] = payload.benefitSummary;
  }

  wsAE.getRange(aeRowIdx + 2, 1, 1, aeWidth).setValues([aeRow]);

  const flagForPPS = isAutomaticOptIn ? demData[demRowIdx][0].toLowerCase() : 'yes';                                                                                               
  const ppsResult = _syncOneEmployeeToPPS(wsPPS, aeColMap, aeRow, ssn, payFreq, flagForPPS);

  return {
    matchedBy, contactId: wantCID || null, email: wantEm || null, ssn,
    spreadsheet: ssId, aeRow: aeRowIdx + 2, demRow: demRowIdx + 2,
    pps: ppsResult, payFreqUsed: payFreq, totalMonthly: payload.totalMonthly || null,
  };
}

// ============================================================
// CONTINUATION OF HELPER LOGIC PIPELINES
// ============================================================
function _resolveSpreadsheetId(payload, props) {
  if (payload && payload.spreadsheetId) return String(payload.spreadsheetId).trim();
  const cid = payload && payload.match && payload.match.contactId;
  if (cid) {
    try {
      const token = props.getProperty('GHL_PIT');
      if (token) {
        const json = _ghlGetContact(cid, token);
        const cf   = json && (json.contact || json) && (json.contact ? json.contact.customFields : json.customFields);
        if (Array.isArray(cf)) {
          for (const f of cf) {
            const v = String(f.value || '').trim();
            if (/^[a-zA-Z0-9_-]{20,}$/.test(v)) return v;
          }
        }
      }
    } catch (err) { console.warn('GHL mapping identification scan bypassed:', err.message); }
  }
  return props.getProperty('SPREADSHEET_ID') || null;
}

function _clearPlanFields(row, m) {
  const cols = [
    'TL Plan Type','TL Insured Option','TL Benefit Amount','TL Premium','TL Issue Date','TL Signed Date',
    'TL Plan Type 2','TL Insured Option 2','TL Benefit Amount 2','TL Premium 2','TL Issue Date 2','TL Signed Date 2',
    'TL Plan Type 3','TL Insured Option 3','TL Benefit Amount 3','TL Premium 3','TL Issue Date 3','TL Signed Date 3',
    'Life Plan Type','Life Insured Option','Life Certificate Amount','Life Premium Amount','Life Issue Date','Life Signed Date',
    'Spouse Life Plan Type','Spouse Life Insured Option','Spouse Life Certificate Amount','Spouse Life Premium Amount','Spouse Life Issue Date','Spouse Life Signed Date',
    'AE Plan Type','AE Insured Option','AE Premium Amount','AE Issue Date','AE Signed Date',
    'CI Plan Type','CI Insured Option','CI Benefit Amount','CI Premium Amount','CI Issue Date','CI Signed Date',
    'HI Plan Type','HI Insured Option','HI Benefit Amount','HI Premium Amount','HI Issue Date','HI Signed Date',
    'DI Plan Type','DI Benefit Amount','DI Premium Amount','DI Issue Date','DI Signed Date',
  ];
  cols.forEach(c => { if (m[c] !== undefined) row[m[c]] = ''; });
}

function _writePlansFromPayload(row, m, payload) {
  const b = payload.benefits || {};
  const today    = new Date();
  const issueRaw = payload.effectiveDate ? new Date(payload.effectiveDate) : today;
  const signRaw  = payload.signDate      ? new Date(payload.signDate)      : today;
  const issueDate = isNaN(issueRaw.getTime()) ? today : issueRaw;
  const signDate  = isNaN(signRaw.getTime())  ? today : signRaw;
  const set = (col, v) => { if (m[col] !== undefined) row[m[col]] = v; };

  if (b.accident && b.accident.on) {
    set('AE Plan Type',      _planNameAccident(b.accident.plan));
    set('AE Insured Option', _tierToOption(b.accident.tier));
    set('AE Premium Amount', _num(b.accident.premium));
    set('AE Issue Date',     issueDate);
    set('AE Signed Date',    signDate);
  }
  if (b.critical && b.critical.on) {
    set('CI Plan Type',      'Critical Tier 2');
    set('CI Insured Option', _tierToOption(b.critical.tier));
    set('CI Benefit Amount', _num(b.critical.benefit));
    set('CI Premium Amount', _num(b.critical.premium));
    set('CI Issue Date',     issueDate);
    set('CI Signed Date',    signDate);
  }
  if (b.hospital && b.hospital.on) {
    set('HI Plan Type',      _planNameHospital(b.hospital.plan));
    set('HI Insured Option', _tierToOption(b.hospital.tier));
    set('HI Benefit Amount', 1000);
    set('HI Premium Amount', _num(b.hospital.premium));
    set('HI Issue Date',     issueDate);
    set('HI Signed Date',    signDate);
  }
  if (b.disability && b.disability.on) {
    set('DI Plan Type',      _diPeriodLabel(b.disability.period));
    set('DI Benefit Amount', _num(b.disability.weekly));
    set('DI Premium Amount', _num(b.disability.premium));
    set('DI Issue Date',     issueDate);
    set('DI Signed Date',    signDate);
  }
  if (b.term && b.term.on) {
    const slot = _termSlotSuffix(b.term.termType);
    set('TL Plan Type'       + slot, _termPlanLabel(b.term.termType));
    set('TL Insured Option' + slot, _tierToOption(b.term.tier));
    set('TL Benefit Amount' + slot, _num(b.term.benefit));
    set('TL Premium'         + slot, _num(b.term.premium));
    set('TL Issue Date'      + slot, issueDate);
    set('TL Signed Date'     + slot, signDate);
  }
  if (b.wholeEmp && b.wholeEmp.on) {
    set('Life Plan Type',          'Whole Life');
    set('Life Insured Option',     'EO');
    set('Life Certificate Amount', _num(b.wholeEmp.benefit));
    set('Life Premium Amount',     _num(b.wholeEmp.premium));
    set('Life Issue Date',         issueDate);
    set('Life Signed Date',        signDate);
  }
  if (b.wholeSpouse && b.wholeSpouse.on) {
    set('Spouse Life Plan Type',          'Whole Life');
    set('Spouse Life Insured Option',     'SP');
    set('Spouse Life Certificate Amount', _num(b.wholeSpouse.benefit));
    set('Spouse Life Premium Amount',     _num(b.wholeSpouse.premium));
    set('Spouse Life Issue Date',         issueDate);
    set('Spouse Life Signed Date',        signDate);
  }
}

function _num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function _tierToOption(t) {
  switch (String(t || '').toLowerCase()) {
    case 'employee':          return 'EE';
    case 'employee_spouse':   return 'ES';
    case 'employee_children': return 'EC';
    case 'family':            return 'F';
    default:                  return 'EE';
  }
}
function _planNameAccident(p) { return p === 'plan2' ? 'Accident Tier 4' : 'Accident Tier 2'; }
function _planNameHospital(p) { return p === 'plan2' ? 'Plan 2' : 'Plan 1'; }
function _diPeriodLabel(p) { if (p === '13wk') return '13 Weeks'; if (p === '52wk') return '52 Weeks'; return '26 Weeks'; }
function _termPlanLabel(t) { if (t === 'term_10') return '10 Year'; if (t === 'term_20') return '20 Year'; return 'To Age 70'; }
function _termSlotSuffix(t) { if (t === 'term_20') return ' 2'; if (t === 'term_to_70') return ' 3'; return ''; }

function _readPPSMappingFromConfig(ss) {
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [k, v] of rows) {
    if (String(k).trim() === 'PPS Mapping' && typeof v === 'string' && v) {
      try { return JSON.parse(v); } catch (e) { return null; }
    }
  }
  return null;
}

function _syncOneEmployeeToPPS(wsPPS, aeColMap, aeRow, ssn, payFreq, currentFlagStatus) {
  const PPS_HEADER_ROW = 6;
  const PPS_DATA_START = 7;
  const ppsLastCol = wsPPS.getLastColumn();
  const ppsHeaders = wsPPS.getRange(PPS_HEADER_ROW, 1, 1, ppsLastCol).getValues()[0];
  const ppsColMap  = buildColumnMap(ppsHeaders);
  const ppsColMapLC = {};
  Object.keys(ppsColMap).forEach(k => { ppsColMapLC[k.toLowerCase().trim()] = ppsColMap[k]; });

  const savedMap = _readPPSMappingFromConfig(wsPPS.getParent());

  const findCol = (names) => {
    for (const n of names) {
      if (ppsColMap[n] !== undefined) return ppsColMap[n];
      const lc = n.toLowerCase().trim();
      if (ppsColMapLC[lc] !== undefined) return ppsColMapLC[lc];
    }
    return -1;
  };

  const DEMO_MAP = [
    { src: 'First Name',       dst: 'First Name'    },
    { src: 'Last Name',        dst: 'Last Name'     },
    { src: 'Gender',           dst: 'Gender'        },
    { src: 'DOB',              dst: 'Date of Birth' },
    { src: 'Mailing Address1', dst: 'Address'       },
    { src: 'City',             dst: 'City'          },
    { src: 'State',            dst: 'State'         },
    { src: 'Zip',              dst: 'Zip'           },
    { src: 'Phone Number',     dst: 'Phone'         },
    { src: 'Email Address',    dst: 'Email'         },
  ];

  const ppsLastRow = wsPPS.getLastRow();
  let ppsData = [];
  if (ppsLastRow >= PPS_DATA_START) {
    ppsData = wsPPS.getRange(PPS_DATA_START, 1, ppsLastRow - PPS_DATA_START + 1, ppsLastCol).getValues();
  }

  const ssnLast4Col = (savedMap && savedMap.ssnLast4Dst)
    ? findCol([savedMap.ssnLast4Dst])
    : findCol(['Employee SSNLast 4 digits','Employee SSN Last 4 digits','Employee SSN Last 4 Digits','SSN Last 4 Digits','SSN Last 4']);

  const ppsFNCol = findCol(['First Name']);
  const ppsLNCol = findCol(['Last Name']);
  const ssn4     = ssn.length >= 4 ? ssn.slice(-4) : '';
  const aeLN     = String(aeRow[aeColMap['Last Name']]  || '').toLowerCase().trim();
  const aeFN     = String(aeRow[aeColMap['First Name']] || '').toLowerCase().trim();
  const matchKey = ssn4 ? ssn4 + '|' + aeLN : aeLN + '|' + aeFN;

  let ppsRowIdx = -1;
  for (let p = 0; p < ppsData.length; p++) {
    const ln = String(ppsLNCol !== -1 ? ppsData[p][ppsLNCol] : '').toLowerCase().trim();
    const fn = String(ppsFNCol !== -1 ? ppsData[p][ppsFNCol] : '').toLowerCase().trim();
    const s4 = ssnLast4Col !== -1 ? String(ppsData[p][ssnLast4Col] || '').trim() : '';
    const key = s4 ? s4 + '|' + ln : ln + '|' + fn;
    if (key === matchKey) { ppsRowIdx = p; break; }
  }
  let action = 'updated';
  if (ppsRowIdx === -1) {
    ppsData.push(new Array(ppsLastCol).fill(''));
    ppsRowIdx = ppsData.length - 1;
    action = 'inserted';
  }
  const row = ppsData[ppsRowIdx];

  DEMO_MAP.forEach(({ src, dst }) => {
    const srcIdx = aeColMap[src];
    const dstIdx = findCol([dst]);
    if (srcIdx === undefined || dstIdx === -1) return;
    row[dstIdx] = aeRow[srcIdx];
  });
  if (ssnLast4Col !== -1 && ssn4) row[ssnLast4Col] = ssn4;

  const effDateCol = findCol(['Effective Date']);
  if (effDateCol !== -1) {
    const issueCols = ['AE Issue Date','DI Issue Date','CI Issue Date','HI Issue Date',
                       'Life Issue Date','TL Issue Date','TL Issue Date 2','TL Issue Date 3'];
    for (const c of issueCols) {
      if (aeColMap[c] === undefined) continue;
      const v = aeRow[aeColMap[c]];
      if (v && String(v).trim() !== '') { row[effDateCol] = v; break; }
    }
  }
  const pfCol = findCol(['Pay Frequency']);
  const dfCol = findCol(['Deduction Frequency']);
  if (pfCol !== -1) row[pfCol] = payFreq;
  if (dfCol !== -1) row[dfCol] = payFreq;

  if (currentFlagStatus === 'customized') {
    const clearTargets = ['Accident After Tax','Critical Illness After Tax','Hospital Indemnity After Tax',
                          'Whole Life After Tax','Short Term Disability After Tax','Disability After Tax',
                          'Term Life 10 Yr After Tax','Term Life 20 Yr After Tax','Term Life To Age 70 After Tax',
                          'Term Life After Tax','Total After Tax Allotment Per Pay Period','Total Spent Per Pay Period'];
    clearTargets.forEach(t => { const idx = findCol([t]); if (idx !== -1) row[idx] = ''; });
    wsPPS.getRange(PPS_DATA_START + ppsRowIdx, 1, 1, ppsLastCol).setValues([row]);
    return { action, ppsRow: PPS_DATA_START + ppsRowIdx, totalSpent: 0 };
  }

  const factor = 12 / payFreq;
  const PREM_MAP = (savedMap && Array.isArray(savedMap.premium) && savedMap.premium.length)
    ? savedMap.premium.map(p => ({ src: p.src, dsts: [p.dst] }))
    : [
        { src: 'AE Premium Amount',   dsts: ['Accident After Tax'] },
        { src: 'CI Premium Amount',   dsts: ['Critical Illness After Tax'] },
        { src: 'HI Premium Amount',   dsts: ['Hospital Indemnity After Tax'] },
        { src: 'Life Premium Amount', dsts: ['Whole Life After Tax'] },
        { src: 'DI Premium Amount',   dsts: ['Short Term Disability After Tax','Disability After Tax'] },
      ];
  let totalSpent = 0;
  PREM_MAP.forEach(({ src, dsts }) => {
    const srcIdx = aeColMap[src];
    const dstIdx = findCol(dsts);
    if (srcIdx === undefined || dstIdx === -1) return;
    const v = parseFloat(aeRow[srcIdx]) || 0;
    if (v > 0) {
      const pv = Math.round(v * factor * 100) / 100;
      row[dstIdx] = pv;
      totalSpent += pv;
    } else row[dstIdx] = '';
  });

  let tlSpecific, tlConsol;
  if (savedMap && Array.isArray(savedMap.tl) && savedMap.tl.length) {
    const consol = savedMap.tl.find(t => t.slot === -1 && t.dst);
    if (consol) { tlSpecific = [-1, -1, -1]; tlConsol = findCol([consol.dst]); }
    else {
      tlSpecific = [0, 1, 2].map(slot => {
        const e = savedMap.tl.find(t => t.slot === slot && t.dst);
        return e ? findCol([e.dst]) : -1;
      });
      tlConsol = -1;
    }
  } else {
    tlSpecific = [
      findCol(['Term Life 10 Yr After Tax']),
      findCol(['Term Life 20 Yr After Tax']),
      findCol(['Term Life To Age 70 After Tax']),
    ];
    tlConsol = findCol(['Term Life After Tax']);
  }
  const tlSrcCols = [aeColMap['TL Premium'], aeColMap['TL Premium 2'], aeColMap['TL Premium 3']];

  if (tlSpecific.some(c => c !== -1)) {
    for (let slot = 0; slot < 3; slot++) {
      if (tlSpecific[slot] === -1 || tlSrcCols[slot] === undefined) continue;
      const v = parseFloat(aeRow[tlSrcCols[slot]]) || 0;
      if (v > 0) {
        const pv = Math.round(v * factor * 100) / 100;
        row[tlSpecific[slot]] = pv;
        totalSpent += pv;
      } else row[tlSpecific[slot]] = '';
    }
  } else if (tlConsol !== -1) {
    let written = false;
    for (let slot = 0; slot < 3; slot++) {
      if (tlSrcCols[slot] === undefined) continue;
      const v = parseFloat(aeRow[tlSrcCols[slot]]) || 0;
      if (v > 0) {
        const pv = Math.round(v * factor * 100) / 100;
        row[tlConsol] = pv;
        totalSpent += pv;
        written = true;
        break;
      }
    }
    if (!written) row[tlConsol] = '';
  }

  const allotCol = findCol(['Total After Tax Allotment Per Pay Period']);
  const spentCol = findCol(['Total Spent Per Pay Period']);
  if (allotCol !== -1 && aeColMap['Allotments'] !== undefined) {
    const allot = parseFloat(aeRow[aeColMap['Allotments']]) || 0;
    row[allotCol] = allot > 0 ? Math.round(allot * factor * 100) / 100 : '';
  }
  if (spentCol !== -1) {
    row[spentCol] = totalSpent > 0 ? Math.round(totalSpent * 100) / 100 : '';
  }

  wsPPS.getRange(PPS_DATA_START + ppsRowIdx, 1, 1, ppsLastCol).setValues([row]);
  if (ssnLast4Col !== -1) {
    wsPPS.getRange(PPS_DATA_START + ppsRowIdx, ssnLast4Col + 1, 1, 1).setNumberFormat('@');
  }
  return { action, ppsRow: PPS_DATA_START + ppsRowIdx, totalSpent };
}

function testSyncCustomizedEmployee() {
  const samplePayload = {
    secret: PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET') || '',
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
    match: { contactId: '', email: 'jane@example.com', firstName: 'Jane', lastName: 'Smith' },
    benefits: {
      accident:   { on: true,  tier: 'employee', plan: 'plan1', premium: 12.72 },
      critical:   { on: true,  tier: 'employee', benefit: 25000, premium: 10.83 },
      hospital:   { on: false, premium: 0 },
      disability: { on: true,  weekly: 500, period: '26wk', premium: 41.53 },
      term:       { on: true,  tier: 'employee', termType: 'term_10', benefit: 50000, premium: 8.48 },
      wholeEmp:   { on: true,  benefit: 25000, childRider: 0, premium: 21.73 },
      wholeSpouse:{ on: false, spouseAge: 35, benefit: 0, premium: 0 },
    },
    benefitSummary: 'Accident Expense\nCritical Illness 25k',
    totalMonthly: 95.29,
  };
  const out = syncCustomizedEmployee(samplePayload);
  Logger.log(JSON.stringify(out, null, 2));
}

// ============================================================
// MISSING CONSTANTS FOR NAME CASING & SYNC ENGINE
// ============================================================
const UPPER_SUFFIXES  = ['JR', 'SR', 'II', 'III', 'IV', 'V', 'MD', 'DDS', 'PHD', 'CPA'];
const CASED_SUFFIXES  = ['PhD', 'VP'];
const LOWER_PARTICLES = ['of', 'the', 'von', 'van', 'de', 'di', 'da', 'le', 'del'];

// Maps your main Demographic tab headers to your internal downstream Employees tab headers
const FIELD_MAP = [
  { src: 'Employee SSN',     dst: 'Employee SSN' },
  { src: 'First Name',       dst: 'First Name' },
  { src: 'Last Name',        dst: 'Last Name' },
  { src: 'DOB',              dst: 'Date of Birth' },
  { src: 'Gender',           dst: 'Gender' },
  { src: 'Mailing Address1', dst: 'Address' },
  { src: 'City',             dst: 'City' },
  { src: 'State',            dst: 'State' },
  { src: 'Zip',              dst: 'Zip' },
  { src: 'Email Address',    dst: 'Email' },
  { src: 'Phone Number',     dst: 'Phone' },
  { src: 'Date of Hire',     dst: 'Date of Hire' }
];

// ============================================================
// SYSTEM DATA PARSING ENGINES
// ============================================================
const STD_SALARY_CAP = [
  [8750,   100], [11000,  125], [13000,  150], [15250,  175], [17500,  200],
  [19500,  225], [21750,  250], [24000,  275], [26000,  300], [28250,  325],
  [30500,  350], [32500,  375], [34750,  400], [37000,  425], [39000,  450],
  [41250,  475], [43500,  500], [45500,  525], [47750,  550], [50000,  575],
  [52000,  600], [54250,  625], [56500,  650], [58500,  675], [60750,  700],
  [63000,  725], [65000,  750], [67250,  775], [69500,  800], [71500,  825],
  [73750,  850], [76000,  875], [78000,  900], [80250,  925], [82500,  950],
  [84500,  975], [86750, 1000]
];

function showDemoSidebar() {
  const html = HtmlService.createHtmlOutput(getDemoHTML())
    .setTitle('Demo Import')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getDemoHTML() {
  return `
  <style>
    body { font-family: Google Sans, Roboto, Arial, sans-serif; padding: 12px; font-size: 13px; }
    h3  { margin-top: 0; color: #1a73e8; }
    h4  { color: #202124; margin: 16px 0 4px 0; }
    textarea {
      width: 100%; height: 180px; font-family: monospace; font-size: 11px;
      border: 1px solid #dadce0; border-radius: 8px; padding: 8px;
      resize: vertical; box-sizing: border-box;
    }
    .btn {
      background: #1a73e8; color: white; border: none; padding: 10px 24px;
      border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 12px; width: 100%;
    }
    .btn:hover    { background: #1557b0; }
    .btn:disabled { background: #dadce0; cursor: not-allowed; }
    .info  { color: #5f6368; font-size: 11px; margin: 4px 0; }
    .count { font-weight: bold; color: #1a73e8; }
    #status { margin-top: 12px; padding: 10px; border-radius: 4px; display: none; font-size: 12px; }
    .success { background: #e6f4ea; color: #137333; }
    .error   { background: #fce8e6; color: #c5221f; }
    .divider { border-top: 1px solid #e0e0e0; margin: 16px 0; }
  </style>

  <h3>Demo Import</h3>
  <h4>1. Savings Data</h4>
  <textarea id="savings" placeholder="First Name&#9;Last Name&#9;$00.00"></textarea>
  <p class="info">Rows: <span class="count" id="savingsCount">0</span></p>

  <div class="divider"></div>

  <h4>2. Salary Data</h4>
  <textarea id="salary" placeholder="Last Name&#9;First Name&#9;$00,000.00"></textarea>
  <p class="info">Rows: <span class="count" id="salaryCount">0</span></p>

  <button class="btn" id="submitBtn" onclick="runImport()">Run Demo Import</button>
  <div id="status"></div>

  <script>
    function countRows(id, countId) {
      document.getElementById(id).addEventListener('input', function() {
        const lines = this.value.trim().split('\\n').filter(l => l.trim());
        document.getElementById(countId).textContent = lines.length;
      });
    }
    countRows('savings', 'savingsCount');
    countRows('salary',  'salaryCount');

    function runImport() {
      const btn    = document.getElementById('submitBtn');
      const status = document.getElementById('status');
      btn.disabled    = true;
      btn.textContent = 'Running import...';
      status.style.display = 'none';

      google.script.run
        .withSuccessHandler(function(msg) {
          status.className = 'success';
          status.innerHTML = msg;
          status.style.display = 'block';
          btn.textContent = 'Done \u2713';
        })
        .withFailureHandler(function(err) {
          status.className = 'error';
          status.textContent = 'Error: ' + err.message;
          status.style.display = 'block';
          btn.disabled    = false;
          btn.textContent = 'Run Demo Import';
        })
        .runDemoImportFromSidebar(
          document.getElementById('savings').value,
          document.getElementById('salary').value
        );
    }
  </script>`;
}

function runDemoImportFromSidebar(savingsRawText, salaryRawText) {
  return executeDemoImport(savingsRawText, salaryRawText);
}

function executeDemoImport(savingsRawText, salaryRawText) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const wsDemo = ss.getSheetByName('Demographic');
  const wsAE   = ss.getSheetByName('Auto Enroll Sheet');

  if (!wsDemo) throw new Error('Missing "Demographic" tab.');
  if (!wsAE)   throw new Error('Missing "Auto Enroll Sheet" tab.');

  const demoHeaders = wsDemo.getRange(1, 1, 1, wsDemo.getLastColumn()).getValues()[0];
  const demoColMap  = buildColumnMap(demoHeaders);

  const requiredDemo = [
    'Employee SSN', 'First Name', 'Last Name', 'DOB', 'Gender',
    'Mailing Address1', 'City', 'State', 'Zip',
    'Email Address', 'Phone Number', 'Date of Hire'
  ];
  const missingDemo = requiredDemo.filter(h => demoColMap[h] === undefined);
  if (missingDemo.length) throw new Error('Demographic missing columns: ' + missingDemo.join(', '));

  const demoLastRow = wsDemo.getLastRow();
  if (demoLastRow < 2) throw new Error('No data rows on Demographic sheet.');
  const demoData = wsDemo.getRange(2, 1, demoLastRow - 1, wsDemo.getLastColumn()).getValues();

  const demoSSNCol      = demoColMap['Employee SSN'];
  const demoFirstCol    = demoColMap['First Name'];
  const demoLastNameCol = demoColMap['Last Name'];

  for (let r = 0; r < demoData.length; r++) {
    demoData[r][demoFirstCol]    = toProperCase(String(demoData[r][demoFirstCol]    || ''));
    demoData[r][demoLastNameCol] = toProperCase(String(demoData[r][demoLastNameCol] || ''));
  }

  const demoFields = [
    'First Name', 'Last Name', 'DOB', 'Gender',
    'Mailing Address1', 'City', 'State', 'Zip',
    'Email Address', 'Phone Number', 'Date of Hire'
  ];
  const ssnQueue = [], ssnSeen = {}, demoBySSN = {};
  for (let r = 0; r < demoData.length; r++) {
    const ssn = normalizeSSN(String(demoData[r][demoSSNCol] || ''));
    if (ssn.length === 9 && !ssnSeen[ssn]) {
      ssnSeen[ssn] = true;
      ssnQueue.push(ssn);
      const record = {};
      demoFields.forEach(f => { record[f] = demoData[r][demoColMap[f]]; });
      demoBySSN[ssn] = record;
    }
  }

  const aeHeaders = wsAE.getRange(1, 1, 1, wsAE.getLastColumn()).getValues()[0];
  const aeColMap  = buildColumnMap(aeHeaders);

  const requiredAE = [
    'Relation', 'Employee SSN', 'Insured SSN',
    'First Name', 'Last Name', 'DOB', 'Gender',
    'Mailing Address1', 'City', 'State', 'Zip',
    'Email Address', 'Phone Number', 'Date of Hire',
    'Allotments', 'Annual Salary'
  ];
  const missingAE = requiredAE.filter(h => aeColMap[h] === undefined);
  if (missingAE.length) throw new Error('Auto Enroll Sheet missing columns: ' + missingAE.join(', '));

  const aeLastRow = wsAE.getLastRow();
  if (aeLastRow < 2) throw new Error('No data rows on Auto Enroll Sheet.');
  const aeData = wsAE.getRange(2, 1, aeLastRow - 1, aeHeaders.length).getValues();

  const relCol    = aeColMap['Relation'];
  const eeSSNCol  = aeColMap['Employee SSN'];
  const insSSNCol = aeColMap['Insured SSN'];
  let ssnIdx = 0, ssnAssigned = 0;

  for (let i = 0; i < aeData.length; i++) {
    if (String(aeData[i][relCol] || '').toUpperCase().trim() !== 'EE') continue;
    if (ssnIdx >= ssnQueue.length) break;
    const ssn = ssnQueue[ssnIdx], formatted = formatSSN(ssn);
    aeData[i][eeSSNCol]  = formatted;
    aeData[i][insSSNCol] = formatted;
    const record = demoBySSN[ssn];
    if (record) {
      demoFields.forEach(f => {
        if (aeColMap[f] !== undefined) aeData[i][aeColMap[f]] = record[f];
      });
    }
    ssnIdx++; ssnAssigned++;
  }

  const savingsMap    = parsePastedData(savingsRawText, 'first-last-value');
  const savingsParsed = Object.keys(savingsMap).length;
  let savingsMatchCount = 0;

  const aeFirstCol      = aeColMap['First Name'];
  const aeLastNameCol   = aeColMap['Last Name'];
  const aeAllotmentsCol = aeColMap['Allotments'];

  for (let i = 0; i < aeData.length; i++) {
    if (String(aeData[i][relCol] || '').toUpperCase().trim() !== 'EE') continue;
    const key = normalizeNameKey(String(aeData[i][aeFirstCol] || ''), String(aeData[i][aeLastNameCol] || ''));
    if (savingsMap[key] !== undefined) { aeData[i][aeAllotmentsCol] = savingsMap[key]; savingsMatchCount++; }
  }

  const salaryMap    = parsePastedData(salaryRawText, 'last-first-value');
  const salaryParsed = Object.keys(salaryMap).length;
  let salaryMatchCount = 0;

  const aeAnnualSalaryCol = aeColMap['Annual Salary'];

  for (let i = 0; i < aeData.length; i++) {
    if (String(aeData[i][relCol] || '').toUpperCase().trim() !== 'EE') continue;
    const key = normalizeNameKey(String(aeData[i][aeFirstCol] || ''), String(aeData[i][aeLastNameCol] || ''));
    if (salaryMap[key] !== undefined) { aeData[i][aeAnnualSalaryCol] = salaryMap[key]; salaryMatchCount++; }
  }

  wsAE.getRange(2, 1, aeData.length, aeHeaders.length).setValues(aeData);

  let demoLastCol = wsDemo.getLastColumn();
  if (demoColMap['Net Tax Savings'] === undefined) {
    demoLastCol++;
    wsDemo.getRange(1, demoLastCol).setValue('Net Tax Savings');
    demoColMap['Net Tax Savings'] = demoLastCol - 1;
  }
  if (demoColMap['Annual Salary'] === undefined) {
    demoLastCol++;
    wsDemo.getRange(1, demoLastCol).setValue('Annual Salary');
    demoColMap['Annual Salary'] = demoLastCol - 1;
  }
  for (let r = 0; r < demoData.length; r++) {
    while (demoData[r].length < demoLastCol) demoData[r].push('');
  }
  for (let r = 0; r < demoData.length; r++) {
    const key = normalizeNameKey(String(demoData[r][demoFirstCol] || ''), String(demoData[r][demoLastNameCol] || ''));
    if (savingsMap[key] !== undefined) demoData[r][demoColMap['Net Tax Savings']] = savingsMap[key];
    if (salaryMap[key]  !== undefined) demoData[r][demoColMap['Annual Salary']]   = salaryMap[key];
  }
  wsDemo.getRange(2, 1, demoData.length, demoLastCol).setValues(demoData);

  return '<b>Demo Import Complete</b><br>' +
    '• AE rows populated (SSN): ' + ssnAssigned      + '<br>' +
    '• Allotments matched: '      + savingsMatchCount + ' of ' + savingsParsed + ' parsed<br>' +
    '• Annual Salary matched: '   + salaryMatchCount  + ' of ' + salaryParsed  + ' parsed';
}

function runAutoPlanBuilder() {
  const ui = SpreadsheetApp.getUi();

  const rateIdResponse = ui.prompt('Rate Tables Sheet ID', 'Enter Rate Tables workbook ID:', ui.ButtonSet.OK_CANCEL);
  if (rateIdResponse.getSelectedButton() !== ui.Button.OK) return;
  const rateSheetId = rateIdResponse.getResponseText().trim();

  let rateSS;
  try { rateSS = SpreadsheetApp.openById(rateSheetId); } catch (e) { ui.alert('Cannot close/open rate file.'); return; }

  const dateResponse = ui.prompt('Effective Date', 'Enter issue date (MM/DD/YYYY):', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const effectiveDate = new Date(dateResponse.getResponseText().trim());

  const signDateResponse = ui.prompt('Sign Date', 'Enter signature date (MM/DD/YYYY):', ui.ButtonSet.OK_CANCEL);
  if (signDateResponse.getSelectedButton() !== ui.Button.OK) return;
  const signDate = new Date(signDateResponse.getResponseText().trim());

  const companyResponse = ui.prompt('Company Name', 'Enter company name:', ui.ButtonSet.OK_CANCEL);
  if (companyResponse.getSelectedButton() !== ui.Button.OK) return;
  const companyName = companyResponse.getResponseText().trim();

  const cutoffResponse = ui.prompt('Enrollment Cutoff', 'Enter cutoff text/date:', ui.ButtonSet.OK_CANCEL);
  if (cutoffResponse.getSelectedButton() !== ui.Button.OK) return;
  const enrollmentCutoff = cutoffResponse.getResponseText().trim();

  const stdSheet = rateSS.getSheetByName('Short Term Disability');
  const stdOptions = detectSTDOptions(stdSheet);
  let selectedSTDOption = stdOptions[0];

  const enrollSS = SpreadsheetApp.getActiveSpreadsheet();
  const wsAE     = enrollSS.getSheetByName('Auto Enroll Sheet');
  const wsDemo   = enrollSS.getSheetByName('Demographic');

  let aeLastCol = wsAE.getLastColumn();
  let aeHeaders = wsAE.getRange(1, 1, 1, aeLastCol).getValues()[0];
  let colMap    = buildColumnMap(aeHeaders);

  if (colMap['Remaining Formula'] === undefined) {
    wsAE.getRange(1, aeHeaders.length + 1).setValue('Remaining Formula');
    aeHeaders = [...aeHeaders, 'Remaining Formula'];
    colMap    = buildColumnMap(aeHeaders);
  }

  const lastRow = wsAE.getLastRow();
  const dataRange = wsAE.getRange(2, 1, lastRow - 1, aeHeaders.length);
  const data      = dataRange.getValues();
  const rates = loadAllRateTables(rateSS, selectedSTDOption);

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const relation = String(row[colMap['Relation']] || '').toUpperCase().trim();

    if (relation === 'EE') {
      const dob = row[colMap['DOB']];
      row[colMap['Age']] = (dob instanceof Date) ? calcAge(dob, effectiveDate) : '';
    } else { row[colMap['Age']] = ''; }

    const age       = row[colMap['Age']];
    const allotment = parseFloat(row[colMap['Allotments']])     || 0;
    const salary    = parseFloat(row[colMap['Annual Salary']])  || 0;
    if (age === '' || allotment <= 0) continue;

    let usedPremium = 0;

    const tlResult = assignTermLife(age, allotment, rates);
    if (tlResult) {
      row[colMap['TL Plan Type']]       = tlResult.plan;
      row[colMap['TL Insured Option']]  = 'EE';
      row[colMap['TL Benefit Amount']]  = tlResult.benefit;
      row[colMap['TL Premium']]         = tlResult.rate;
    }
    normalizeTLRow(row, colMap);
    usedPremium = totalTLPremium(row, colMap);

    const wlResult = assignWholeLife(age, allotment - usedPremium, rates);
    if (wlResult) {
      row[colMap['Life Plan Type']]          = 'Whole Life';
      row[colMap['Life Insured Option']]     = 'EO';
      row[colMap['Life Certificate Amount']] = wlResult.benefit;
      row[colMap['Life Premium Amount']]     = wlResult.rate;
      usedPremium += wlResult.rate;
    }

    const aeResult = assignAccident(allotment - usedPremium, rates);
    if (aeResult) {
      row[colMap['AE Plan Type']]      = aeResult.plan;
      row[colMap['AE Insured Option']] = 'EE';
      row[colMap['AE Premium Amount']] = aeResult.rate;
      usedPremium += aeResult.rate;
    }

    const ciResult = assignCriticalIllness(age, allotment - usedPremium, rates);
    if (ciResult) {
      row[colMap['CI Plan Type']]      = 'Critical Tier 2';
      row[colMap['CI Insured Option']] = 'EE';
      row[colMap['CI Benefit Amount']] = ciResult.benefit;
      row[colMap['CI Premium Amount']] = ciResult.rate;
      usedPremium += ciResult.rate;
    }

    const hiResult = assignHospitalIndemnity(allotment - usedPremium, rates);
    if (hiResult) {
      row[colMap['HI Plan Type']]      = hiResult.plan;
      row[colMap['HI Insured Option']] = 'EE';
      row[colMap['HI Premium Amount']] = hiResult.rate;
      row[colMap['HI Benefit Amount']] = 1000;
      usedPremium += hiResult.rate;
    }

    const diResult = assignSTD(age, salary, allotment - usedPremium, rates);
    if (diResult) {
      row[colMap['DI Plan Type']]      = selectedSTDOption.label;
      row[colMap['DI Benefit Amount']] = diResult.benefit;
      row[colMap['DI Premium Amount']] = diResult.rate;
      usedPremium += diResult.rate;
    }

    const getVal = (col) => (col !== undefined ? parseFloat(row[col]) || 0 : 0);
    row[colMap['Remaining Formula']] =
      getVal(colMap['Allotments'])          - getVal(colMap['TL Premium 3'])        -
      getVal(colMap['TL Premium 2'])        - getVal(colMap['TL Premium'])          -
      getVal(colMap['Life Premium Amount']) - getVal(colMap['HI Premium Amount'])   -
      getVal(colMap['CI Premium Amount'])   - getVal(colMap['DI Premium Amount'])   -
      getVal(colMap['AE Premium Amount']);

    populatePolicyDates(row, colMap, effectiveDate, signDate);
  }

  dataRange.setValues(data);
  writeGroupMetadataToDemographic(wsDemo, companyName, enrollmentCutoff);
  ui.alert('Auto Plan Builder Engine Run Finished Successfully.');
}

function writeGroupMetadataToDemographic(wsDemo, companyName, enrollmentCutoff) {
  let lastCol    = wsDemo.getLastColumn();
  let headers    = wsDemo.getRange(1, 1, 1, lastCol).getValues()[0];
  let demoColMap = buildColumnMap(headers);

  const ensureCol = (name) => {
    if (demoColMap[name] === undefined) {
      lastCol++;
      wsDemo.getRange(1, lastCol).setValue(name);
      demoColMap[name] = lastCol - 1;
    }
  };
  ensureCol('Company Name');
  ensureCol('Enrollment Cutoff');

  const lastRow = wsDemo.getLastRow();
  if (lastRow < 2) return;

  const demoData = wsDemo.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let r = 0; r < demoData.length; r++) {
    demoData[r][demoColMap['Company Name']]      = companyName;
    demoData[r][demoColMap['Enrollment Cutoff']] = enrollmentCutoff;
  }
  wsDemo.getRange(2, 1, demoData.length, lastCol).setValues(demoData);
}
function runBenefitPackageBuilder() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsAE   = ss.getSheetByName('Auto Enroll Sheet');
  const wsDemo = ss.getSheetByName('Demographic');

  let aeLastCol = wsAE.getLastColumn();
  let aeHeaders = wsAE.getRange(1, 1, 1, aeLastCol).getValues()[0];
  let aeColMap  = buildColumnMap(aeHeaders);
  if (aeColMap['Benefit Package'] === undefined) {
    wsAE.getRange(1, aeHeaders.length + 1).setValue('Benefit Package');
    aeHeaders = [...aeHeaders, 'Benefit Package'];
    aeColMap  = buildColumnMap(aeHeaders);
  }

  let demoLastCol = wsDemo.getLastColumn();
  let demoHeaders = wsDemo.getRange(1, 1, 1, demoLastCol).getValues()[0];
  let demoColMap  = buildColumnMap(demoHeaders);

  const ensureDemoCol = (name) => {
    if (demoColMap[name] === undefined) {
      demoLastCol++;
      wsDemo.getRange(1, demoLastCol).setValue(name);
      demoColMap[name] = demoLastCol - 1;
    }
  };
  ensureDemoCol('Benefits Package');
  ensureDemoCol('Age');
  ensureDemoCol('Savings');

  const ssnColDemo = demoColMap['Employee SSN'];
  const lastRowDemo = wsDemo.getLastRow();
  const demoData    = wsDemo.getRange(2, 1, lastRowDemo - 1, demoLastCol).getValues();
  const ssnToRow    = {};
  for (let r = 0; r < demoData.length; r++) {
    const ssn = normalizeSSN(String(demoData[r][ssnColDemo] || '')).padStart(9, '0');
    if (ssn) ssnToRow[ssn] = r;
  }

  const ssnColAE    = aeColMap['Insured SSN'];
  const ageColAE    = aeColMap['Age'];
  const remColAE    = aeColMap['Remaining Formula'];
  const bpColAE     = aeColMap['Benefit Package'];
  const AEPlanCol    = aeColMap['AE Plan Type'];
  const CIPlanCol    = aeColMap['CI Plan Type'];
  const CIBenCol     = aeColMap['CI Benefit Amount'];
  const HIPlanCol    = aeColMap['HI Plan Type'];
  const DIPlanCol    = aeColMap['DI Plan Type'];
  const DIBenCol     = aeColMap['DI Benefit Amount'];
  const LifePlanCol  = aeColMap['Life Plan Type'];
  const LifeCertCol  = aeColMap['Life Certificate Amount'];
  const tlPlanCols   = [aeColMap['TL Plan Type'],      aeColMap['TL Plan Type 2'],       aeColMap['TL Plan Type 3']];
  const tlAmtCols    = [aeColMap['TL Benefit Amount'],  aeColMap['TL Benefit Amount 2'],  aeColMap['TL Benefit Amount 3']];
  
  const lastRowAE = wsAE.getLastRow();
  const aeData    = wsAE.getRange(2, 1, lastRowAE - 1, aeHeaders.length).getValues();
  
  for (let i = 0; i < aeData.length; i++) {
    const row = aeData[i];

    // FIX: Removed relation !== 'EE' check. Builds text blobs for all rows uniformly.
    const parts = [];
    if (AEPlanCol !== undefined && row[AEPlanCol]) parts.push('Accident Expense');
    if (CIPlanCol !== undefined && row[CIPlanCol]) {
      const amt = formatK(CIBenCol !== undefined ? row[CIBenCol] : '');
      if (amt) parts.push('Critical Illness ' + amt);
    }
    if (HIPlanCol !== undefined && row[HIPlanCol]) parts.push('Hospital Indemnity');
    if (DIPlanCol !== undefined && row[DIPlanCol]) {
      const ben = parseFloat(DIBenCol !== undefined ? row[DIBenCol] : 0);
      if (ben) parts.push('Short Term Disability ' + ben + ' Weekly');
    }
    for (let j = 0; j < tlPlanCols.length; j++) {
      if (tlPlanCols[j] !== undefined && row[tlPlanCols[j]]) {
        const amt = formatK(tlAmtCols[j] !== undefined ? row[tlAmtCols[j]] : '');
        if (amt) parts.push('Term Life ' + row[tlPlanCols[j]] + ' ' + amt);
        break;
      }
    }
    if (LifePlanCol !== undefined && row[LifePlanCol]) {
      const amt = formatK(LifeCertCol !== undefined ? row[LifeCertCol] : '');
      if (amt) parts.push('Whole Life ' + amt);
    }

    const summary = parts.join('\n');
    row[bpColAE] = summary;
    const ssn        = normalizeSSN(String(ssnColAE !== undefined ? row[ssnColAE] : '')).padStart(9, '0');
    const demoRowIdx = ssnToRow[ssn];

    if (demoRowIdx !== undefined) {
      demoData[demoRowIdx][demoColMap['Benefits Package']] = summary;
      if (ageColAE !== undefined) demoData[demoRowIdx][demoColMap['Age']]     = row[ageColAE];
      if (remColAE !== undefined) demoData[demoRowIdx][demoColMap['Savings']] = row[remColAE];
    }
  }

  wsAE.getRange(2, 1, aeData.length, aeHeaders.length).setValues(aeData);
  wsDemo.getRange(2, 1, demoData.length, demoLastCol).setValues(demoData);
}

function detectSTDOptions(stdSheet) {
  const allData = stdSheet.getDataRange().getValues();
  const options = [];
  for (let r = 0; r < allData.length; r++) {
    const cellA = String(allData[r][0] || '');
    if (cellA.toLowerCase().includes('group disability income') && cellA.toLowerCase().includes('benefit period')) {
      const match     = cellA.match(/Benefit Period:\s*(\d+\s*Weeks?)/i);
      const elimMatch = cellA.match(/Elimination:\s*([^\s|]+(?:\s+days?)?)/i);
      let label = match ? match[1] : 'Unknown';
      if (elimMatch) label += ' | Elim: ' + elimMatch[1];
      options.push({ label: label, headerRow: r + 1, dataStartRow: r + 2, rawData: allData });
    }
  }
  return options;
}

function loadAllRateTables(rateSS, selectedSTDOption) {
  const rates = {};

  ['10 Year', '20 Year', 'To Age 70'].forEach(name => {
    const ws = rateSS.getSheetByName(name);
    if (!ws) return;
    const d        = ws.getRange(2, 1, 11, 5).getValues();
    const benefits = [];
    for (let c = 1; c < d[0].length; c++) benefits.push(parseBenefitAmount(d[0][c]));
    const bands = [];
    for (let r = 1; r < d.length; r++) {
      const bandRates = [];
      for (let c = 1; c < d[r].length; c++) {
        const v = d[r][c];
        bandRates.push(isValidRate(v) ? parseFloat(v) : null);
      }
      bands.push({ band: String(d[r][0]), rates: bandRates });
    }
    rates[name] = { benefits, bands };
  });

  const wlSheet = rateSS.getSheetByName('Whole Life');
  if (wlSheet) {
    const wlFullWidth = wlSheet.getLastColumn();
    const wlAllRows   = wlSheet.getRange(2, 1, wlSheet.getLastRow() - 1, wlFullWidth).getValues();

    let wlColCount = 1;
    while (wlColCount < wlAllRows[0].length) {
      if (parseBenefitAmount(wlAllRows[0][wlColCount]) <= 0) break;
      wlColCount++;
    }

    const wlBenefits = [];
    for (let c = 1; c < wlColCount; c++) wlBenefits.push(parseBenefitAmount(wlAllRows[0][c]));

    const wlLookup = {};
    let wlDataStarted = false;
    for (let r = 1; r < wlAllRows.length; r++) {
      const a = parseInt(wlAllRows[r][0]);
      if (isNaN(a)) { if (wlDataStarted) break; continue; }
      wlDataStarted = true;
      const ratesArr = [];
      for (let c = 1; c < wlColCount; c++) {
        const v = wlAllRows[r][c];
        ratesArr.push(isValidRate(v) ? parseFloat(v) : null);
      }
      wlLookup[a] = ratesArr;
    }
    rates['Whole Life'] = { benefits: wlBenefits, lookup: wlLookup };
  }

  const accSheet = rateSS.getSheetByName('Accident');
  if (accSheet) {
    const accData = accSheet.getRange(3, 1, 2, 2).getValues();
    rates['Accident'] = {
      plan1: { rate: parseFloat(accData[0][1]) || 0 },
      plan2: { rate: parseFloat(accData[1][1]) || 0 }
    };
  }

  const hiSheet = rateSS.getSheetByName('Hospital Indemnity');
  if (hiSheet) {
    const hiData = hiSheet.getRange(3, 1, 2, 2).getValues();
    rates['Hospital Indemnity'] = {
      plan1: { rate: parseFloat(hiData[0][1]) || 0 },
      plan2: { rate: parseFloat(hiData[1][1]) || 0 }
    };
  }

  const ciSheet = rateSS.getSheetByName('Critical Illness');
  if (ciSheet) {
    const ciData     = ciSheet.getRange(2, 1, ciSheet.getLastRow() - 1, ciSheet.getLastColumn()).getValues();
    const ciBenefits = [];
    for (let c = 1; c < ciData[0].length; c++) ciBenefits.push(parseBenefitAmount(ciData[0][c]));
    const ciBands = [];
    for (let r = 1; r < ciData.length; r++) {
      const bandRates = [];
      for (let c = 1; c < ciData[r].length; c++) {
        const v = ciData[r][c];
        bandRates.push(isValidRate(v) ? parseFloat(v) : null);
      }
      ciBands.push({ band: String(ciData[r][0]), rates: bandRates });
    }
    rates['Critical Illness'] = { benefits: ciBenefits, bands: ciBands };
  }

  const stdAllData   = selectedSTDOption.rawData;
  const stdHeaderRow = selectedSTDOption.headerRow;
  const stdDataStart = selectedSTDOption.dataStartRow;

  const stdBenefits = [];
  for (let c = 1; c < stdAllData[stdHeaderRow].length; c++) {
    const val = String(stdAllData[stdHeaderRow][c] || '');
    const m   = val.match(/\$?([\d,]+)/);
    if (m) stdBenefits.push(parseInt(m[1].replace(/,/g, '')));
  }

  const stdBands = [];
  for (let r = stdDataStart; r < stdAllData.length; r++) {
    const bandVal = String(stdAllData[r][0] || '').trim();
    if (!bandVal || bandVal.toLowerCase().includes('group disability')) break;
    const bandRates = [];
    for (let c = 1; c <= stdBenefits.length; c++) {
      const v = stdAllData[r][c];
      bandRates.push(isValidRate(v) ? parseFloat(v) : null);
    }
    stdBands.push({ band: bandVal, rates: bandRates });
  }
  rates['STD'] = { benefits: stdBenefits, bands: stdBands };

  return rates;
}

function assignTermLife(age, allotment, rates) {
  const priority = age < 50 ? ['To Age 70', '20 Year', '10 Year'] : ['20 Year', '10 Year'];
  for (const planName of priority) {
    const table = rates[planName];
    if (!table) continue;
    const result = findBestBenefit(age, allotment, table);
    if (result) return { plan: planName, benefit: result.benefit, rate: result.rate };
  }
  return null;
}

function assignWholeLife(age, remaining, rates) {
  const wl = rates['Whole Life'];
  if (!wl || remaining <= 0) return null;
  const ageRates = wl.lookup[age];
  if (!ageRates) return null;
  let bestBenefit = 0, bestRate = 0;
  for (let c = 0; c < ageRates.length; c++) {
    const rate = ageRates[c];
    if (rate !== null && rate > 0 && remaining >= rate && wl.benefits[c] > bestBenefit) {
      bestBenefit = wl.benefits[c];
      bestRate    = rate;
    }
  }
  return bestBenefit > 0 ? { benefit: bestBenefit, rate: bestRate } : null;
}

function assignAccident(remaining, rates) {
  const acc = rates['Accident'];
  if (!acc || remaining <= 0) return null;
  if (remaining >= acc.plan2.rate && acc.plan2.rate > 0) return { plan: 'Accident Tier 4', rate: acc.plan2.rate };
  if (remaining >= acc.plan1.rate && acc.plan1.rate > 0) return { plan: 'Accident Tier 2', rate: acc.plan1.rate };
  return null;
}

function assignCriticalIllness(age, remaining, rates) {
  const ci = rates['Critical Illness'];
  if (!ci || remaining <= 0) return null;
  return findBestBenefit(age, remaining, ci);
}

function assignHospitalIndemnity(remaining, rates) {
  const hi = rates['Hospital Indemnity'];
  if (!hi || remaining <= 0) return null;
  if (remaining >= hi.plan2.rate && hi.plan2.rate > 0) return { plan: 'Plan 2', rate: hi.plan2.rate };
  if (remaining >= hi.plan1.rate && hi.plan1.rate > 0) return { plan: 'Plan 1', rate: hi.plan1.rate };
  return null;
}

function assignSTD(age, salary, remaining, rates) {
  const std = rates['STD'];
  if (!std || remaining <= 0) return null;

  let maxWeeklyBenefit = 0;
  for (const [salThreshold, weeklyBen] of STD_SALARY_CAP) {
    if (salary >= salThreshold) maxWeeklyBenefit = weeklyBen;
  }
  if (maxWeeklyBenefit === 0) return null;

  const bandIdx = findAgeBandIndex(age, std.bands);
  if (bandIdx === -1) return null;

  let bestBenefit = 0, bestRate = 0;
  for (let c = 0; c < std.benefits.length; c++) {
    const weeklyBen = std.benefits[c];
    const rate      = std.bands[bandIdx].rates[c];
    if (rate !== null && rate > 0 && remaining >= rate && weeklyBen <= maxWeeklyBenefit && weeklyBen > bestBenefit) {
      bestBenefit = weeklyBen;
      bestRate    = rate;
    }
  }
  return bestBenefit > 0 ? { benefit: bestBenefit, rate: bestRate } : null;
}

function normalizeTLRow(row, colMap) {
  const planType = String(row[colMap['TL Plan Type']] || '').toUpperCase().trim();
  const fields   = ['TL Plan Type', 'TL Insured Option', 'TL Benefit Amount', 'TL Premium', 'TL Issue Date', 'TL Signed Date'];

  if (planType === '20 YEAR') {
    fields.forEach(f => {
      const src = colMap[f], tgt = colMap[f + ' 2'];
      if (src !== undefined && tgt !== undefined) { row[tgt] = row[src]; row[src] = ''; }
    });
  } else if (planType === 'TO AGE 70') {
    fields.forEach(f => {
      const src = colMap[f], tgt = colMap[f + ' 3'];
      if (src !== undefined && tgt !== undefined) { row[tgt] = row[src]; row[src] = ''; }
    });
  }
}

function totalTLPremium(row, colMap) {
  return (parseFloat(row[colMap['TL Premium']])   || 0) +
         (parseFloat(row[colMap['TL Premium 2']]) || 0) +
         (parseFloat(row[colMap['TL Premium 3']]) || 0);
}

function populatePolicyDates(row, colMap, effectiveDate, signDate) {
  const dateFields = [
    { check: 'TL Insured Option',   issue: 'TL Issue Date',   signed: 'TL Signed Date'   },
    { check: 'TL Insured Option 2', issue: 'TL Issue Date 2', signed: 'TL Signed Date 2' },
    { check: 'TL Insured Option 3', issue: 'TL Issue Date 3', signed: 'TL Signed Date 3' },
    { check: 'Life Insured Option', issue: 'Life Issue Date', signed: 'Life Signed Date'  },
    { check: 'Spouse Life Insured Option', issue: 'Spouse Life Issue Date', signed: 'Spouse Life Signed Date' },
    { check: 'AE Insured Option',   issue: 'AE Issue Date',   signed: 'AE Signed Date'   },
    { check: 'CI Insured Option',   issue: 'CI Issue Date',   signed: 'CI Signed Date'   },
    { check: 'HI Insured Option',   issue: 'HI Issue Date',   signed: 'HI Signed Date'   },
    { check: 'DI Plan Type',        issue: 'DI Issue Date',   signed: 'DI Signed Date'   }
  ];

  dateFields.forEach(df => {
    if (colMap[df.check] !== undefined && row[colMap[df.check]]) {
      if (colMap[df.issue]  !== undefined) row[colMap[df.issue]]  = effectiveDate;
      if (colMap[df.signed] !== undefined) row[colMap[df.signed]] = signDate;
    }
  });
}

function toProperCase(str) {
  if (!str) return '';
  str = String(str).trim();
  const tokens = str.split(/(\s+|-)/);
  return tokens.map((token, idx) => {
    if (/^\s+$/.test(token) || token === '-') return token;
    const upper = token.toUpperCase();
    if (UPPER_SUFFIXES.indexOf(upper) !== -1) return upper;
    for (const suf of CASED_SUFFIXES) { if (upper === suf.toUpperCase()) return suf; }
    const lower = token.toLowerCase();
    if (idx > 0 && LOWER_PARTICLES.indexOf(lower) !== -1) return lower;
    if (token.length === 1) return upper;
    if (/^mc/i.test(token) && token.length > 2) return 'Mc' + token.charAt(2).toUpperCase() + token.slice(3).toLowerCase();
    if (/^o'/i.test(token) && token.length > 2) return "O'" + token.charAt(2).toUpperCase() + token.slice(3).toLowerCase();
    return token.charAt(0).toUpperCase() + token.slice(1).toLowerCase();
  }).join('');
}

function normalizeNameKey(first, last) { return (toProperCase(first) + ' ' + toProperCase(last)).trim().toLowerCase(); }
function buildColumnMap(headers) { const map = {}; headers.forEach((h, idx) => { map[String(h).trim()] = idx; }); return map; }
function normalizeSSN(s) { return s.replace(/[^0-9]/g, ''); }
function padSSN(s)       { return normalizeSSN(String(s)).padStart(9, '0'); }
function formatSSN(digits) { return digits.length === 9 ? digits.substring(0, 3) + '-' + digits.substring(3, 5) + '-' + digits.substring(5) : digits; }
function calcAge(dob, refDate) { let age = refDate.getFullYear() - dob.getFullYear(); const mDiff = refDate.getMonth() - dob.getMonth(); if (mDiff < 0 || (mDiff === 0 && refDate.getDate() < dob.getDate())) age--; return age; }
function parseBenefitAmount(val) { const s = String(val).replace(/[$,\/wk]/g, '').trim(); const kMatch = s.match(/([\d.]+)\s*K/i); if (kMatch) return parseFloat(kMatch[1]) * 1000; return parseFloat(s.replace(/[^0-9.]/g, '')) || 0; }
var isValidRate = (v) => v !== null && v !== undefined && v !== '' && String(v).toLowerCase() !== 'n/a' && !isNaN(parseFloat(v)) && parseFloat(v) > 0;
function formatK(amount) { const n = parseFloat(String(amount).replace(/[$,]/g, '').trim()); if (!n || isNaN(n)) return ''; const k = n / 1000; return Math.abs(k - Math.round(k)) < 1e-7 ? Math.round(k) + 'k' : k.toFixed(1) + 'k'; }
function findAgeBandIndex(age, bands) { for (let i = 0; i < bands.length; i++) { if (ageInBand(age, bands[i].band)) return i; } return -1; }

function ageInBand(age, band) {
  const b = String(band).trim();
  const underMatch = b.match(/Under\s+(\d+)/i); if (underMatch) return age < parseInt(underMatch[1]);
  const plusMatch  = b.match(/^(\d+)\s*\+$/); if (plusMatch) return age >= parseInt(plusMatch[1]);
  const rangeMatch = b.match(/(\d+)\s*[-–]\s*(\d+)/); if (rangeMatch) return age >= parseInt(rangeMatch[1]) && age <= parseInt(rangeMatch[2]);
  return false;
}

function findBestBenefit(age, budget, table) {
  const bandIdx = findAgeBandIndex(age, table.bands); if (bandIdx === -1) return null;
  let bestBenefit = 0, bestRate = 0;
  for (let c = 0; c < table.benefits.length; c++) {
    const rate = table.bands[bandIdx].rates[c];
    if (rate !== null && rate > 0 && budget >= rate && table.benefits[c] > bestBenefit) { bestBenefit = table.benefits[c]; bestRate = rate; }
  }
  return bestBenefit > 0 ? { benefit: bestBenefit, rate: bestRate } : null;
}

function parsePastedData(text, format) {
  const map = {}; if (!text) return map;
  for (const line of text.split('\n')) {
    const trimmed = line.trim(); if (!trimmed) continue;
    let parts;
    if (trimmed.indexOf('\t') !== -1)       parts = trimmed.split('\t');
    else if (trimmed.indexOf(',') !== -1)   parts = trimmed.split(',');
    else                                    parts = trimmed.split(/\s{2,}/);
    if (parts.length < 3) continue;
    let first, last, value;
    if (format === 'first-last-value') { first = parts[0].trim(); last = parts[1].trim(); value = parseNumericValue(parts[2].trim()); }
    else { last = parts[0].trim(); first = parts[1].trim(); value = parseNumericValue(parts[2].trim()); }
    const key = normalizeNameKey(first, last); if (key && value !== null) map[key] = value;
  }
  return map;
}
function parseNumericValue(str) { if (!str) return null; const num = parseFloat(str.replace(/[$,]/g, '').trim()); return isNaN(num) ? null : num; }

function standardPlanName(rawName, hdr) {
  const txt = rawName.toLowerCase(); const h = hdr.toLowerCase();
  if (txt.includes('accident'))   return 'Assurity Group Accident Expense';
  if (txt.includes('critical'))   return 'Assurity Group Critical Illness';
  if (txt.includes('whole life')) return 'Assurity Group Whole Life';
  if (txt.includes('to age 70'))  return 'Assurity Group Term To Age 70';
  if (txt.includes('10 year'))    return 'Assurity Group Term 10 year';
  if (txt.includes('20 year'))    return 'Assurity Group Term 20 year';
  if (h.includes('di plan type') || txt.includes('week')) return 'Assurity Group Disability Income';
  if (h.includes('hi plan type') && (txt.includes('plan 1') || txt.includes('plan 2'))) return 'Assurity Group Hospital Indemnity';
  return rawName;
}

var tierFromOption = (opt) => String(opt).toUpperCase() === 'EE' ? 'EO' : opt;
var findColByName = (headerRow, name) => headerRow.map(h => String(h).trim().toLowerCase()).indexOf(name.trim().toLowerCase());

function getApprovedSSNs(ss) {
  const wsDem = ss.getSheetByName('Demographic'); if (!wsDem) return null;
  const demData = wsDem.getDataRange().getValues(); if (demData.length < 2) return null;
  const ssnCol = findColByName(demData[0], 'Employee SSN'); if (ssnCol === -1) return null;

  const approved = {};
  for (let d = 1; d < demData.length; d++) {
    const flag = String(demData[d][0]).replace(/\s+/g, ' ').trim().toLowerCase();
    if (flag.includes('yes') || flag.includes('customized')) {
      const ssn = padSSN(String(demData[d][ssnCol])); if (ssn) approved[ssn] = true;
    }
  }
  return approved;
}

function stackAllPlans() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsSrc = ss.getSheetByName('Auto Enroll Sheet');
  let   wsDst = ss.getSheetByName('StackedPlans');
  if (!wsSrc) return;

  const approvedSSNs = getApprovedSSNs(ss); if (approvedSSNs === null) return;
  if (!wsDst) wsDst = ss.insertSheet('StackedPlans'); else wsDst.clearContents();
  wsDst.getRange('A:A').setNumberFormat('@');

  const stackHeaders = [
    'Employee SSN', 'EID', 'Location', 'Coverage Number', 'Plan name', 'Product name',
    'Coverage Tier', 'Coverage Option', 'Face Coverage Option', 'Benefit Amount', 'Benefit Frequency',
    'Face Amount', 'Policy Number', 'Policy Date', 'Deduction Date', 'Deduction End Date', 'Effective Date',
    'Termination Date', 'Application Date', 'Deduction Frequency', 'EE Cost', 'Pretax', 'ER Cost',
    'Imputed Income', 'EE Extra Cost', 'Change Pending', 'Option Pending', 'Face Option Pending',
    'Benefit Amount Pending', 'Face Amount Pending', 'Calculate', 'EE Cost Pending', 'ER Cost Pending',
    'EE Extra Cost Pending', 'Lock', 'Event Description', 'Event Code', 'Repair Effective Date'
  ];
  wsDst.getRange(1, 1, 1, stackHeaders.length).setValues([stackHeaders]);

  const srcData   = wsSrc.getDataRange().getValues();
  const headerRow = srcData[0];
  const ssnColSrc = findColByName(headerRow, 'Employee SSN');
  const output    = [];

  for (let r = 1; r < srcData.length; r++) {
    const empSSN = padSSN(String(ssnColSrc >= 0 ? srcData[r][ssnColSrc] : srcData[r][0]));
    if (!approvedSSNs[empSSN]) continue;

    for (let c = 0; c < headerRow.length; c++) {
      const hdr = String(headerRow[c]);
      if (!/Plan Type(\s+\d+)?$/.test(hdr) || /^Life Plan Type\s+[2-4]$/i.test(hdr)) continue;

      const planNameRaw = srcData[r][c]; if (!planNameRaw || String(planNameRaw).trim() === '') continue;
      const planPrefix = String(hdr).split(/\s/)[0];
      let eeCost = '', effDate = '', benAmt = '', covTier = '';

      for (let s = c + 1; s <= Math.min(c + 5, headerRow.length - 1); s++) {
        const sHdr = String(headerRow[s]);
        if (covTier === '' && sHdr.toLowerCase().includes('insured option')) covTier = tierFromOption(srcData[r][s]);
        if (eeCost === '' && sHdr.toLowerCase().includes('premium') && sHdr.toLowerCase().startsWith(planPrefix.toLowerCase())) eeCost = srcData[r][s];
        if (effDate === '' && sHdr.toLowerCase().includes('issue date') && sHdr.toLowerCase().startsWith(planPrefix.toLowerCase())) effDate = srcData[r][s];
        if (benAmt === '' && (sHdr.toLowerCase().includes('benefit amount') || sHdr.toLowerCase().includes('certificate amount')) && sHdr.toLowerCase().startsWith(planPrefix.toLowerCase())) benAmt = srcData[r][s];
      }

      const mappedPlan = standardPlanName(String(planNameRaw), hdr);
      const row = new Array(stackHeaders.length).fill('');
      row[0] = empSSN; row[4] = mappedPlan; row[6] = (mappedPlan === 'Assurity Group Disability Income' && !covTier) ? 'EO' : covTier;
      row[9] = benAmt; row[16] = effDate; row[20] = eeCost;
      output.push(row);
    }
  }
  if (output.length > 0) wsDst.getRange(2, 1, output.length, stackHeaders.length).setValues([output]);
}

function syncDemographics() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsDem = ss.getSheetByName('Demographic');
  const wsEmp = ss.getSheetByName('Employees');
  if (!wsDem || !wsEmp) return;

  const demData = wsDem.getDataRange().getValues();
  const empLastCol = wsEmp.getLastColumn();
  const empHeaders = wsEmp.getRange(1, 1, 1, empLastCol).getValues()[0];
  let empData = wsEmp.getLastRow() >= 2 ? wsEmp.getRange(2, 1, wsEmp.getLastRow() - 1, empLastCol).getValues() : [];

  const srcCols = {}, dstCols = {};
  FIELD_MAP.forEach(fm => { srcCols[fm.src] = findColByName(demData[0], fm.src); dstCols[fm.src] = findColByName(empHeaders, fm.dst); });

  const deptCol = findColByName(empHeaders, 'Departments');
  const hoursCol = findColByName(empHeaders, 'Hours Per Week');
  const ssnColEmp = findColByName(empHeaders, 'Employee SSN');

  const empIndex = {};
  for (let e = 0; e < empData.length; e++) { const ssn = padSSN(String(empData[e][ssnColEmp])); if (ssn) empIndex[ssn] = e; }

  for (let d = 1; d < demData.length; d++) {
    const flag = String(demData[d][0]).trim().toLowerCase();
    if (flag !== 'yes' && flag !== 'customized') continue;

    const ssn = padSSN(String(demData[d][srcCols['Employee SSN']])); if (!ssn) continue;

    if (ssn in empIndex) {
      const empRow = empIndex[ssn];
      FIELD_MAP.forEach(fm => { empData[empRow][dstCols[fm.src]] = demData[d][srcCols[fm.src]]; });
      if (deptCol >= 0) empData[empRow][deptCol] = 'All Departments';
      if (hoursCol >= 0) empData[empRow][hoursCol] = 40;
    } else {
      const newRow = new Array(empLastCol).fill('');
      FIELD_MAP.forEach(fm => { newRow[dstCols[fm.src]] = demData[d][srcCols[fm.src]]; });
      if (deptCol >= 0) newRow[deptCol] = 'All Departments';
      if (hoursCol >= 0) newRow[hoursCol] = 40;
      empData.push(newRow); empIndex[ssn] = empData.length - 1;
    }
  }
  if (empData.length > 0) wsEmp.getRange(2, 1, empData.length, empLastCol).setValues(empData);
}

// ============================================================
// §11  PPS VISUAL SIDEBAR COMPONENT DEFINITIONS
// ============================================================
function getPPSMappingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const wsAE = ss.getSheetByName('Auto Enroll Sheet'); const wsPPS = ss.getSheetByName('PPS');
  const aeHeaders = wsAE.getRange(1, 1, 1, wsAE.getLastColumn()).getValues()[0]; const aeColMap = buildColumnMap(aeHeaders);
  const ppsRawHdrs = wsPPS.getRange(6, 1, 1, wsPPS.getLastColumn()).getValues()[0]; const ppsColMap = buildColumnMap(ppsRawHdrs);

  const matchPPSName = (candidates) => { for (const n of candidates) { if (ppsColMap[n] !== undefined) return n; } return null; };
  const premiumMappings = PPS_PREMIUM_FIELD_MAP.map(({ src, label, dsts }) => ({ src, label, dsts, matched: !!(aeColMap[src] !== undefined && matchPPSName(dsts)), matchedDst: matchPPSName(dsts) }));

  return {
    demoMappings: PPS_DEMO_FIELD_MAP.map(({ src, dst }) => ({ src, dst, matched: !!(aeColMap[src] !== undefined && matchPPSName([dst])), matchedDst: matchPPSName([dst]) })),
    ssnLast4Matched: matchPPSName(SSN_LAST4_CANDIDATES),
    premiumMappings,
    tlMappings: [{ src: 'TL Premium', label: 'TL Premium — 10 Year', dsts: ['Term Life 10 Yr After Tax'] }],
    consolidatedTLDst: matchPPSName(['Term Life After Tax']),
    ppsCols: ppsRawHdrs.map(h => String(h).trim()).filter(h => h !== ''),
    detectedPayFreqLabel: 'Bi-Weekly (26 Periods)', detectedPayFreq: 26, multipleFreqs: false
  };
}

function buildPPSSidebarHTML(data) {
  return `<body><h3>Sync to PPS</h3><p class="intro">Review matrix layout generation constraints before processing execution data pipelines.</p><button class="btn" onclick="google.script.run.runPPSSyncFromSidebar(JSON.stringify({premium:[],tl:[],ssnLast4Dst:''}),'26')">Run PPS Sync Matrix</button></body>`;
}

function runPPSSyncFromSidebar(mappingsJson, payFreqInput) {
    const pf = parsePayFrequency(payFreqInput);   // sidebar/config freq, no longer ignored
    syncToPPS(pf);
    return "✓ Operations finished executing.";
  }

  function syncToPPS(payFreqOverride) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const wsAE  = ss.getSheetByName('Auto Enroll Sheet');
    const wsDem = ss.getSheetByName('Demographic');
    const wsPPS = ss.getSheetByName('PPS');
    if (!wsAE || !wsDem || !wsPPS) return;

    const cfg     = getConfig(ss);
    const payFreq = Number(payFreqOverride) || Number(cfg['Pay Frequency']) || 26;
    const factor  = 12 / payFreq;                  // honors Weekly / Semi-Monthly / Monthly

    const aeData = wsAE.getDataRange().getValues();
    const aeColMap = buildColumnMap(aeData[0]);
    const demData = wsDem.getDataRange().getValues();
    const ppsHeaders = wsPPS.getRange(6, 1, 1, wsPPS.getLastColumn()).getValues()[0];
    const ppsColMap = buildColumnMap(ppsHeaders);

    const ssnFlagMap = {};
    const ssnColDem = findColByName(demData[0], 'Employee SSN');
    for (let d = 1; d < demData.length; d++) {
      const flag = String(demData[d][0]).trim().toLowerCase();
      if (flag === 'yes' || flag === 'customized') { const ssn = padSSN(String(demData[d][ssnColDem])); if (ssn) ssnFlagMap[ssn] = flag; }
    }

    let ppsData = wsPPS.getLastRow() >= 7 ? wsPPS.getRange(7, 1, wsPPS.getLastRow() - 6, ppsHeaders.length).getValues() : [];
    const ppsLNCol = ppsColMap['Last Name'];

    for (let r = 1; r < aeData.length; r++) {
      if (String(aeData[r][aeColMap['Relation']]).toUpperCase().trim() !== 'EE') continue;
      const ssn = padSSN(String(aeData[r][aeColMap['Employee SSN']])); if (!ssn || !ssnFlagMap[ssn]) continue;

      const matchKey = (ssn.slice(-4)) + '|' + String(aeData[r][aeColMap['Last Name']]).toLowerCase().trim();
      let rowIdx = -1;
      for (let p = 0; p < ppsData.length; p++) {
        const pKey = String(ppsData[p][ppsColMap['Employee SSNLast 4 digits'] || 0]) + '|' + String(ppsData[p][ppsLNCol]).toLowerCase().trim();
        if (pKey === matchKey) { rowIdx = p; break; }
      }
      if (rowIdx === -1) { ppsData.push(new Array(ppsHeaders.length).fill('')); rowIdx = ppsData.length - 1; }
      const row = ppsData[rowIdx];

      PPS_DEMO_FIELD_MAP.forEach(m => { if (aeColMap[m.src] !== undefined && ppsColMap[m.dst] !== undefined) row[ppsColMap[m.dst]] = aeData[r][aeColMap[m.src]]; });
      if (ppsColMap['Employee SSNLast 4 digits'] !== undefined) row[ppsColMap['Employee SSNLast 4 digits']] = ssn.slice(-4);

      if (ssnFlagMap[ssn] === 'yes') {
        let spent = 0;
        PPS_PREMIUM_FIELD_MAP.forEach(m => {
          const ppsIdx = findColByName(ppsHeaders, m.dsts[0]);
          if (aeColMap[m.src] !== undefined && ppsIdx !== -1) {
            const val = parseFloat(aeData[r][aeColMap[m.src]]) || 0;
            if (val > 0) { const pv = Math.round(val * factor * 100) / 100; row[ppsIdx] = pv; spent += pv; }
          }
        });
        if (ppsColMap['Total Spent Per Pay Period'] !== undefined) row[ppsColMap['Total Spent Per Pay Period']] = spent;
      }
    }
    if (ppsData.length > 0) wsPPS.getRange(7, 1, ppsData.length, ppsHeaders.length).setValues(ppsData);
  }

function parsePayFrequency(val) {
  if (!val) return null; const num = parseFloat(String(val).replace(/[^0-9.]/g, '')); if (!isNaN(num) && num > 0) return num;
  const s = String(val).trim().toLowerCase();
  if (s.includes('weekly') && !s.includes('bi')) return 52; if (s.includes('bi-weekly') || s.includes('bi weekly')) return 26;
  if (s.includes('semi-monthly') || s.includes('semi monthly')) return 24; if (s.includes('monthly')) return 12;
  return null;
}
// ============================================================
// SYSTEM UTILITY — RESET CLONED WORKBOOK
// ============================================================
function clearAllClientData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Hard core confirmation check to prevent disasters
  const response = ui.alert(
    '🚨 CRITICAL WORKBOOK WIPE',
    'This will completely erase all employee records, allotments, premium calculations, and configuration settings inside this cloned workbook.\n\nAre you absolutely sure you want to reset this file?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Operation cancelled. Your data is untouched.');
    return;
  }
  
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Skip sheets that have absolutely no contents
    if (lastRow < 1 || lastCol < 1) return;
    
    // Default rule: Data begins on Row 2 (Demographic, Auto Enroll, Employees, StackedPlans)
    let dataStartRow = 2; 
    
    // Apply structural configuration offsets
    if (name === 'PPS') {
      dataStartRow = 7; // Structural headers live on row 6; data rows are 7+
    } else if (name === CONFIG_SHEET_NAME) {
      dataStartRow = 2; // Key/Value labels occupy row 1
    }
    
    // Clear the contents ONLY if transactional data exists below the headers
    if (lastRow >= dataStartRow) {
      const totalRowsToWipe = lastRow - dataStartRow + 1;
      sheet.getRange(dataStartRow, 1, totalRowsToWipe, lastCol).clearContent();
    }
  });
  
  ui.alert(
    '✨ Workbook Reset Complete', 
    'All individual employee data, calculations, and setup parameters have been safely cleared.\n\nAll structural sheet headers have been preserved for the next onboarding run.', 
    ui.ButtonSet.OK
  );
}
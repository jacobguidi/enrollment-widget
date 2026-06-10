# TBG Enrollment System — Project README

> **Maintainer:** Jacob Thrive Benefits Group  
> **Last Updated:** June 2026  
> **Purpose:** Consolidated reference for all interconnected scripts powering the TBG group enrollment and eligibility pipeline.

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Architecture & Data Flow](#architecture--data-flow)
3. [Directory Structure](#directory-structure)
4. [Version Control Process](#version-control-process)
5. [Scripts Reference](#scripts-reference)
   - [MasterEngine](#masterengine)
   - [MasterEngine Menu](#masterengine-menu)
   - [Enrollment Widget](#enrollment-widget)
   - [Consolidator](#consolidator)
6. [External Integrations](#external-integrations)
7. [Key Constants & Configuration](#key-constants--configuration)
8. [Shared Data Model](#shared-data-model)
9. [Known Issues & Upgrade Notes](#known-issues--upgrade-notes)
10. [Onboarding Checklist — New Group](#onboarding-checklist--new-group)
11. [Glossary](#glossary)

---

## System Overview

The TBG enrollment system is a Google Apps Script + browser-based pipeline that:

1. Onboards employer groups into a **Plan Builder workbook** (Google Sheets)
2. Auto-assigns benefit plans to employees based on their demographics and allotments
3. Publishes a **browser-based Enrollment Widget** (HTML/JS) that employees use to review and customize their benefits
4. Writes confirmed elections back to the workbook via a **webhook** and into a connected **PPS tracker** (payroll deduction sheet)
5. Pushes contact records to **GoHighLevel (GHL)** for enrollment communications and automation
6. Syncs eligibility data to a **Master Workbook**, which the Consolidator uses to generate monthly **Recuro eligibility reports**

All groups share the same engine (MasterEngine library), the same widget codebase, and the same master workbook. Each client gets their own cloned Plan Builder workbook.

---

## Architecture & Data Flow

```
┌─────────────────────────────────────────────────────────────┐
│                     NEW GROUP SETUP                         │
│  HR sends demographic file  →  Paste into Demographic tab   │
│  Create / connect PPS Workbook  →  Enter Rate Table Sheet ID│
└────────────────────┬────────────────────────────────────────┘
                     │ Setup & Import Wizard (Step 1)
                     ▼
┌─────────────────────────────────────────────────────────────┐
│              PLAN BUILDER WORKBOOK (per client)             │
│                                                             │
│  Tabs:  Demographic  |  Auto Enroll Sheet  |  Employees     │
│         PPS          |  StackedPlans       |  __config      │
│                                                             │
│  Step 1 — Setup & Import  (saveSetupAndImport)              │
│    • Saves config to __config tab                           │
│    • Imports savings + salary data (Demo Import)            │
│    • Registers client row in Master Workbook > Clients tab  │
│                                                             │
│  Step 2 — Build Plans  (runBuildPlans)                      │
│    • Reads Rate Tables Sheet (separate Google Sheet)        │
│    • Assigns TL, WL, AE, CI, HI, STD premiums per employee │
│    • Respects "Leave in Paycheck" cushion                   │
│    • Populates Benefit Package strings                      │
│                                                             │
│  Step 3 — Sync Out  (runSyncOut)                            │
│    • Stack All Plans → StackedPlans tab (Selerix upload)    │
│    • Sync Demographics → Employees tab                      │
│    • Sync to PPS → PPS workbook (deduction tracker)         │
│    • GHL push is manual: Tools > Push to GHL only           │
└─────────────────────┬──────────────────┬───────────────────┘
                      │                  │
          GHL Push    │                  │  Webhook
          (manual)    │                  │  (doPost)
                      ▼                  ▼
┌──────────────────────────┐  ┌──────────────────────────────┐
│  GoHighLevel (GHL)       │  │  Enrollment Widget           │
│                          │  │  (Browser HTML/JS)           │
│  • Contact upsert        │  │                              │
│  • Custom fields written │  │  Step 1: Employee info       │
│    (allotment, benefits  │  │  Step 2: Customize benefits  │
│    package, savings,     │  │  Step 3: Dependents +        │
│    company, SSN last 4,  │  │          beneficiaries +     │
│    Plan Builder sheet ID)│  │          GHL consent form    │
│                          │  │                              │
│  GHL Automations →       │  │  Rates loaded live from      │
│  Send enrollment link    │  │  Rate Tables Google Sheet    │
│  to employee             │  │  (CSV/JSONP)                 │
│                          │  │                              │
│  On opt-in or customize: │  │  On submit:                  │
│  POST to doPost webhook  │◄─┘  POST to WEBHOOK_URL         │
└──────────┬───────────────┘                                  │
           │ Webhook payload routed by syncCustomizedEmployee()│
           ▼                                                  │
┌─────────────────────────────────────────────────────────────┘
│  Plan Builder Workbook (webhook write-back)                 │
│    • Demographic tab → flag set to "Yes" or "Customized"    │
│    • Auto Enroll Sheet → plan fields updated if Customized  │
│    • PPS tab → deduction amounts written                    │
└──────────────────────────┬──────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│              MASTER WORKBOOK (single shared)                │
│                                                             │
│  Clients tab  — one row per group                           │
│  Master Eligibility tab  — aggregated PPS data              │
│                                                             │
│  ← Consolidator reads this workbook                         │
└──────────────────────────┬──────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                     CONSOLIDATOR                            │
│                                                             │
│  runMasterConsolidation()  — nightly cron (2am)             │
│    Reads PPS tab from every Active client workbook          │
│    Writes normalized rows to Master Eligibility tab         │
│                                                             │
│  generateRecuroFile()  — monthly cron (5th of month, 6am)  │
│    Reads Master Eligibility                                 │
│    Generates Recuro_Eligibility_YYYY-MM-DD.csv              │
│    Saves to Drive + emails to jacob@thrivebg.com            │
└─────────────────────────────────────────────────────────────┘
```

---

## Directory Structure

This is the intended Claude Code working directory layout. Each script maps to a folder with a `current/` version and a `previous_versions/` archive.

```
tbg-enrollment-system/
│
├── README.md                          ← this file
│
├── MasterEngine/
│   ├── current/
│   │   └── MasterEngine.gs
│   └── previous_versions/
│       └── MasterEngine_YYYY-MM-DD_vN.gs
│
├── MasterEngine_Menu/
│   ├── current/
│   │   └── MasterEngine_Menu.gs
│   └── previous_versions/
│       └── MasterEngine_Menu_YYYY-MM-DD_vN.gs
│
├── Enrollment_Widget/
│   ├── current/
│   │   └── Enrollment_Widget.html
│   └── previous_versions/
│       └── Enrollment_Widget_YYYY-MM-DD_vN.html
│
└── Consolidator/
    ├── current/
    │   └── Consolidator.gs
    └── previous_versions/
        └── Consolidator_YYYY-MM-DD_vN.gs
```

---

## Version Control Process

> **This is a hard rule. Claude Code must follow these steps every time, without being asked, before touching any file.**

---

### Rule: Save Before You Edit

Before modifying any file in `current/`, you must copy the existing version into `previous_versions/` with a dated, versioned filename. This is non-negotiable — even for small changes.

---

### Naming Convention

```
<ScriptName>_YYYY-MM-DD_v<N>.<ext>
```

| Part | Meaning |
|---|---|
| `ScriptName` | Matches the folder name exactly: `MasterEngine`, `MasterEngine_Menu`, `Enrollment_Widget`, `Consolidator` |
| `YYYY-MM-DD` | Today's date |
| `vN` | Version number starting at `v1`. If a save already exists for today, increment: `v2`, `v3`, etc. |
| `.ext` | `.gs` for Apps Script files, `.html` for the widget |

**Examples:**
```
MasterEngine_2026-06-09_v1.gs
MasterEngine_2026-06-09_v2.gs        ← second change same day
MasterEngine_Menu_2026-06-09_v1.gs
Enrollment_Widget_2026-06-09_v1.html
Consolidator_2026-06-09_v1.gs
```

---

### Step-by-Step Process (Claude Code must follow this exactly)

**Step 1 — Check what versions already exist today**
```bash
ls <ScriptName>/previous_versions/
```
Look for any file dated today. If `v1` already exists, your archive will be `v2`. If none exist, use `v1`.

**Step 2 — Copy current file to previous_versions with correct name**
```bash
cp <ScriptName>/current/<filename> \
   <ScriptName>/previous_versions/<ScriptName>_YYYY-MM-DD_vN.<ext>
```

Real examples for each script:
```bash
# MasterEngine
cp MasterEngine/current/MasterEngine.gs \
   MasterEngine/previous_versions/MasterEngine_2026-06-09_v1.gs

# MasterEngine Menu
cp MasterEngine_Menu/current/MasterEngine_Menu.gs \
   MasterEngine_Menu/previous_versions/MasterEngine_Menu_2026-06-09_v1.gs

# Enrollment Widget
cp Enrollment_Widget/current/Enrollment_Widget.html \
   Enrollment_Widget/previous_versions/Enrollment_Widget_2026-06-09_v1.html

# Consolidator
cp Consolidator/current/Consolidator.gs \
   Consolidator/previous_versions/Consolidator_2026-06-09_v1.gs
```

**Step 3 — Confirm the archive was created**
```bash
ls -lh <ScriptName>/previous_versions/
```
Verify the file exists and the file size matches the original before proceeding.

**Step 4 — Make your edits to the file in `current/`**

Edit `<ScriptName>/current/<filename>` only. Never edit files inside `previous_versions/`.

**Step 5 — Add a change note at the top of the edited file**

Every edited file should have a comment block at the very top noting what changed and when:

```javascript
// ============================================================
// CHANGE LOG
// ============================================================
// 2026-06-09 v2 — Description of what was changed and why
// 2026-06-09 v1 — Initial version saved to version control
// ============================================================
```

---

### Recovering a Previous Version

To restore a previous version, copy from `previous_versions/` back to `current/` — but **archive the current file first** before overwriting it:

```bash
# 1. Archive the current (broken) version before overwriting
cp MasterEngine/current/MasterEngine.gs \
   MasterEngine/previous_versions/MasterEngine_2026-06-09_v3_broken.gs

# 2. Restore the version you want
cp MasterEngine/previous_versions/MasterEngine_2026-06-09_v1.gs \
   MasterEngine/current/MasterEngine.gs
```

Use the `_broken` suffix on the archived file so it's clear why it was saved.

---

### Quick Reference — All Four Scripts

| Script | Current File | Previous Versions Folder |
|---|---|---|
| MasterEngine | `MasterEngine/current/MasterEngine.gs` | `MasterEngine/previous_versions/` |
| MasterEngine Menu | `MasterEngine_Menu/current/MasterEngine_Menu.gs` | `MasterEngine_Menu/previous_versions/` |
| Enrollment Widget | `Enrollment_Widget/current/Enrollment_Widget.html` | `Enrollment_Widget/previous_versions/` |
| Consolidator | `Consolidator/current/Consolidator.gs` | `Consolidator/previous_versions/` |

---

## Scripts Reference

---

### MasterEngine

**File:** `MasterEngine/current/MasterEngine.gs`  
**Type:** Google Apps Script Library  
**Deployed as:** A Script Library — other workbooks call its functions via `MasterEngine.<functionName>()`

#### Purpose
Central production engine. All business logic lives here. Client Plan Builder workbooks have no logic of their own — they only contain the MasterEngine Menu script (which passes calls through to this library) and the data tabs.

#### Sections (§)

| Section | Name | Key Functions |
|---|---|---|
| §1 | Constants + Menu | `onOpen()`, `MASTER_WORKBOOK_ID`, mapping tables |
| §2 | `__config` Tab Helpers | `getConfig()`, `saveConfig()`, `configValue()` |
| §3 | Setup & Import Wizard | `showSetupWizard()`, `loadCurrentConfig()`, `saveSetupAndImport()` |
| §3a | Master Workbook Registration | `registerWithMaster()` |
| §3b | Wizard HTML | `_buildWizardHTML()` — 5-tab sidebar UI |
| §4 | Build Plans | `runBuildPlans()` — assigns all benefit premiums |
| §5 | Sync Out | `runSyncOut()` — orchestrates stack + demo sync + PPS sync |
| §6 | Push to GHL | `pushDemographicToGHL()` — upserts contacts to GoHighLevel |
| §7 | GHL API Helpers | `_ghlFetchCustomFields()`, `_ghlUpsertContact()` |
| §8 | Field Normalizers | `_normalizePhone()`, `_normalizeDOB()` |
| §9 | Webhook — doPost | `doPost()`, `syncFromBenefitSummary()` |
| §10 | Webhook Helpers | `syncCustomizedEmployee()`, `_writePlansFromPayload()` |
| §11 | PPS Visual Sidebar | `showPPSSidebar()`, `getPPSMappingData()`, `runPPSSyncFromSidebar()` |
| — | Rate Table Loaders | `loadAllRateTables()`, `detectSTDOptions()` |
| — | Plan Assigners | `assignTermLife()`, `assignWholeLife()`, `assignAccident()`, `assignCriticalIllness()`, `assignHospitalIndemnity()`, `assignSTD()` |
| — | Stack + Demo Sync | `stackAllPlans()`, `syncDemographics()`, `executeDemoImport()` |
| — | Utilities | `buildColumnMap()`, `normalizeSSN()`, `padSSN()`, `toProperCase()`, `calcAge()` |
| — | Reset | `clearAllClientData()` — wipes all data rows, preserves headers |

#### Important Constants

```javascript
const MASTER_WORKBOOK_ID = '1jqzmXQPbI0jlIvVN7dMgVL6YNHz8oxTEwboh8ROVbtU';
const GHL_BASE           = 'https://services.leadconnectorhq.com';
const GHL_API_VERSION    = '2021-07-28';
const CONFIG_SHEET_NAME  = '__config';
```

#### `__config` Tab Keys

| Key | Description |
|---|---|
| `Company Name` | Group's company name |
| `Group ID` | Recuro-assigned group ID (fill in after onboarding) |
| `Enrollment Cutoff` | Display text/date for enrollment deadline |
| `Effective Date` | Policy effective date (MM/DD/YYYY) |
| `Sign Date` | Signature date for issued policies |
| `Rate Tables Sheet ID` | Google Sheet ID of the group's rate table workbook |
| `STD Benefit Period` | Selected STD option label from rate sheet |
| `Pay Frequency` | Numeric periods per year (52/26/24/12) |
| `PPS Workbook ID` | Sheet ID of the PPS deduction tracker |
| `PPS Mapping` | JSON blob — maps AE Sheet columns → PPS columns |
| `Leave in Paycheck` | Dollar amount to exclude from build budget per employee |

#### Script Properties (set manually in Apps Script)

| Property | Description |
|---|---|
| `GHL_PIT` | GoHighLevel Private Integration Token |
| `GHL_LOCATION_ID` | GHL Location/Sub-account ID |
| `WEBHOOK_SECRET` | Secret string validated on incoming webhook POSTs |
| `SPREADSHEET_ID` | Fallback spreadsheet ID if payload doesn't include one |
| `PAY_FREQUENCY` | Global fallback pay frequency for webhook-triggered syncs |

#### Plan Assignment Logic (Build Plans)

Premiums are assigned in this priority order per employee:

1. **Term Life** — highest priority; assigned first against full budget
2. **Whole Life (Employee)** — assigned from remaining budget
3. **Accident Expense** — plan1 or plan2 based on remaining budget
4. **Critical Illness** — age-banded; largest affordable benefit
5. **Hospital Indemnity** — plan1 or plan2
6. **STD (Short Term Disability)** — salary-capped weekly benefit

The `Leave in Paycheck` cushion reduces the **build budget** but not the `Remaining Formula` column (so the cushion is visible in the sheet).

#### Webhook Routing (`doPost`)

Incoming POST requests are routed based on payload shape:

- **`benefit_summary` key present** → `syncFromBenefitSummary()` (GHL form submission text)
- **`benefits` key present** → `syncCustomizedEmployee()` (structured widget payload)

Employee matching priority: `contactId` → `email` → `firstName + lastName`

Flag written to Demographic column A:
- `"Yes"` — opt-in (no customization; default plan preserved)
- `"Customized"` — employee changed at least one selection

---

### MasterEngine Menu

**File:** `MasterEngine_Menu/current/MasterEngine_Menu.gs`  
**Type:** Google Apps Script — bound directly to each client Plan Builder workbook  
**Purpose:** Thin pass-through layer. Creates the `TBG Plan Builder` menu and routes every menu item to the corresponding `MasterEngine.<function>()` call.

#### Why it exists
Apps Script libraries can't directly bind menus to a host spreadsheet's UI. This script sits in the workbook and forwards every call to the library. When you update MasterEngine, all workbooks automatically get the new behavior with no changes needed here.

#### Menu Structure

```
TBG Plan Builder
  1. Setup & Import         → showSetupWizard()
  2. Build Plans            → runBuildPlans()
  3. Sync Out               → runSyncOut()
  ─────
  Tools >
    Edit Setup              → showSetupWizard()
    Stack All Plans         → stackAllPlans()
    Sync Demographics       → syncDemographics()
    Sync to PPS (sidebar)   → showPPSSidebar()
    Push to GHL only        → pushDemographicToGHL()
    ─────
    Demo Import (legacy)    → showDemoSidebar()
    Auto Plan Builder (legacy) → runAutoPlanBuilder()
    Build Benefit Package (legacy) → runBenefitPackageBuilder()
    ─────
    Test GHL Connection     → testGHLConnection()
    Test Widget Sync        → testSyncCustomizedEmployee()
    ─────
    Clear All Workbook Data → clearAllClientData()
```

#### Critical Web Wizard Callbacks
The following functions must exist in this file (not just the library) because Google's `google.script.run` can only call functions bound to the active document's script context:

```javascript
function loadCurrentConfig()
function loadPPSColumnsForWizard()
function loadSTDOptionsForWizard(rateSheetId)
function saveSetupAndImport(formJSON)
function runDemoImportFromSidebar(savingsRawText, salaryRawText)
function runPPSSyncFromSidebar(mappingsJson, payFreqInput)
```

These all delegate immediately to `MasterEngine.*` — do not put logic here.

---

### Enrollment Widget

**File:** `Enrollment_Widget/current/Enrollment_Widget.html`  
**Type:** Standalone HTML file (single file, CSS + JS inline)  
**Deployed:** Embedded in GHL funnel pages via iframe  
**Purpose:** Browser-based 3-step enrollment flow for employees

#### Step Flow

```
Step 1 — About You
  • Company select (maps to Rate Table Sheet ID)
  • First name, last name, email, phone, age
  • Tobacco status toggle
  • On advance → loads live rates from company's Rate Tables Sheet

Step 2 — Your Benefits Package
  • Shows one card per benefit (7 total)
  • Each card has a toggle (on/off), tier selector, and plan options
  • Live allotment banner shows budget used vs. remaining
  • Rates loaded from Google Sheets (CSV/JSONP) or fallback tables

Step 3 — Details & Beneficiaries
  • Spouse info (if any tiered benefit includes spouse)
  • Children info (if any benefit includes children/family)
  • Beneficiary name, relationship, % allocation
  • Enrollment summary table
  • GHL consent form embedded as iframe on submit
```

#### Company → Rate Sheet Mapping

Defined in the `COMPANIES` object at the top of the script:

```javascript
const COMPANIES = {
  'Company Name': {
    sheetId: '<Google Sheet ID of Rate Tables workbook>',
    wholeEmpMaxBenefit: 75000
  },
  ...
};
```

To add a new company: add a key here with their Rate Tables sheet ID.

#### Rate Loading Architecture

The widget fetches 8 tabs from the Rate Tables sheet:

| Tab Name | Fallback Name | Parser |
|---|---|---|
| `Accident` | `ACC` | `parsePlanTab()` |
| `Hospital Indemnity` | `Hospital` | `parsePlanTab()` |
| `Critical Illness` | `CI` | `parseCITab()` |
| `Short Term Disability` | `STD` | `parseDisabilityTab()` |
| `10 Year` | `Term 10` | `parseTermTab()` |
| `20 Year` | `Term 20` | `parseTermTab()` |
| `To Age 70` | `Age 70` | `parseTermTab()` |
| `Whole Life` | `Whole Life Employee` | `parseWholeLifeTab()` |

If any tab fails to load, that benefit uses embedded fallback rate tables. A rate source badge is shown on Step 2 (`Live rates loaded` vs. `Using embedded rates`).

**Tab name debugging:** Open browser console (F12) and check `window._debugRates` for per-tab load status.

#### Spouse Whole Life Logic

- Rates are read from the **same Whole Life tab as the employee**, keyed by spouse age
- Hard cap: `WHOLE_SPOUSE_MAX_BENEFIT = 25000`
- Soft cap: spouse benefit cannot exceed the employee's elected WL benefit
- All spouse amounts show an underwriting notice

#### Submission Flow

1. Employee clicks "Review & Sign"
2. Enrollment summary assembled as plain text (`benefit_summary` string)
3. GHL consent form loaded as iframe (pre-filled with name/email/phone/summary)
4. On GHL form submit → `window.postMessage` event → redirect to `https://gothrivebg.com/whats-next`
5. GHL automation fires → sends `benefit_summary` payload → POSTs to `WEBHOOK_URL`
6. MasterEngine `doPost()` receives it → writes back to Plan Builder workbook

#### URL Parameters (pre-fill from GHL)

The widget accepts these URL parameters to pre-fill the form:

| Param | Description |
|---|---|
| `first_name` | Employee first name |
| `last_name` | Employee last name |
| `email` | Employee email |
| `phone` | Employee phone |
| `age` | Employee age (integer) |
| `allotment` | Monthly tax savings allotment (shows banner) |
| `contact_id` | GHL contact ID (passed to webhook for matching) |
| `company` | Company name (auto-selects from dropdown) |
| `benefits_package` | Pre-built package string (auto-selects benefit toggles) |

---

### Consolidator

**File:** `Consolidator/current/Consolidator.gs`  
**Type:** Standalone Google Apps Script (not a library)  
**Deployed:** Bound to the Master Workbook  
**Purpose:** Two jobs — (1) aggregate PPS data from all active clients into one eligibility sheet, (2) generate and email monthly Recuro CSV

#### Constants (set once at top of file)

```javascript
const MASTER_WORKBOOK_ID = '1jqzmXQPbI0jlIvVN7dMgVL6YNHz8oxTEwboh8ROVbtU';
const NOTIFY_EMAIL       = 'jacob@thrivebg.com';
const PPS_HEADER_ROW  = 6;    // PPS headers always live on row 6
const PPS_DATA_START  = 7;    // PPS data rows start on row 7
```

#### Master Workbook Structure

**Clients tab** — one row per group, headers expected:

| Column | Description |
|---|---|
| `Company Name` | Group display name |
| `Plan Builder ID` | Google Sheet ID of the Plan Builder workbook |
| `Workbook ID` | Google Sheet ID of the PPS Tracker workbook |
| `Rate Table ID` | Google Sheet ID of the Rate Tables workbook |
| `Group ID` | Recuro-assigned group ID |
| `Service Offering` | Product/tier description for Recuro |
| `Enrollment Cutoff` | Enrollment deadline text |
| `Effective Date` | Policy effective date |
| `Pay Frequency` | Payroll frequency |
| `Status` | `Active` or anything else (non-active rows are skipped) |

**Master Eligibility tab** — aggregated output, headers in row 1 (`RECURO_HEADERS` constant)

#### `runMasterConsolidation()`

- Iterates all rows in `Clients` tab where `Status === 'active'`
- For each client, opens their PPS workbook, reads the `PPS` tab
- Normalizes PPS headers to Recuro field names using `PPS_TO_RECURO` map
- Skips blank rows, rows missing first+last name, rows missing effective date
- Skips terminated rows older than 60 days
- Writes all rows to `Master Eligibility` tab (full overwrite)

#### `generateRecuroFile()`

- Reads Master Eligibility tab
- Validates required fields: `LastName`, `FirstName`, `DateOfBirth`, `EmailAddress`, `EffectiveStart`, `MemberType`, `ClientMemberID`, `ServiceOffering`, `GroupID`, `GroupName`
- Skips rows with `_TBG_Status = 'pending'` or missing required fields
- Generates CSV with only the 26 Recuro-spec columns (strips internal `_TBG_*` columns)
- Saves file to the same Drive folder as the Master Workbook
- Emails CSV attachment + summary to `NOTIFY_EMAIL`

#### Triggers (set up once with `createTriggers()`)

| Trigger | Schedule | Function |
|---|---|---|
| Nightly consolidation | Every day at 2am | `runMasterConsolidation()` |
| Monthly Recuro file | 5th of each month at 6am | `generateRecuroFile()` |

Run `createTriggers()` once manually to install. Re-running it safely deletes and recreates both triggers.

---

## External Integrations

### GoHighLevel (GHL)

- **Sub-account/Location ID:** stored in Script Properties as `GHL_LOCATION_ID`
- **Auth:** Private Integration Token (`GHL_PIT`) in Script Properties
- **Upsert endpoint:** `POST /contacts/upsert`
- **Custom fields written per contact:**
  - `Allotment` (Net Tax Savings monthly)
  - `Enrollment Cutoff`
  - `Benefits Package`
  - `Savings`
  - `Company Name`
  - `Gender`, `Age`
  - `SSN Last 4` / `Employee SSN Last 4 digits`
  - `Plan Builder Sheet ID`

### Recuro

- Receives monthly eligibility CSV via email attachment
- File format defined by `RECURO_HEADERS` constant in Consolidator
- `ClientMemberID` = `{Plan Builder ID}{SSN Last 4}` (truncated to 15 chars)
- Effective start required; terminations included for 60 days post-termination

### Selerix (Benefit Admin System)

- Receives `StackedPlans` output tab as an uploadable file
- Headers defined by `stackHeaders` array in `stackAllPlans()`
- One row per policy per employee (employees with multiple TL plans get multiple rows)

### Rate Tables Google Sheet (per client)

- Separate workbook for each client
- Tabs: `10 Year`, `20 Year`, `To Age 70`, `Whole Life`, `Accident`, `Hospital Indemnity`, `Critical Illness`, `Short Term Disability`
- Published to web (no login required) — widget fetches live CSVs
- Sheet ID stored in `__config` as `Rate Tables Sheet ID`

---

## Key Constants & Configuration

### Shared Across Scripts

| Constant | Value | Where |
|---|---|---|
| `MASTER_WORKBOOK_ID` | `1jqzmXQPbI0jlIvVN7dMgVL6YNHz8oxTEwboh8ROVbtU` | MasterEngine §1, Consolidator |
| `PPS_HEADER_ROW` | `6` | MasterEngine, Consolidator |
| `PPS_DATA_START` | `7` | MasterEngine, Consolidator |
| `CONFIG_SHEET_NAME` | `'__config'` | MasterEngine |
| `WHOLE_SPOUSE_MAX_BENEFIT` | `25000` | Enrollment Widget |

### SSN Last 4 — Column Name Candidates

The engine tries these column names in order when looking for SSN Last 4 in the PPS tab:

```javascript
'Employee SSNLast 4 digits'
'Employee SSN Last 4 digits'
'Employee SSN Last 4 Digits'
'SSN Last 4 Digits'
'SSN Last 4'
```

---

## Shared Data Model

### Auto Enroll Sheet — Key Column Groups

**Demographics**
`Relation` | `Employee SSN` | `Insured SSN` | `First Name` | `Last Name` | `DOB` | `Gender` | `Mailing Address1` | `City` | `State` | `Zip` | `Email Address` | `Phone Number` | `Date of Hire` | `Age` | `Pay Frequency`

**Budget**
`Allotments` | `Annual Salary` | `Remaining Formula`

**Term Life (3 slots)**
`TL Plan Type` | `TL Insured Option` | `TL Benefit Amount` | `TL Premium` | `TL Issue Date` | `TL Signed Date`
*(+ `TL Plan Type 2` ... `TL Signed Date 2`, `TL Plan Type 3` ... `TL Signed Date 3`)*

**Whole Life (Employee + Spouse)**
`Life Plan Type` | `Life Insured Option` | `Life Certificate Amount` | `Life Premium Amount` | `Life Issue Date` | `Life Signed Date`
`Spouse Life Plan Type` | `Spouse Life Insured Option` | `Spouse Life Certificate Amount` | `Spouse Life Premium Amount` | `Spouse Life Issue Date` | `Spouse Life Signed Date`

**Other Benefits**
`AE Plan Type` | `AE Insured Option` | `AE Premium Amount` | `AE Issue Date` | `AE Signed Date`
`CI Plan Type` | `CI Insured Option` | `CI Benefit Amount` | `CI Premium Amount` | `CI Issue Date` | `CI Signed Date`
`HI Plan Type` | `HI Insured Option` | `HI Benefit Amount` | `HI Premium Amount` | `HI Issue Date` | `HI Signed Date`
`DI Plan Type` | `DI Benefit Amount` | `DI Premium Amount` | `DI Issue Date` | `DI Signed Date`

**Output**
`Benefit Package` | `Contact ID` | `GHL Sync Status` | `GHL Sync Last Run`

### Demographic Tab — Key Columns

Column A is the **enrollment flag**: `Yes` | `Customized` | (blank)

`Employee SSN` | `First Name` | `Last Name` | `DOB` | `Gender` | `Mailing Address1` | `City` | `State` | `Zip` | `Email Address` | `Phone Number` | `Date of Hire` | `Company Name` | `Enrollment Cutoff` | `Benefits Package` | `Age` | `Savings` (= Remaining Formula) | `Net Tax Savings` | `Annual Salary` | `Contact ID` | `GHL Sync Status` | `GHL Sync Last Run`

### PPS Tab — Structure

- Row 6: Column headers
- Row 7+: Data rows
- Key columns: `First Name` | `Last Name` | `Date of Birth` | `Gender` | `Address` | `City` | `State` | `Zip` | `Phone` | `Email` | `Employee SSNLast 4 digits` | `Effective Date` | `Pay Frequency` | `Deduction Frequency` | benefit premium columns | `Total After Tax Allotment Per Pay Period` | `Total Spent Per Pay Period`

---

## Known Issues & Upgrade Notes

### Active Patch Notes (as of current build)

1. **§3a — registerWithMaster() TDZ Fix:** `const ss = ss || ...` was a TDZ bug. Fixed to `const ss = SpreadsheetApp.getActiveSpreadsheet()`.
2. **§3a — Workbook ID header tolerance:** Now checks for both `"Workbook ID"` and `"PPS Workbook ID"` in Master Clients tab.
3. **§3 — leaveInPaycheck empty state:** `loadCurrentConfig()` no longer writes `"0"` into a blank input on first load; uses `!= null` check.
4. **§3 — loadSTDOptionsForWizard re-added:** Was dropped in a prior pass and caused `MasterEngine.loadSTDOptionsForWizard is not a function`.
5. **§3 — loadPPSColumnsForWizard re-added:** Same as above.
6. **Consolidator — Plan Builder ID tracking:** Bug fix — now uses `Plan Builder ID` as the primary key column, not `Client ID`.
7. **Enrollment Widget — Spouse WL rates:** Reads from same `Whole Life` tab as employee, keyed by spouse age. Hard-capped at $25,000. Cannot exceed employee's elected benefit.
8. **Enrollment Widget — wholeEmp tab fallbacks:** `fetchSheetTabWithFallbacks()` now tries `'Whole Life'` then `'Whole Life Employee'`. Check `window._debugRates` in console if tab not found.

### Planned Upgrades

- [ ] Rate Table Sheet IDs should be pulled dynamically from Master Workbook `Clients` tab instead of hardcoded in the Enrollment Widget `COMPANIES` object
- [ ] GHL push should optionally trigger automatically after Sync Out (currently manual)
- [ ] Consolidator should write a `_TBG_LastSeen` timestamp and allow soft-delete detection
- [ ] Widget should support being embedded in GHL natively (not just iframe)
- [ ] Add a `clearAllClientData()` confirmation step that logs to Master Workbook before wipe

---

## Onboarding Checklist — New Group

**In Google Drive:**
- [ ] Clone the Plan Builder template workbook → rename to company name
- [ ] Clone the PPS template workbook → rename to company name, configure benefit headers in row 6
- [ ] Confirm client has a Rate Tables workbook (or create one from template)

**In the Plan Builder workbook:**
- [ ] Open `TBG Plan Builder > 1. Setup & Import`
- [ ] Tab 1: Enter Company Name, Group ID (if available), Enrollment Cutoff, Effective Date, Sign Date
- [ ] Tab 2: Paste Rate Tables Sheet ID, select STD Benefit Period
- [ ] Tab 3: Select "New" or "Existing" group, paste PPS Workbook Sheet ID
- [ ] Tab 4: Set Pay Frequency, set Leave in Paycheck cushion, confirm PPS column mappings
- [ ] Tab 5: Paste savings data (First | Last | $Amount) and salary data (Last | First | $Amount)
- [ ] Click `Save & Import` → verify confirmation message

**In Apps Script (Script Properties):**
- [ ] Set `GHL_PIT` (GoHighLevel Private Integration Token)
- [ ] Set `GHL_LOCATION_ID`
- [ ] Set `WEBHOOK_SECRET`
- [ ] Set `SPREADSHEET_ID` (this workbook's ID, fallback for webhook)

**After setup:**
- [ ] Run `TBG Plan Builder > 2. Build Plans` — verify Remaining Formula column is populated
- [ ] Run `TBG Plan Builder > Tools > Test GHL Connection` — verify custom fields load
- [ ] Run `TBG Plan Builder > 3. Sync Out` — verify StackedPlans, Employees, PPS tabs populated
- [ ] Run `TBG Plan Builder > Tools > Push to GHL only` — verify contacts created in GHL
- [ ] Add company and Rate Tables sheet ID to `COMPANIES` object in Enrollment Widget
- [ ] Test widget: open in browser with `?company=<Company Name>&age=35`
- [ ] Verify rates load from sheet (check `Live rates loaded` badge)
- [ ] Test submit → verify webhook write-back to workbook

---

## Glossary

| Term | Meaning |
|---|---|
| **Plan Builder** | The per-client Google Sheet workbook containing Demographic, Auto Enroll Sheet, PPS, and related tabs |
| **MasterEngine** | The Apps Script library containing all business logic |
| **MasterEngine Menu** | The thin Apps Script file bound to each Plan Builder workbook that routes menu calls to the library |
| **PPS** | Payroll Product Summary — the deduction tracker sheet. Headers in row 6, data starts row 7 |
| **AE Sheet** | Auto Enroll Sheet — the main working tab where plan assignments are written per employee row |
| **Rate Tables** | A separate Google Sheet (one per client) containing premium rate tables for all benefit products |
| **Master Workbook** | The single shared Google Sheet containing the `Clients` registry and `Master Eligibility` aggregated data |
| **Consolidator** | Standalone Apps Script bound to the Master Workbook — runs nightly to pull PPS data and monthly to generate Recuro CSV |
| **Recuro** | Third-party health plan administrator; receives monthly eligibility files to manage billing |
| **Selerix** | Benefit admin system; receives StackedPlans output for policy record upload |
| **GHL** | GoHighLevel — CRM/automation platform used for enrollment communications and consent form collection |
| **doPost** | The Apps Script web app endpoint that receives enrollment submissions from the widget (via GHL automation) |
| **allotment** | The monthly dollar amount the employer contributes toward the employee's benefit premiums (= net tax savings) |
| **cushion / Leave in Paycheck** | Amount subtracted from build budget so a portion of the allotment stays in the employee's paycheck |
| **SSN Last 4** | Last 4 digits of employee SSN — used as a lightweight match key in PPS and Recuro records |
| **Customized** | Demographic flag indicating the employee changed their plan selections in the widget |
| **Yes** | Demographic flag indicating the employee opted in with the default plan (no changes) |

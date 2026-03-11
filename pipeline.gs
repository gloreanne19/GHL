/******************************************************
 * MENU
 ******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Data Organization")
    .addItem("Task 1: Build Organized Tab", "runFullPipeline")
    .addItem("Task 2: Clean Data",          "cleanData")
    .addToUi();

  ui.createMenu("Skip Tracing Steps")
    .addItem("Task 1: Ready For XLeads",      "readyForXLeads")
    .addItem("Task 2: XLeads Results",        "ST_part1_readyVsXleads_noMatchOrNoPhone_pullForImport")
    .addItem("Task 3: Ready For Mojo",        "readyForMojo")
    .addItem("Task 4: Mojo Results",          "ST_part2_notInXleadsVsMojo_noMatchOrNoPhone_pullForImport")
    .addItem("Task 5: Ready For Direct Skip", "readyForDirect")
    .addItem("Task 6: Direct Skip Results",   "ST_part3_notInMojoVsDirect_noMatchOrNoPhone_pullForImport")
    .addToUi();

  ui.createMenu("Skip Trace Done (Merge)")
    .addItem("Task 1: Append XLeads", "ST_merge_A_xleads_buildHeaderAndWrite")
    .addItem("Task 2: Append Mojo",   "ST_merge_B_appendMojo")
    .addItem("Task 3: Append Direct", "ST_merge_C_appendDirect")
    .addToUi();
}

/******************************************************
 * PIPELINE (NO Organized, NO Cleaned)
 *
 * Master source: Supabase (contacts table via REST API)
 *
 * Outputs:
 * - Not A Fit (Zipcodes)     (Raw Data rows as-is)
 * - For Import               (Formatted 17 headers + blacklist removed)
 * - Skip Import              (Raw Data rows as-is)
 * - Needs Update             (Raw Data rows as-is; skipped if any match PT has PO)
 * - Match Found              (Master rows + Source; only for Needs Update)
 * - Debug - Needs Update Reasons
 *
 * ZIP Not-A-Fit checked ONLY ONCE (in audit).
 * For Import is built in the 17-column "Organized" header format.
 * Blacklist filtering is applied ONLY to For Import (Decedent).
 ******************************************************/

// ─────────────────────────────────────────────
// SUPABASE CONFIG  (replaces the two master sheet IDs)
// ─────────────────────────────────────────────
const SUPABASE_URL = 'https://fmspuwcfygbsklgnrray.supabase.co';
const SUPABASE_KEY = 'sb_publishable_y2oAsYVOdrjyDcxAa5FX1w_6saruoJY';
const SUPABASE_TABLE = 'contacts';

// ⚠️  To truly fetch ALL rows in one call, go to:
// Supabase Dashboard → Project Settings → API → "Max Rows"
// and set it to a number larger than your total contacts count.

// This will be populated dynamically from Supabase
let DB_LABELS = {}; 
let DYNAMIC_SELECT_COLS = '';

/******************************************************
 * FETCH DYNAMIC SCHEMA
 ******************************************************/
function loadDynamicSchema_() {
  const url = `${SUPABASE_URL}/rest/v1/schema_config?select=excel_header,db_column&is_active=eq.true`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'apikey':        SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('⚠️ Failed to load dynamic schema, using fallback.');
    return;
  }

  const schema = JSON.parse(response.getContentText());
  const labels = {};
  const cols = [];

  schema.forEach(s => {
    labels[s.db_column] = s.excel_header;
    cols.push(s.db_column);
  });

  // Always include id and created_at if not in schema
  if (!cols.includes('id')) cols.push('id');
  if (!cols.includes('created_date')) cols.push('created_date');

  DB_LABELS = labels;
  DYNAMIC_SELECT_COLS = cols.join(',');
}

/******************************************************
 * ENTRY POINT
 ******************************************************/
function runFullPipeline() {
  loadDynamicSchema_(); // First, get the latest mapping
  auditRawDataAgainstMasters_();
}

/******************************************************
 * AUDIT "Raw Data" AGAINST SUPABASE MASTER
 * Fixed input columns (Raw Data):
 *   C = Prospect Type
 *   E = Property Address
 *   H = Property Zip
 ******************************************************/
function auditRawDataAgainstMasters_() {
  const INPUT_SHEET_NAME     = 'Raw Data';
  const NOT_FIT_SHEET_NAME   = 'Not A Fit (Zipcodes)';
  const NEEDS_UPDATE_SHEET   = 'Needs Update';
  const SKIP_IMPORT_SHEET    = 'Skip Import';
  const FOR_IMPORT_SHEET     = 'For Import';
  const MATCH_FOUND_SHEET    = 'Match Found';
  const DEBUG_SHEET_NAME     = 'Debug - Needs Update Reasons';

  // Blacklist spreadsheet
  const blacklistSpreadsheetId = "10GuRF-vFgLG3YRYhTNf0qwx2bNjdZnqZYWvlj7w7EaY";
  const blacklistSheetName     = "";  // blank => first sheet
  const blacklistDecedentCol   = 4;   // Column D in blacklist file

  // Fixed columns in Raw Data for matching
  const COL_PROSPECT_TYPE = 3; // C
  const COL_ADDRESS       = 5; // E
  const COL_ZIP           = 8; // H

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  // ── Input sheet ──
  const input = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!input) { ui.alert(`⚠️ Sheet not found: "${INPUT_SHEET_NAME}"`); return; }

  const lastRow = input.getLastRow();
  const lastCol = input.getLastColumn();
  if (lastRow < 2) { ui.alert(`No data in "${INPUT_SHEET_NAME}".`); return; }

  const rangeAll   = input.getRange(1, 1, lastRow, lastCol);
  const valuesAll  = rangeAll.getValues();
  const displayAll = rangeAll.getDisplayValues();
  const inputHeader = valuesAll[0];

  // ── ZIP Not-A-Fit ──
  const NOT_A_FIT = getNotAFitZipSet_();

  // ── Blacklisted decedents ──
  const blacklisted = getBlacklistedDecedentsById_(
    blacklistSpreadsheetId, blacklistSheetName, blacklistDecedentCol
  );

  // ── Supabase master ──
  ss.toast('⏳ Loading master records from Supabase…', 'Pipeline', 30);

  const masterBundle = buildMasterBundleFromSupabase_(ui);
  if (!masterBundle) return;

  ss.toast(`✅ Master loaded: ${Object.keys(masterBundle.map).length.toLocaleString()} unique addresses`, 'Pipeline', 5);

  const masterMap    = masterBundle.map;       // key → [{rowValues, rowDisplay, idxProspectTypeMaster, sourceLabel}]
  const masterDbCols = masterBundle.dbCols;    // ordered array of DB column names  e.g. ['id','contact_id',...]
  const masterHeader = masterDbCols.map(c => DB_LABELS[c] || c);  // human-readable

  // Index of prospect_type in masterDbCols
  const idxProspectTypeOut = masterDbCols.indexOf('prospect_type');

  // ── Output accumulators ──
  const notFitRows       = [];
  const needsUpdate      = [];
  const skipImport       = [];
  const forImportFormatted = [];

  const matchFoundHeader = masterHeader.concat(['Source']);
  const matchFoundRows   = [];
  const matchFoundKeys   = new Set();

  const debugHeader = [
    'Input Row #', 'Key (Address|Zip)',
    'Input Prospect Type (raw)', 'Input Prospect Type (norm)',
    'Master Source',
    'Master Prospect Type (raw)', 'Master Prospect Type (norm)',
    'Why'
  ];
  const debugRows = [];

  const forImportHeader = [
    "PR File Date","County","Type","Decedent","Second POC Name",
    "Property Address","Property City","Property State","Property Zip",
    "Appraised Value","Representative Name","First Name","Last Name",
    "Representative Address","Representative City","Representative State","Representative Zip",
  ];

  // ── Row loop ──
  for (let i = 1; i < valuesAll.length; i++) {
    const rowV = valuesAll[i];
    const rowD = displayAll[i];

    // ZIP filter
    const zip = normalizeZip_(rowV[COL_ZIP - 1]);
    if (NOT_A_FIT.has(zip)) { notFitRows.push(rowV); continue; }

    const addr = String(rowD[COL_ADDRESS - 1] ?? '').trim();
    const key  = `${addr}|${zip}`;
    const matches = masterMap[key];

    if (!matches || matches.length === 0) {
      // Candidate for For Import
      const formattedRow   = mapRawToForImportRow_(rowV);
      const decedentKey    = normalizeKey_(formattedRow[3]);
      if (decedentKey && blacklisted.has(decedentKey)) continue;  // blacklisted
      forImportFormatted.push(formattedRow);
      continue;
    }

    // Has a match — check Prospect Type
    const ptNewRaw = rowD[COL_PROSPECT_TYPE - 1];
    const ptNew    = normProspectType_(ptNewRaw);

    const anyMatchHasPO = matches.some(m =>
      hasProspectType_(m.rowDisplay[m.idxProspectTypeMaster], 'PO')
    );

    let hasExactMatch = false;
    let debugAdded    = false;

    for (const m of matches) {
      const ptMainRaw = m.rowDisplay[m.idxProspectTypeMaster];
      const ptMain    = normProspectType_(ptMainRaw);

      if (ptNew === ptMain) { hasExactMatch = true; break; }

      if (!debugAdded) {
        debugRows.push([
          i + 1, key,
          String(ptNewRaw ?? ''), ptNew,
          m.sourceLabel,
          String(ptMainRaw ?? ''), ptMain,
          anyMatchHasPO ? 'Skipped Needs Update (master has PO)' : 'Prospect Type differs'
        ]);
        debugAdded = true;
      }
    }

    if (hasExactMatch)   { skipImport.push(rowV); continue; }
    if (anyMatchHasPO)   { continue; }

    needsUpdate.push(rowV);

    // Match Found rows (only once per key)
    if (!matchFoundKeys.has(key)) {
      matchFoundKeys.add(key);
      for (const m of matches) {
        const outRow = m.rowValues.slice();
        while (outRow.length < masterHeader.length) outRow.push('');
        outRow.length = masterHeader.length;

        if (idxProspectTypeOut >= 0) {
          const importPT = rowD[COL_PROSPECT_TYPE - 1];
          const matchPT  = outRow[idxProspectTypeOut];
          outRow[idxProspectTypeOut] = mergeProspectTypesPrefix_(importPT, matchPT);
        }

        outRow.push(m.sourceLabel);
        matchFoundRows.push(outRow);
      }
    }
  }

  // ── Write outputs ──
  outputSheet_(ss, NOT_FIT_SHEET_NAME,  inputHeader,      notFitRows);
  outputSheet_(ss, NEEDS_UPDATE_SHEET,  inputHeader,      needsUpdate);
  outputSheet_(ss, SKIP_IMPORT_SHEET,   inputHeader,      skipImport);
  outputSheet_(ss, FOR_IMPORT_SHEET,    forImportHeader,  forImportFormatted);
  outputSheet_(ss, MATCH_FOUND_SHEET,   matchFoundHeader, matchFoundRows);
  outputSheet_(ss, DEBUG_SHEET_NAME,    debugHeader,      debugRows);

  ui.alert(
    `✅ Complete\n\n` +
    `Not A Fit (Zipcodes): ${notFitRows.length}\n` +
    `For Import (formatted + blacklist-free): ${forImportFormatted.length}\n` +
    `Skip Import: ${skipImport.length}\n` +
    `Needs Update: ${needsUpdate.length}\n` +
    `Match Found rows: ${matchFoundRows.length}\n` +
    `Debug rows: ${debugRows.length}`
  );
}

/******************************************************
 * BUILD MASTER FROM SUPABASE
 * Fetches all rows from the Supabase `contacts` table
 * via the REST API with pagination (1000 rows/batch).
 *
 * Returns:
 *   { map, dbCols }
 *   map    : { "address|zip" : [{rowValues, rowDisplay, idxProspectTypeMaster, sourceLabel}] }
 *   dbCols : ordered array of DB column names (for building header)
 ******************************************************/
function buildMasterBundleFromSupabase_(ui) {

  // Use the dynamic column list loaded from schema_config
  const SELECT_COLS = DYNAMIC_SELECT_COLS || 'contact_id,full_name,phone,email,prospect_type,street_address,postal_code';

  const BATCH = 5000;  // rows per request — fast enough to stay under the 8s timeout

  const map      = {};
  let dbCols     = null;
  let offset     = 0;
  let totalLoaded = 0;

  while (true) {
    const url = `${SUPABASE_URL}/rest/v1/${SUPABASE_TABLE}` +
                `?select=${SELECT_COLS}&limit=${BATCH}&offset=${offset}`;

    let response;
    try {
      response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          'apikey':        SUPABASE_KEY,
          'Authorization': 'Bearer ' + SUPABASE_KEY,
          'Accept':        'application/json',
        },
        muteHttpExceptions: true,
      });
    } catch (e) {
      ui.alert('❌ Network error reaching Supabase:\n' + e);
      return null;
    }

    const code = response.getResponseCode();
    if (code !== 200 && code !== 206) {
      ui.alert(
        `❌ Supabase returned HTTP ${code}\n\n` +
        response.getContentText().substring(0, 500)
      );
      return null;
    }

    let rows;
    try {
      rows = JSON.parse(response.getContentText());
    } catch (e) {
      ui.alert('❌ Failed to parse Supabase response:\n' + e);
      return null;
    }

    if (!rows || rows.length === 0) break; // no more data

    // Capture column order from first batch
    if (!dbCols) dbCols = Object.keys(rows[0]);

    const idxAddr    = dbCols.indexOf('street_address');
    const idxAddrAlt = dbCols.indexOf('subject_property_address');
    const idxPT      = dbCols.indexOf('prospect_type');

    for (const row of rows) {
      const rowValues  = dbCols.map(c => row[c] ?? '');
      const rowDisplay = dbCols.map(c => String(row[c] ?? ''));

      const addr = (
        String(rowDisplay[idxAddr] || '').trim() ||
        String(idxAddrAlt >= 0 ? rowDisplay[idxAddrAlt] : '').trim()
      );
      const zip = normalizeZip_(row['postal_code'] || '');
      const key = `${addr}|${zip}`;

      if (!key || key === '|') continue;

      if (!map[key]) map[key] = [];
      map[key].push({
        rowValues,
        rowDisplay,
        sourceLabel:           'SUPABASE',
        idxProspectTypeMaster: idxPT,
      });
    }

    totalLoaded += rows.length;
    Logger.log(`Loaded ${totalLoaded} master rows so far…`);

    if (rows.length < BATCH) break; // last page
    offset += BATCH;
  }

  Logger.log(`✅ Supabase master loaded: ${totalLoaded} total rows`);

  if (!dbCols) dbCols = [];
  return { map, dbCols };
}



/******************************************************
 * Map Raw Data A:N into the 17-col "For Import" format
 ******************************************************/
function mapRawToForImportRow_(row) {
  const date       = row[0];   // A - PR File Date
  const county     = row[1];   // B - County
  const type       = row[2];   // C - Type
  const decedent   = row[3];   // D - Decedent
  const propAddr   = row[4];   // E - Property Address
  const propCity   = row[5];   // F - Property City
  const propState  = row[6];   // G - Property State
  const propZip    = normalizeZip_(row[7]); // H - Property Zip
  const appraised  = row[8];   // I - Appraised Value
  const repName    = row[9];   // J - Representative Name
  const repAddr    = row[10];  // K - Representative Address
  const repCity    = row[11];  // L - Representative City
  const repState   = row[12];  // M - Representative State
  const repZip     = row[13];  // N - Representative Zip

  const secondPOC = (decedent ? String(decedent) : '') + ' - Deceased';

  const parts     = repName ? String(repName).trim().split(/\s+/).filter(Boolean) : [];
  const firstName = parts.length ? parts[0] : '';
  const lastName  = parts.length > 1 ? parts[parts.length - 1] : '';

  return [
    date, county, type, decedent, secondPOC,
    propAddr, propCity, propState, propZip, appraised,
    repName, firstName, lastName,
    repAddr, repCity, repState, repZip,
  ];
}

/******************************************************
 * BLACKLIST HELPERS
 ******************************************************/
function getBlacklistedDecedentsById_(spreadsheetId, sheetName, decedentCol) {
  const blacklistSS = SpreadsheetApp.openById(spreadsheetId);
  let sheet = null;

  if (sheetName && sheetName.trim()) {
    sheet = blacklistSS.getSheetByName(sheetName);
    if (!sheet) {
      const target = sheetName.trim().toLowerCase();
      sheet = blacklistSS.getSheets().find(s =>
        s.getName().trim().toLowerCase().includes(target)
      ) || null;
    }
    if (!sheet) {
      throw new Error(
        `Blacklist sheet "${sheetName}" not found. ` +
        `Available: [${blacklistSS.getSheets().map(s => s.getName()).join(', ')}]`
      );
    }
  } else {
    sheet = blacklistSS.getSheets()[0];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const values = sheet.getRange(2, decedentCol, lastRow - 1, 1).getValues();
  const set = new Set();
  values.forEach(r => { const k = normalizeKey_(r[0]); if (k) set.add(k); });
  return set;
}

function normalizeKey_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().replace(/\s+/g, ' ').toLowerCase();
}

/******************************************************
 * GENERAL HELPERS
 ******************************************************/
function getNotAFitZipSet_() {
  return new Set([
    '76051','76092','76099','76248','75201','75204','75214','75229',
    '75022','75028','76226','76259','75093','75070','75071','75034',
    '75035','75094','75424','75135',
  ]);
}

function normalizeZip_(zipVal) {
  if (zipVal === null || zipVal === undefined || zipVal === '') return '';
  if (typeof zipVal === 'number') return String(Math.trunc(zipVal)).padStart(5, '0');
  const s = String(zipVal).trim();
  const m = s.match(/^(\d{5})/);
  return m ? m[1] : s;
}

function outputSheet_(spreadsheet, sheetName, outHeader, rows) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  else sheet.clear();

  sheet.getRange(1, 1, 1, outHeader.length).setValues([outHeader]);

  if (rows.length > 0) {
    const normalized = rows.map(r => {
      const rr = r.slice();
      while (rr.length < outHeader.length) rr.push('');
      rr.length = outHeader.length;
      return rr;
    });
    sheet.getRange(2, 1, normalized.length, outHeader.length).setValues(normalized);
  }

  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, outHeader.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, outHeader.length);
  } catch (e) {}
}

function normalizeHeader_(h) {
  return String(h ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function buildHeaderIndexMap_(headersRow) {
  const map = new Map();
  for (let i = 0; i < headersRow.length; i++) {
    const key = normalizeHeader_(headersRow[i]);
    if (key && !map.has(key)) map.set(key, i);
  }
  return map;
}

function mustGetIndex_(map, candidates, label) {
  for (const c of candidates) {
    const key = normalizeHeader_(c);
    if (map.has(key)) return map.get(key);
  }
  throw new Error(`Missing column for ${label}. Looked for: ${candidates.join(', ')}`);
}

function normProspectType_(v) {
  return String(v ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/\s*\/\s*/g, ' / ')
    .toUpperCase();
}

/******************************************************
 * Prospect Type helpers
 ******************************************************/
function splitProspectTypes_(pt) {
  return String(pt ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim().toUpperCase()
    .split('/').map(x => x.trim()).filter(Boolean);
}

function hasProspectType_(ptValue, token) {
  const t = String(token || '').trim().toUpperCase();
  if (!t) return false;
  return splitProspectTypes_(ptValue).includes(t);
}

function joinProspectTypes_(tokens) {
  return tokens.filter(Boolean).join(' / ');
}

function mergeProspectTypesPrefix_(importPT, matchPT) {
  const importTokens = splitProspectTypes_(importPT);
  if (importTokens.length === 0) return String(matchPT ?? '').trim();

  const matchTokens = splitProspectTypes_(matchPT);
  const seen   = new Set(matchTokens);
  const prefix = importTokens.filter(t => !seen.has(t));

  if (prefix.length === 0) return String(matchPT ?? '').trim();
  return joinProspectTypes_([...prefix, ...matchTokens]);
}

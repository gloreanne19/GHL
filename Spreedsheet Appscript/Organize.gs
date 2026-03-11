/******************************************************
 * PIPELINE (NO Organized, NO Cleaned) — UPGRADED
 *
 * Outputs:
 * - Not A Fit (Zipcodes)     (Raw Data rows as-is)
 * - Above 550K ARV           (Raw Data rows as-is; Appraised Value > 550000)
 * - Blacklisted              (17-col formatted rows removed from For Import)
 * - For Import               (Formatted 17 headers + blacklist removed)
 * - Skip Import              (Raw Data rows as-is)
 * - Needs Update             (Raw Data rows as-is BUT:
 *                               - Prospect Type (Col C) is merged: Input + Master
 *                               - Contact Id appended (from Master "Contact Id", col B)
 *                             skipped if any match PT has PO)
 * - Match Found              (Master rows + Source; only for Needs Update)
 * - Debug - Needs Update Reasons
 *
 * ZIP Not-A-Fit checked ONLY ONCE (in audit).
 * Above 550K ARV checked ONLY ONCE (in audit).
 * For Import is built in the 17-column "Organized" header format.
 * Blacklist filtering is applied ONLY to For Import candidates (Decedent),
 * and every removed record is logged to the "Blacklisted" tab.
 *
 * CHANGE:
 * - For Import header "County" is now "Property County"
 * - Property County is guaranteed to be filled from Raw Data:
 *   prefers header match ("Property County" or "County"), otherwise falls back to column B.
 ******************************************************/

function runFullPipeline() {
  auditRawDataAgainstMasters_();
}

/******************************************************
 * AUDIT "Raw Data" AGAINST MASTERS
 * Fixed input columns (Raw Data):
 *   C = Prospect Type
 *   E = Property Address
 *   H = Property Zip
 *   I = Appraised Value
 ******************************************************/
function auditRawDataAgainstMasters_() {
  const MASTER_MAIN_FILE_ID = '1dOVkeCxDsP12-56n9-Io0uWu1FQXk_BFTTa8HWtPFzY';
  const MASTER_2026_FILE_ID = '1rUNbirm7bIrQYWsOAVrqmN9d7BlHMOhngoeIMNPjKNU';

  const MASTER_MAIN_SHEET_NAME = 'GHL ALL Records';
  const MASTER_2026_SHEET_NAME = 'GHL ALL Records_2';

  const INPUT_SHEET_NAME = 'Raw Data';

  const NOT_FIT_SHEET_NAME = 'Not A Fit (Zipcodes)';
  const ABOVE_550K_SHEET_NAME = 'Above 550K ARV';
  const BLACKLISTED_SHEET_NAME = 'Blacklisted';
  const NEEDS_UPDATE_SHEET = 'Needs Update';
  const SKIP_IMPORT_SHEET = 'Skip Import';
  const FOR_IMPORT_SHEET = 'For Import';
  const MATCH_FOUND_SHEET = 'Match Found';
  const DEBUG_SHEET_NAME = 'Debug - Needs Update Reasons';

  // Blacklist spreadsheet
  const blacklistSpreadsheetId = "10GuRF-vFgLG3YRYhTNf0qwx2bNjdZnqZYWvlj7w7EaY";
  const blacklistSheetName = "";  // blank => first sheet
  const blacklistDecedentCol = 4; // Column D in blacklist file

  // Fixed columns in Raw Data for matching
  const COL_PROSPECT_TYPE = 3;   // C
  const COL_ADDRESS = 5;         // E
  const COL_ZIP = 8;             // H
  const COL_APPRAISED_VALUE = 9; // I

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const input = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!input) {
    ui.alert(`⚠️ Sheet not found: "${INPUT_SHEET_NAME}"`);
    return;
  }

  const lastRow = input.getLastRow();
  const lastCol = input.getLastColumn();
  if (lastRow < 2) {
    ui.alert(`No data in "${INPUT_SHEET_NAME}".`);
    return;
  }

  const rangeAll = input.getRange(1, 1, lastRow, lastCol);
  const valuesAll = rangeAll.getValues();
  const displayAll = rangeAll.getDisplayValues();
  const inputHeader = valuesAll[0];

  // Build header index for Raw Data
  const inputHdrMap = buildInputHeaderIndexMap_(inputHeader);

  // ZIP Not-A-Fit list
  const NOT_A_FIT = getNotAFitZipSet_();

  // Load blacklist once
  const blacklisted = getBlacklistedDecedentsById_(
    blacklistSpreadsheetId,
    blacklistSheetName,
    blacklistDecedentCol
  );

  // Build masters
  const masterBundle = buildMasterBundle_([
    { id: MASTER_MAIN_FILE_ID, label: 'MAIN', sheetName: MASTER_MAIN_SHEET_NAME },
    { id: MASTER_2026_FILE_ID, label: '2026', sheetName: MASTER_2026_SHEET_NAME },
  ], ui);
  if (!masterBundle) return;

  const masterMap = masterBundle.map;
  const masterHeader = masterBundle.headerRow;

  // Optional: index of Prospect Type in master output rows
  let idxProspectTypeOut = -1;
  {
    const hdrMap = buildHeaderIndexMap_(masterHeader.map(String));
    try {
      idxProspectTypeOut = mustGetIndex_(hdrMap, ['Prospect Type'], 'Prospect Type (Master Header)');
    } catch (e) {
      idxProspectTypeOut = -1;
    }
  }

  // Contact Id column in Master
  let idxContactIdOut = -1;
  {
    const hdrMap = buildHeaderIndexMap_(masterHeader.map(String));
    try {
      idxContactIdOut = mustGetIndex_(hdrMap, ['Contact Id'], 'Contact Id (Master Header)');
    } catch (e) {
      idxContactIdOut = 1; // fallback to Column B
    }
  }

  // Outputs
  const notFitRows = [];
  const above550kRows = [];
  const skipImport = [];

  const needsUpdateRows = [];
  const needsUpdateHeader = inputHeader.slice();
  needsUpdateHeader.push('Contact Id');

  const forImportFormatted = [];
  const blacklistedFormatted = [];

  const matchFoundHeader = masterHeader.slice();
  matchFoundHeader.push('Source');

  const matchFoundRows = [];
  const matchFoundKeys = new Set();

  const debugHeader = [
    'Input Row #',
    'Key (Address|Zip)',
    'Input Prospect Type (raw)',
    'Input Prospect Type (norm)',
    'Master Source',
    'Master Prospect Type (raw)',
    'Master Prospect Type (norm)',
    'Why'
  ];
  const debugRows = [];

  // For Import exact header
  const forImportHeader = [
    "PR File Date",
    "Property County",
    "Type",
    "Decedent",
    "Second POC Name",
    "Property Address",
    "Property City",
    "Property State",
    "Property Zip",
    "Appraised Value",
    "Representative Name",
    "First Name",
    "Last Name",
    "Representative Address",
    "Representative City",
    "Representative State",
    "Representative Zip",
  ];

  for (let i = 1; i < valuesAll.length; i++) {
    const rowV = valuesAll[i];
    const rowD = displayAll[i];

    const zip = normalizeZip_(rowV[COL_ZIP - 1]);

    // 1) Not A Fit ZIP
    if (NOT_A_FIT.has(zip)) {
      notFitRows.push(rowV);
      continue;
    }

    // 2) Above 550K ARV by Appraised Value > 550000
    const appraisedValue = parseMoneyNumber_(rowV[COL_APPRAISED_VALUE - 1]);
    if (appraisedValue > 550000) {
      above550kRows.push(rowV);
      continue;
    }

    // 3) Compare to masters
    const addr = String(rowD[COL_ADDRESS - 1] ?? '').trim();
    const key = `${addr}|${zip}`;
    const matches = masterMap[key];

    if (!matches || matches.length === 0) {
      // candidate for For Import
      const formattedRow = mapRawToForImportRow_(rowV, inputHdrMap);

      // Blacklist check BEFORE adding to For Import
      const decedentKey = normalizeKey_(formattedRow[3]); // Decedent
      if (decedentKey && blacklisted.has(decedentKey)) {
        blacklistedFormatted.push(formattedRow);
        continue;
      }

      forImportFormatted.push(formattedRow);
      continue;
    }

    // Exact match: Prospect Type ONLY
    const ptNewRaw = rowD[COL_PROSPECT_TYPE - 1];
    const ptNew = normProspectType_(ptNewRaw);

    const anyMatchHasPO = matches.some(m => hasProspectType_(m.rowDisplay[m.idxProspectTypeMaster], 'PO'));

    let hasExactMatch = false;
    let debugAdded = false;

    for (const m of matches) {
      const ptMainRaw = m.rowDisplay[m.idxProspectTypeMaster];
      const ptMain = normProspectType_(ptMainRaw);

      if (ptNew === ptMain) {
        hasExactMatch = true;
        break;
      }

      if (!debugAdded) {
        debugRows.push([
          i + 1,
          key,
          String(ptNewRaw ?? ''),
          ptNew,
          m.sourceLabel,
          String(ptMainRaw ?? ''),
          ptMain,
          anyMatchHasPO ? 'Skipped Needs Update (master has PO)' : 'Prospect Type differs'
        ]);
        debugAdded = true;
      }
    }

    if (hasExactMatch) {
      skipImport.push(rowV);
      continue;
    }

    if (anyMatchHasPO) {
      continue;
    }

    // Needs Update output row
    let best = matches[0];
    for (const m of matches) {
      if (m.sourceLabel === 'MAIN') {
        best = m;
        break;
      }
    }

    const masterPtRaw = best.rowDisplay[best.idxProspectTypeMaster];
    const mergedPT = mergeProspectTypesPrefix_(ptNewRaw, masterPtRaw);

    const contactId = (idxContactIdOut >= 0 && best.rowValues.length > idxContactIdOut)
      ? (best.rowValues[idxContactIdOut] ?? '')
      : '';

    const outNeeds = rowV.slice();
    outNeeds[COL_PROSPECT_TYPE - 1] = mergedPT;
    outNeeds.push(contactId);

    needsUpdateRows.push(outNeeds);

    // Match Found rows
    if (!matchFoundKeys.has(key)) {
      matchFoundKeys.add(key);

      for (const m of matches) {
        const outRow = m.rowValues.slice();
        while (outRow.length < masterHeader.length) outRow.push('');
        outRow.length = masterHeader.length;

        if (idxProspectTypeOut >= 0) {
          const importPT = rowD[COL_PROSPECT_TYPE - 1];
          const matchPT = outRow[idxProspectTypeOut];
          outRow[idxProspectTypeOut] = mergeProspectTypesPrefix_(importPT, matchPT);
        }

        outRow.push(m.sourceLabel);
        matchFoundRows.push(outRow);
      }
    }
  }

  // Write outputs
  outputSheet_(ss, NOT_FIT_SHEET_NAME, inputHeader, notFitRows);
  outputSheet_(ss, ABOVE_550K_SHEET_NAME, inputHeader, above550kRows);
  outputSheet_(ss, NEEDS_UPDATE_SHEET, needsUpdateHeader, needsUpdateRows);
  outputSheet_(ss, SKIP_IMPORT_SHEET, inputHeader, skipImport);
  outputSheet_(ss, BLACKLISTED_SHEET_NAME, forImportHeader, blacklistedFormatted);
  outputSheet_(ss, FOR_IMPORT_SHEET, forImportHeader, forImportFormatted);
  outputSheet_(ss, MATCH_FOUND_SHEET, matchFoundHeader, matchFoundRows);
  outputSheet_(ss, DEBUG_SHEET_NAME, debugHeader, debugRows);

  ui.alert(
    `✅ Complete

Not A Fit (Zipcodes): ${notFitRows.length}
Above 550K ARV: ${above550kRows.length}
Blacklisted (removed from For Import): ${blacklistedFormatted.length}
For Import (formatted + blacklist-free): ${forImportFormatted.length}
Skip Import: ${skipImport.length}
Needs Update (merged PT + Contact Id): ${needsUpdateRows.length}
Match Found rows: ${matchFoundRows.length}
Debug rows: ${debugRows.length}`
  );
}

/******************************************************
 * Map Raw Data A:N into the 17-col "For Import" format
 ******************************************************/
function mapRawToForImportRow_(row, inputHdrMap) {
  const date = row[0]; // A

  const propertyCounty =
    getByHeader_(row, inputHdrMap, ['Property County', 'County']) ||
    row[1] || ''; // B fallback

  const type = row[2];     // C
  const decedent = row[3]; // D

  const propAddr = row[4];  // E
  const propCity = row[5];  // F
  const propState = row[6]; // G
  const propZip = normalizeZip_(row[7]); // H

  const appraised = row[8]; // I

  const repName = row[9];   // J
  const repAddr = row[10];  // K
  const repCity = row[11];  // L
  const repState = row[12]; // M
  const repZip = row[13];   // N

  const secondPOC = (decedent ? String(decedent) : "") + " - Deceased";

  const parts = repName ? String(repName).trim().split(/\s+/).filter(Boolean) : [];
  const firstName = parts.length ? parts[0] : "";
  const lastName = parts.length ? parts[parts.length - 1] : "";

  return [
    date,
    propertyCounty,
    type,
    decedent,
    secondPOC,
    propAddr,
    propCity,
    propState,
    propZip,
    appraised,
    repName,
    firstName,
    lastName,
    repAddr,
    repCity,
    repState,
    repZip
  ];
}

/******************************************************
 * MASTER LOADING
 ******************************************************/
function buildMasterBundle_(configs, ui) {
  const map = {};
  let headerRow = null;

  for (const cfg of configs) {
    let masterSS, masterSheet;
    try {
      masterSS = SpreadsheetApp.openById(cfg.id);
      masterSheet = masterSS.getSheetByName(cfg.sheetName);
    } catch (e) {
      ui.alert(`⚠️ Failed to open master (${cfg.label}).\n${e}`);
      return null;
    }

    if (!masterSheet) {
      const names = masterSS.getSheets().map(s => s.getName()).join('\n');
      ui.alert(
        `⚠️ Master sheet not found in (${cfg.label}): "${cfg.sheetName}"\n\nSheets in that file:\n${names}`
      );
      return null;
    }

    const r = masterSheet.getDataRange();
    const masterValues = r.getValues();
    const masterDisplay = r.getDisplayValues();
    if (masterValues.length < 2) continue;

    if (!headerRow) headerRow = masterValues[0];

    const hdrMap = buildHeaderIndexMap_(masterDisplay[0]);

    const idxAddrMaster = mustGetIndex_(
      hdrMap,
      ['Street address', 'Subject Property Address', 'Property Address', 'Address'],
      `Master Address (${cfg.label})`
    );
    const idxZipMaster = mustGetIndex_(
      hdrMap,
      ['Postal Code', 'Zip', 'Zipcode', 'Mailing Zipcode'],
      `Master Postal Code (${cfg.label})`
    );
    const idxProspectTypeMaster = mustGetIndex_(
      hdrMap,
      ['Prospect Type'],
      `Master Prospect Type (${cfg.label})`
    );

    for (let i = 1; i < masterValues.length; i++) {
      const rowV = masterValues[i];
      const rowD = masterDisplay[i];

      const addr = String(rowD[idxAddrMaster] ?? '').trim();
      const zip = String(rowD[idxZipMaster] ?? '').trim();
      const key = `${addr}|${zip}`;

      if (!map[key]) map[key] = [];
      map[key].push({
        rowValues: rowV,
        rowDisplay: rowD,
        sourceLabel: cfg.label,
        idxProspectTypeMaster
      });
    }
  }

  if (!headerRow) headerRow = [];
  return { map, headerRow };
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
      sheet = blacklistSS.getSheets().find(s => s.getName().trim().toLowerCase() === target) || null;
    }
    if (!sheet) {
      const target = sheetName.trim().toLowerCase();
      sheet = blacklistSS.getSheets().find(s => s.getName().trim().toLowerCase().includes(target)) || null;
    }
    if (!sheet) {
      const available = blacklistSS.getSheets().map(s => s.getName()).join(", ");
      throw new Error(
        `Blacklist sheet "${sheetName}" not found in spreadsheet ID ${spreadsheetId}. ` +
        `Available sheets: [${available}]`
      );
    }
  } else {
    sheet = blacklistSS.getSheets()[0];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const values = sheet.getRange(2, decedentCol, lastRow - 1, 1).getValues();

  const set = new Set();
  values.forEach(r => {
    const key = normalizeKey_(r[0]);
    if (key) set.add(key);
  });

  return set;
}

function normalizeKey_(v) {
  if (v === null || v === undefined) return "";
  const s = String(v).trim().replace(/\s+/g, " ");
  return s ? s.toLowerCase() : "";
}

/******************************************************
 * GENERAL HELPERS
 ******************************************************/
function getNotAFitZipSet_() {
  return new Set([
    "76051","76092","76099","76248","75201","75204","75214","75229",
    "75022","75028","76226","76259","75093","75070","75071","75034",
    "75035","75094","75424","75135"
  ]);
}

function normalizeZip_(zipVal) {
  if (zipVal === null || zipVal === undefined || zipVal === "") return "";
  if (typeof zipVal === "number") return String(Math.trunc(zipVal)).padStart(5, "0");
  const s = String(zipVal).trim();
  const m = s.match(/^(\d{5})/);
  return m ? m[1] : s;
}

function parseMoneyNumber_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const cleaned = String(v).replace(/[^0-9.\-]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
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
    sheet.getRange(1, 1, 1, outHeader.length).setFontWeight("bold");
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
  throw new Error(`Missing required column for ${label}. Looked for: ${candidates.join(', ')}`);
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
  const s = String(pt ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toUpperCase();
  if (!s) return [];
  return s.split('/').map(x => x.trim()).filter(Boolean);
}

function hasProspectType_(ptValue, token) {
  const t = String(token || '').trim().toUpperCase();
  if (!t) return false;
  const tokens = splitProspectTypes_(ptValue);
  return tokens.includes(t);
}

function joinProspectTypes_(tokens) {
  return tokens.filter(Boolean).join(' / ');
}

function mergeProspectTypesPrefix_(importPT, matchPT) {
  const importTokens = splitProspectTypes_(importPT);
  if (importTokens.length === 0) return String(matchPT ?? '').trim();

  const matchTokens = splitProspectTypes_(matchPT);
  const seen = new Set(matchTokens);

  const prefix = [];
  for (const t of importTokens) {
    if (!seen.has(t)) prefix.push(t);
  }

  if (prefix.length === 0) return String(matchPT ?? '').trim();
  return joinProspectTypes_([...prefix, ...matchTokens]);
}

/******************************************************
 * Raw Data header helpers
 ******************************************************/
function buildInputHeaderIndexMap_(headersRow) {
  const map = new Map();
  for (let i = 0; i < headersRow.length; i++) {
    const key = String(headersRow[i] ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
    if (key && !map.has(key)) map.set(key, i);
  }
  return map;
}

function getByHeader_(row, headerMap, names) {
  for (const name of names) {
    const key = String(name).trim().toLowerCase().replace(/\s+/g, ' ');
    if (headerMap.has(key)) return row[headerMap.get(key)];
  }
  return '';
}
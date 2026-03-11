function ST_appendForImportRepData() {
  const ss = SpreadsheetApp.getActive();
  const stf = ss.getSheetByName("Skip Trace Finished");
  const fin = ss.getSheetByName("For Import");

  if (!stf || !fin) return;

  // --------------------------------------------------
  // STEP 0 — Delete "Property County" column in STF
  // --------------------------------------------------
  const firstRow = stf.getRange(1, 1, 1, stf.getLastColumn()).getValues()[0];
  const propertyCountyIdx = firstRow.findIndex(h =>
    String(h || "").trim().toLowerCase() === "property county"
  );

  if (propertyCountyIdx !== -1) {
    stf.deleteColumn(propertyCountyIdx + 1); // sheet columns are 1-based
  }

  // --------------------------------------------------
  // STEP 1 — Rename headers in Skip Trace Finished
  // --------------------------------------------------
  const headers = stf.getRange(1, 1, 1, stf.getLastColumn()).getValues()[0];

  const renameMap = {
    "Property Address": "Mailing Add",
    "Property City": "Mailing Cit",
    "Property State": "Mailing ST",
    "Property Zip": "Mailing Zip"
  };

  const newHeaders = headers.map(h => renameMap[h] || h);
  stf.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

  // --------------------------------------------------
  // STEP 2 — Load data
  // --------------------------------------------------
  if (stf.getLastRow() < 2 || fin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No data to process.");
    return;
  }

  const stfData = stf.getRange(2, 1, stf.getLastRow() - 1, stf.getLastColumn()).getValues();
  const finData = fin.getRange(2, 1, fin.getLastRow() - 1, fin.getLastColumn()).getValues();
  const finHeaders = fin.getRange(1, 1, 1, fin.getLastColumn()).getValues()[0];

  const START_COL = 239; // IE

  // --------------------------------------------------
  // STEP 3 — Write For Import headers starting at IE
  // --------------------------------------------------
  stf.getRange(1, START_COL, 1, finHeaders.length).setValues([finHeaders]);

  // --------------------------------------------------
  // STEP 4 — Build lookup map from For Import
  // Match:
  // STF A + D
  // vs
  // For Import N + Q
  // --------------------------------------------------
  const map = new Map();

  for (const r of finData) {
    const addr = normalizeAddr_(r[13]); // N
    const zip = normalizeZip_(r[16]);   // Q

    if (!addr || !zip) continue;

    const key = addr + "|" + zip;
    map.set(key, r);
  }

  // --------------------------------------------------
  // STEP 5 — Append matching For Import rows into STF at IE
  // --------------------------------------------------
  let matches = 0;

  for (let i = 0; i < stfData.length; i++) {
    const addr = normalizeAddr_(stfData[i][0]); // A
    const zip = normalizeZip_(stfData[i][3]);   // D

    const key = addr + "|" + zip;
    const row = map.get(key);

    if (row) {
      stf.getRange(i + 2, START_COL, 1, row.length).setValues([row]);
      matches++;
    }
  }

  // --------------------------------------------------
  // STEP 6 — Force appended phone/zip/LTV columns to plain text
  // --------------------------------------------------
  forceAppendedPlainTextColumns_(stf, finHeaders, START_COL, stf.getLastRow() - 1, [
    "Phone",
    "Additional Phone",
    "Landline",
    "Additional Landline",
    "Additional Landlines",
    "Property Zip",
    "Mailing Zipcode",
    "Representative Zip",
    "LTV"
  ]);

  // Optional: force ZIP-like appended columns to 5 digits
  forceAppendedZip5_(stf, finHeaders, START_COL, stf.getLastRow() - 1, [
    "Property Zip",
    "Mailing Zipcode",
    "Representative Zip"
  ]);

  SpreadsheetApp.getUi().alert("Representative rows appended: " + matches);
}

function normalizeAddr_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeZip_(z) {
  const m = String(z || '').match(/(\d{5})/);
  return m ? m[1] : '';
}

function normHeader_(h) {
  return String(h || "").trim().toLowerCase();
}

function findHeaderIdx_(headers, names) {
  const wants = names.map(normHeader_);
  for (let i = 0; i < headers.length; i++) {
    if (wants.includes(normHeader_(headers[i]))) return i;
  }
  return -1;
}

function forceAppendedPlainTextColumns_(sheet, headers, startCol, rowCount, headerNames) {
  if (rowCount <= 0) return;

  for (const headerName of headerNames) {
    const idx0 = findHeaderIdx_(headers, [headerName]);
    if (idx0 === -1) continue;

    const col = startCol + idx0;
    const rng = sheet.getRange(2, col, rowCount, 1);

    // Plain text format
    rng.setNumberFormat('@');

    const vals = rng.getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === "" || vals[i][0] == null) continue;
      vals[i][0] = String(vals[i][0]).trim();
    }
    rng.setValues(vals);
  }
}

function forceAppendedZip5_(sheet, headers, startCol, rowCount, headerNames) {
  if (rowCount <= 0) return;

  for (const headerName of headerNames) {
    const idx0 = findHeaderIdx_(headers, [headerName]);
    if (idx0 === -1) continue;

    const col = startCol + idx0;
    const rng = sheet.getRange(2, col, rowCount, 1);

    rng.setNumberFormat('@');

    const vals = rng.getValues();
    for (let i = 0; i < vals.length; i++) {
      const m = String(vals[i][0] || "").match(/(\d{5})/);
      vals[i][0] = m ? m[1] : "";
    }
    rng.setValues(vals);
  }
}
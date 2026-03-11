/**
 * SET DIRECT SKIP RESULTS COLUMNS (STANDARDIZE + DELETE EXTRAS)
 *
 * Sheet: "Direct Skip Results"
 *
 * Behavior:
 * 1) Renames any matching/synonym headers to the exact STANDARD headers (in place).
 * 2) Inserts any missing STANDARD columns in the correct order (shifts right as needed).
 * 3) Deletes any columns NOT in the STANDARD list (this is where "extra columns" are removed).
 *
 * ✅ End state: The sheet will contain ONLY these columns, in this exact order:
 * Input Last Name, Input First Name, Input Mailing Address, Input Mailing City, Input Mailing State, Input Mailing Zip,
 * Input Property Address, Input Property City, Input Property State, Input Property Zip,
 * Input Custom Field 1, Input Custom Field 2, Input Custom Field 3,
 * ResultCode, Matched First Name, Matched Last Name, Age, Deceased,
 * Phone1, Phone1 Type, Phone2, Phone2 Type, Phone3, Phone3 Type, Phone4, Phone4 Type, Phone5, Phone5 Type, Phone6, Phone6 Type, Phone7, Phone7 Type,
 * Email1, Email2
 *
 * ⚠️ IMPORTANT:
 * - Deleting extra columns WILL permanently remove those columns and their data.
 * - Inserting missing columns will shift columns to the right (records stay in the same rows).
 */
function setDirectSkipResultsColumns_DeleteExtras() {
  const SHEET_NAME = "Direct Skip Results";

  const STANDARD_HEADERS = [
    "Input Last Name",
    "Input First Name",
    "Input Mailing Address",
    "Input Mailing City",
    "Input Mailing State",
    "Input Mailing Zip",
    "Input Property Address",
    "Input Property City",
    "Input Property State",
    "Input Property Zip",
    "Input Custom Field 1",
    "Input Custom Field 2",
    "Input Custom Field 3",
    "ResultCode",
    "Matched First Name",
    "Matched Last Name",
    "Age",
    "Deceased",
    "Phone1",
    "Phone1 Type",
    "Phone2",
    "Phone2 Type",
    "Phone3",
    "Phone3 Type",
    "Phone4",
    "Phone4 Type",
    "Phone5",
    "Phone5 Type",
    "Phone6",
    "Phone6 Type",
    "Phone7",
    "Phone7 Type",
    "Email1",
    "Email2"
  ];

  // Optional synonym mapping (add variants you see from exports)
  const SYNONYMS = {
    "ResultCode": ["result code", "result_code", "result"],
    "Deceased": ["is deceased", "deceased flag"],
    "Input Mailing Zip": ["input mailing zipcode", "input mailing zip code", "mailing zip", "mailing zipcode", "mailing zip code"],
    "Input Property Zip": ["input property zipcode", "input property zip code", "property zip", "property zipcode", "property zip code"],
    "Email1": ["email 1", "email_1", "email"],
    "Email2": ["email 2", "email_2"]
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const hdrRange = sh.getRange(1, 1, 1, lastCol);
  let headers = hdrRange.getValues()[0].map(h => String(h || "").trim());

  // -------------------------
  // 1) Rename known headers IN PLACE to standard names
  // -------------------------
  const stdNormToStd = new Map(STANDARD_HEADERS.map(h => [ds_norm_(h), h]));
  const synNormToStd = new Map();
  for (const std of Object.keys(SYNONYMS)) {
    for (const s of SYNONYMS[std]) synNormToStd.set(ds_norm_(s), std);
  }

  for (let c = 0; c < headers.length; c++) {
    const cur = headers[c];
    const curN = ds_norm_(cur);
    if (!curN) continue;

    // already standard (maybe different case/spacing)
    if (stdNormToStd.has(curN)) {
      const std = stdNormToStd.get(curN);
      if (cur !== std) sh.getRange(1, c + 1).setValue(std);
      continue;
    }

    // synonym -> standard
    if (synNormToStd.has(curN)) {
      const std = synNormToStd.get(curN);
      sh.getRange(1, c + 1).setValue(std);
      headers[c] = std;
    }
  }

  // Refresh headers after renames
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());

  // -------------------------
  // 2) Insert missing standard columns IN ORDER (no data rewrite)
  // -------------------------
  for (let i = 0; i < STANDARD_HEADERS.length; i++) {
    const needed = STANDARD_HEADERS[i];
    const neededN = ds_norm_(needed);

    // refresh to handle shifting as we insert
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());

    const curIndex = headers.findIndex(h => ds_norm_(h) === neededN);
    if (curIndex === -1) {
      const insertAt = i + 1; // 1-based desired position
      if (insertAt <= sh.getLastColumn()) {
        sh.insertColumnBefore(insertAt);
        sh.getRange(1, insertAt).setValue(needed);
      } else {
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, sh.getLastColumn()).setValue(needed);
      }
    }
  }

  // -------------------------
  // 3) Delete extra columns (NOT in standard list)
  //    Do this right-to-left to keep indexes stable
  // -------------------------
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const allowed = new Set(STANDARD_HEADERS.map(h => ds_norm_(h)));

  for (let col = sh.getLastColumn(); col >= 1; col--) {
    const h = headers[col - 1];
    const hn = ds_norm_(h);
    if (!hn) {
      // if header is blank, treat as extra and delete
      sh.deleteColumn(col);
      continue;
    }
    if (!allowed.has(hn)) {
      sh.deleteColumn(col);
    }
  }

  // -------------------------
  // 4) Ensure final header order matches STANDARD exactly
  //    (We do NOT move data. We only rewrite header row values across existing columns.)
  // -------------------------
  const finalCols = sh.getLastColumn();
  if (finalCols !== STANDARD_HEADERS.length) {
    // If something unexpected happened, expand/shrink to match
    if (finalCols < STANDARD_HEADERS.length) {
      sh.insertColumnsAfter(finalCols, STANDARD_HEADERS.length - finalCols);
    } else if (finalCols > STANDARD_HEADERS.length) {
      // delete extra at end
      for (let c = finalCols; c > STANDARD_HEADERS.length; c--) sh.deleteColumn(c);
    }
  }

  sh.getRange(1, 1, 1, STANDARD_HEADERS.length).setValues([STANDARD_HEADERS]);

  // Formatting
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, STANDARD_HEADERS.length).setFontWeight("bold");
  sh.autoResizeColumns(1, STANDARD_HEADERS.length);
}

/** normalize header for matching */
function ds_norm_(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ");
}function myFunction() {
  
}

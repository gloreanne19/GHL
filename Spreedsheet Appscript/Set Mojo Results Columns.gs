function setMojoResultsColumns_InsertMissingInOrder() {
  const SHEET_NAME = "Mojo Results";

  const STANDARD_HEADERS = [
    "Property Address","Property City","Property State","Property Zip Code",
    "Full Name","First Name","Last Name","Second Name",
    "Mailing Address","Mailing City","Mailing State","Mailing Zip Code",
    "Email 1","Email 2","Email 3","Email 4","Email 5","Email 6",
    "Primary Phone",
    "Mobile 1","Mobile 2","Mobile 3","Mobile 4","Mobile 5","Mobile 6","Mobile 7","Mobile 8","Mobile 9","Mobile 10",
    "Phone 1","Phone 2","Phone 3","Phone 4","Phone 5","Phone 6","Phone 7","Phone 8","Phone 9","Phone 10"
  ];

  // Add common export variants here (optional). The script also matches exact names.
  const SYNONYMS = {
    "Property Address": ["address", "street address", "subject property address", "propertyaddress"],
    "Property City": ["city", "propertycity"],
    "Property State": ["state", "propertystate"],
    "Property Zip Code": ["zip", "zipcode", "zip code", "postal code", "property zip", "property zipcode"],

    "Full Name": ["name", "fullname"],
    "First Name": ["firstname", "first"],
    "Last Name": ["lastname", "last"],
    "Second Name": ["secondname", "secondary name"],

    "Mailing Address": ["mail address", "mailingaddress"],
    "Mailing City": ["mailingcity"],
    "Mailing State": ["mailingstate"],
    "Mailing Zip Code": ["mailing zip", "mailing zipcode", "mailing zip code", "mailing postal code"],

    "Primary Phone": ["primaryphone", "main phone", "primary number"]
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  // Read current header row
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());

  // Build map: normalized existing header -> column index (0-based)
  const existing = new Map();
  for (let c = 0; c < headers.length; c++) {
    const n = normHeader_(headers[c]);
    if (n && !existing.has(n)) existing.set(n, c);
  }

  // 1) Rename known columns IN PLACE (header cell only)
  // We do this first so subsequent “missing?” checks are more accurate.
  for (let c = 0; c < headers.length; c++) {
    const cur = headers[c];
    const curN = normHeader_(cur);
    if (!curN) continue;

    // If it already matches a standard header, keep as-is
    const directStd = STANDARD_HEADERS.find(h => normHeader_(h) === curN);
    if (directStd) {
      if (cur !== directStd) sheet.getRange(1, c + 1).setValue(directStd);
      continue;
    }

    // Try to map via synonyms to a standard header
    let mapped = null;
    for (const std of STANDARD_HEADERS) {
      const syns = SYNONYMS[std] || [];
      for (const s of syns) {
        if (normHeader_(s) === curN) {
          mapped = std;
          break;
        }
      }
      if (mapped) break;
    }

    if (mapped) {
      sheet.getRange(1, c + 1).setValue(mapped);
      headers[c] = mapped; // update local copy
    }
  }

  // Helper to refresh headers array as we insert columns
  function refreshHeaders_() {
    const lc = sheet.getLastColumn();
    return sheet.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  }

  // 2) Enforce STANDARD_HEADERS order by inserting missing columns at the correct positions
  //    We only insert blanks. Existing columns remain, but may shift right when we insert.
  let currentHeaders = refreshHeaders_();

  for (let i = 0; i < STANDARD_HEADERS.length; i++) {
    const needed = STANDARD_HEADERS[i];
    const neededN = normHeader_(needed);

    // Recompute current position of this needed header (after any inserts)
    currentHeaders = refreshHeaders_();
    const curIndex = currentHeaders.findIndex(h => normHeader_(h) === neededN);

    if (curIndex === -1) {
      // Missing => insert at position i+1 (1-based column index)
      // If sheet currently has fewer than i columns, add to the end until we reach i.
      const insertAtCol = Math.min(i + 1, sheet.getLastColumn() + 1); // 1-based insertion position

      // Insert ONE column BEFORE insertAtCol (or at end)
      if (insertAtCol <= sheet.getLastColumn()) {
        sheet.insertColumnBefore(insertAtCol);
        sheet.getRange(1, insertAtCol).setValue(needed);
      } else {
        // insert at end
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue(needed);
      }
    } else {
      // Present: do nothing. We are not reordering existing columns (only inserting missing).
      // (If it’s in a different position, we leave it—this avoids moving data.)
    }
  }

  // Formatting
  const finalCols = sheet.getLastColumn();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, finalCols).setFontWeight("bold");
  sheet.autoResizeColumns(1, Math.min(finalCols, 40));
}

function normHeader_(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ");
}
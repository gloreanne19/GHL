/**
 * READY FOR XLEADS (FROM "For Import")
 * Creates/updates a tab named "Ready for XLeads" using data from "For Import".
 *
 * Output headers (in order):
 *  Street Address  <- Column N (Representative Address)
 *  City            <- Column O (Representative City)
 *  State           <- Column P (Representative State)
 *  Zip             <- Column Q (Representative Zip)
 *
 * - Drops rows where all 4 fields are blank
 * - Forces Zip column to PLAIN TEXT to prevent date conversion
 */
function readyForXLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceName = "For Import";
  const targetName = "Ready 4 XLeads";

  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error(`Sheet not found: ${sourceName}`);

  const target = ss.getSheetByName(targetName) || ss.insertSheet(targetName);

  const headers = ["Street Address", "City", "State", "Zip"];

  // Reset output sheet
  target.clearContents();
  target.getRange(1, 1, 1, headers.length).setValues([headers]);

  const lastRow = source.getLastRow();
  if (lastRow < 2) {
    forceZipTextFormat_(target);
    formatSheet_(target, headers.length);
    return;
  }

  // Column references (1-based):
  // N = 14 (Representative Address)
  // O = 15 (Representative City)
  // P = 16 (Representative State)
  // Q = 17 (Representative Zip)

  const values = source.getRange(2, 14, lastRow - 1, 4).getValues();

  const output = [];

  values.forEach(row => {
    const cleanedRow = [
      row[0], // N
      row[1], // O
      row[2], // P
      normalizeZipText_(row[3]) // Q
    ];

    if (cleanedRow.some(v => String(v || "").trim() !== "")) {
      output.push(cleanedRow);
    }
  });

  // Force Zip column to plain text BEFORE writing
  forceZipTextFormat_(target);

  if (output.length) {
    target.getRange(2, 1, output.length, headers.length).setValues(output);
  }

  formatSheet_(target, headers.length);
}

/** Forces column D (Zip) to Plain Text */
function forceZipTextFormat_(sheet) {
  sheet.getRange(2, 4, Math.max(sheet.getMaxRows() - 1, 1), 1)
       .setNumberFormat("@");
}

/** Ensure ZIP stays a 5-digit text value (keeps leading zeros). */
function normalizeZipText_(zipVal) {
  if (!zipVal) return "";

  if (Object.prototype.toString.call(zipVal) === "[object Date]" && !isNaN(zipVal)) {
    return "";
  }

  if (typeof zipVal === "number") {
    return String(Math.trunc(zipVal)).padStart(5, "0");
  }

  const s = String(zipVal).trim();
  const m = s.match(/^(\d{5})/);
  return m ? m[1] : s;
}

/** Formatting helper */
function formatSheet_(sheet, headerCount) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headerCount).setFontWeight("bold");
  sheet.autoResizeColumns(1, headerCount);
}
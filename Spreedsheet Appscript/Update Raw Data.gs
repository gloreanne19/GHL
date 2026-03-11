/**
 * Update Raw Data headers and set Prospect Type values to PO
 *
 * Header changes:
 * Date   -> PR File Date
 * County -> Property County
 * Type   -> Prospect Type
 *
 * Then updates every row under Prospect Type to "PO"
 */

function RD_updateRawDataHeadersAndPT() {

  const SHEET_NAME = "Raw Data";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastCol === 0) return;

  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  const header = headerRange.getValues()[0];

  const norm = v => String(v ?? "").trim().toLowerCase();

  let idxCounty = -1;
  let idxType = -1;

  for (let i = 0; i < header.length; i++) {

    const h = norm(header[i]);

    if (h === "date") {
      header[i] = "PR File Date";
    }

    if (h === "county") {
      idxCounty = i;
      header[i] = "Property County";
    }

    if (h === "type" || h === "prospect type") {
      idxType = i;
      header[i] = "Prospect Type";
    }
  }

  headerRange.setValues([header]);

  // Fill Prospect Type column with PO
  if (idxType >= 0 && lastRow > 1) {
    const values = Array(lastRow - 1).fill(["PO"]);
    sheet.getRange(2, idxType + 1, lastRow - 1, 1).setValues(values);
  }

  SpreadsheetApp.getUi().alert(
    "✅ Headers updated:\n" +
    "Date → PR File Date\n" +
    "County → Property County\n" +
    "Type → Prospect Type\n\n" +
    "All Prospect Type rows set to PO."
  );
}
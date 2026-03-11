function readyForMojo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceName = "Not in XLeads";
  const targetName = "Ready for Mojo";

  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error(`Sheet not found: ${sourceName}`);

  const target = ss.getSheetByName(targetName) || ss.insertSheet(targetName);

  const headers = ["Street Address", "City", "State", "Zip"];

  // Clear and write headers
  target.clearContents();
  target.getRange(1, 1, 1, headers.length).setValues([headers]);

  const lastRow = source.getLastRow();
  if (lastRow < 2) {
    formatSheet_(target, headers.length);
    return;
  }

  // Column references:
  // N = 14 (Representative Address)
  // O = 15 (Representative City)
  // P = 16 (Representative State)
  // Q = 17 (Representative Zip)

  const repAddress = source.getRange(2, 14, lastRow - 1, 1).getValues();
  const repCity    = source.getRange(2, 15, lastRow - 1, 1).getValues();
  const repState   = source.getRange(2, 16, lastRow - 1, 1).getValues();
  const repZip     = source.getRange(2, 17, lastRow - 1, 1).getValues();

  const output = [];

  for (let i = 0; i < repAddress.length; i++) {
    const row = [
      repAddress[i][0],
      repCity[i][0],
      repState[i][0],
      repZip[i][0]
    ];

    // Remove completely blank rows
    if (row.some(v => String(v || "").trim() !== "")) {
      output.push(row);
    }
  }

  if (output.length) {
    target.getRange(2, 1, output.length, headers.length).setValues(output);
  }

  formatSheet_(target, headers.length);
}
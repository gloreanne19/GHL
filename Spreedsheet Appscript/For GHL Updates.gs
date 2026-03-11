function createForGHLUpdates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet = ss.getSheetByName("Needs Update");
  const afterSheet = ss.getSheetByName("Need Manual Dm");

  if (!sourceSheet) {
    throw new Error("Needs Update sheet not found.");
  }

  // Delete old sheet if it exists
  const existing = ss.getSheetByName("For GHL Updates");
  if (existing) ss.deleteSheet(existing);

  // Create new sheet
  const targetSheet = ss.insertSheet("For GHL Updates");

  // Move it after "Need Manual Dm"
  if (afterSheet) {
    ss.setActiveSheet(targetSheet);
    ss.moveActiveSheet(afterSheet.getIndex() + 1);
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const output = [];

  // Headers
  output.push([
    "Contact ID",       // A
    "Prospect Type",    // B
    "Second POC Name",  // C
    "First Name",       // D
    "Last Name",        // E
    "Mailing Street",   // F
    "Mailing City",     // G
    "Mailing State",    // H
    "Mailing Zip"       // I
  ]);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const contactId = row[16]; // Column Q
    const prospectType = row[2]; // Column C
    const decedent = row[3] ? row[3] + " - Deceased" : ""; // Column D
    const firstName = row[10]; // Column K
    const lastName = row[11]; // Column L
    const street = row[12]; // Column M
    const city = row[13]; // Column N
    const state = row[14]; // Column O
    const zip = row[15]; // Column P

    output.push([
      contactId,
      prospectType,
      decedent,
      firstName,
      lastName,
      street,
      city,
      state,
      zip
    ]);
  }

  targetSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}function myFunction() {
  
}

/**
 * Prepend New Import
 * ─────────────────
 * Reads data from "Reorder Columns For Prepend" sheet and
 * upserts it directly into Supabase (replaces old hardcoded
 * "GHL ALL Records_2" spreadsheet approach).
 *
 * Requires DB_Connection.gs to be in the same Apps Script project.
 */
function prependNewImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const importSheet = ss.getSheetByName('Reorder Columns For Prepend');
  if (!importSheet) {
    ui.alert('⚠️ "Reorder Columns For Prepend" sheet not found.');
    return;
  }

  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data to import.');
    return;
  }

  // Read headers from row 1
  const headers = importSheet
    .getRange(1, 1, 1, importSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());

  // Use DisplayValues to preserve date formatting (avoids timezone shifts)
  const dataRows = importSheet
    .getRange(2, 1, lastRow - 1, importSheet.getLastColumn())
    .getDisplayValues();

  // Load the Excel→DB column mapping from Supabase schema_config
  const colMap = DB_getColMap(); // defined in DB_Connection.gs

  // Map each row to a DB object using the schema
  const rows = dataRows.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      const dbCol = colMap[header];
      if (dbCol && row[i] !== '') {
        obj[dbCol] = row[i];
      }
    });
    return obj;
  }).filter(obj => Object.keys(obj).length > 0);

  if (rows.length === 0) {
    ui.alert(
      'No mappable rows found.\n\n' +
      'Check that the sheet headers match the column names in your Schema Manager.'
    );
    return;
  }

  // Upsert to Supabase (insert new, update existing by contact_id)
  const result = DB_upsertContacts(rows); // defined in DB_Connection.gs

  ui.alert(
    '✅ Import to Supabase Complete\n\n' +
    `Rows sent:               ${rows.length}\n` +
    `Inserted / Updated:      ${result.inserted}\n` +
    `Errors:                  ${result.errors}`
  );
}

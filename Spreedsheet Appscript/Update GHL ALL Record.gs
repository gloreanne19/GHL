/**
 * ====================================================
 * Update GHL ALL Record.gs
 * ====================================================
 * Pushes Prospect Type and other field updates from
 * the active Google Sheet directly into Supabase.
 *
 * Replaces the old two-spreadsheet approach:
 *   MAIN spreadsheet "GHL ALL Records"
 *   2026 spreadsheet "GHL ALL Records_2"
 *
 * All DB calls go through DB_Connection.gs functions.
 *
 * Source sheet layout (configure below):
 *   Column A = Contact ID
 *   Column B = Value to update  → maps to DB_UPDATE_COL_1
 *   Column C = Value to update  → maps to DB_UPDATE_COL_2
 * ====================================================
 */

// ─────────────────────────────────────────────
// CONFIG — Adjust these to match what you want
// to update in Supabase
// ─────────────────────────────────────────────
const GHL_SOURCE_SHEET = '';             // Leave blank to use the active sheet
const DB_UPDATE_COL_1  = 'prospect_type';  // Which contact DB column gets value from col B
const DB_UPDATE_COL_2  = 'marketing_stage'; // Which contact DB column gets value from col C

/**
 * Reads source sheet (A = contact_id, B = col1 value, C = col2 value)
 * and pushes updates to the Supabase contacts table.
 */
function pushUpdatesToSupabase() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const srcSh = GHL_SOURCE_SHEET
    ? ss.getSheetByName(GHL_SOURCE_SHEET)
    : ss.getActiveSheet();

  if (!srcSh) {
    ui.alert(`⚠️ Source sheet not found: "${GHL_SOURCE_SHEET}"`);
    return;
  }

  const lastRow = srcSh.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data found. Need at least 1 row under the header.');
    return;
  }

  // Read A:C (ContactID, col1 value, col2 value)
  const srcVals = srcSh.getRange(2, 1, lastRow - 1, 3).getValues();

  const updates = [];

  for (const row of srcVals) {
    const id = String(row[0] || '').trim();
    if (!id) continue;

    const update = { contact_id: id };

    // Only include non-empty fields
    if (row[1] !== '' && row[1] !== null && row[1] !== undefined) {
      update[DB_UPDATE_COL_1] = row[1];
    }
    if (row[2] !== '' && row[2] !== null && row[2] !== undefined) {
      update[DB_UPDATE_COL_2] = row[2];
    }

    if (Object.keys(update).length > 1) { // more than just contact_id
      updates.push(update);
    }
  }

  if (updates.length === 0) {
    ui.alert('No valid Contact IDs or updates found in the sheet.');
    return;
  }

  // Push to Supabase via DB_Connection.gs
  const result = DB_updateContactFields(updates);

  ui.alert(
    '✅ Supabase Update Complete\n\n' +
    `Rows processed: ${updates.length}\n` +
    `Successfully updated: ${result.updated}\n` +
    `Skipped (not found / nothing to update): ${result.notFound}\n` +
    `Errors: ${result.errors}`
  );
}

/**
 * Prepend rows from "Reorder Columns For Prepend" sheet
 * directly into Supabase (replaces the old "GHL ALL Records_2" prepend).
 * Rows are INSERTED (upserted), not prepended to a sheet.
 */
function prependNewImport_Supabase() {
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

  const headers = importSheet.getRange(1, 1, 1, importSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());

  const dataRows = importSheet.getRange(2, 1, lastRow - 1, importSheet.getLastColumn())
    .getDisplayValues(); // use DisplayValues to preserve date formatting

  // Load the column map from Supabase schema_config
  const colMap = DB_getColMap(); // { ExcelHeader -> db_column }

  const rows = dataRows.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      const dbCol = colMap[header];
      if (dbCol && row[i] !== '') obj[dbCol] = row[i];
    });
    return obj;
  }).filter(obj => Object.keys(obj).length > 0);

  if (rows.length === 0) {
    ui.alert('No mappable rows found. Check that the sheet headers match the Schema Manager.');
    return;
  }

  const result = DB_upsertContacts(rows);

  ui.alert(
    `✅ Import to Supabase Complete\n\n` +
    `Rows sent: ${rows.length}\n` +
    `Upserted (inserted/updated): ${result.inserted}\n` +
    `Errors: ${result.errors}`
  );
}

/**
 * Pull contacts from Supabase and write them to a "GHL ALL Records" sheet.
 * Useful for reviewing the master data in the spreadsheet without
 * maintaining a separate master file.
 */
function pullAllRecordsFromSupabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ui.alert('📥 Fetching all contacts from Supabase... This may take a moment.');

  const labelMap = DB_getLabelMap(); // { db_column -> ExcelHeader }
  const rows     = DB_fetchContacts();

  if (!rows || rows.length === 0) {
    ui.alert('No contacts found in Supabase.');
    return;
  }

  // Build ordered headers from the first row keys
  const dbCols = Object.keys(rows[0]);
  const headers = dbCols.map(col => labelMap[col] || col);

  // Write to sheet
  let sheet = ss.getSheetByName('GHL ALL Records');
  if (!sheet) sheet = ss.insertSheet('GHL ALL Records');
  else sheet.clearContents();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const dataOut = rows.map(row => dbCols.map(col => row[col] || ''));
  if (dataOut.length > 0) {
    sheet.getRange(2, 1, dataOut.length, headers.length).setValues(dataOut);
  }

  sheet.autoResizeColumns(1, Math.min(headers.length, 20));

  ui.alert(
    `✅ Done\n\nLoaded ${rows.length} contacts from Supabase\n` +
    `Written to: "GHL ALL Records" tab`
  );
}

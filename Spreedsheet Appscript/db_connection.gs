/**
 * ============================================================
 * DB_Connection.gs
 * ============================================================
 * Central Supabase connection layer for ALL Google Apps Scripts.
 * All other .gs files should call these functions instead of
 * talking to Supabase or master spreadsheets directly.
 *
 * USAGE IN OTHER FILES:
 *   const rows = DB_fetchContacts('full_name,phone,prospect_type');
 *   DB_upsertContacts([{ contact_id: '123', prospect_type: 'PO' }]);
 *   DB_updateContactFields([{ contact_id: '123', prospect_type: 'PO/HQ' }]);
 * ============================================================
 */

// ─────────────────────────────────────────────
// CONFIG — Edit here only
// ─────────────────────────────────────────────
const DB_URL   = 'https://fmspuwcfygbsklgnrray.supabase.co';
const DB_KEY   = 'sb_publishable_y2oAsYVOdrjyDcxAa5FX1w_6saruoJY';
const DB_TABLE = 'contacts';

// ─────────────────────────────────────────────
// SCHEMA — loads column mapping from schema_config
// ─────────────────────────────────────────────
let _schemaCache = null;

/**
 * Loads the active schema from Supabase schema_config table.
 * Returns array of { excel_header, db_column } objects.
 * Results are cached per script execution.
 */
function DB_loadSchema() {
  if (_schemaCache) return _schemaCache;
  const res = DB_request_('GET', '/schema_config?select=excel_header,db_column&is_active=eq.true');
  _schemaCache = res || [];
  return _schemaCache;
}

/**
 * Returns a { ExcelHeader: db_column } map.
 */
function DB_getColMap() {
  const schema = DB_loadSchema();
  const map = {};
  schema.forEach(s => { map[s.excel_header] = s.db_column; });
  return map;
}

/**
 * Returns a { db_column: ExcelHeader } reverse map.
 */
function DB_getLabelMap() {
  const schema = DB_loadSchema();
  const map = {};
  schema.forEach(s => { map[s.db_column] = s.excel_header; });
  return map;
}

// ─────────────────────────────────────────────
// READ — Fetch contacts from Supabase
// ─────────────────────────────────────────────

/**
 * Fetch all contacts from Supabase with pagination.
 * @param {string} select  Comma-separated DB column names. Defaults to all schema columns.
 * @param {Object} filters Optional key-value pairs for exact-match filtering. e.g. { state: 'TX' }
 * @returns {Array} Array of row objects.
 */
function DB_fetchContacts(select, filters) {
  if (!select) {
    const schema = DB_loadSchema();
    select = schema.map(s => s.db_column).join(',') || '*';
  }

  const BATCH   = 1000;
  let offset    = 0;
  const allRows = [];

  while (true) {
    let endpoint = `/contacts?select=${select}&limit=${BATCH}&offset=${offset}`;

    if (filters) {
      Object.keys(filters).forEach(col => {
        endpoint += `&${col}=eq.${encodeURIComponent(filters[col])}`;
      });
    }

    const rows = DB_request_('GET', endpoint);
    if (!rows || rows.length === 0) break;

    allRows.push(...rows);
    if (rows.length < BATCH) break;
    offset += BATCH;
  }

  Logger.log(`DB_fetchContacts: loaded ${allRows.length} rows`);
  return allRows;
}

/**
 * Fetch a single contact by contact_id.
 * @param {string} contactId
 * @returns {Object|null}
 */
function DB_fetchContactById(contactId) {
  const rows = DB_request_('GET', `/contacts?contact_id=eq.${encodeURIComponent(contactId)}&limit=1`);
  return (rows && rows.length) ? rows[0] : null;
}

// ─────────────────────────────────────────────
// WRITE — Push data to Supabase
// ─────────────────────────────────────────────

/**
 * Upsert (insert or update) rows into the contacts table.
 * Matches on contact_id; if no match → insert; if match → update.
 * @param {Array}  rows      Array of objects with DB column keys.
 * @param {number} batchSize Rows per request. Default 100.
 * @returns {{ inserted: number, errors: number }}
 */
function DB_upsertContacts(rows, batchSize) {
  batchSize = batchSize || 100;
  let inserted = 0, errors = 0;

  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    const res = DB_request_('POST', '/contacts', JSON.stringify(batch), {
      'Prefer': 'resolution=merge-duplicates,return=minimal'
    });
    if (res === null) errors += batch.length;
    else inserted += batch.length;
  }

  Logger.log(`DB_upsertContacts: upserted ${inserted}, errors ${errors}`);
  return { inserted, errors };
}

/**
 * Update specific fields on existing contacts matched by contact_id.
 * Only the columns provided in each row object are updated.
 * @param {Array} updates Array of { contact_id, field1, field2, ... }
 * @returns {{ updated: number, notFound: number, errors: number }}
 */
function DB_updateContactFields(updates) {
  let updated = 0, notFound = 0, errors = 0;

  for (const item of updates) {
    const id = String(item.contact_id || '').trim();
    if (!id) { notFound++; continue; }

    const cleanPayload = {};
    for (const key of Object.keys(item)) {
      if (key === 'contact_id') continue;
      const val = item[key];
      if (val !== '' && val !== null && val !== undefined) {
        cleanPayload[key] = val;
      }
    }

    if (Object.keys(cleanPayload).length === 0) { notFound++; continue; }

    const res = DB_request_(
      'PATCH',
      `/contacts?contact_id=eq.${encodeURIComponent(id)}`,
      JSON.stringify(cleanPayload),
      { 'Prefer': 'return=minimal' }
    );

    if (res === null) errors++;
    else updated++;
  }

  Logger.log(`DB_updateContactFields: updated ${updated}, not found ${notFound}, errors ${errors}`);
  return { updated, notFound, errors };
}

/**
 * Dynamically add a new column to the contacts table AND register it in schema_config.
 * @param {string} excelHeader  Display name (shown in Excel/web app)
 * @param {string} dbColumn     DB column name (snake_case)
 * @param {string} dataType     SQL type, default 'TEXT'
 */
function DB_addColumn(excelHeader, dbColumn, dataType) {
  dataType = dataType || 'TEXT';
  DB_request_('POST', '/rpc/add_column_to_contacts', JSON.stringify({ col_name: dbColumn, col_type: dataType }));
  DB_request_('POST', '/schema_config', JSON.stringify({ excel_header: excelHeader, db_column: dbColumn, data_type: dataType }), {
    'Prefer': 'resolution=ignore-duplicates'
  });
  _schemaCache = null; // reset cache
  Logger.log(`DB_addColumn: "${excelHeader}" -> "${dbColumn}" (${dataType})`);
}

// ─────────────────────────────────────────────
// INTERNAL HTTP HELPER
// ─────────────────────────────────────────────

/**
 * Internal: sends a request to Supabase REST API.
 * @param {string} method   GET | POST | PATCH | DELETE
 * @param {string} endpoint Path starting with /  e.g. '/contacts?limit=10'
 * @param {string} payload  JSON string body (optional)
 * @param {Object} extra    Extra headers (optional)
 * @returns {Array|Object|null} Parsed JSON, or null on error.
 */
function DB_request_(method, endpoint, payload, extra) {
  const url = DB_URL + '/rest/v1' + endpoint;
  const headers = Object.assign({
    'apikey':        DB_KEY,
    'Authorization': 'Bearer ' + DB_KEY,
    'Content-Type':  'application/json',
    'Accept':        'application/json'
  }, extra || {});

  const opts = { method: method, headers: headers, muteHttpExceptions: true };
  if (payload) opts.payload = payload;

  let response;
  try {
    response = UrlFetchApp.fetch(url, opts);
  } catch (e) {
    Logger.log(`DB network error [${method} ${endpoint}]: ${e.message}`);
    return null;
  }

  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code < 200 || code > 299) {
    Logger.log(`DB error (HTTP ${code}) [${method} ${endpoint}]: ${body.substring(0, 300)}`);
    return null;
  }

  if (!body || body.trim() === '' || code === 204) return {};

  try {
    return JSON.parse(body);
  } catch (e) {
    Logger.log(`DB parse error: ${e.message}`);
    return null;
  }
}

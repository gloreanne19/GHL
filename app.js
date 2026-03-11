// ─────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────
const SUPABASE_URL  = 'https://fmspuwcfygbsklgnrray.supabase.co';
const SUPABASE_KEY  = 'sb_publishable_y2oAsYVOdrjyDcxAa5FX1w_6saruoJY';
const TABLE         = 'contacts';
const BATCH         = 100;
const PAGE_SIZE     = 50; 

// Excel header → DB column name
// mapping will be loaded from DB
let COL      = {};
let LABELS   = {};
let ALL_COLS = [];
let SCHEMA   = []; // raw schema records {id, excel_header, db_column, ...}

function updateLookups() {
  const active = SCHEMA.filter(s => s.is_active);
  COL = Object.fromEntries(active.map(s => [s.excel_header, s.db_column]));
  LABELS = Object.fromEntries(active.map(s => [s.db_column, s.excel_header]));
  ALL_COLS = Object.values(COL);
}

const DEFAULT_COLS = [
  'prospect_type','full_name','phone','email',
  'city','state','postal_code','marketing_stage','source','contact_type'
];

// ─────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────
let db           = null;
let parsedRows   = [];
let failedImports = [];
let currentPage  = 1;
let totalRows    = 0;
let filters      = { prospect:'', state:'', stage:'', q:'', dStart:'', dEnd:'' };
let visibleCols  = [];
let searchTimer  = null;
let selectedIds  = new Set();
let currentRows  = [];
let currentUserRole = 'user'; // 'admin' or 'user'

// ─────────────────────────────────────────────
// AUTH
// ─────────────────────────────────────────────
async function init() {
  db = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

  // Check existing session
  const { data: { session } } = await db.auth.getSession();
  if (session) {
    showApp();
  } else {
    showLogin();
  }

  // Listen for auth changes
  db.auth.onAuthStateChange((_event, session) => {
    if (session) showApp();
    else showLogin();
  });
}

function showLogin() {
  document.getElementById('login-screen').style.display = 'flex';
  document.getElementById('main-app').style.display = 'none';
  // Clear form fields and reset button so autocomplete doesn't auto-submit
  document.getElementById('login-email').value    = '';
  document.getElementById('login-password').value = '';
  document.getElementById('login-error').style.display = 'none';
  const btn = document.getElementById('login-btn');
  btn.disabled = false;
  btn.textContent = 'Sign In';
}

async function showApp() {
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('main-app').style.display = 'flex';
  try { await connectDB(); } catch(e) { console.warn('connectDB error:', e); }
  try { await loadCurrentUser(); } catch(e) { console.warn('loadCurrentUser error:', e); }
  try { await loadSchema(); } catch(e) { console.warn('loadSchema error:', e); }
}

async function loadCurrentUser() {
  const { data: { user } } = await db.auth.getUser();
  if (!user) return;

  // Try to get role from RPC, fall back to raw_app_meta_data, then default to 'user'
  let role = 'user';
  try {
    const { data: rpcRole, error } = await db.rpc('get_my_role');
    if (!error && rpcRole) role = rpcRole;
  } catch (_) {
    // RPC not created yet — fall back silently
    role = 'user';
  }
  currentUserRole = role;

  // Display user in sidebar
  const email     = user.email || '';
  const initials  = email.charAt(0).toUpperCase() || '?';
  const roleLbl   = currentUserRole === 'admin' ? 'Admin' : 'User';
  const roleColor = currentUserRole === 'admin' ? 'var(--accent-dark)' : 'var(--muted)';

  const avatarEl = document.getElementById('sidebar-user-avatar');
  const emailEl  = document.getElementById('sidebar-user-email');
  const roleEl   = document.getElementById('sidebar-user-role');
  if (avatarEl) avatarEl.textContent = initials;
  if (emailEl)  emailEl.textContent  = email;
  if (roleEl)  { roleEl.textContent  = roleLbl; roleEl.style.color = roleColor; }

  // Access control: hide User Accounts nav for non-admins
  const usersNav = document.getElementById('nav-users');
  if (usersNav) usersNav.style.display = currentUserRole === 'admin' ? 'flex' : 'none';
}

async function loadSchema() {
  const { data, error } = await db.from('schema_config').select('*').order('excel_header');
  if (error) {
    if (error.code === '42P01') {
       // schema_config doesn't exist yet, show setup modal handled in connectDB
    } else {
       toast('Failed to load schema: ' + error.message, 'error');
    }
    return;
  }
  SCHEMA = data || [];
  updateLookups();
  // By default, visible columns are either the defaults or all active columns if few exist
  if (visibleCols.length === 0) {
    visibleCols = [...new Set([...DEFAULT_COLS, ...ALL_COLS])].filter(c => ALL_COLS.includes(c)).slice(0, 15);
    if (visibleCols.length === 0) visibleCols = ALL_COLS.slice(0, 15);
  }
  renderSchemaManager();
  if (document.getElementById('tab-records').classList.contains('active')) {
    loadRecords(true);
  }
}

// Login form
document.getElementById('login-form').addEventListener('submit', async e => {
  e.preventDefault();
  const email    = document.getElementById('login-email').value.trim();
  const password = document.getElementById('login-password').value;
  const btn      = document.getElementById('login-btn');
  const errEl    = document.getElementById('login-error');

  btn.disabled = true;
  btn.textContent = 'Signing in…';
  errEl.style.display = 'none';

  const { error } = await db.auth.signInWithPassword({ email, password });

  if (error) {
    errEl.textContent = error.message;
    errEl.style.display = 'block';
    btn.disabled = false;
    btn.textContent = 'Sign In';
  }
  // success → onAuthStateChange fires → showApp()
});

// Logout
document.getElementById('logout-btn').addEventListener('click', async () => {
  currentUserRole = 'user';
  visibleCols = [];
  await db.auth.signOut();
  // showLogin() will be called by onAuthStateChange
});


// ─────────────────────────────────────────────
// DB CONNECTION
// ─────────────────────────────────────────────
const CREATE_SQL = `
-- 1. Contacts Table
CREATE TABLE IF NOT EXISTS contacts (
  id BIGSERIAL PRIMARY KEY,
  created_date TEXT, contact_id TEXT UNIQUE, prospect_type TEXT,
  marketing_stage TEXT, contact_type TEXT, full_name TEXT,
  first_name TEXT, last_name TEXT, phone TEXT, email TEXT,
  subject_property_address TEXT, street_address TEXT, city TEXT,
  state TEXT, postal_code TEXT, property_county TEXT,
  mailing_street TEXT, mailing_city TEXT, mailing_state TEXT, mailing_zipcode TEXT,
  source TEXT, auction_date TEXT, property_type TEXT, house_style TEXT,
  year_built TEXT, square_footage TEXT, beds TEXT, baths TEXT, pool TEXT,
  last_sales_date TEXT, deed_date TEXT, record_status TEXT, bad_call_kill TEXT,
  owner_type TEXT, retail_score TEXT, rental_score TEXT, loan_to_value TEXT,
  loan_balance_15k TEXT, original_loan TEXT, second_poc_name TEXT,
  deceased_owner TEXT, pr_file_date TEXT, first_date_contact_added TEXT,
  equity_of_property TEXT, date_of_death TEXT, loan_date TEXT,
  foreclosing_lien TEXT, complaint_type TEXT, ros_offer TEXT,
  foreclosures TEXT, absentee_owner TEXT, high_equity TEXT,
  pre_foreclosure TEXT, deceased_probate TEXT, last_sales_price TEXT,
  recording_date TEXT, avm TEXT, estimated_value TEXT, market_value TEXT,
  assessed_total TEXT, rental_estimate_low TEXT, rental_estimate_high TEXT,
  total_loans TEXT, estimated_mortgage_balance TEXT,
  estimated_mortgage_payment TEXT, mortgage_interest_rate TEXT,
  ltv TEXT, maturity_date TEXT, free_and_clear TEXT, equity_percent TEXT,
  additional_phones TEXT, additional_emails TEXT,
  imported_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Schema Configuration Table
CREATE TABLE IF NOT EXISTS schema_config (
  id SERIAL PRIMARY KEY,
  excel_header TEXT UNIQUE,
  db_column TEXT,
  data_type TEXT DEFAULT 'TEXT',
  is_active BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Dynamic Column Function
CREATE OR REPLACE FUNCTION add_column_to_contacts(col_name TEXT, col_type TEXT)
RETURNS void AS $$
BEGIN
  EXECUTE format('ALTER TABLE contacts ADD COLUMN IF NOT EXISTS %I %s', col_name, col_type);
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- 4. Initial Config (Seed)
INSERT INTO schema_config (excel_header, db_column) VALUES
('Created','created_date'),('Contact Id','contact_id'),('Prospect Type','prospect_type'),
('Marketing Stage of Contact','marketing_stage'),('Type','contact_type'),('Name','full_name'),
('First Name','first_name'),('Last Name','last_name'),('Phone','phone'),('Email','email'),
('Subject Property Address','subject_property_address'),('Street address','street_address'),
('City','city'),('State','state'),('Postal Code','postal_code'),('Property County','property_county')
ON CONFLICT (excel_header) DO NOTHING;

ALTER TABLE contacts DISABLE ROW LEVEL SECURITY;
ALTER TABLE schema_config DISABLE ROW LEVEL SECURITY;
`;

async function connectDB() {
  try {
    const { error } = await db.from(TABLE).select('id', { count:'exact', head:true });
    if (error) {
      if (error.code === '42P01') {
        document.getElementById('setup-sql').textContent = CREATE_SQL;
        document.getElementById('setup-modal').style.display = 'flex';
        setStatus('err', 'Table missing');
      } else throw error;
      return;
    }
    setStatus('ok', 'Connected');
      toast('Connected', 'success');
  } catch(e) {
    setStatus('err', 'Error');
    toast('DB error: ' + e.message, 'error');
  }
}

function setStatus(cls, txt) {
  document.getElementById('status-dot').className = 'status-dot ' + cls;
  document.getElementById('status-text').textContent = txt;
}

// ─────────────────────────────────────────────
// NAVIGATION
// ─────────────────────────────────────────────
document.querySelectorAll('.nav-item').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    btn.classList.add('active');
    const tab = btn.dataset.tab;
    document.getElementById('tab-' + tab).classList.add('active');
    if (tab === 'records') loadRecords(true);
    if (tab === 'history') loadHistory();
    if (tab === 'users')   loadUsers();
  });
});


// ─────────────────────────────────────────────
// FILE UPLOAD
// ─────────────────────────────────────────────
const dz = document.getElementById('dropzone');
const fi = document.getElementById('file-input');

dz.addEventListener('click', () => fi.click());
dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', ()  => dz.classList.remove('over'));
dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('over'); handleFile(e.dataTransfer.files[0]); });
fi.addEventListener('change', e => handleFile(e.target.files[0]));
document.getElementById('remove-file').addEventListener('click', resetImport);
document.getElementById('import-another').addEventListener('click', resetImport);

let _lastFile = null; // keep raw File object for upload

function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) { toast('Please upload .xlsx, .xls or .csv', 'error'); return; }
  _lastFile = file; // store for later upload
  const reader = new FileReader();
  reader.onload = e => parseFile(e.target.result, file);
  reader.readAsArrayBuffer(file);
}

function parseFile(buf, file) {
  try {
    const wb    = XLSX.read(buf, { type:'array', cellDates:true });
    const ws    = wb.Sheets[wb.SheetNames[0]];
    const data  = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
    if (data.length < 2) { toast('File appears empty', 'error'); return; }

    const headers = data[0].map(h => String(h).trim());
    const rawRows = data.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));

    // Detect missing columns
    const missing = headers.filter(h => !COL[h] && h !== '');
    if (missing.length > 0) {
       promptForNewColumns(missing, headers, rawRows, file);
       return;
    }

    finalizeParsedData(headers, rawRows, file);
  } catch(e) {
    toast('Failed to parse file: ' + e.message, 'error');
  }
}

let pendingNewCols = [];
let pendingHeaders = [];
let pendingRawRows = [];
let pendingFile    = null;

function promptForNewColumns(missing, headers, rawRows, file) {
  pendingNewCols = missing;
  pendingHeaders = headers;
  pendingRawRows = rawRows;
  pendingFile    = file;
  showNextNewCol();
}

function showNextNewCol() {
  if (pendingNewCols.length === 0) {
    finalizeParsedData(pendingHeaders, pendingRawRows, pendingFile);
    return;
  }
  const next = pendingNewCols[0];
  document.getElementById('new-col-name').textContent = next;
  document.getElementById('new-col-db-name').value = next.toLowerCase().replace(/[^a-z0-9]/g, '_');
  document.getElementById('new-col-modal').style.display = 'flex';
}

document.getElementById('new-col-add').addEventListener('click', async () => {
    const header = pendingNewCols[0];
    const dbName = document.getElementById('new-col-db-name').value.trim();
    if (!dbName) return;

    // Add to DB
    const { error: rpcErr } = await db.rpc('add_column_to_contacts', { col_name: dbName, col_type: 'TEXT' });
    if (rpcErr) { toast('RPC Error: ' + rpcErr.message, 'error'); return; }

    const { error: insErr } = await db.from('schema_config').insert({ excel_header: header, db_column: dbName });
    if (insErr) { toast('Config Error: ' + insErr.message, 'error'); return; }

    toast(`Added column ${dbName}`, 'success');
    await loadSchema(); // refresh COL
    
    document.getElementById('new-col-modal').style.display = 'none';
    pendingNewCols.shift();
    showNextNewCol();
});

document.getElementById('new-col-skip').addEventListener('click', () => {
    document.getElementById('new-col-modal').style.display = 'none';
    pendingNewCols.shift();
    showNextNewCol();
});

function finalizeParsedData(headers, rawRows, file) {
    parsedRows = rawRows.map(row => {
      const obj = {};
      const raw = {}; // keep raw for error review
      headers.forEach((h, i) => {
        raw[h] = row[i];
        const dbCol = COL[h];
        if (dbCol) {
          let val = row[i];
          if (val instanceof Date) val = val.toISOString().split('T')[0];
          obj[dbCol] = val !== null && val !== undefined ? String(val).trim() : '';
        }
      });
      return { mapped: obj, raw: raw };
    }).filter(o => Object.keys(o.mapped).length > 0);

    document.getElementById('file-bar').style.display = 'flex';
    document.getElementById('file-name').textContent = file.name;
    document.getElementById('file-meta').textContent = `${rawRows.length.toLocaleString()} rows · ${headers.length} columns · ${fmtSize(file.size)}`;
    document.getElementById('import-options').style.display = 'block';
    document.getElementById('import-actions').style.display = 'block';
    document.getElementById('preview-badge').textContent = rawRows.length.toLocaleString() + ' rows';
    showPreview(headers, parsedRows.map(p => p.mapped));
}

function showPreview(headers, mapped) {
  const visH = headers.filter(h => COL[h]).slice(0, 10);
  const rows  = mapped.slice(0, 5);
  document.getElementById('preview-head').innerHTML = '<tr>' + visH.map(h => `<th>${h}</th>`).join('') + '</tr>';
  document.getElementById('preview-body').innerHTML = rows.map(row =>
    '<tr>' + visH.map(h => { const v = row[COL[h]] || ''; return `<td title="${escHtml(v)}">${escHtml(v.substring(0,28))}${v.length>28?'…':''}</td>`; }).join('') + '</tr>'
  ).join('');
  document.getElementById('preview-card').style.display = 'block';
}

function resetImport() {
  parsedRows = []; fi.value = '';
  ['file-bar','import-options','preview-card','import-actions','progress-card','result-card']
    .forEach(id => { document.getElementById(id).style.display = 'none'; });
}

// ─────────────────────────────────────────────
// IMPORT
// ─────────────────────────────────────────────
document.getElementById('import-btn').addEventListener('click', doImport);

async function doImport() {
  if (!parsedRows.length) return;
  if (!db) { toast('Not connected', 'error'); return; }

  const mode = document.querySelector('input[name="dup"]:checked').value;
  const btn  = document.getElementById('import-btn');
  btn.disabled = true;
  document.getElementById('progress-card').style.display = 'block';
  document.getElementById('result-card').style.display   = 'none';

  let inserted = 0, errors = 0;
  failedImports = []; 
  const total  = parsedRows.length;

  for (let i = 0; i < total; i += BATCH) {
    const batchData = parsedRows.slice(i, i + BATCH);
    const mappedBatch = batchData.map(d => d.mapped);
    
    const { error } = mode === 'update'
      ? await db.from(TABLE).upsert(mappedBatch, { onConflict:'contact_id', ignoreDuplicates:false })
      : await db.from(TABLE).upsert(mappedBatch, { onConflict:'contact_id', ignoreDuplicates:true  });
    
    if (error) {
      // If batch fails, try individual rows to capture errors
      for (const item of batchData) {
        const { error: singleErr } = await db.from(TABLE).upsert([item.mapped], { onConflict:'contact_id' });
        if (singleErr) {
          errors++;
          failedImports.push({ ...item, error: singleErr.message });
        } else {
          inserted++;
        }
      }
    } else {
      inserted += batchData.length;
    }

    const done = Math.min(i + BATCH, total);
    const pct  = Math.round((done / total) * 100);
    document.getElementById('prog-fill').style.width  = pct + '%';
    document.getElementById('prog-pct').textContent   = pct + '%';
    document.getElementById('prog-label').textContent = `Importing batch ${Math.ceil((i+1)/BATCH)}…`;
    document.getElementById('prog-stats').textContent = `${done.toLocaleString()} of ${total.toLocaleString()} rows processed`;
  }

  btn.disabled = false;
  const ok = errors === 0;
  document.getElementById('result-icon').textContent  = ok ? '✅' : '⚠️';
  document.getElementById('result-title').textContent = ok ? 'Import Successful' : 'Import Completed with Errors';
  document.getElementById('result-msg').textContent   = `Processed ${total.toLocaleString()} rows from Excel.`;
  document.getElementById('result-stats').innerHTML   = `
    <div class="rs"><strong style="color:var(--success)">${inserted.toLocaleString()}</strong><small>Imported</small></div>
    <div class="rs"><strong style="color:var(--error)">${errors.toLocaleString()}</strong><small>Errors</small></div>`;
  
  if (errors > 0) {
    document.getElementById('view-errors-btn').style.display = 'block';
    document.getElementById('nav-errors').style.display = 'flex';
    document.getElementById('err-count').textContent = errors;
    renderErrorTable();
  }

  document.getElementById('result-card').style.display = 'block';
  toast(ok ? `Import complete: ${inserted.toLocaleString()} rows` : `Import done with ${errors} errors`, ok ? 'success' : 'error');

  // Upload the original file to Supabase Storage, then log history
  const fileName = document.getElementById('file-name')?.textContent || 'unknown';
  uploadImportFile(_lastFile, fileName).then(fileUrl => {
    logImport(fileName, total, inserted, errors, parsedRows.slice(0, 200).map(p => p.raw), fileUrl);
  });
}

document.getElementById('view-errors-btn').addEventListener('click', () => {
    document.getElementById('nav-errors').click();
});

// ─────────────────────────────────────────────
// RECORDS
// ─────────────────────────────────────────────
async function loadRecords(reset) {
  if (!db) return;
  if (reset) { currentPage = 1; selectedIds.clear(); updateDeleteBtn(); await populateFilters(); }

  const from = (currentPage - 1) * PAGE_SIZE;
  const to   = from + PAGE_SIZE - 1;

  let q = db.from(TABLE).select('*', { count:'exact' });
  if (filters.prospect) q = q.eq('prospect_type', filters.prospect);
  if (filters.state)    q = q.eq('state', filters.state);
  if (filters.stage)    q = q.eq('marketing_stage', filters.stage);
  if (filters.q) {
    const s = filters.q;
    q = q.or(`full_name.ilike.%${s}%,phone.ilike.%${s}%,email.ilike.%${s}%,street_address.ilike.%${s}%,city.ilike.%${s}%,contact_id.ilike.%${s}%`);
  }
  if (filters.dStart) {
    q = q.gte('imported_at', filters.dStart + 'T00:00:00.000Z');
  }
  if (filters.dEnd) {
    q = q.lte('imported_at', filters.dEnd + 'T23:59:59.999Z');
  }
  q = q.order('id', { ascending:false }).range(from, to);

  const { data, count, error } = await q;
  if (error) { toast('Failed to load: ' + error.message, 'error'); return; }

  totalRows    = count || 0;
  currentRows  = data || [];

  document.getElementById('total-count').textContent = totalRows.toLocaleString() + ' records';
  document.getElementById('showing-text').textContent =
    `Showing ${from+1}–${Math.min(to+1, totalRows)} of ${totalRows.toLocaleString()}`;

  renderTable(currentRows);
  renderPagination();
  if (reset) renderStats();
}

async function populateFilters() {
  const { data: pt } = await db.from(TABLE).select('prospect_type').not('prospect_type','is',null);
  const prospects = [...new Set((pt||[]).map(r=>r.prospect_type).filter(Boolean))].sort();
  document.getElementById('filter-prospect').innerHTML = '<option value="">All Prospect Types</option>' +
    prospects.map(p => `<option value="${escHtml(p)}">${escHtml(p)}</option>`).join('');

  const { data: st } = await db.from(TABLE).select('state').not('state','is',null);
  const states = [...new Set((st||[]).map(r=>r.state).filter(Boolean))].sort();
  document.getElementById('filter-state').innerHTML = '<option value="">All States</option>' +
    states.map(s => `<option value="${escHtml(s)}">${escHtml(s)}</option>`).join('');

  const { data: mg } = await db.from(TABLE).select('marketing_stage').not('marketing_stage','is',null);
  const stages = [...new Set((mg||[]).map(r=>r.marketing_stage).filter(Boolean))].sort();
  document.getElementById('filter-stage').innerHTML = '<option value="">All Stages</option>' +
    stages.map(s => `<option value="${escHtml(s)}">${escHtml(s)}</option>`).join('');

  buildColsPanel();
}

function renderTable(rows) {
  // Header with select-all checkbox
  document.getElementById('rec-head').innerHTML =
    '<tr><th class="cb-col"><input type="checkbox" id="select-all-cb" title="Select all"></th>' +
    visibleCols.map(c => `<th>${LABELS[c] || c}</th>`).join('') +
    '<th style="width:40px"></th></tr>';

  document.getElementById('select-all-cb').addEventListener('change', function() {
    const checked = this.checked;
    rows.forEach(row => { if (checked) selectedIds.add(row.id); else selectedIds.delete(row.id); });
    document.querySelectorAll('.row-cb').forEach(cb => { cb.checked = checked; });
    document.querySelectorAll('#rec-body tr').forEach(tr => tr.classList.toggle('selected', checked));
    updateDeleteBtn();
  });

  if (!rows.length) {
    document.getElementById('rec-body').innerHTML =
      `<tr><td colspan="${visibleCols.length + 2}" class="empty-cell">No records found</td></tr>`;
    return;
  }

  document.getElementById('rec-body').innerHTML = rows.map(row => {
    const isSelected = selectedIds.has(row.id);
    return `<tr class="${isSelected ? 'selected' : ''}" data-id="${row.id}">
      <td class="cb-col">
        <input type="checkbox" class="row-cb" data-id="${row.id}" ${isSelected ? 'checked' : ''}>
      </td>
      ${visibleCols.map(c => {
        let val = row[c] || '';
        if (c === 'prospect_type' && val) {
          return `<td><span class="pb" style="background:rgba(124,106,247,.18);color:#a79fff">${escHtml(val)}</span></td>`;
        }
        return `<td title="${escHtml(String(val))}">${escHtml(String(val).substring(0,30))}${String(val).length>30?'…':''}</td>`;
      }).join('')}
      <td><button class="row-del-btn" data-id="${row.id}" title="Delete this record">Delete</button></td>
    </tr>`;
  }).join('');

  // Row checkboxes
  document.querySelectorAll('.row-cb').forEach(cb => {
    cb.addEventListener('change', function() {
      const id = +this.dataset.id;
      if (this.checked) selectedIds.add(id); else selectedIds.delete(id);
      const tr = this.closest('tr');
      tr.classList.toggle('selected', this.checked);
      updateDeleteBtn();
    });
  });

  // Row delete buttons
  document.querySelectorAll('.row-del-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const id = +btn.dataset.id;
      const row = rows.find(r => r.id === id);
      const name = row ? (row.full_name || row.contact_id || `ID ${id}`) : `ID ${id}`;
      showDeleteModal([id], `Delete record: ${escHtml(name)}?`);
    });
  });
}

function renderPagination() {
  const pages = Math.ceil(totalRows / PAGE_SIZE);
  if (pages <= 1) { document.getElementById('pagination').innerHTML = ''; return; }

  let html = `<button class="pg-btn" ${currentPage===1?'disabled':''} id="pg-prev">‹ Prev</button>`;
  const start = Math.max(1, currentPage - 2);
  const end   = Math.min(pages, currentPage + 2);
  if (start > 1) html += `<button class="pg-btn" data-p="1">1</button>${start>2?'<span style="color:var(--muted)">…</span>':''}`;
  for (let p = start; p <= end; p++) html += `<button class="pg-btn ${p===currentPage?'on':''}" data-p="${p}">${p}</button>`;
  if (end < pages) html += `${end<pages-1?'<span style="color:var(--muted)">…</span>':''}<button class="pg-btn" data-p="${pages}">${pages}</button>`;
  html += `<button class="pg-btn" ${currentPage===pages?'disabled':''} id="pg-next">Next ›</button>`;

  const pag = document.getElementById('pagination');
  pag.innerHTML = html;
  pag.querySelectorAll('[data-p]').forEach(btn => btn.addEventListener('click', () => { currentPage = +btn.dataset.p; loadRecords(); }));
  const prev = document.getElementById('pg-prev');
  const next = document.getElementById('pg-next');
  if (prev) prev.addEventListener('click', () => { currentPage--; loadRecords(); });
  if (next) next.addEventListener('click', () => { currentPage++; loadRecords(); });
}

async function renderStats() {
  const { data } = await db.from(TABLE).select('prospect_type').not('prospect_type','is',null);
  const counts = {};
  (data||[]).forEach(r => { if (r.prospect_type) counts[r.prospect_type] = (counts[r.prospect_type]||0)+1; });
  const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10);

  document.getElementById('stats-row').innerHTML = sorted.map(([pt, cnt]) =>
    `<div class="stat-pill" onclick="filterByProspect('${escHtml(pt)}')" title="Filter: ${escHtml(pt)}" style="cursor:pointer;flex-direction:column;align-items:flex-start;gap:2px;padding:10px 16px;min-width:110px">
       <strong style="font-size:18px;font-weight:800;color:var(--text)">${cnt.toLocaleString()}</strong>
       <span style="font-size:11px;color:var(--muted);font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:140px">${escHtml(pt)}</span>
     </div>`
  ).join('');
}

window.filterByProspect = function(pt) {
  filters.prospect = pt;
  document.getElementById('filter-prospect').value = pt;
  currentPage = 1; loadRecords();
};

// ─────────────────────────────────────────────
// EXPORT TO EXCEL
// ─────────────────────────────────────────────
async function exportToExcel() {
  const btn   = document.getElementById('export-btn');
  const BATCH = 1000;

  // Inject progress bar under button once
  let progBar = document.getElementById('export-prog-bar');
  if (!progBar) {
    progBar = document.createElement('div');
    progBar.id = 'export-prog-bar';
    progBar.style.cssText = 'height:3px;border-radius:3px;background:var(--border);margin-top:6px;overflow:hidden;display:none;width:100%';
    progBar.innerHTML = '<div id="export-prog-fill" style="height:100%;width:0%;background:var(--accent);border-radius:3px;transition:width .25s ease"></div>';
    btn.insertAdjacentElement('afterend', progBar);
  }
  const setP = (pct, label) => {
    btn.textContent = `${label}… ${pct}%`;
    document.getElementById('export-prog-fill').style.width = pct + '%';
    progBar.style.display = 'block';
  };

  btn.disabled = true;
  setP(0, 'Starting');

  try {
    const activeSchema = SCHEMA.filter(s => s.is_active);

    // Step 1: count
    setP(2, 'Counting');
    let cq = db.from(TABLE).select('id', { count: 'exact', head: true });
    if (filters.prospect) cq = cq.eq('prospect_type', filters.prospect);
    if (filters.state)    cq = cq.eq('state', filters.state);
    if (filters.stage)    cq = cq.eq('marketing_stage', filters.stage);
    if (filters.q) { const s = filters.q; cq = cq.or(`full_name.ilike.%${s}%,phone.ilike.%${s}%,email.ilike.%${s}%,street_address.ilike.%${s}%,city.ilike.%${s}%,contact_id.ilike.%${s}%`); }
    if (filters.dStart) cq = cq.gte('imported_at', filters.dStart + 'T00:00:00.000Z');
    if (filters.dEnd)   cq = cq.lte('imported_at', filters.dEnd   + 'T23:59:59.999Z');
    const { count, error: cErr } = await cq;
    if (cErr) throw cErr;
    if (!count) { toast('No records to export.', 'error'); return; }

    // Step 2: fetch in batches
    const totalBatches = Math.ceil(count / BATCH);
    let allData = [];
    for (let i = 0; i < totalBatches; i++) {
      const from = i * BATCH; const to = from + BATCH - 1;
      const fetched = Math.min((i + 1) * BATCH, count);
      setP(Math.round(5 + (i / totalBatches) * 75), `Fetching ${fetched.toLocaleString()} / ${count.toLocaleString()}`);
      let bq = db.from(TABLE).select('*').order('id', { ascending: false }).range(from, to);
      if (filters.prospect) bq = bq.eq('prospect_type', filters.prospect);
      if (filters.state)    bq = bq.eq('state', filters.state);
      if (filters.stage)    bq = bq.eq('marketing_stage', filters.stage);
      if (filters.q) { const s = filters.q; bq = bq.or(`full_name.ilike.%${s}%,phone.ilike.%${s}%,email.ilike.%${s}%,street_address.ilike.%${s}%,city.ilike.%${s}%,contact_id.ilike.%${s}%`); }
      if (filters.dStart) bq = bq.gte('imported_at', filters.dStart + 'T00:00:00.000Z');
      if (filters.dEnd)   bq = bq.lte('imported_at', filters.dEnd   + 'T23:59:59.999Z');
      const { data: batch, error: bErr } = await bq;
      if (bErr) throw bErr;
      allData = allData.concat(batch || []);
    }

    // Step 3: map to Excel headers
    setP(82, 'Building');
    await new Promise(r => setTimeout(r, 0));
    const rows = allData.map(row => {
      const out = {};
      activeSchema.forEach(s => { out[s.excel_header] = row[s.db_column] ?? ''; });
      return out;
    });

    // Step 4: create xlsx
    setP(92, 'Writing');
    await new Promise(r => setTimeout(r, 0));
    const ws = XLSX.utils.json_to_sheet(rows, { header: activeSchema.map(s => s.excel_header) });
    ws['!cols'] = activeSchema.map(s => ({ wch: Math.max(s.excel_header.length, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'GHL All Records');

    setP(98, 'Saving');
    await new Promise(r => setTimeout(r, 80));

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `GHL_Records_${stamp}.xlsx`);

    setP(100, 'Done!');
    toast(`Exported ${allData.length.toLocaleString()} records`, 'success');

  } catch (e) {
    toast('Export error: ' + e.message, 'error');
  } finally {
    setTimeout(() => {
      btn.disabled = false;
      btn.textContent = '\u2b07 Export Excel';
      progBar.style.display = 'none';
      document.getElementById('export-prog-fill').style.width = '0%';
    }, 2000);
  }
}

document.getElementById('export-btn').addEventListener('click', exportToExcel);



// ─────────────────────────────────────────────
// DELETE
// ─────────────────────────────────────────────
function updateDeleteBtn() {
  const btn = document.getElementById('delete-selected-btn');
  const cnt = selectedIds.size;
  document.getElementById('selected-count').textContent = cnt;
  btn.style.display = cnt > 0 ? 'inline-flex' : 'none';
}

document.getElementById('delete-selected-btn').addEventListener('click', () => {
  const ids = [...selectedIds];
  showDeleteModal(ids, `Delete ${ids.length} selected record${ids.length>1?'s':''}?`);
});

let pendingDeleteIds = [];

function showDeleteModal(ids, title) {
  pendingDeleteIds = ids;
  document.getElementById('delete-modal-title').textContent = title;
  document.getElementById('delete-modal-msg').textContent = `This will permanently remove ${ids.length} record${ids.length>1?'s':''} from the database. This cannot be undone.`;
  document.getElementById('delete-modal').style.display = 'flex';
  document.getElementById('mass-delete-warning').style.display = 'none';
  document.getElementById('delete-confirm').style.display = 'inline-block';
  document.getElementById('mass-delete-confirm').style.display = 'none';
}

document.getElementById('delete-cancel').addEventListener('click', () => {
  document.getElementById('delete-modal').style.display = 'none';
  pendingDeleteIds = [];
});

document.getElementById('delete-confirm').addEventListener('click', async () => {
  document.getElementById('delete-modal').style.display = 'none';
  if (!pendingDeleteIds.length) return;

  const ids = [...pendingDeleteIds];
  pendingDeleteIds = [];

  const { error } = await db.from(TABLE).delete().in('id', ids);
  if (error) {
    toast('Delete failed: ' + error.message, 'error');
  } else {
    ids.forEach(id => selectedIds.delete(id));
    updateDeleteBtn();
      toast(`Deleted ${ids.length} record${ids.length>1?'s':''}`, 'success');
    loadRecords(false);
    renderStats();
  }
});

// Mass Delete Event
document.getElementById('mass-delete-btn').addEventListener('click', () => {
  if (totalRows === 0) return toast('No records to delete', 'info');
  
  document.getElementById('delete-modal-title').textContent = 'Confirm Mass Delete';
  document.getElementById('delete-modal-msg').textContent = `This will permanently remove ${totalRows.toLocaleString()} record${totalRows>1?'s':''} from the database. This cannot be undone.`;
  document.getElementById('delete-modal').style.display = 'flex';
  document.getElementById('mass-delete-warning').style.display = 'block';
  document.getElementById('delete-confirm').style.display = 'none';
  document.getElementById('mass-delete-confirm').style.display = 'inline-block';
});

document.getElementById('mass-delete-confirm').addEventListener('click', async () => {
  document.getElementById('delete-modal').style.display = 'none';
  const btn = document.getElementById('mass-delete-btn');
  btn.textContent = 'Deleting...';
  btn.disabled = true;

  try {
    let q = db.from(TABLE).delete();
    let hasFilter = false;

    if (filters.prospect) { q = q.eq('prospect_type', filters.prospect); hasFilter = true; }
    if (filters.state)    { q = q.eq('state', filters.state); hasFilter = true; }
    if (filters.stage)    { q = q.eq('marketing_stage', filters.stage); hasFilter = true; }
    if (filters.dStart)   { q = q.gte('imported_at', filters.dStart + 'T00:00:00.000Z'); hasFilter = true; }
    if (filters.dEnd)     { q = q.lte('imported_at', filters.dEnd + 'T23:59:59.999Z'); hasFilter = true; }
    if (filters.q) {
      const s = filters.q;
      q = q.or(`full_name.ilike.%${s}%,phone.ilike.%${s}%,email.ilike.%${s}%,street_address.ilike.%${s}%,city.ilike.%${s}%,contact_id.ilike.%${s}%`);
      hasFilter = true;
    }

    // Supabase requires a filter when doing multi-row deletes
    if (!hasFilter) q = q.neq('id', 0);

    const { error } = await q;
    
    if (error) {
      toast('Mass delete failed: ' + error.message, 'error');
    } else {
      toast(`Deleted ${totalRows.toLocaleString()} records`, 'success');
      selectedIds.clear();
      updateDeleteBtn();
      loadRecords(true);
    }
  } catch (err) {
    toast('Error: ' + err.message, 'error');
  } finally {
    btn.textContent = 'Mass Delete';
    btn.disabled = false;
  }
});

// ─────────────────────────────────────────────
// FILTER EVENTS
// ─────────────────────────────────────────────
document.getElementById('filter-prospect').addEventListener('change', e => { filters.prospect = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('filter-state').addEventListener('change',    e => { filters.state    = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('filter-stage').addEventListener('change',    e => { filters.stage    = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('filter-date-start').addEventListener('change', e => { filters.dStart = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('filter-date-end').addEventListener('change',   e => { filters.dEnd   = e.target.value; currentPage=1; loadRecords(); });

document.getElementById('search-input').addEventListener('input', e => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { filters.q = e.target.value.trim(); currentPage=1; loadRecords(); }, 400);
});

document.getElementById('reset-btn').addEventListener('click', () => {
  filters = { prospect:'', state:'', stage:'', q:'', dStart:'', dEnd:'' };
  ['filter-prospect','filter-state','filter-stage','filter-date-start','filter-date-end'].forEach(id => { document.getElementById(id).value = ''; });
  document.getElementById('search-input').value = '';
  currentPage = 1; loadRecords();
});

// ─────────────────────────────────────────────
// COLUMN TOGGLE
// ─────────────────────────────────────────────
function buildColsPanel() {
  const grid = document.getElementById('cols-grid');
  grid.innerHTML = ALL_COLS.map(c => `
    <label class="col-check">
      <input type="checkbox" value="${c}" ${visibleCols.includes(c)?'checked':''}>
      ${escHtml(LABELS[c] || c)}
    </label>`).join('');
  grid.querySelectorAll('input').forEach(cb =>
    cb.addEventListener('change', () => {
      visibleCols = [...grid.querySelectorAll('input:checked')].map(i => i.value);
      if (!visibleCols.length) visibleCols = DEFAULT_COLS;
      loadRecords();
    })
  );
}

document.getElementById('cols-btn').addEventListener('click', e => {
  e.stopPropagation();
  const p = document.getElementById('cols-panel');
  p.style.display = p.style.display === 'none' ? 'block' : 'none';
});
document.addEventListener('click', () => {
  document.getElementById('cols-panel').style.display = 'none';
});

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────
function toast(msg, type = 'info') {
  const el = document.getElementById('toast');
  el.textContent = msg; el.className = `toast ${type} show`;
  setTimeout(() => el.classList.remove('show'), 3500);
}
function fmtSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b/1024).toFixed(1) + ' KB';
  return (b/1048576).toFixed(1) + ' MB';
}
function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─────────────────────────────────────────────
// ERROR REVIEW
// ─────────────────────────────────────────────
function renderErrorTable() {
  const container = document.getElementById('err-body');
  if (!failedImports.length) {
    container.innerHTML = '<tr><td colspan="5" class="empty-cell">No errors to review</td></tr>';
    return;
  }

  // Determine headers to show (first 5 DB columns + Error)
  const cols = SCHEMA.filter(s => s.is_active).slice(0, 5);
  
  document.getElementById('err-head').innerHTML = 
    '<tr>' + cols.map(c => `<th>${c.excel_header}</th>`).join('') + '<th>Error Reason</th></tr>';

  container.innerHTML = failedImports.map((item, idx) => `
    <tr data-idx="${idx}">
      ${cols.map(c => `
        <td contenteditable="true" data-col="${c.db_column}" class="editable-cell">
          ${escHtml(item.mapped[c.db_column] || '')}
        </td>`).join('')}
      <td class="err-reason">${escHtml(item.error)}</td>
    </tr>
  `).join('');

  // Update data on edit
  container.querySelectorAll('.editable-cell').forEach(td => {
    td.addEventListener('blur', () => {
      const idx = td.closest('tr').dataset.idx;
      const col = td.dataset.col;
      failedImports[idx].mapped[col] = td.textContent.trim();
    });
  });
}

document.getElementById('retry-import-btn').addEventListener('click', async () => {
  const btn = document.getElementById('retry-import-btn');
  btn.disabled = true;
  btn.textContent = 'Retrying...';

  const toRetry = [...failedImports];
  failedImports = [];
  let success = 0;

  for (const item of toRetry) {
    const { error } = await db.from(TABLE).upsert([item.mapped], { onConflict: 'contact_id' });
    if (error) {
      failedImports.push({ ...item, error: error.message });
    } else {
      success++;
    }
  }

  btn.disabled = false;
  btn.textContent = 'Retry All Corrections';
  toast(`Successfully imported ${success} records. ${failedImports.length} still failing.`, success > 0 ? 'success' : 'error');
  
  document.getElementById('err-count').textContent = failedImports.length;
  if (failedImports.length === 0) {
    document.getElementById('nav-errors').style.display = 'none';
    document.getElementById('nav-records').click();
  } else {
    renderErrorTable();
  }
});

// ─────────────────────────────────────────────
// SCHEMA MANAGER
// ─────────────────────────────────────────────
function renderSchemaManager() {
  const container = document.getElementById('schema-body');
  if (!SCHEMA.length) {
    container.innerHTML = '<tr><td colspan="5" class="empty-cell">No schema mappings yet. Click "+ Add Row" to get started.</td></tr>';
    return;
  }
  container.innerHTML = SCHEMA.map(s => `
    <tr data-id="${s.id}">
      <td>
        <input type="text" class="field-input sm" value="${escHtml(s.excel_header)}" data-key="excel_header" style="min-width:140px">
      </td>
      <td>
        <input type="text" class="field-input sm" value="${escHtml(s.db_column)}" data-key="db_column" style="min-width:120px">
      </td>
      <td>
        <select class="flt-select sm" data-key="data_type" style="min-width:100px">
          <option value="TEXT"        ${s.data_type==='TEXT'       ?'selected':''}>TEXT</option>
          <option value="INTEGER"     ${s.data_type==='INTEGER'    ?'selected':''}>INTEGER</option>
          <option value="BOOLEAN"     ${s.data_type==='BOOLEAN'    ?'selected':''}>BOOLEAN</option>
          <option value="TIMESTAMPTZ" ${s.data_type==='TIMESTAMPTZ'?'selected':''}>TIMESTAMPTZ</option>
        </select>
      </td>
      <td style="text-align:center">
        <label class="toggle">
          <input type="checkbox" data-key="is_active" ${s.is_active ? 'checked' : ''}>
          <span class="toggle-track"></span>
        </label>
      </td>
      <td style="text-align:center">
        <button class="btn-danger" style="padding:4px 10px;font-size:12px" onclick="deleteMapping(${s.id})">Delete</button>
      </td>
    </tr>
  `).join('');
}

document.getElementById('save-schema-btn').addEventListener('click', async () => {
    const rows = document.querySelectorAll('#schema-body tr');
    const updates = [];
    rows.forEach(tr => {
        const id = +tr.dataset.id;
        if (!id) return; // skip new ones handled differently or not yet saved
        updates.push({
            id,
            excel_header: tr.querySelector('[data-key="excel_header"]').value.trim(),
            db_column   : tr.querySelector('[data-key="db_column"]').value.trim(),
            data_type   : tr.querySelector('[data-key="data_type"]').value,
            is_active   : tr.querySelector('[data-key="is_active"]').checked
        });
    });

    const { error } = await db.from('schema_config').upsert(updates);
    if (error) toast('Save failed: ' + error.message, 'error');
    else {
        toast('Schema saved', 'success');
        await loadSchema();
    }
});

document.getElementById('add-mapping-btn').addEventListener('click', async () => {
    const { data, error } = await db.from('schema_config').insert({ excel_header: 'New Column', db_column: 'new_col' }).select();
    if (error) toast('Add failed: ' + error.message, 'error');
    else {
        SCHEMA.push(data[0]);
        renderSchemaManager();
    }
});

window.deleteMapping = async (id) => {
    if (!confirm('Remove this mapping? (Database column will remain)')) return;
    const { error } = await db.from('schema_config').delete().eq('id', id);
    if (error) toast('Delete failed: ' + error.message, 'error');
    else {
        SCHEMA = SCHEMA.filter(s => s.id !== id);
        renderSchemaManager();
    }
};

// ─────────────────────────────────────────────
// START
// ─────────────────────────────────────────────

// ══════════════════════════════════════════════════════════
//  IMPORT HISTORY
// ══════════════════════════════════════════════════════════
// Upload the raw file to Supabase Storage
async function uploadImportFile(file, fileName) {
  if (!file) return null;
  try {
    const { data: { user } } = await db.auth.getUser();
    const ts       = Date.now();
    const safeName = fileName.replace(/[^a-z0-9._-]/gi, '_');
    const path     = `${user?.id || 'anon'}/${ts}_${safeName}`;

    const { error } = await db.storage
      .from('import-files')
      .upload(path, file, { upsert: false, contentType: file.type || 'application/octet-stream' });

    if (error) {
      console.warn('File upload failed:', error.message);
      return null;
    }

    // Get a signed URL valid for 1 year (31536000 seconds)
    const { data: signed } = await db.storage
      .from('import-files')
      .createSignedUrl(path, 31536000);

    return signed?.signedUrl || null;
  } catch (e) {
    console.warn('Upload error:', e);
    return null;
  }
}

async function logImport(fileName, totalRows, imported, errors, previewData, fileUrl = null) {
  const { data: { user } } = await db.auth.getUser();
  const payload = {
    user_email:   user?.email || 'unknown',
    file_name:    fileName,
    total_rows:   totalRows,
    imported:     imported,
    errors:       errors,
    status:       errors > 0 ? (imported > 0 ? 'partial' : 'failed') : 'success',
    preview_json: JSON.stringify(previewData || []).slice(0, 50000),
    file_url:     fileUrl
  };
  await db.from('import_history').insert(payload);
}

async function loadHistory() {
  const tbody = document.getElementById('history-body');
  tbody.innerHTML = '<tr><td colspan="8" class="empty-cell">Loading…</td></tr>';

  const { data, error } = await db
    .from('import_history')
    .select('*')
    .order('created_at', { ascending: false })
    .limit(100);

  if (error) {
    tbody.innerHTML = `<tr><td colspan="8" class="empty-cell" style="color:var(--error)">${error.message}</td></tr>`;
    return;
  }

  if (!data || data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-cell">No import history yet.</td></tr>';
    return;
  }

  tbody.innerHTML = data.map(row => {
    const dt = new Date(row.created_at).toLocaleString();
    const statusColor = row.status === 'success' ? 'var(--success)' :
                        row.status === 'partial'  ? 'var(--warning)' : 'var(--error)';

    const viewBtn = row.preview_json
      ? `<button class="btn-secondary" style="padding:4px 10px;font-size:12px;margin-right:4px" onclick="showHistoryPreview(${row.id})">📊 View Data</button>`
      : '';
    const dlBtn = row.file_url
      ? `<a href="${row.file_url}" download target="_blank" class="btn-secondary" style="padding:4px 10px;font-size:12px;text-decoration:none">⬇ Download</a>`
      : '';

    return `<tr>
      <td>${dt}</td>
      <td>${row.user_email}</td>
      <td style="font-weight:500">${row.file_name}</td>
      <td>${row.total_rows ?? '–'}</td>
      <td style="color:var(--success);font-weight:600">${row.imported ?? '–'}</td>
      <td style="color:var(--error)">${row.errors ?? '–'}</td>
      <td><span style="color:${statusColor};font-weight:600;text-transform:capitalize">${row.status}</span></td>
      <td style="white-space:nowrap">${viewBtn}${dlBtn}${!viewBtn && !dlBtn ? '–' : ''}</td>
    </tr>`;
  }).join('');

  // store for preview lookup
  window._historyCache = Object.fromEntries(data.map(r => [r.id, r]));
}

window.showHistoryPreview = function(id) {
  const row = window._historyCache?.[id];
  if (!row || !row.preview_json) return;

  let preview;
  try { preview = JSON.parse(row.preview_json); } catch { preview = []; }

  if (!Array.isArray(preview) || preview.length === 0) {
    toast('No preview data available.', 'error');
    return;
  }

  const headers = Object.keys(preview[0]);
  document.getElementById('history-preview-title').textContent = `Preview: ${row.file_name}`;
  document.getElementById('history-preview-head').innerHTML =
    '<tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr>';
  document.getElementById('history-preview-body').innerHTML =
    preview.slice(0, 100).map(r =>
      '<tr>' + headers.map(h => `<td>${r[h] ?? ''}</td>`).join('') + '</tr>'
    ).join('');

  document.getElementById('history-preview-modal').style.display = 'flex';
};

document.getElementById('refresh-history-btn').addEventListener('click', loadHistory);


// ══════════════════════════════════════════════════════════
//  USER ACCOUNT MANAGEMENT
// ══════════════════════════════════════════════════════════
async function loadUsers() {
  const tbody = document.getElementById('users-body');
  tbody.innerHTML = '<tr><td colspan="4" class="empty-cell">Loading…</td></tr>';

  // Use Supabase admin API via an edge function OR the auth.admin methods
  // Since we use the anon key, we list users via a custom DB view/function
  const { data, error } = await db.rpc('list_users');

  if (error) {
    tbody.innerHTML = `<tr><td colspan="4" class="empty-cell" style="color:var(--error)">
      ${error.message}<br><small>Make sure the <code>list_users</code> function is created in Supabase.</small>
    </td></tr>`;
    return;
  }

  // Get current user and exclude them from the list
  const { data: { user: currentUser } } = await db.auth.getUser();
  const others = (data || []).filter(u => u.id !== currentUser?.id);

  if (others.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-cell">No other users found. Use "+ Add User" to create accounts.</td></tr>';
    return;
  }

  tbody.innerHTML = others.map(u => {
    const created  = u.created_at     ? new Date(u.created_at).toLocaleDateString()      : '–';
    const lastSign = u.last_sign_in_at ? new Date(u.last_sign_in_at).toLocaleDateString() : 'Never';
    const userRole  = u.role || 'user';
    const roleStyle = userRole === 'admin'
      ? 'background:var(--accent-bg);color:var(--accent-dark);border:1px solid rgba(52,211,153,.3)'
      : 'background:#f1f5f9;color:var(--muted);border:1px solid var(--border-dark)';
    return `<tr>
      <td>
        <div style="display:flex;align-items:center;gap:8px">
          <div style="width:30px;height:30px;border-radius:50%;background:var(--accent-bg);color:var(--accent-dark);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;flex-shrink:0">
            ${u.email.charAt(0).toUpperCase()}
          </div>
          <span>${u.email}</span>
        </div>
      </td>
      <td><span style="font-size:11px;font-weight:600;padding:3px 9px;border-radius:20px;${roleStyle};text-transform:capitalize">${userRole}</span></td>
      <td>${created}</td>
      <td>${lastSign}</td>
      <td style="white-space:nowrap">
        <button class="btn-secondary" style="padding:4px 12px;font-size:12px;margin-right:6px" onclick="openEditUser('${u.id}','${u.email}','${userRole}')">&#9998; Edit</button>
        <button class="btn-danger" style="padding:4px 12px;font-size:12px" onclick="deleteUser('${u.id}','${u.email}')">&#128465; Delete</button>
      </td>
    </tr>`;
  }).join('');
}


window.openEditUser = function(id, email, role = 'user') {
  document.getElementById('user-modal-title').textContent = 'Edit User';
  document.getElementById('user-modal-email').value    = email;
  document.getElementById('user-modal-password').value = '';
  document.getElementById('user-modal-id').value       = id;
  document.getElementById('user-modal-role').value     = role;
  document.getElementById('user-modal').style.display  = 'flex';
};


window.deleteUser = async function(id, email) {
  if (!confirm(`Are you sure you want to delete the account:\n\n${email}\n\nThis cannot be undone.`)) return;

  const { error } = await db.rpc('delete_user', { user_id: id });

  if (error) {
    toast('Delete failed: ' + error.message, 'error');
  } else {
    toast(`User "${email}" deleted.`, 'success');
    loadUsers(); // refresh the list
  }
};


document.getElementById('add-user-btn').addEventListener('click', () => {
  document.getElementById('user-modal-title').textContent = 'Add New User';
  document.getElementById('user-modal-email').value    = '';
  document.getElementById('user-modal-password').value = '';
  document.getElementById('user-modal-id').value       = '';
  document.getElementById('user-modal-role').value     = 'user';
  document.getElementById('user-modal').style.display  = 'flex';
});


document.getElementById('user-modal-cancel').addEventListener('click', () => {
  document.getElementById('user-modal').style.display = 'none';
});

document.getElementById('user-modal-save').addEventListener('click', async () => {
  const id       = document.getElementById('user-modal-id').value.trim();
  const email    = document.getElementById('user-modal-email').value.trim();
  const password = document.getElementById('user-modal-password').value.trim();
  const role     = document.getElementById('user-modal-role').value;

  if (!email) { toast('Email is required.', 'error'); return; }

  const btn = document.getElementById('user-modal-save');
  btn.disabled = true;
  btn.textContent = 'Saving…';

  let error, savedId = id;

  if (!id) {
    // Create new user
    if (!password) { toast('Password is required for new users.', 'error'); btn.disabled = false; btn.textContent = 'Save'; return; }
    const res = await db.rpc('create_user', { user_email: email, user_password: password });
    error = res.error;
    // Get the new user's ID to set role
    if (!error) {
      const { data: users } = await db.rpc('list_users');
      const newUser = (users || []).find(u => u.email === email);
      if (newUser) savedId = newUser.id;
    }
  } else {
    // Update existing user
    const payload = { user_id: id, new_email: email };
    if (password) payload.new_password = password;
    const res = await db.rpc('update_user', payload);
    error = res.error;
  }

  // Set role if no error
  if (!error && savedId) {
    await db.rpc('set_user_role', { user_id: savedId, new_role: role });
  }

  btn.disabled = false;
  btn.textContent = 'Save';

  if (error) {
    toast('Error: ' + error.message, 'error');
  } else {
    toast(id ? 'User updated!' : 'User created!', 'success');
    document.getElementById('user-modal').style.display = 'none';
    loadUsers();
  }
});



// ─────────────────────────────────────────────
// START
// ─────────────────────────────────────────────
init();


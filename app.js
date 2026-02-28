// ─────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────
const SUPABASE_URL  = 'https://fmspuwcfygbsklgnrray.supabase.co';
const SUPABASE_KEY  = 'sb_publishable_y2oAsYVOdrjyDcxAa5FX1w_6saruoJY';
const TABLE         = 'contacts';
const BATCH         = 100;
const PAGE_SIZE     = 50;

// Excel header → DB column name
const COL = {
  'Created'                                        : 'created_date',
  'Contact Id'                                     : 'contact_id',
  'Prospect Type'                                  : 'prospect_type',
  'Marketing Stage of Contact'                     : 'marketing_stage',
  'Type'                                           : 'contact_type',
  'Name'                                           : 'full_name',
  'First Name'                                     : 'first_name',
  'Last Name'                                      : 'last_name',
  'Phone'                                          : 'phone',
  'Email'                                          : 'email',
  'Subject Property Address'                       : 'subject_property_address',
  'Street address'                                 : 'street_address',
  'City'                                           : 'city',
  'State'                                          : 'state',
  'Postal Code'                                    : 'postal_code',
  'Property County'                                : 'property_county',
  'Mailing Street'                                 : 'mailing_street',
  'Mailing City'                                   : 'mailing_city',
  'Mailing State'                                  : 'mailing_state',
  'Mailing Zipcode'                                : 'mailing_zipcode',
  'Source'                                         : 'source',
  'Auction Date'                                   : 'auction_date',
  'Property Type'                                  : 'property_type',
  'House Style'                                    : 'house_style',
  'Year Built'                                     : 'year_built',
  'Square Footage'                                 : 'square_footage',
  'Beds'                                           : 'beds',
  'Baths'                                          : 'baths',
  'Pool'                                           : 'pool',
  'Last Sales Date'                                : 'last_sales_date',
  'Deed Date'                                      : 'deed_date',
  '-- Record - Believed Bad OR Correct OR Deceased --': 'record_status',
  '4. Bad Call, Kill or NOT A Fit'                 : 'bad_call_kill',
  'Owner Type'                                     : 'owner_type',
  'Retail Score'                                   : 'retail_score',
  'Rental Score'                                   : 'rental_score',
  'Loan To Value'                                  : 'loan_to_value',
  'Loan Balance +15K'                              : 'loan_balance_15k',
  'Original Loan'                                  : 'original_loan',
  'Second POC Name'                                : 'second_poc_name',
  'Deceased Owner'                                 : 'deceased_owner',
  'PR File Date'                                   : 'pr_file_date',
  '1st Date Contact Added'                         : 'first_date_contact_added',
  'Equity of Property'                             : 'equity_of_property',
  'Date of Death'                                  : 'date_of_death',
  'Loan Date'                                      : 'loan_date',
  'Foreclosing Lien'                               : 'foreclosing_lien',
  'Complaint Type'                                 : 'complaint_type',
  'ROS OFFER'                                      : 'ros_offer',
  'Foreclosures'                                   : 'foreclosures',
  'Absentee Owner'                                 : 'absentee_owner',
  'High Equity'                                    : 'high_equity',
  'Pre-Foreclosure'                                : 'pre_foreclosure',
  'Deceased Probate'                               : 'deceased_probate',
  'Last Sales Price'                               : 'last_sales_price',
  'Recording Date'                                 : 'recording_date',
  'AVM'                                            : 'avm',
  'Estimated Value'                                : 'estimated_value',
  'Market Value'                                   : 'market_value',
  'Assessed Total'                                 : 'assessed_total',
  'Rental Estimate Low'                            : 'rental_estimate_low',
  'Rental Estimate High'                           : 'rental_estimate_high',
  'Total Loans'                                    : 'total_loans',
  'Estimated Mortgage Balance'                     : 'estimated_mortgage_balance',
  'Estimated Mortgage Payment'                     : 'estimated_mortgage_payment',
  'Mortgage Interest Rate'                         : 'mortgage_interest_rate',
  'LTV'                                            : 'ltv',
  'Maturity Date'                                  : 'maturity_date',
  'Free And Clear'                                 : 'free_and_clear',
  'Equity Percent'                                 : 'equity_percent',
  'Additional Phones'                              : 'additional_phones',
  'Additional Emails'                              : 'additional_emails',
};

const LABELS = Object.fromEntries(Object.entries(COL).map(([k,v]) => [v, k]));
const ALL_COLS = Object.values(COL);
const DEFAULT_COLS = [
  'prospect_type','full_name','phone','email',
  'city','state','postal_code','marketing_stage','source','contact_type'
];

// ─────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────
let db           = null;
let parsedRows   = [];
let currentPage  = 1;
let totalRows    = 0;
let filters      = { prospect:'', state:'', stage:'', q:'' };
let visibleCols  = [...DEFAULT_COLS];
let searchTimer  = null;
let selectedIds  = new Set();   // DB row ids selected for deletion
let currentRows  = [];          // rows currently shown in the table

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
}

async function showApp() {
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('main-app').style.display = 'flex';
  await connectDB();
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
  await db.auth.signOut();
  toast('Signed out successfully', 'info');
});

// ─────────────────────────────────────────────
// DB CONNECTION
// ─────────────────────────────────────────────
const CREATE_SQL = `CREATE TABLE IF NOT EXISTS contacts (
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
ALTER TABLE contacts DISABLE ROW LEVEL SECURITY;`;

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

function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) { toast('Please upload .xlsx, .xls or .csv', 'error'); return; }
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

    parsedRows = rawRows.map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const dbCol = COL[h];
        if (dbCol) {
          let val = row[i];
          if (val instanceof Date) val = val.toISOString().split('T')[0];
          obj[dbCol] = val !== null && val !== undefined ? String(val).trim() : '';
        }
      });
      return obj;
    }).filter(o => Object.keys(o).length > 0);

    document.getElementById('file-bar').style.display = 'flex';
    document.getElementById('file-name').textContent = file.name;
    document.getElementById('file-meta').textContent = `${rawRows.length.toLocaleString()} rows · ${headers.length} columns · ${fmtSize(file.size)}`;
    document.getElementById('import-options').style.display = 'block';
    document.getElementById('import-actions').style.display = 'block';
    document.getElementById('preview-badge').textContent = rawRows.length.toLocaleString() + ' rows';
    showPreview(headers, parsedRows);
  } catch(e) {
    toast('Failed to parse file: ' + e.message, 'error');
  }
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

  let inserted = 0, skipped = 0, errors = 0;
  const total  = parsedRows.length;

  for (let i = 0; i < total; i += BATCH) {
    const batch = parsedRows.slice(i, i + BATCH);
    try {
      const { error } = mode === 'update'
        ? await db.from(TABLE).upsert(batch, { onConflict:'contact_id', ignoreDuplicates:false })
        : await db.from(TABLE).upsert(batch, { onConflict:'contact_id', ignoreDuplicates:true  });
      if (error) { errors += batch.length; console.error(error); }
      else        { inserted += batch.length; }
    } catch(e) { errors += batch.length; }

    const done = Math.min(i + BATCH, total);
    const pct  = Math.round((done / total) * 100);
    document.getElementById('prog-fill').style.width  = pct + '%';
    document.getElementById('prog-pct').textContent   = pct + '%';
    document.getElementById('prog-label').textContent = `Importing batch ${Math.ceil((i+1)/BATCH)} of ${Math.ceil(total/BATCH)}…`;
    document.getElementById('prog-stats').textContent = `${done.toLocaleString()} of ${total.toLocaleString()} rows processed`;
  }

  btn.disabled = false;
  const ok = errors === 0;
  document.getElementById('result-icon').textContent  = ok ? 'Complete' : 'Warning';
  document.getElementById('result-title').textContent = ok ? 'Import Successful' : 'Import Completed with Errors';
  document.getElementById('result-msg').textContent   = `Processed ${total.toLocaleString()} rows from Excel.`;
  document.getElementById('result-stats').innerHTML   = `
    <div class="rs"><strong style="color:var(--success)">${inserted.toLocaleString()}</strong><small>Imported</small></div>
    <div class="rs"><strong style="color:var(--warning)">${skipped.toLocaleString()}</strong><small>Skipped</small></div>
    <div class="rs"><strong style="color:var(--error)">${errors.toLocaleString()}</strong><small>Errors</small></div>`;
  document.getElementById('result-card').style.display = 'block';
  toast(ok ? `Import complete: ${inserted.toLocaleString()} rows` : `Import done with ${errors} errors`, ok ? 'success' : 'error');
}

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
  const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,8);
  document.getElementById('stats-row').innerHTML = sorted.map(([pt, cnt]) =>
    `<div class="stat-chip" onclick="filterByProspect('${escHtml(pt)}')">
       <strong>${cnt.toLocaleString()}</strong><small>${escHtml(pt)}</small>
     </div>`
  ).join('');
}

window.filterByProspect = function(pt) {
  filters.prospect = pt;
  document.getElementById('filter-prospect').value = pt;
  currentPage = 1; loadRecords();
};

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

// ─────────────────────────────────────────────
// FILTER EVENTS
// ─────────────────────────────────────────────
document.getElementById('filter-prospect').addEventListener('change', e => { filters.prospect = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('filter-state').addEventListener('change',    e => { filters.state    = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('filter-stage').addEventListener('change',    e => { filters.stage    = e.target.value; currentPage=1; loadRecords(); });
document.getElementById('search-input').addEventListener('input', e => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { filters.q = e.target.value.trim(); currentPage=1; loadRecords(); }, 400);
});
document.getElementById('reset-btn').addEventListener('click', () => {
  filters = { prospect:'', state:'', stage:'', q:'' };
  ['filter-prospect','filter-state','filter-stage'].forEach(id => { document.getElementById(id).value = ''; });
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
// START
// ─────────────────────────────────────────────
init();

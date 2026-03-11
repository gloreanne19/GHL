/***************************************************************
 * 05_SkipTrace_Merge_NoDuplicates.gs
 *
 * FULL MERGE SCRIPT (3 parts) — NO duplicates for Mojo + Direct
 * - Part A builds header + writes Xleads rows
 * - Part B appends Mojo rows WITHOUT creating duplicates (updates existing rows)
 * - Part C appends Direct rows WITHOUT creating duplicates (updates existing rows)
 *
 * Match rule for “already exists” (Mojo + Direct):
 * - Normalize Property Address only
 * - If Direct “Input Property Address” (or mailing address) matches Skip Trace Finished col A
 *   → DO NOT append; update contact fields in place.
 *
 * Excludes these For Import columns from Skip Trace Finished entirely:
 * Decedent, Second POC Name, Appraised Value, Representative Name,
 * First Name, Last Name, Representative Address, Representative City,
 * Representative State, Representative Zip, PR File Date, County, Type
 ***************************************************************/

// =============================================================
// Shared Helpers
// =============================================================
function STM_headers_(sh) {
  const lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
}
function STM_upsertSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function STM_clearAndHeaders_(sh, H) {
  sh.clearContents();
  if (H && H.length) sh.getRange(1, 1, 1, H.length).setValues([H]);
  sh.setFrozenRows(1);
}
function STM_findIdx_(H, syns) {
  const want = (syns || []).map(s => String(s || '').toLowerCase().trim());
  for (let i = 0; i < H.length; i++) {
    const h = String(H[i] || '').toLowerCase().trim();
    if (!h) continue;
    if (want.indexOf(h) !== -1) return i;
  }
  return -1;
}
function STM_colToIndex0_(letter) {
  const s = String(letter || '').toUpperCase().trim();
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return Math.max(0, n - 1);
}
function STM_safe_(row, idx) {
  return (idx >= 0 && idx < row.length) ? row[idx] : '';
}
function STM_norm_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
function STM_normAddrOnly_(s) {
  return STM_norm_(s);
}
function STM_key_(a, c, s) {
  const A = STM_norm_(a), C = STM_norm_(c), S = STM_norm_(s);
  if (!A && !C && !S) return '';
  return `${A}|${C}|${S}`;
}
function STM_digits5_(v) {
  const m = String(v || '').match(/(\d{5})/);
  return m ? m[1] : '';
}
function STM_bestZip_(rawZip, addr, city, state) {
  const z1 = STM_digits5_(rawZip);
  if (z1) return { zip: z1, backfilled: false };

  const combo = [addr, city, state].map(x => String(x || '')).join(' ');
  const z2 = STM_digits5_(combo);
  if (z2) return { zip: z2, backfilled: true };

  const z3 = STM_digits5_(addr);
  if (z3) return { zip: z3, backfilled: true };

  return { zip: '', backfilled: false };
}
function STM_has7digits_(v) {
  return String(v || '').replace(/\D+/g, '').length >= 7;
}
function STM_join_(arr) {
  return (arr || []).map(x => String(x || '').trim()).filter(Boolean).join(', ');
}
function STM_unique_(arr) {
  const seen = new Set();
  const out = [];
  for (const x of (arr || [])) {
    const v = String(x || '').trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  return out;
}
function STM_appendRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}
function STM_forceZipText5_(sheet, startRow, numRows) {
  if (!numRows) return;
  // Property Zip is column D (4)
  sheet.getRange(startRow, 4, numRows, 1).setNumberFormat('@');
  const zVals = sheet.getRange(startRow, 4, numRows, 1).getValues();
  for (let i = 0; i < zVals.length; i++) zVals[i][0] = STM_digits5_(zVals[i][0]) || '';
  sheet.getRange(startRow, 4, numRows, 1).setValues(zVals);
}
function STM_splitComma_(s) {
  return String(s || '')
    .split(',')
    .map(x => String(x || '').trim())
    .filter(Boolean);
}
function STM_mergeLists_(a, b) {
  return STM_unique_([].concat(a || [], b || []));
}

// =============================================================
// Exclude For Import Columns (NOT added to Skip Trace Finished)
// =============================================================
function STM_excludeForImportHeaders_() {
  return new Set([
    'Decedent',
    'Second POC Name',
    'Appraised Value',
    'Representative Name',
    'First Name',
    'Last Name',
    'Representative Address',
    'Representative City',
    'Representative State',
    'Representative Zip',
    'PR File Date',
    'County',
    'Type'
  ]);
}

// =============================================================
// For Import enrichment map + dyn headers (EXCLUDING fields above)
// =============================================================
function STM_buildForImportEnrichment_() {
  const ss = SpreadsheetApp.getActive();
  const fin = ss.getSheetByName('For Import');
  if (!fin || fin.getLastRow() < 2) throw new Error('Missing or empty: "For Import"');

  const HF = STM_headers_(fin);

  const finAddr = STM_findIdx_(HF, ['property address','propertyaddress','street address','property street address','address']);
  const finCity = STM_findIdx_(HF, ['property city','propertycity','city']);
  const finState= STM_findIdx_(HF, ['property state','propertystate','state']);
  const finZip  = STM_findIdx_(HF, ['property zip','property zip code','zip','zip code','zipcode','postal','postal code','property postal','propertypostal']);

  if (finAddr === -1 || finCity === -1 || finState === -1) {
    throw new Error('For Import must have Address + City + State headers (or add synonyms).');
  }

  // All For Import columns EXCEPT: address/city/state/zip + excluded headers set
  const finExclude = new Set([finAddr, finCity, finState, finZip].filter(i => i !== -1));
  const EXCL = STM_excludeForImportHeaders_();

  const finDynIdx = [];
  const finDynHdr = [];
  for (let i = 0; i < HF.length; i++) {
    if (finExclude.has(i)) continue;
    const h = String(HF[i] || '').trim();
    if (!h) continue;
    if (EXCL.has(h)) continue;
    finDynIdx.push(i);
    finDynHdr.push(h);
  }

  const finBody = fin.getRange(2, 1, fin.getLastRow() - 1, fin.getLastColumn()).getValues();
  const finMap = new Map();
  for (const r of finBody) {
    const key = STM_key_(r[finAddr], r[finCity], r[finState]);
    if (!key) continue;
    if (!finMap.has(key)) finMap.set(key, finDynIdx.map(i => STM_safe_(r, i)));
  }

  return { fin, HF, finAddr, finCity, finState, finZip, finDynIdx, finDynHdr, finMap };
}

// =============================================================
// Xleads tail: Xleads columns C..end, skipping G..J
// =============================================================
function STM_buildXleadsTail_() {
  const ss = SpreadsheetApp.getActive();
  const xres = ss.getSheetByName('Xleads Results');
  if (!xres || xres.getLastRow() < 2) return { HX: [], xTailHdr: [], xTailMap: new Map() };

  const HX = STM_headers_(xres);

  const xTailIdx = [];
  const xTailHdr = [];
  for (let i = 2; i < HX.length; i++) {
    if (i >= 6 && i <= 9) continue; // skip G..J
    xTailIdx.push(i);
    xTailHdr.push(HX[i]);
  }

  let xAddr = STM_findIdx_(HX, ['propertyaddress','property address','address']); if (xAddr === -1) xAddr = STM_colToIndex0_('G');
  let xCity = STM_findIdx_(HX, ['propertycity','property city','city']);
  let xState= STM_findIdx_(HX, ['propertystate','property state','state']);

  const body = xres.getRange(2, 1, xres.getLastRow() - 1, xres.getLastColumn()).getValues();
  const xTailMap = new Map();
  for (const r of body) {
    const key = STM_key_(STM_safe_(r, xAddr), STM_safe_(r, xCity), STM_safe_(r, xState));
    if (!key) continue;
    xTailMap.set(key, xTailIdx.map(i => STM_safe_(r, i)));
  }

  return { HX, xTailHdr, xTailMap };
}

// =============================================================
// Header plan (NO Contact2Name)
// Fixed A–M -> Xleads Tail -> For Import dyn
// =============================================================
function STM_buildHeaderPlan_(finDynHdr, xTailHdr) {
  const fixedHdr = [
    'Property Address','Property City','Property State','Property Zip',
    'First Name','Last Name','Full Name',
    'Phone','Additional Phone','Landline','Additional Landlines',
    'Email','Additional Emails'
  ];
  const HOUT = fixedHdr.concat(xTailHdr).concat(finDynHdr);
  return { fixedHdr, xTailHdr, finDynHdr, HOUT };
}

// =============================================================
// PART A — Build header + write Xleads rows
// =============================================================
function ST_merge_A_xleads_buildHeaderAndWrite() {
  const ss = SpreadsheetApp.getActive();

  const xres = ss.getSheetByName('Xleads Results');
  if (!xres || xres.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Xleads Results"');
    return;
  }

  const dest = STM_upsertSheet_(ss, 'Skip Trace Finished');
  const qa   = STM_upsertSheet_(ss, 'Merge QA');

  const finPack  = STM_buildForImportEnrichment_();
  const tailPack = STM_buildXleadsTail_();
  const plan     = STM_buildHeaderPlan_(finPack.finDynHdr, tailPack.xTailHdr);

  STM_clearAndHeaders_(dest, plan.HOUT);

  qa.clearContents(); qa.clearFormats();
  qa.getRange(1,1,1,5).setValues([['Metric','Value','Notes','Timestamp','Source']]);
  qa.setFrozenRows(1);

  const HX = STM_headers_(xres);
  const IX = (function(){
    let addr = STM_findIdx_(HX, ['propertyaddress','property address','address']); if (addr === -1) addr = STM_colToIndex0_('G');
    let city = STM_findIdx_(HX, ['propertycity','property city','city']);
    let state= STM_findIdx_(HX, ['propertystate','property state','state']);

    let zip = STM_findIdx_(HX, [
      'propertypostalcode','property postal code','propertypostal','property postal',
      'property zipcode','property zip code','property zip',
      'zip','zip code','zipcode','postal','postal code'
    ]);
    if (zip === -1 && HX.length > 9) {
      const hJ = String(HX[9] || '').toLowerCase();
      if (/zip|postal/.test(hJ)) zip = 9;
    }

    let first = STM_findIdx_(HX, ['firstname','first name']);
    let last  = STM_findIdx_(HX, ['lastname','last name']);

    let p1  = STM_findIdx_(HX, ['contact1phone_1','phone 1','phone1','primary phone']); if (p1 === -1) p1 = STM_colToIndex0_('V');
    let p1t = STM_findIdx_(HX, ['contact1phone_1_type','phone 1 type']);
    let p2  = STM_findIdx_(HX, ['contact1phone_2','phone 2','phone2']);                 if (p2 === -1) p2 = STM_colToIndex0_('AA');
    let p2t = STM_findIdx_(HX, ['contact1phone_2_type','phone 2 type']);
    let p3  = STM_findIdx_(HX, ['contact1phone_3','phone 3','phone3']);                 if (p3 === -1) p3 = STM_colToIndex0_('AG');
    let p3t = STM_findIdx_(HX, ['contact1phone_3_type','phone 3 type']);

    let e1 = STM_findIdx_(HX, ['contact1email_1','email 1','email']);
    let e2 = STM_findIdx_(HX, ['contact1email_2','email 2']);
    let e3 = STM_findIdx_(HX, ['contact1email_3','email 3']);

    return {addr, city, state, zip, first, last, p1,p1t,p2,p2t,p3,p3t, e1,e2,e3};
  })();

  function routeXPhones_(r) {
    const out = { phone:'', addlPhones:[], landline:'', addlLandlines:[] };
    function place(num, typ, primary) {
      const n = String(num || '').trim();
      if (!STM_has7digits_(n)) return;
      const t = String(typ || '').toLowerCase().trim();
      if (t === 'mobile') {
        if (primary && !out.phone) out.phone = n; else out.addlPhones.push(n);
      } else if (t === 'residential') {
        if (primary && !out.landline) out.landline = n; else out.addlLandlines.push(n);
      }
    }
    place(STM_safe_(r, IX.p1), STM_safe_(r, IX.p1t), true);
    place(STM_safe_(r, IX.p2), STM_safe_(r, IX.p2t), false);
    place(STM_safe_(r, IX.p3), STM_safe_(r, IX.p3t), false);
    out.addlPhones = STM_unique_(out.addlPhones);
    out.addlLandlines = STM_unique_(out.addlLandlines);
    return out;
  }
  function routeXEmails_(r) {
    const email = String(STM_safe_(r, IX.e1) || '').trim();
    const addl = STM_unique_([STM_safe_(r, IX.e2), STM_safe_(r, IX.e3)]);
    return { email, addlEmails: addl };
  }

  const body = xres.getRange(2, 1, xres.getLastRow() - 1, xres.getLastColumn()).getValues();

  const outRows = [];
  let zipBackfilled = 0;
  let missingFin = 0;
  let missingTail = 0;

  for (const r of body) {
    const A = STM_safe_(r, IX.addr);
    const B = STM_safe_(r, IX.city);
    const C = STM_safe_(r, IX.state);

    const z = STM_bestZip_(STM_safe_(r, IX.zip), A, B, C);
    if (z.backfilled && z.zip) zipBackfilled++;
    const D = z.zip;

    const E = STM_safe_(r, IX.first);
    const F = STM_safe_(r, IX.last);
    const G = ''; // Full Name blank by spec for Xleads

    const ph = routeXPhones_(r);
    const em = routeXEmails_(r);

    const key = STM_key_(A, B, C);

    let dyn = [];
    if (finPack.finDynHdr.length) {
      const got = finPack.finMap.get(key);
      if (got) dyn = got;
      else { missingFin++; dyn = new Array(finPack.finDynHdr.length).fill(''); }
    }

    let tailOut = [];
    if (tailPack.xTailHdr.length) {
      const t = tailPack.xTailMap.get(key);
      if (t) tailOut = t;
      else { missingTail++; tailOut = new Array(tailPack.xTailHdr.length).fill(''); }
    }

    outRows.push(
      [
        A,B,C,D,
        E,F,G,
        ph.phone, STM_join_(ph.addlPhones),
        ph.landline, STM_join_(ph.addlLandlines),
        em.email, STM_join_(em.addlEmails)
      ]
      .concat(tailOut)
      .concat(dyn)
    );
  }

  if (outRows.length) dest.getRange(2, 1, outRows.length, plan.HOUT.length).setValues(outRows);
  if (outRows.length) STM_forceZipText5_(dest, 2, outRows.length);

  const now = new Date();
  const rows = [
    ['Rows written (Xleads)', outRows.length, 'Part A built header + wrote Xleads rows', now, 'Merge Part A'],
    ['ZIP backfilled (Xleads)', zipBackfilled, 'Parsed 5-digit ZIP from text when ZIP blank', now, 'Merge Part A'],
    ['Missing For Import enrichment matches (Xleads)', finPack.finDynHdr.length ? missingFin : 0, 'Key=address+city+state', now, 'Merge Part A'],
    ['Missing Xleads tail matches (Xleads)', tailPack.xTailHdr.length ? missingTail : 0, 'Should be near zero for Xleads', now, 'Merge Part A']
  ];
  qa.getRange(2,1,rows.length,5).setValues(rows);
  qa.getRange(2,4,rows.length,1).setNumberFormat('MM/dd/yyyy HH:mm');

  SpreadsheetApp.getUi().alert(
    'Part A complete: header built + Xleads written.\n' +
    `Rows: ${outRows.length}\nZIP backfilled: ${zipBackfilled}\nMissing For Import matches: ${finPack.finDynHdr.length ? missingFin : 0}\nMissing tail matches: ${tailPack.xTailHdr.length ? missingTail : 0}`
  );
}

// =============================================================
// PART B — Mojo: NO duplicates (update if address exists)
// =============================================================
function ST_merge_B_appendMojo() {
  const ss = SpreadsheetApp.getActive();

  const mojo = ss.getSheetByName('Mojo Results');
  if (!mojo || mojo.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Mojo Results"');
    return;
  }

  const dest = ss.getSheetByName('Skip Trace Finished');
  if (!dest || dest.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished missing or empty. Run Part A first.');
    return;
  }

  const qa = STM_upsertSheet_(ss, 'Merge QA');

  const finPack  = STM_buildForImportEnrichment_();
  const tailPack = STM_buildXleadsTail_();
  STM_buildHeaderPlan_(finPack.finDynHdr, tailPack.xTailHdr); // plan not needed directly here but keeps logic consistent

  const existingVals = dest.getRange(2, 1, dest.getLastRow() - 1, dest.getLastColumn()).getValues();
  const existingMap = new Map(); // addrNorm -> sheetRowNumber
  for (let i = 0; i < existingVals.length; i++) {
    const key = STM_normAddrOnly_(existingVals[i][0]);
    if (!key) continue;
    if (!existingMap.has(key)) existingMap.set(key, i + 2);
  }

  const outFirstIdx = 4, outLastIdx = 5, outFullIdx = 6;
  const outPhoneIdx = 7, outAddPhoneIdx = 8;
  const outLandIdx = 9, outAddLandIdx = 10;
  const outEmailIdx = 11, outAddEmailIdx = 12;

  const HM = STM_headers_(mojo);
  const IM = (function(){
    let addr = STM_findIdx_(HM, ['street address','propertyaddress','property address','address']); if (addr === -1) addr = STM_colToIndex0_('A');
    let city = STM_findIdx_(HM, ['city','property city','propertycity']);                         if (city === -1) city = STM_colToIndex0_('B');
    let state= STM_findIdx_(HM, ['state','property state','propertystate']);                      if (state=== -1) state= STM_colToIndex0_('C');
    let zip  = STM_findIdx_(HM, ['zip','zip code','zipcode','postal','postal code','property zip','property zip code','propertypostalcode']);
    if (zip === -1) zip = STM_colToIndex0_('D');

    let first = STM_findIdx_(HM, ['first name','firstname','first']);
    let last  = STM_findIdx_(HM, ['last name','lastname','last']);
    let full  = STM_findIdx_(HM, ['full name','fullname','name']);

    let primary = STM_findIdx_(HM, ['primary phone','primaryphone']); if (primary === -1) primary = STM_colToIndex0_('N');

    const mobiles = [];
    for (let i=0;i<HM.length;i++){
      const h = String(HM[i]||'').toLowerCase().replace(/\s+/g,'').trim();
      if (/^mobile\d+$/.test(h)) mobiles.push(i);
    }

    const lands = [];
    for (let i=0;i<HM.length;i++){
      const h = String(HM[i]||'').toLowerCase().replace(/\s+/g,'').trim();
      if (/^phone\d+$/.test(h) || h.includes('landline')) lands.push(i);
    }

    const emails = [];
    for (let i=0;i<HM.length;i++){
      const h = String(HM[i]||'').toLowerCase();
      if (h.includes('email')) emails.push(i);
    }

    return {addr, city, state, zip, first, last, full, primary, mobiles, lands, emails};
  })();

  function routeMojoPhones_(r) {
    let primary = STM_safe_(r, IM.primary);
    if (!STM_has7digits_(primary)) {
      for (const idx of IM.mobiles) {
        const v = STM_safe_(r, idx);
        if (STM_has7digits_(v)) { primary = v; break; }
      }
    }

    const addlPhones = [];
    for (const idx of IM.mobiles) {
      const v = STM_safe_(r, idx);
      if (!STM_has7digits_(v)) continue;
      if (String(v).trim() === String(primary).trim()) continue;
      addlPhones.push(String(v).trim());
    }

    const lands = [];
    for (const idx of IM.lands) {
      const v = STM_safe_(r, idx);
      if (!STM_has7digits_(v)) continue;
      lands.push(String(v).trim());
    }

    const uniqAddl = STM_unique_(addlPhones);
    const uniqLand = STM_unique_(lands);

    return {
      phone: String(primary || '').trim(),
      addlPhones: uniqAddl,
      landline: uniqLand[0] || '',
      addlLandlines: uniqLand.length > 1 ? uniqLand.slice(1) : []
    };
  }

  function routeMojoEmails_(r) {
    const list = [];
    for (const idx of IM.emails) {
      const v = String(STM_safe_(r, idx) || '').trim();
      if (v) list.push(v);
    }
    const uniq = STM_unique_(list);
    return { email: uniq[0] || '', addlEmails: uniq.slice(1) };
  }

  const body = mojo.getRange(2, 1, mojo.getLastRow() - 1, mojo.getLastColumn()).getValues();

  const rowsToAppend = [];
  let updated = 0;
  let appended = 0;
  let zipBackfilled = 0;
  let missingFinForNew = 0;
  let missingTailForNew = 0;

  for (const r of body) {
    const A = STM_safe_(r, IM.addr);
    const addrKey = STM_normAddrOnly_(A);
    if (!addrKey) continue;

    const B = STM_safe_(r, IM.city);
    const C = STM_safe_(r, IM.state);

    const z = STM_bestZip_(STM_safe_(r, IM.zip), A, B, C);
    if (z.backfilled && z.zip) zipBackfilled++;
    const D = z.zip;

    const E = STM_safe_(r, IM.first);
    const F = STM_safe_(r, IM.last);
    const G = STM_safe_(r, IM.full);

    const ph = routeMojoPhones_(r);
    const em = routeMojoEmails_(r);

    const existingRowNum = existingMap.get(addrKey);

    if (existingRowNum) {
      const rowRange = dest.getRange(existingRowNum, 1, 1, dest.getLastColumn());
      const rowVals = rowRange.getValues()[0];

      if (!rowVals[outFirstIdx] && E) rowVals[outFirstIdx] = E;
      if (!rowVals[outLastIdx]  && F) rowVals[outLastIdx]  = F;
      if (!rowVals[outFullIdx]  && G) rowVals[outFullIdx]  = G;
      if (!rowVals[3] && D) rowVals[3] = D;

      if (!rowVals[outPhoneIdx] && ph.phone) rowVals[outPhoneIdx] = ph.phone;

      const curAddPhones = STM_splitComma_(rowVals[outAddPhoneIdx]);
      rowVals[outAddPhoneIdx] = STM_mergeLists_(curAddPhones, ph.addlPhones).join(', ');

      if (!rowVals[outLandIdx] && ph.landline) rowVals[outLandIdx] = ph.landline;

      const curAddLands = STM_splitComma_(rowVals[outAddLandIdx]);
      rowVals[outAddLandIdx] = STM_mergeLists_(curAddLands, ph.addlLandlines).join(', ');

      if (!rowVals[outEmailIdx] && em.email) rowVals[outEmailIdx] = em.email;

      const curAddEmails = STM_splitComma_(rowVals[outAddEmailIdx]);
      rowVals[outAddEmailIdx] = STM_mergeLists_(curAddEmails, em.addlEmails).join(', ');

      rowRange.setValues([rowVals]);
      updated++;
      continue;
    }

    const key = STM_key_(A, B, C);

    let dyn = [];
    if (finPack.finDynHdr.length) {
      const got = finPack.finMap.get(key);
      if (got) dyn = got;
      else { dyn = new Array(finPack.finDynHdr.length).fill(''); missingFinForNew++; }
    }

    let tailOut = [];
    if (tailPack.xTailHdr.length) {
      const t = tailPack.xTailMap.get(key);
      if (t) tailOut = t;
      else { tailOut = new Array(tailPack.xTailHdr.length).fill(''); missingTailForNew++; }
    }

    rowsToAppend.push(
      [
        A,B,C,D,
        E,F,G,
        ph.phone, STM_join_(ph.addlPhones),
        ph.landline, STM_join_(ph.addlLandlines),
        em.email, STM_join_(em.addlEmails)
      ]
      .concat(tailOut)
      .concat(dyn)
    );

    appended++;
    existingMap.set(addrKey, -1);
  }

  if (rowsToAppend.length) {
    const startRow = dest.getLastRow() + 1;
    STM_appendRows_(dest, rowsToAppend);
    STM_forceZipText5_(dest, startRow, rowsToAppend.length);
  }

  const now = new Date();
  const qaStart = qa.getLastRow() + 1;
  qa.getRange(qaStart, 1, 1, 5).setValues([[
    'Mojo merge (no duplicates)',
    `Updated ${updated}, appended ${appended}`,
    `ZIP backfilled: ${zipBackfilled} | Missing FI (new): ${missingFinForNew} | Missing tail (new): ${missingTailForNew}`,
    now,
    'Merge Part B'
  ]]);
  qa.getRange(qaStart, 4, 1, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  SpreadsheetApp.getUi().alert(
    'Part B complete: Mojo merged (no duplicates).\n' +
    `Updated existing rows: ${updated}\nAppended new rows: ${appended}\nZIP backfilled: ${zipBackfilled}`
  );
}

// =============================================================
// PART C — Direct: NO duplicates + phone type routing to Landline
// =============================================================
function ST_merge_C_appendDirect() {
  const ss = SpreadsheetApp.getActive();

  const dskip = ss.getSheetByName('Direct Skip Results');
  if (!dskip || dskip.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Direct Skip Results"');
    return;
  }

  const dest = ss.getSheetByName('Skip Trace Finished');
  if (!dest || dest.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished missing or empty. Run Part A first.');
    return;
  }

  const qa = STM_upsertSheet_(ss, 'Merge QA');

  const finPack  = STM_buildForImportEnrichment_();
  const tailPack = STM_buildXleadsTail_();
  STM_buildHeaderPlan_(finPack.finDynHdr, tailPack.xTailHdr); // plan not needed directly here but keeps logic consistent

  const existingVals = dest.getRange(2, 1, dest.getLastRow() - 1, dest.getLastColumn()).getValues();
  const existingMap = new Map(); // addrNorm -> sheetRowNumber
  for (let i = 0; i < existingVals.length; i++) {
    const key = STM_normAddrOnly_(existingVals[i][0]);
    if (!key) continue;
    if (!existingMap.has(key)) existingMap.set(key, i + 2);
  }

  const outFirstIdx = 4, outLastIdx = 5, outFullIdx = 6;
  const outPhoneIdx = 7, outAddPhoneIdx = 8;
  const outLandIdx = 9, outAddLandIdx = 10;
  const outEmailIdx = 11, outAddEmailIdx = 12;

  const HD = STM_headers_(dskip);
  const ID = (function(){
    let addr = STM_findIdx_(HD, ['input property address','property address','input mailing address','mailing address','input address','address','street address']);
    let city = STM_findIdx_(HD, ['input property city','property city','input mailing city','mailing city','city']);
    let state= STM_findIdx_(HD, ['input property state','property state','input mailing state','mailing state','state']);
    let zip  = STM_findIdx_(HD, ['input property zip','property zip','input mailing zip','input mailing zip code','mailing zip','zip','zip code','zipcode','postal','postal code']);

    let first = STM_findIdx_(HD, ['input first name','first name','firstname']);
    let last  = STM_findIdx_(HD, ['input last name','last name','lastname']);
    let full  = STM_findIdx_(HD, ['input custom field 1','custom field 1','full name','fullname','name']);

    function findPhoneN(n){
      let idx = STM_findIdx_(HD, [`phone ${n}`, `phone${n}`, `phone_${n}`]);
      if (idx !== -1) return idx;
      for (let i=0;i<HD.length;i++){
        const h = String(HD[i]||'').toLowerCase().replace(/\s+/g,'').trim();
        if (h === `phone${n}`) return i;
      }
      return -1;
    }
    function findPhoneTypeN(n){
      const candidates = [
        `phone ${n} type`, `phone${n}type`, `phone${n} type`,
        `phone_${n}_type`, `phone_${n}type`
      ];
      let idx = STM_findIdx_(HD, candidates);
      if (idx !== -1) return idx;
      for (let i=0;i<HD.length;i++){
        const h = String(HD[i]||'').toLowerCase().replace(/\s+/g,'').trim();
        if (h === `phone${n}type`) return i;
      }
      return -1;
    }

    const phones = [];
    const phoneTypes = [];
    for (let n=1;n<=7;n++){
      phones.push(findPhoneN(n));
      phoneTypes.push(findPhoneTypeN(n));
    }

    const emails = [];
    for (let i=0;i<HD.length;i++){
      const h = String(HD[i]||'').toLowerCase();
      if (h.includes('email')) emails.push(i);
    }

    return {addr, city, state, zip, first, last, full, phones, phoneTypes, emails};
  })();

  if (ID.addr === -1) {
    SpreadsheetApp.getUi().alert('Direct Skip Results must have Input Property/Mailing Address (or add synonym).');
    return;
  }

  function normType_(t){
    const s = String(t||'').toLowerCase().trim();
    if (!s) return '';
    if (s.includes('res')) return 'residential';
    if (s.includes('land')) return 'residential';
    if (s.includes('mob')) return 'mobile';
    if (s.includes('cell')) return 'mobile';
    return s;
  }

  function extractDirectPhones_(r){
    const mob = [];
    const lan = [];
    const unk = [];

    for (let k=0;k<ID.phones.length;k++){
      const pIdx = ID.phones[k];
      if (pIdx === -1) continue;

      const numRaw = STM_safe_(r, pIdx);
      if (!STM_has7digits_(numRaw)) continue;
      const num = String(numRaw).trim();

      const tIdx = ID.phoneTypes[k];
      const typ = (tIdx !== -1) ? normType_(STM_safe_(r, tIdx)) : '';

      if (typ === 'residential') lan.push(num);
      else if (typ === 'mobile') mob.push(num);
      else unk.push(num);
    }

    const MOB = STM_unique_(mob);
    const LAN = STM_unique_(lan);
    const UNK = STM_unique_(unk);

    return {
      phone: MOB[0] || '',
      addlPhones: STM_unique_([].concat(MOB.slice(1), UNK)),
      landline: LAN[0] || '',
      addlLandlines: STM_unique_(LAN.slice(1))
    };
  }

  function extractDirectEmails_(r){
    const list = [];
    for (const idx of ID.emails) {
      const v = String(STM_safe_(r, idx) || '').trim();
      if (v) list.push(v);
    }
    const uniq = STM_unique_(list);
    return { email: uniq[0] || '', addlEmails: uniq.slice(1) };
  }

  const body = dskip.getRange(2, 1, dskip.getLastRow() - 1, dskip.getLastColumn()).getValues();

  const rowsToAppend = [];
  let updated = 0;
  let appended = 0;
  let zipBackfilled = 0;
  let missingFinForNew = 0;
  let missingTailForNew = 0;

  for (const r of body) {
    const A = STM_safe_(r, ID.addr);
    const addrKey = STM_normAddrOnly_(A);
    if (!addrKey) continue;

    const B = STM_safe_(r, ID.city);
    const C = STM_safe_(r, ID.state);

    const z = STM_bestZip_(STM_safe_(r, ID.zip), A, B, C);
    if (z.backfilled && z.zip) zipBackfilled++;
    const D = z.zip;

    const E = STM_safe_(r, ID.first);
    const F = STM_safe_(r, ID.last);
    const G = STM_safe_(r, ID.full);

    const ph = extractDirectPhones_(r);
    const em = extractDirectEmails_(r);

    const existingRowNum = existingMap.get(addrKey);

    if (existingRowNum) {
      const rowRange = dest.getRange(existingRowNum, 1, 1, dest.getLastColumn());
      const rowVals = rowRange.getValues()[0];

      if (!rowVals[outFirstIdx] && E) rowVals[outFirstIdx] = E;
      if (!rowVals[outLastIdx]  && F) rowVals[outLastIdx]  = F;
      if (!rowVals[outFullIdx]  && G) rowVals[outFullIdx]  = G;
      if (!rowVals[3] && D) rowVals[3] = D;

      if (!rowVals[outPhoneIdx] && ph.phone) rowVals[outPhoneIdx] = ph.phone;

      const curAddPhones = STM_splitComma_(rowVals[outAddPhoneIdx]);
      rowVals[outAddPhoneIdx] = STM_mergeLists_(curAddPhones, ph.addlPhones).join(', ');

      if (!rowVals[outLandIdx] && ph.landline) rowVals[outLandIdx] = ph.landline;

      const curAddLands = STM_splitComma_(rowVals[outAddLandIdx]);
      rowVals[outAddLandIdx] = STM_mergeLists_(curAddLands, ph.addlLandlines).join(', ');

      if (!rowVals[outEmailIdx] && em.email) rowVals[outEmailIdx] = em.email;

      const curAddEmails = STM_splitComma_(rowVals[outAddEmailIdx]);
      rowVals[outAddEmailIdx] = STM_mergeLists_(curAddEmails, em.addlEmails).join(', ');

      rowRange.setValues([rowVals]);
      updated++;
      continue;
    }

    const key = STM_key_(A, B, C);

    let dyn = [];
    if (finPack.finDynHdr.length) {
      const got = finPack.finMap.get(key);
      if (got) dyn = got;
      else { dyn = new Array(finPack.finDynHdr.length).fill(''); missingFinForNew++; }
    }

    let tailOut = [];
    if (tailPack.xTailHdr.length) {
      const t = tailPack.xTailMap.get(key);
      if (t) tailOut = t;
      else { tailOut = new Array(tailPack.xTailHdr.length).fill(''); missingTailForNew++; }
    }

    rowsToAppend.push(
      [
        A,B,C,D,
        E,F,G,
        ph.phone, STM_join_(ph.addlPhones),
        ph.landline, STM_join_(ph.addlLandlines),
        em.email, STM_join_(em.addlEmails)
      ]
      .concat(tailOut)
      .concat(dyn)
    );

    appended++;
    existingMap.set(addrKey, -1);
  }

  if (rowsToAppend.length) {
    const startRow = dest.getLastRow() + 1;
    STM_appendRows_(dest, rowsToAppend);
    STM_forceZipText5_(dest, startRow, rowsToAppend.length);
  }

  const now = new Date();
  const qaStart = qa.getLastRow() + 1;
  qa.getRange(qaStart, 1, 1, 5).setValues([[
    'Direct merge (no duplicates)',
    `Updated ${updated}, appended ${appended}`,
    `ZIP backfilled: ${zipBackfilled} | Missing FI (new): ${missingFinForNew} | Missing tail (new): ${missingTailForNew}`,
    now,
    'Merge Part C'
  ]]);
  qa.getRange(qaStart, 4, 1, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  SpreadsheetApp.getUi().alert(
    'Part C complete: Direct merged (no duplicates).\n' +
    `Updated existing rows: ${updated}\nAppended new rows: ${appended}\nZIP backfilled: ${zipBackfilled}`
  );
}

// =============================================================
// OPTIONAL: One runner you can add to the menu
// =============================================================
function runFullSkipTracePipeline() {
  ST_merge_A_xleads_buildHeaderAndWrite();
  ST_merge_B_appendMojo();
  ST_merge_C_appendDirect();
  SpreadsheetApp.getUi().alert('Done: Part A + Part B + Part C.');
}
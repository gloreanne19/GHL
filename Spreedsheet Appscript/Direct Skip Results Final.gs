/** PART 3 — Direct Skip
 *
 * Compare: "Not In Mojo" vs "Direct Skip Results"
 *
 * Include in "Not In Direct" when:
 *   (A) Property from Not In Mojo has NO match in Direct Skip, OR
 *   (B) Property matches Direct Skip BUT Direct Skip has NO phone (7+ digits across Phone1..Phone7)
 *
 * OUTPUT:
 * - "Not In Direct" uses ALL headers from "For Import"
 * - Rows written are FULL rows pulled from "For Import" unchanged
 * - Addresses stay exactly where they already are in "For Import"
 *
 * MATCH KEY:
 * - normalized Address + City + State
 *
 * DIRECT SKIP PHONE TEST:
 * - checks Phone1..Phone7 by header pattern
 */

// ------------------------- Helpers -------------------------
function STP3_headers_(sh) {
  const lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
}

function STP3_upsertSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function STP3_clearAndHeaders_(sh, H) {
  sh.clearContents();
  if (H && H.length) sh.getRange(1, 1, 1, H.length).setValues([H]);
  sh.setFrozenRows(1);
}

function STP3_norm_(s) {
  let v = String(s || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const replacements = [
    [/\bnorth\b/g, 'n'],
    [/\bsouth\b/g, 's'],
    [/\beast\b/g, 'e'],
    [/\bwest\b/g, 'w'],
    [/\bnortheast\b/g, 'ne'],
    [/\bnorthwest\b/g, 'nw'],
    [/\bsoutheast\b/g, 'se'],
    [/\bsouthwest\b/g, 'sw'],

    [/\bfirst\b/g, '1st'],
    [/\bsecond\b/g, '2nd'],
    [/\bthird\b/g, '3rd'],
    [/\bfourth\b/g, '4th'],
    [/\bfifth\b/g, '5th'],
    [/\bsixth\b/g, '6th'],
    [/\bseventh\b/g, '7th'],
    [/\beighth\b/g, '8th'],
    [/\bninth\b/g, '9th'],
    [/\btenth\b/g, '10th'],

    [/\bstreet\b/g, 'st'],
    [/\bst\b/g, 'st'],

    [/\bavenue\b/g, 'ave'],
    [/\bave\b/g, 'ave'],

    [/\broad\b/g, 'rd'],
    [/\brd\b/g, 'rd'],

    [/\bdrive\b/g, 'dr'],
    [/\bdr\b/g, 'dr'],

    [/\blane\b/g, 'ln'],
    [/\bln\b/g, 'ln'],

    [/\bcourt\b/g, 'ct'],
    [/\bct\b/g, 'ct'],

    [/\bcircle\b/g, 'cir'],
    [/\bcir\b/g, 'cir'],

    [/\bplace\b/g, 'pl'],
    [/\bpl\b/g, 'pl'],

    [/\bparkway\b/g, 'pkwy'],
    [/\bpkwy\b/g, 'pkwy'],

    [/\bhighway\b/g, 'hwy'],
    [/\bhwy\b/g, 'hwy'],

    [/\btrail\b/g, 'trl'],
    [/\btrl\b/g, 'trl'],

    [/\btrace\b/g, 'trce'],
    [/\btrce\b/g, 'trce'],

    [/\bterrace\b/g, 'ter'],
    [/\bter\b/g, 'ter'],

    [/\bboulevard\b/g, 'blvd'],
    [/\bblvd\b/g, 'blvd'],

    [/\bridge\b/g, 'rdg'],
    [/\brdg\b/g, 'rdg'],

    [/\bhill\b/g, 'hl'],
    [/\bhl\b/g, 'hl'],

    [/\bmountain\b/g, 'mtn'],
    [/\bmtn\b/g, 'mtn'],

    [/\bpoint\b/g, 'pt'],
    [/\bpt\b/g, 'pt'],

    [/\bvalley\b/g, 'vly'],
    [/\bvly\b/g, 'vly'],

    [/\bheights\b/g, 'hts'],
    [/\bhts\b/g, 'hts']
  ];

  for (let i = 0; i < replacements.length; i++) {
    v = v.replace(replacements[i][0], replacements[i][1]);
  }

  return v.replace(/\s+/g, ' ').trim();
}

function STP3_key_(addr, city, state) {
  const a = STP3_norm_(addr);
  const c = STP3_norm_(city);
  const s = STP3_norm_(state);
  if (!a && !c && !s) return '';
  return `${a}|${c}|${s}`;
}

function STP3_findIdx_(H, syns) {
  const want = (syns || []).map(s => String(s || '').toLowerCase().trim());
  for (let i = 0; i < H.length; i++) {
    const h = String(H[i] || '').toLowerCase().trim();
    if (!h) continue;
    if (want.indexOf(h) !== -1) return i;
  }
  return -1;
}

function STP3_has7digits_(v) {
  return String(v || '').replace(/\D+/g, '').length >= 7;
}

function STP3_digits5_(v) {
  const s = String(v || '').replace(/\D+/g, '').trim();
  if (!s) return '';
  if (s.length === 4) return '0' + s;
  if (s.length >= 5) return s.substring(0, 5);
  return s;
}

// ------------------------- Main -------------------------
function ST_part3_notInMojoVsDirect_noMatchOrNoPhone_pullForImport() {
  const ss = SpreadsheetApp.getActive();

  const src   = ss.getSheetByName('Not In Mojo');
  const dskip = ss.getSheetByName('Direct Skip Results');
  const fin   = ss.getSheetByName('For Import');

  if (!src || src.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Not In Mojo"');
    return;
  }
  if (!dskip || dskip.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Direct Skip Results"');
    return;
  }
  if (!fin || fin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "For Import"');
    return;
  }

  const out = STP3_upsertSheet_(ss, 'Not In Direct');

  const HS = STP3_headers_(src);
  const HD = STP3_headers_(dskip);
  const HF = STP3_headers_(fin);

  // Source: Not In Mojo
  const sAddr = (function () {
    let i = STP3_findIdx_(HS, ['property address', 'propertyaddress', 'street address', 'property street address', 'address']);
    return (i !== -1) ? i : 0;
  })();
  const sCity = (function () {
    let i = STP3_findIdx_(HS, ['property city', 'propertycity', 'city']);
    return (i !== -1) ? i : 1;
  })();
  const sState = (function () {
    let i = STP3_findIdx_(HS, ['property state', 'propertystate', 'state']);
    return (i !== -1) ? i : 2;
  })();

  // Direct Skip Results
  const dAddr = STP3_findIdx_(HD, ['input mailing address', 'mailing address', 'input address', 'address', 'property address', 'propertyaddress', 'street address']);
  const dCity = STP3_findIdx_(HD, ['input mailing city', 'mailing city', 'city']);
  const dState = STP3_findIdx_(HD, ['input mailing state', 'mailing state', 'state']);

  if (dAddr === -1 || dCity === -1 || dState === -1) {
    SpreadsheetApp.getUi().alert(
      'Direct Skip Results must have Input Mailing Address + City + State (or equivalent headers).'
    );
    return;
  }

  // Direct phone columns
  const phoneIdxs = [];
  for (let i = 0; i < HD.length; i++) {
    const h = String(HD[i] || '').toLowerCase().replace(/\s+/g, '').trim();
    if (/^phone[1-7]$/.test(h)) phoneIdxs.push(i);
  }

  function directHasAnyPhone_(row) {
    for (const idx of phoneIdxs) {
      if (STP3_has7digits_(row[idx])) return true;
    }
    return false;
  }

  // For Import key columns
  const fAddr = STP3_findIdx_(HF, ['property address', 'propertyaddress', 'street address', 'property street address', 'address']);
  const fCity = STP3_findIdx_(HF, ['property city', 'propertycity', 'city']);
  const fState = STP3_findIdx_(HF, ['property state', 'propertystate', 'state']);
  const fZip = STP3_findIdx_(HF, ['property zip', 'property zip code', 'zip', 'zip code', 'zipcode', 'postal', 'postal code']);

  if (fAddr === -1 || fCity === -1 || fState === -1) {
    SpreadsheetApp.getUi().alert(
      'For Import must have Property Address + Property City + Property State headers.'
    );
    return;
  }

  const sBody = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues();
  const dBody = dskip.getRange(2, 1, dskip.getLastRow() - 1, dskip.getLastColumn()).getValues();
  const fBody = fin.getRange(2, 1, fin.getLastRow() - 1, fin.getLastColumn()).getValues();

  // Build Direct Skip lookup
  const directHasPhoneByKey = new Map();
  for (const r of dBody) {
    const key = STP3_key_(r[dAddr], r[dCity], r[dState]);
    if (!key) continue;

    const hasPhone = directHasAnyPhone_(r);
    const prev = directHasPhoneByKey.get(key);

    if (prev === true) continue;
    if (hasPhone === true) directHasPhoneByKey.set(key, true);
    else if (prev === undefined) directHasPhoneByKey.set(key, false);
  }

  // Decide which keys to include
  const includeKeys = new Set();
  let srcBlankAddr = 0;
  let included_noMatch = 0;
  let included_matchNoPhone = 0;
  let excluded_hasPhone = 0;

  for (const r of sBody) {
    const A = r[sAddr];
    const B = r[sCity];
    const C = r[sState];

    if (!String(A || '').trim()) {
      srcBlankAddr++;
      continue;
    }

    const key = STP3_key_(A, B, C);
    if (!key) continue;

    const hasPhone = directHasPhoneByKey.get(key);

    if (hasPhone === true) {
      excluded_hasPhone++;
      continue;
    } else if (hasPhone === false) {
      includeKeys.add(key);
      included_matchNoPhone++;
    } else {
      includeKeys.add(key);
      included_noMatch++;
    }
  }

  // Build For Import lookup by Property Address + City + State
  const finByKey = new Map();
  for (const r of fBody) {
    const key = STP3_key_(r[fAddr], r[fCity], r[fState]);
    if (!key) continue;

    if (!finByKey.has(key)) finByKey.set(key, []);
    finByKey.get(key).push(r);
  }

  // Pull FULL For Import rows unchanged
  const outRows = [];
  const seenRowKeys = new Set();
  let missingInForImport = 0;

  includeKeys.forEach(key => {
    const rows = finByKey.get(key);
    if (!rows || !rows.length) {
      missingInForImport++;
      return;
    }

    for (const rr of rows) {
      const rowKey = JSON.stringify(rr);
      if (seenRowKeys.has(rowKey)) continue;
      seenRowKeys.add(rowKey);
      outRows.push(rr);
    }
  });

  // Write output exactly as For Import
  STP3_clearAndHeaders_(out, HF);
  if (outRows.length) {
    out.getRange(2, 1, outRows.length, HF.length).setValues(outRows);
  }

  // ZIP safety only for existing For Import zip column
  if (fZip !== -1 && outRows.length) {
    out.getRange(2, fZip + 1, outRows.length, 1).setNumberFormat('@');
    const zVals = out.getRange(2, fZip + 1, outRows.length, 1).getValues();
    for (let i = 0; i < zVals.length; i++) {
      zVals[i][0] = STP3_digits5_(zVals[i][0]) || '';
    }
    out.getRange(2, fZip + 1, outRows.length, 1).setValues(zVals);
  }

  SpreadsheetApp.getUi().alert(
    'Not In Direct rebuilt.\n' +
    `Source rows scanned: ${sBody.length}\n` +
    `Source blank address skipped: ${srcBlankAddr}\n` +
    `Included keys (no match): ${included_noMatch}\n` +
    `Included keys (match but no phone): ${included_matchNoPhone}\n` +
    `Excluded because Direct Skip match has phone: ${excluded_hasPhone}\n` +
    `Total included keys (unique): ${includeKeys.size}\n` +
    `Rows written (from For Import): ${outRows.length}\n` +
    `Included keys not found in For Import: ${missingInForImport}`
  );
}
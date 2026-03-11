/** PART 1 — Compare "Ready 4 Xleads" vs "Xleads Results"
 *
 * INCLUDE IN "Not In Xleads" when:
 *  (A) Ready property has NO match in Xleads, OR
 *  (B) Ready property matches Xleads BUT Xleads has NO phone (7+ digits in any phone column)
 *
 * READY vs XLEADS MATCH:
 *  - Ready 4 Xleads:
 *      Column A = Street Address
 *      Column D = Zip
 *  - Xleads Results:
 *      Column G = PropertyAddress
 *      Column J = PropertyPostalCode
 *
 * PHONE CHECK IN XLEADS:
 *  - W  = Contact1Phone_1
 *  - AB = Contact1Phone_2
 *  - AG = Contact1Phone_3
 *
 * OUTPUT ROW SOURCE:
 *  - Pull rows from "For Import"
 *  - Match:
 *      Ready Street Address -> For Import Representative Address
 *      Ready Zip            -> For Import Representative Zip
 *
 * OUTPUT TAB:
 *  - Not In Xleads
 */

// ------------------------- Helpers -------------------------
function STP1c_headers_(sh) {
  var lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
}

function STP1c_upsertSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function STP1c_clearAndHeaders_(sh, H) {
  sh.clearContents();
  if (H && H.length) sh.getRange(1, 1, 1, H.length).setValues([H]);
  sh.setFrozenRows(1);
}

function STP1c_colToIndex0_(letter) {
  var s = String(letter || '').toUpperCase().trim();
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return Math.max(0, n - 1);
}

function STP1c_norm_(s) {
  var v = String(s || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var replacements = [
    // directions
    [/\bnorth\b/g, 'n'],
    [/\bsouth\b/g, 's'],
    [/\beast\b/g, 'e'],
    [/\bwest\b/g, 'w'],
    [/\bnortheast\b/g, 'ne'],
    [/\bnorthwest\b/g, 'nw'],
    [/\bsoutheast\b/g, 'se'],
    [/\bsouthwest\b/g, 'sw'],

    // ordinal words
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

    // street types
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

  for (var i = 0; i < replacements.length; i++) {
    v = v.replace(replacements[i][0], replacements[i][1]);
  }

  return v.replace(/\s+/g, ' ').trim();
}

function STP1c_digits5_(v) {
  var s = String(v || '').replace(/\D+/g, '').trim();
  if (!s) return '';
  if (s.length === 4) return '0' + s;
  if (s.length >= 5) return s.substring(0, 5);
  return s;
}

function STP1c_keyAddrZip_(addr, zip) {
  var a = STP1c_norm_(addr);
  var z = STP1c_digits5_(zip);
  if (!a && !z) return '';
  return a + '|' + z;
}

function STP1c_hasAnyPhone_(row, idxs) {
  for (var i = 0; i < idxs.length; i++) {
    var idx = idxs[i];
    if (idx === -1) continue;
    var digits = String(row[idx] || '').replace(/\D+/g, '');
    if (digits.length >= 7) return true;
  }
  return false;
}

function STP1c_findIdx_(H, syns) {
  var want = (syns || []).map(function(s) {
    return String(s || '').toLowerCase().trim();
  });

  for (var i = 0; i < H.length; i++) {
    var h = String(H[i] || '').toLowerCase().trim();
    if (!h) continue;
    if (want.indexOf(h) !== -1) return i;
  }
  return -1;
}

// ------------------------- Main -------------------------
function ST_part1_readyVsXleads_noMatchOrNoPhone_pullForImport() {
  var ss = SpreadsheetApp.getActive();

  var ready = ss.getSheetByName('Ready 4 Xleads');
  var xres  = ss.getSheetByName('Xleads Results');
  var fin   = ss.getSheetByName('For Import');

  if (!ready || ready.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Ready 4 Xleads"');
    return;
  }
  if (!xres || xres.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Xleads Results"');
    return;
  }
  if (!fin || fin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "For Import"');
    return;
  }

  var out = STP1c_upsertSheet_(ss, 'Not In Xleads');

  var HF = STP1c_headers_(fin);

  // ---------- Ready 4 Xleads fixed columns ----------
  var rAddr = STP1c_colToIndex0_('A'); // Street Address
  var rZip  = STP1c_colToIndex0_('D'); // Zip

  // ---------- Xleads Results fixed columns ----------
  var xAddr = STP1c_colToIndex0_('G'); // PropertyAddress
  var xZip  = STP1c_colToIndex0_('J'); // PropertyPostalCode

  // ---------- Xleads phone columns ----------
  var p1 = STP1c_colToIndex0_('W');  // Contact1Phone_1
  var p2 = STP1c_colToIndex0_('AB'); // Contact1Phone_2
  var p3 = STP1c_colToIndex0_('AG'); // Contact1Phone_3
  var phoneIdxs = [p1, p2, p3];

  // ---------- For Import columns ----------
  var fRepAddr = STP1c_findIdx_(HF, ['representative address', 'rep address', 'mailing street', 'mailing address']);
  var fRepZip  = STP1c_findIdx_(HF, ['representative zip', 'representative zipcode', 'rep zip', 'rep zipcode', 'mailing zip', 'mailing zipcode']);
  var fZipOut  = STP1c_findIdx_(HF, ['property zip', 'property zip code', 'zip', 'zip code', 'zipcode', 'postal', 'postal code', 'property postal', 'propertypostal']);

  if (fRepAddr === -1 || fRepZip === -1) {
    SpreadsheetApp.getUi().alert(
      'For Import must have Representative Address and Representative Zip headers.'
    );
    return;
  }

  // ---------- Load data ----------
  var readyBody = ready.getRange(2, 1, ready.getLastRow() - 1, ready.getLastColumn()).getValues();
  var xBody     = xres.getRange(2, 1, xres.getLastRow() - 1, xres.getLastColumn()).getValues();
  var fBody     = fin.getRange(2, 1, fin.getLastRow() - 1, fin.getLastColumn()).getValues();

  // ---------- Build Xleads map: Address + Zip -> hasPhone ----------
  var xHasPhoneByAddrZip = new Map();

  for (var i = 0; i < xBody.length; i++) {
    var xr = xBody[i];
    var key = STP1c_keyAddrZip_(xr[xAddr], xr[xZip]);
    if (!key) continue;

    var hasPhone = STP1c_hasAnyPhone_(xr, phoneIdxs);
    var prev = xHasPhoneByAddrZip.get(key);

    if (prev === true) continue;
    if (hasPhone === true) xHasPhoneByAddrZip.set(key, true);
    else if (prev === undefined) xHasPhoneByAddrZip.set(key, false);
  }

  // ---------- Decide which Ready rows to include ----------
  var includeReadyRows = [];
  var includePullKeys = new Set();

  var readyBlankAddr = 0;
  var included_noMatch = 0;
  var included_matchNoPhone = 0;
  var excluded_hasPhone = 0;

  for (var j = 0; j < readyBody.length; j++) {
    var rr = readyBody[j];
    var addr = rr[rAddr];
    var zip  = rr[rZip];

    if (!String(addr || '').trim()) {
      readyBlankAddr++;
      continue;
    }

    var key = STP1c_keyAddrZip_(addr, zip);
    if (!key) continue;

    var hasPhone = xHasPhoneByAddrZip.get(key);

    if (hasPhone === true) {
      excluded_hasPhone++;
      continue;
    } else if (hasPhone === false) {
      includeReadyRows.push(rr);
      includePullKeys.add(key);
      included_matchNoPhone++;
    } else {
      includeReadyRows.push(rr);
      includePullKeys.add(key);
      included_noMatch++;
    }
  }

  // ---------- Build For Import lookup: Representative Address + Representative Zip ----------
  var finByRepAddrZip = new Map();

  for (var k = 0; k < fBody.length; k++) {
    var fr = fBody[k];
    var fKey = STP1c_keyAddrZip_(fr[fRepAddr], fr[fRepZip]);
    if (!fKey) continue;

    if (!finByRepAddrZip.has(fKey)) finByRepAddrZip.set(fKey, []);
    finByRepAddrZip.get(fKey).push(fr);
  }

  // ---------- Pull For Import rows ----------
  var outRows = [];
  var seenRowKeys = new Set();
  var missingInForImport = 0;

  for (var m = 0; m < includeReadyRows.length; m++) {
    var rrow = includeReadyRows[m];
    var pullKey = STP1c_keyAddrZip_(rrow[rAddr], rrow[rZip]);
    var matches = finByRepAddrZip.get(pullKey);

    if (!matches || !matches.length) {
      missingInForImport++;
      continue;
    }

    for (var n = 0; n < matches.length; n++) {
      var rowKey = JSON.stringify(matches[n]);
      if (seenRowKeys.has(rowKey)) continue;
      seenRowKeys.add(rowKey);
      outRows.push(matches[n]);
    }
  }

  // ---------- Write output ----------
  STP1c_clearAndHeaders_(out, HF);

  if (outRows.length) {
    out.getRange(2, 1, outRows.length, HF.length).setValues(outRows);
  }

  // ---------- ZIP safety ----------
  if (fZipOut !== -1) {
    out.getRange(2, fZipOut + 1, Math.max(1, outRows.length), 1).setNumberFormat('@');
    if (outRows.length) {
      var zVals = out.getRange(2, fZipOut + 1, outRows.length, 1).getValues();
      for (var z = 0; z < zVals.length; z++) {
        zVals[z][0] = STP1c_digits5_(zVals[z][0]) || '';
      }
      out.getRange(2, fZipOut + 1, outRows.length, 1).setValues(zVals);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Not In Xleads rebuilt.\n' +
    'Ready rows scanned: ' + readyBody.length + '\n' +
    'Ready blank address skipped: ' + readyBlankAddr + '\n' +
    'Included keys (no match): ' + included_noMatch + '\n' +
    'Included keys (match but no phone): ' + included_matchNoPhone + '\n' +
    'Excluded because Xleads match has phone: ' + excluded_hasPhone + '\n' +
    'Total included keys (unique): ' + includePullKeys.size + '\n' +
    'Rows written (from For Import): ' + outRows.length + '\n' +
    'Included Ready records not found in For Import: ' + missingInForImport
  );
}
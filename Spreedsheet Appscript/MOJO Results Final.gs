/** PART 2 — Mojo (same concept as updated Xleads)
 *
 * Compare:
 *   "Not In Xleads" vs "Mojo Results"
 *
 * Include in "Not In Mojo" when:
 *   (A) Source property has NO match in Mojo, OR
 *   (B) Property matches Mojo BUT Mojo has NO phone (7+ digits)
 *
 * OUTPUT:
 *   - "Not In Mojo" uses ALL headers from "For Import"
 *   - Rows written are FULL rows pulled from "For Import"
 *
 * MATCHING RULES:
 *   Source ("Not In Xleads"):
 *     - Representative Address
 *     - Representative Zip
 *
 *   Mojo Results:
 *     - Column A = Street Address
 *     - Column D = Zip
 *
 * PHONE RULES (Mojo):
 *   - Mobile 1..10 = columns T:AC
 *   - Phone 1..10  = columns AD:AN
 *
 * FOR IMPORT PULL:
 *   - Representative Address
 *   - Representative Zip
 */

// ------------------------- Helpers -------------------------
function STP2_headers_(sh) {
  var lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
}

function STP2_upsertSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function STP2_clearAndHeaders_(sh, H) {
  sh.clearContents();
  if (H && H.length) sh.getRange(1, 1, 1, H.length).setValues([H]);
  sh.setFrozenRows(1);
}

function STP2_colToIndex0_(letter) {
  var s = String(letter || '').toUpperCase().trim();
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return Math.max(0, n - 1);
}

function STP2_norm_(s) {
  var v = String(s || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var replacements = [
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

  for (var i = 0; i < replacements.length; i++) {
    v = v.replace(replacements[i][0], replacements[i][1]);
  }

  return v.replace(/\s+/g, ' ').trim();
}

function STP2_digits5_(v) {
  var s = String(v || '').replace(/\D+/g, '').trim();
  if (!s) return '';
  if (s.length === 4) return '0' + s;
  if (s.length >= 5) return s.substring(0, 5);
  return s;
}

function STP2_keyAddrZip_(addr, zip) {
  var a = STP2_norm_(addr);
  var z = STP2_digits5_(zip);
  if (!a && !z) return '';
  return a + '|' + z;
}

function STP2_findIdx_(H, syns) {
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

function STP2_has7digits_(v) {
  return String(v || '').replace(/\D+/g, '').length >= 7;
}

// ------------------------- Main -------------------------
function ST_part2_notInXleadsVsMojo_noMatchOrNoPhone_pullForImport() {
  var ss = SpreadsheetApp.getActive();

  var src  = ss.getSheetByName('Not In Xleads');
  var mojo = ss.getSheetByName('Mojo Results');
  var fin  = ss.getSheetByName('For Import');

  if (!src || src.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Not In Xleads"');
    return;
  }
  if (!mojo || mojo.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Mojo Results"');
    return;
  }
  if (!fin || fin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "For Import"');
    return;
  }

  var out = STP2_upsertSheet_(ss, 'Not In Mojo');

  var HS = STP2_headers_(src);
  var HF = STP2_headers_(fin);

  var sRepAddr = STP2_findIdx_(HS, ['representative address', 'rep address', 'mailing street', 'mailing address']);
  var sRepZip  = STP2_findIdx_(HS, ['representative zip', 'representative zipcode', 'rep zip', 'rep zipcode', 'mailing zip', 'mailing zipcode']);

  if (sRepAddr === -1 || sRepZip === -1) {
    SpreadsheetApp.getUi().alert(
      'Not In Xleads must have Representative Address and Representative Zip headers.'
    );
    return;
  }

  var mAddr = STP2_colToIndex0_('A');
  var mZip  = STP2_colToIndex0_('D');

  var mojoPhoneIdxs = [
    STP2_colToIndex0_('T'),
    STP2_colToIndex0_('U'),
    STP2_colToIndex0_('V'),
    STP2_colToIndex0_('W'),
    STP2_colToIndex0_('X'),
    STP2_colToIndex0_('Y'),
    STP2_colToIndex0_('Z'),
    STP2_colToIndex0_('AA'),
    STP2_colToIndex0_('AB'),
    STP2_colToIndex0_('AC'),
    STP2_colToIndex0_('AD'),
    STP2_colToIndex0_('AE'),
    STP2_colToIndex0_('AF'),
    STP2_colToIndex0_('AG'),
    STP2_colToIndex0_('AH'),
    STP2_colToIndex0_('AI'),
    STP2_colToIndex0_('AJ'),
    STP2_colToIndex0_('AK'),
    STP2_colToIndex0_('AL'),
    STP2_colToIndex0_('AM'),
    STP2_colToIndex0_('AN')
  ];

  function mojoHasAnyPhone_(row) {
    for (var i = 0; i < mojoPhoneIdxs.length; i++) {
      if (STP2_has7digits_(row[mojoPhoneIdxs[i]])) return true;
    }
    return false;
  }

  var fRepAddr = STP2_findIdx_(HF, ['representative address', 'rep address', 'mailing street', 'mailing address']);
  var fRepZip  = STP2_findIdx_(HF, ['representative zip', 'representative zipcode', 'rep zip', 'rep zipcode', 'mailing zip', 'mailing zipcode']);
  var fZipOut  = STP2_findIdx_(HF, ['property zip', 'property zip code', 'zip', 'zip code', 'zipcode', 'postal', 'postal code', 'property postal', 'propertypostal']);

  if (fRepAddr === -1 || fRepZip === -1) {
    SpreadsheetApp.getUi().alert(
      'For Import must have Representative Address and Representative Zip headers.'
    );
    return;
  }

  var sBody = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues();
  var mBody = mojo.getRange(2, 1, mojo.getLastRow() - 1, mojo.getLastColumn()).getValues();
  var fBody = fin.getRange(2, 1, fin.getLastRow() - 1, fin.getLastColumn()).getValues();

  var mojoHasPhoneByKey = new Map();

  for (var r = 0; r < mBody.length; r++) {
    var row = mBody[r];
    var key = STP2_keyAddrZip_(row[mAddr], row[mZip]);
    if (!key) continue;

    var hasPhone = mojoHasAnyPhone_(row);
    var prev = mojoHasPhoneByKey.get(key);

    if (prev === true) continue;
    if (hasPhone === true) mojoHasPhoneByKey.set(key, true);
    else if (prev === undefined) mojoHasPhoneByKey.set(key, false);
  }

  var includeKeys = new Set();
  var srcBlankAddr = 0;
  var included_noMatch = 0;
  var included_matchNoPhone = 0;
  var excluded_hasPhone = 0;

  for (var j = 0; j < sBody.length; j++) {
    var sr = sBody[j];
    var addr = sr[sRepAddr];
    var zip  = sr[sRepZip];

    if (!String(addr || '').trim()) {
      srcBlankAddr++;
      continue;
    }

    var key = STP2_keyAddrZip_(addr, zip);
    if (!key) continue;

    var hasPhone = mojoHasPhoneByKey.get(key);

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

  var finByKey = new Map();

  for (var k = 0; k < fBody.length; k++) {
    var fr = fBody[k];
    var fKey = STP2_keyAddrZip_(fr[fRepAddr], fr[fRepZip]);
    if (!fKey) continue;

    if (!finByKey.has(fKey)) finByKey.set(fKey, []);
    finByKey.get(fKey).push(fr);
  }

  var outRows = [];
  var seenRowKeys = new Set();
  var missingInForImport = 0;

  includeKeys.forEach(function(key) {
    var rows = finByKey.get(key);
    if (!rows || !rows.length) {
      missingInForImport++;
      return;
    }

    for (var i = 0; i < rows.length; i++) {
      var rowKey = JSON.stringify(rows[i]);
      if (seenRowKeys.has(rowKey)) continue;
      seenRowKeys.add(rowKey);
      outRows.push(rows[i]);
    }
  });

  STP2_clearAndHeaders_(out, HF);
  if (outRows.length) {
    out.getRange(2, 1, outRows.length, HF.length).setValues(outRows);
  }

  if (fZipOut !== -1) {
    out.getRange(2, fZipOut + 1, Math.max(1, outRows.length), 1).setNumberFormat('@');
    if (outRows.length) {
      var zVals = out.getRange(2, fZipOut + 1, outRows.length, 1).getValues();
      for (var z = 0; z < zVals.length; z++) {
        zVals[z][0] = STP2_digits5_(zVals[z][0]) || '';
      }
      out.getRange(2, fZipOut + 1, outRows.length, 1).setValues(zVals);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Not In Mojo rebuilt.\n' +
    'Source rows scanned (Not In Xleads): ' + sBody.length + '\n' +
    'Source blank address skipped: ' + srcBlankAddr + '\n' +
    'Included keys (no match): ' + included_noMatch + '\n' +
    'Included keys (match but no phone): ' + included_matchNoPhone + '\n' +
    'Excluded because Mojo match has phone: ' + excluded_hasPhone + '\n' +
    'Total included keys (unique): ' + includeKeys.size + '\n' +
    'Rows written (from For Import): ' + outRows.length + '\n' +
    'Included keys not found in For Import: ' + missingInForImport
  );
}
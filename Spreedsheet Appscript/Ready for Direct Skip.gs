/** PART 3 — Direct Skip (same concept as updated Xleads / Mojo)
 *
 * Compare:
 *   "Not In Mojo" vs "Direct Skip Results"
 *
 * Include in "Not In Direct" when:
 *   (A) Source property has NO match in Direct Skip, OR
 *   (B) Property matches Direct Skip BUT Direct Skip has NO phone (7+ digits across Phone1..Phone7)
 *
 * OUTPUT:
 *   - "Not In Direct" uses ALL headers from "For Import"
 *   - Rows written are FULL rows pulled from "For Import"
 *
 * MATCHING RULES:
 *   Source ("Not In Mojo"):
 *     - Representative Address
 *     - Representative Zip
 *
 *   Direct Skip Results:
 *     - Input Mailing Address
 *     - Input Mailing Zip
 *
 * PHONE RULES (Direct Skip):
 *   - checks Phone1..Phone7 by header pattern
 *
 * FOR IMPORT PULL:
 *   - Representative Address
 *   - Representative Zip
 *
 * Run:
 *   - ST_part3_notInMojoVsDirect_noMatchOrNoPhone_pullForImport()
 */

// ------------------------- Helpers -------------------------
function STP3_headers_(sh) {
  var lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
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
  return String(s || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function STP3_digits5_(v) {
  var m = String(v || '').match(/(\d{5})/);
  return m ? m[1] : '';
}

function STP3_keyAddrZip_(addr, zip) {
  var a = STP3_norm_(addr);
  var z = STP3_digits5_(zip);
  if (!a && !z) return '';
  return a + '|' + z;
}

function STP3_findIdx_(H, syns) {
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

function STP3_has7digits_(v) {
  return String(v || '').replace(/\D+/g, '').length >= 7;
}

// ------------------------- Main -------------------------
function ST_part3_notInMojoVsDirect_noMatchOrNoPhone_pullForImport() {
  var ss = SpreadsheetApp.getActive();

  var src   = ss.getSheetByName('Not In Mojo');
  var dskip = ss.getSheetByName('Direct Skip Results');
  var fin   = ss.getSheetByName('For Import');

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

  var out = STP3_upsertSheet_(ss, 'Not In Direct');

  var HS = STP3_headers_(src);
  var HD = STP3_headers_(dskip);
  var HF = STP3_headers_(fin);

  // ---------- Source columns: Not In Mojo ----------
  // Since Not In Mojo is built from For Import rows, use Representative Address + Representative Zip
  var sRepAddr = STP3_findIdx_(HS, ['representative address', 'rep address', 'mailing street', 'mailing address']);
  var sRepZip  = STP3_findIdx_(HS, ['representative zip', 'representative zipcode', 'rep zip', 'rep zipcode', 'mailing zip', 'mailing zipcode']);

  if (sRepAddr === -1 || sRepZip === -1) {
    SpreadsheetApp.getUi().alert(
      'Not In Mojo must have Representative Address and Representative Zip headers.'
    );
    return;
  }

  // ---------- Direct Skip columns ----------
  var dAddr = STP3_findIdx_(HD, [
    'input mailing address',
    'mailing address',
    'input address',
    'address',
    'street address',
    'property address',
    'propertyaddress'
  ]);

  var dZip = STP3_findIdx_(HD, [
    'input mailing zip',
    'input mailing zipcode',
    'mailing zip',
    'mailing zipcode',
    'zip',
    'zip code',
    'zipcode',
    'postal',
    'postal code'
  ]);

  if (dAddr === -1 || dZip === -1) {
    SpreadsheetApp.getUi().alert(
      'Direct Skip Results must have Input Mailing Address and Input Mailing Zip (or equivalent headers).'
    );
    return;
  }

  // ---------- Direct Skip phones: Phone1..Phone7 ----------
  var phoneIdxs = [];
  for (var i = 0; i < HD.length; i++) {
    var h = String(HD[i] || '').toLowerCase().replace(/\s+/g, '').trim();
    if (/^phone[1-7]$/.test(h)) phoneIdxs.push(i);
  }

  function directHasAnyPhone_(row) {
    for (var i = 0; i < phoneIdxs.length; i++) {
      if (STP3_has7digits_(row[phoneIdxs[i]])) return true;
    }
    return false;
  }

  // ---------- For Import columns ----------
  var fRepAddr = STP3_findIdx_(HF, ['representative address', 'rep address', 'mailing street', 'mailing address']);
  var fRepZip  = STP3_findIdx_(HF, ['representative zip', 'representative zipcode', 'rep zip', 'rep zipcode', 'mailing zip', 'mailing zipcode']);

  // Optional ZIP output formatting target
  var fZipOut = STP3_findIdx_(HF, [
    'property zip',
    'property zip code',
    'zip',
    'zip code',
    'zipcode',
    'postal',
    'postal code',
    'property postal',
    'propertypostal'
  ]);

  if (fRepAddr === -1 || fRepZip === -1) {
    SpreadsheetApp.getUi().alert(
      'For Import must have Representative Address and Representative Zip headers.'
    );
    return;
  }

  // ---------- Load bodies ----------
  var sBody = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues();
  var dBody = dskip.getRange(2, 1, dskip.getLastRow() - 1, dskip.getLastColumn()).getValues();
  var fBody = fin.getRange(2, 1, fin.getLastRow() - 1, fin.getLastColumn()).getValues();

  // ---------- Build Direct map: Address + Zip -> hasPhone ----------
  var directHasPhoneByKey = new Map();

  for (var r = 0; r < dBody.length; r++) {
    var row = dBody[r];
    var key = STP3_keyAddrZip_(row[dAddr], row[dZip]);
    if (!key) continue;

    var hasPhone = directHasAnyPhone_(row);
    var prev = directHasPhoneByKey.get(key);

    if (prev === true) continue;
    if (hasPhone === true) directHasPhoneByKey.set(key, true);
    else if (prev === undefined) directHasPhoneByKey.set(key, false);
  }

  // ---------- Decide which source keys to include ----------
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

    var key = STP3_keyAddrZip_(addr, zip);
    if (!key) continue;

    var hasPhone = directHasPhoneByKey.get(key);

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

  // ---------- Build For Import lookup ----------
  var finByKey = new Map();

  for (var k = 0; k < fBody.length; k++) {
    var fr = fBody[k];
    var fKey = STP3_keyAddrZip_(fr[fRepAddr], fr[fRepZip]);
    if (!fKey) continue;

    if (!finByKey.has(fKey)) finByKey.set(fKey, []);
    finByKey.get(fKey).push(fr);
  }

  // ---------- Pull rows ----------
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

  // ---------- Write ----------
  STP3_clearAndHeaders_(out, HF);
  if (outRows.length) {
    out.getRange(2, 1, outRows.length, HF.length).setValues(outRows);
  }

  // ---------- ZIP safety ----------
  if (fZipOut !== -1) {
    out.getRange(2, fZipOut + 1, Math.max(1, outRows.length), 1).setNumberFormat('@');
    if (outRows.length) {
      var zVals = out.getRange(2, fZipOut + 1, outRows.length, 1).getValues();
      for (var z = 0; z < zVals.length; z++) {
        zVals[z][0] = STP3_digits5_(zVals[z][0]) || '';
      }
      out.getRange(2, fZipOut + 1, outRows.length, 1).setValues(zVals);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Not In Direct rebuilt.\n' +
    'Source rows scanned: ' + sBody.length + '\n' +
    'Source blank address skipped: ' + srcBlankAddr + '\n' +
    'Included keys (no match): ' + included_noMatch + '\n' +
    'Included keys (match but no phone): ' + included_matchNoPhone + '\n' +
    'Excluded because Direct Skip match has phone: ' + excluded_hasPhone + '\n' +
    'Total included keys (unique): ' + includeKeys.size + '\n' +
    'Rows written: ' + outRows.length + '\n' +
    'Missing in For Import: ' + missingInForImport
  );
}

function myFunction() {}
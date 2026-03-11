// ==============================
// READY GHL DUPLICATION SCRIPT
// Mobile, Landline, Email
// Multi-number cells supported
// ==============================

// ------------------------------
// MOBILE DUPLICATION (R + S -> R)
// also sets Type Of Phone Number = Mobile
// ------------------------------
function ST_buildMobileReady4GHL() {
  explodeNumbersToReadyTab_(
    'Skip Trace Finished',
    'Mobile Ready 4 GHL',
    ['R', 'S'],
    'R',
    'Mobile'
  );
}

// ------------------------------
// LANDLINE DUPLICATION (T + U -> T)
// also sets Type Of Phone Number = Not Mobile
// ------------------------------
function ST_buildLandlineReady4GHL() {
  explodeNumbersToReadyTab_(
    'Skip Trace Finished',
    'Landline Ready 4 GHL',
    ['T', 'U'],
    'T',
    'Not Mobile'
  );
}

// ------------------------------
// EMAIL DUPLICATION
// ------------------------------
function ST_buildEmailReady4GHL() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Skip Trace Finished');

  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished empty.');
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const normH = h => String(h || '').trim().toLowerCase();

  const idxPrimary = headers.findIndex(h =>
    ['email', 'primary email'].includes(normH(h))
  );

  const idxAddl = headers.findIndex(h =>
    ['additional email', 'additional emails', 'other emails'].includes(normH(h))
  );

  if (idxPrimary === -1 && idxAddl === -1) {
    SpreadsheetApp.getUi().alert('No email columns found.');
    return;
  }

  const seen = new Set();
  const outRows = [];

  for (const row of data) {
    let emails = [];

    if (idxPrimary !== -1 && row[idxPrimary]) {
      emails.push(String(row[idxPrimary]).trim());
    }

    if (idxAddl !== -1 && row[idxAddl]) {
      emails.push(
        ...String(row[idxAddl])
          .split(/[,;\s]+/)
          .map(e => e.trim())
          .filter(Boolean)
      );
    }

    for (const e of emails) {
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) continue;

      const lower = e.toLowerCase();
      if (seen.has(lower)) continue;
      seen.add(lower);

      const newRow = row.slice();
      if (idxPrimary !== -1) newRow[idxPrimary] = e;
      if (idxAddl !== -1) newRow[idxAddl] = '';
      outRows.push(newRow);
    }
  }

  const outSheetName = 'Email Ready 4 GHL';
  const out = ss.getSheetByName(outSheetName) || ss.insertSheet(outSheetName);

  out.clearContents();
  out.getRange(1, 1, 1, lastCol).setValues([headers]);

  if (outRows.length) {
    out.getRange(2, 1, outRows.length, lastCol).setValues(outRows);
  }

  out.setFrozenRows(1);
  out.getRange(1, 1, 1, lastCol).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(outSheetName + ' rebuilt. Rows written: ' + outRows.length);
}

// ------------------------------
// RUN ALL READY TABS
// ------------------------------
function ST_buildAllReadyTabs_GHL() {
  ST_buildMobileReady4GHL();
  ST_buildLandlineReady4GHL();
  ST_buildEmailReady4GHL();
  SpreadsheetApp.getUi().alert('Done: Mobile/Landline/Email Ready tabs rebuilt.');
}

// ------------------------------
// CORE ENGINE FOR MOBILE/LANDLINE
// ------------------------------
function explodeNumbersToReadyTab_(sourceSheetName, targetSheetName, sourceCols, targetColLetter, phoneTypeValue) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(sourceSheetName);

  if (!src || src.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(sourceSheetName + ' empty or not found.');
    return;
  }

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  const headers = src.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const data = src.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const normH = h => String(h || '').trim().toLowerCase();

  const idxTypeOfPhone = headers.findIndex(h =>
    ['type of phone number', 'type of phone', 'phone type'].includes(normH(h))
  );

  const sourceIndices = sourceCols.map(colLetter => {
    let n = 0;
    const s = String(colLetter).toUpperCase().trim();
    for (let i = 0; i < s.length; i++) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n - 1;
  });

  let targetCol = 0;
  const t = String(targetColLetter).toUpperCase().trim();
  for (let i = 0; i < t.length; i++) {
    targetCol = targetCol * 26 + (t.charCodeAt(i) - 64);
  }
  targetCol = targetCol - 1;

  const outRows = [];

  for (const row of data) {
    const numbers = [];

    for (const c of sourceIndices) {
      if (!row[c]) continue;

      const parts = String(row[c])
        .split(/[,;\s]+/)
        .map(p => p.trim())
        .filter(Boolean);

      for (const part of parts) {
        const digitsOnly = part.replace(/\D+/g, '');
        if (digitsOnly.length >= 7 && digitsOnly.length <= 11) {
          numbers.push(digitsOnly);
        }
      }
    }

    for (const n of numbers) {
      const newRow = row.slice();
      newRow[targetCol] = n;

      for (const c of sourceIndices) {
        if (c !== targetCol) newRow[c] = '';
      }

      if (idxTypeOfPhone !== -1) {
        newRow[idxTypeOfPhone] = phoneTypeValue;
      }

      outRows.push(newRow);
    }
  }

  const out = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);

  out.clearContents();
  out.getRange(1, 1, 1, lastCol).setValues([headers]);

  if (outRows.length) {
    out.getRange(2, 1, outRows.length, lastCol).setValues(outRows);
  }

  out.setFrozenRows(1);
  out.getRange(1, 1, 1, lastCol).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(targetSheetName + ' rebuilt. Rows written: ' + outRows.length);
}
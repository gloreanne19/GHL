function ST_resolveDuplicates_M_vs_F_and_RemoveAbove550kMatches() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Skip Trace Finished');
  const arv = ss.getSheetByName('Above 550K ARV');

  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Skip Trace Finished"');
    return;
  }

  if (!arv || arv.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Missing or empty: "Above 550K ARV"');
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();

  const headers = data[0];
  let body = data.slice(1);

  // Skip Trace Finished columns (0-based)
  // F = 5, M = 12, P = 15
  const STF_COL_F = 5;
  const STF_COL_M = 12;
  const STF_COL_P = 15;

  // Above 550K ARV columns (0-based)
  // F = 5, I = 8
  const ARV_COL_F = 5;
  const ARV_COL_I = 8;

  function norm_(v) {
    return String(v || '')
      .toLowerCase()
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function filledCount_(row) {
    let count = 0;
    for (let i = 0; i < row.length; i++) {
      if (String(row[i] || '').trim() !== '') count++;
    }
    return count;
  }

  function splitList_(v) {
    return String(v || '')
      .split(',')
      .map(s => String(s || '').trim())
      .filter(Boolean);
  }

  function uniqList_(arr) {
    const seen = new Set();
    const out = [];
    for (const x of arr) {
      const k = String(x || '').trim();
      if (!k || seen.has(k)) continue;
      seen.add(k);
      out.push(k);
    }
    return out;
  }

  function mergeCell_(keepVal, dropVal) {
    const a = String(keepVal || '').trim();
    const b = String(dropVal || '').trim();

    if (!a && b) return b;
    if (a && !b) return a;
    if (!a && !b) return '';

    if (a === b) return a;

    if (a.includes(',') || b.includes(',')) {
      return uniqList_([].concat(splitList_(a), splitList_(b))).join(', ');
    }

    return a;
  }

  function mergeRows_(keepRow, dropRow) {
    const merged = keepRow.slice();
    for (let c = 0; c < merged.length; c++) {
      merged[c] = mergeCell_(merged[c], dropRow[c]);
    }
    return merged;
  }

  function pairKey_(a, b) {
    const x = norm_(a);
    const y = norm_(b);
    if (!x && !y) return '';
    return x + '|' + y;
  }

  // ============================================================
  // PART 1 — Resolve duplicates in Skip Trace Finished
  // Match: Column M of one row = Column F of another row
  // Keep row with most filled cells, merge values, delete other row
  // ============================================================
  const fMap = new Map();
  for (let i = 0; i < body.length; i++) {
    const fVal = norm_(body[i][STF_COL_F]);
    if (!fVal) continue;
    if (!fMap.has(fVal)) fMap.set(fVal, []);
    fMap.get(fVal).push(i);
  }

  const rowsToDeleteAfterMerge = new Set();
  let merges = 0;

  for (let i = 0; i < body.length; i++) {
    if (rowsToDeleteAfterMerge.has(i)) continue;

    const mVal = norm_(body[i][STF_COL_M]);
    if (!mVal) continue;

    const matches = fMap.get(mVal);
    if (!matches || !matches.length) continue;

    for (const j of matches) {
      if (i === j) continue;
      if (rowsToDeleteAfterMerge.has(i) || rowsToDeleteAfterMerge.has(j)) continue;

      const rowA = body[i];
      const rowB = body[j];

      const scoreA = filledCount_(rowA);
      const scoreB = filledCount_(rowB);

      let keepIdx, dropIdx, keepRow, dropRow;

      if (scoreA >= scoreB) {
        keepIdx = i;
        dropIdx = j;
        keepRow = rowA;
        dropRow = rowB;
      } else {
        keepIdx = j;
        dropIdx = i;
        keepRow = rowB;
        dropRow = rowA;
      }

      body[keepIdx] = mergeRows_(keepRow, dropRow);
      rowsToDeleteAfterMerge.add(dropIdx);
      merges++;

      if (dropIdx === i) break;
    }
  }

  // Rebuild body after duplicate merge
  const mergedBody = [];
  for (let i = 0; i < body.length; i++) {
    if (!rowsToDeleteAfterMerge.has(i)) mergedBody.push(body[i]);
  }

  // ============================================================
  // PART 2 — Delete STF rows when:
  // Skip Trace Finished (M + P) matches Above 550K ARV (F + I)
  // ============================================================
  const arvData = arv.getRange(2, 1, arv.getLastRow() - 1, arv.getLastColumn()).getValues();
  const arvKeySet = new Set();

  for (let i = 0; i < arvData.length; i++) {
    const key = pairKey_(arvData[i][ARV_COL_F], arvData[i][ARV_COL_I]);
    if (key) arvKeySet.add(key);
  }

  const finalRows = [];
  let deletedByArvMatch = 0;

  for (let i = 0; i < mergedBody.length; i++) {
    const row = mergedBody[i];
    const stfKey = pairKey_(row[STF_COL_M], row[STF_COL_P]);

    if (stfKey && arvKeySet.has(stfKey)) {
      deletedByArvMatch++;
      continue;
    }

    finalRows.push(row);
  }

  // ============================================================
  // WRITE BACK
  // ============================================================
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (finalRows.length) {
    sh.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  }

  sh.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'Skip Trace Finished cleanup complete.\n' +
    'Original rows: ' + body.length + '\n' +
    'Merged duplicates: ' + merges + '\n' +
    'Rows deleted from duplicate merge: ' + rowsToDeleteAfterMerge.size + '\n' +
    'Rows deleted from Above 550K ARV match: ' + deletedByArvMatch + '\n' +
    'Final remaining rows: ' + finalRows.length
  );
}
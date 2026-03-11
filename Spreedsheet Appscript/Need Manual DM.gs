function ST_buildNeedManualDM() {
  const ss = SpreadsheetApp.getActive();

  const stf = ss.getSheetByName("Skip Trace Finished");
  const nid = ss.getSheetByName("Not In Direct");

  if (!stf) {
    SpreadsheetApp.getUi().alert("Skip Trace Finished not found.");
    return;
  }

  const need = ss.getSheetByName("Need Manual DM") || ss.insertSheet("Need Manual DM");
  need.clearContents();

  const stfLastRow = stf.getLastRow();
  const stfLastCol = stf.getLastColumn();

  if (stfLastRow < 1 || stfLastCol < 1) {
    SpreadsheetApp.getUi().alert("Skip Trace Finished is empty.");
    return;
  }

  const stfHeaders = stf.getRange(1, 1, 1, stfLastCol).getValues()[0].map(h => String(h || "").trim());
  const stfData = stfLastRow > 1
    ? stf.getRange(2, 1, stfLastRow - 1, stfLastCol).getValues()
    : [];

  // Output uses the exact same headers as Skip Trace Finished
  need.getRange(1, 1, 1, stfHeaders.length).setValues([stfHeaders]);

  // Required STF columns for move logic
  const phoneIdx = findHeaderIdx_(stfHeaders, ['Phone']);
  const landIdx = findHeaderIdx_(stfHeaders, ['Landline']);

  if (phoneIdx === -1 || landIdx === -1) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished must contain "Phone" and "Landline" headers.');
    return;
  }

  // --------------------------------------------------
  // STEP 1 — Move rows from STF where Phone and Landline are blank
  // --------------------------------------------------
  const rowsFromSTF = [];
  const rowsToDelete = [];
  const seenRows = new Set();

  for (let i = 0; i < stfData.length; i++) {
    const row = stfData[i];

    const phone = String(row[phoneIdx] || "").trim();
    const land = String(row[landIdx] || "").trim();

    if (!phone && !land) {
      const rowCopy = row.slice();
      const rowKey = buildRowKey_(rowCopy);

      if (!seenRows.has(rowKey)) {
        seenRows.add(rowKey);
        rowsFromSTF.push(rowCopy);
      }

      rowsToDelete.push(i + 2); // actual sheet row number in Skip Trace Finished
    }
  }

  // --------------------------------------------------
  // STEP 2 — Append rows from Not In Direct mapped into STF structure
  // --------------------------------------------------
  const rowsFromNID = [];

  if (nid && nid.getLastRow() > 1 && nid.getLastColumn() > 0) {
    const nidHeaders = nid.getRange(1, 1, 1, nid.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const nidData = nid.getRange(2, 1, nid.getLastRow() - 1, nid.getLastColumn()).getValues();

    const nidNormMap = new Map();
    nidHeaders.forEach((h, i) => {
      const key = normalizeHeader_(h);
      if (key && !nidNormMap.has(key)) nidNormMap.set(key, i);
    });

    for (const srcRow of nidData) {
      const outRow = new Array(stfHeaders.length).fill("");

      // 1) Copy exact same-header matches first
      for (let c = 0; c < stfHeaders.length; c++) {
        const targetHeader = stfHeaders[c];
        const srcIdx = nidNormMap.get(normalizeHeader_(targetHeader));
        if (srcIdx !== undefined) outRow[c] = srcRow[srcIdx];
      }

      // 2) Special mappings
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap, ['Property Address'], ['Street Address']);
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap, ['Property City'], ['City']);
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap, ['Property State'], ['State']);
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap, ['Property Zip'], ['Zip']);

      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap,
        ['Representative Address', 'Mailing Street', 'Mailing Address', 'Mailing Add'],
        ['Representative Address', 'Mailing Address', 'Mailing Street']
      );
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap,
        ['Representative City', 'Mailing City', 'Mailing Cit'],
        ['Representative City', 'Mailing City']
      );
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap,
        ['Representative State', 'Mailing State', 'Mailing ST'],
        ['Representative State', 'Mailing State']
      );
      setMappedValue_(outRow, stfHeaders, srcRow, nidNormMap,
        ['Representative Zip', 'Mailing Zipcode', 'Mailing Zip', 'Mailing Zip Code'],
        ['Representative Zip', 'Mailing Zipcode']
      );

      normalizeOutZip_(outRow, stfHeaders, ['Property Zip', 'Zip']);
      normalizeOutZip_(outRow, stfHeaders, ['Representative Zip', 'Mailing Zipcode']);

      const rowKey = buildRowKey_(outRow);
      if (!seenRows.has(rowKey)) {
        seenRows.add(rowKey);
        rowsFromNID.push(outRow);
      }
    }
  }

  // --------------------------------------------------
  // STEP 3 — Write Need Manual DM
  // --------------------------------------------------
  const finalRows = rowsFromSTF.concat(rowsFromNID);

  if (finalRows.length) {
    need.getRange(2, 1, finalRows.length, stfHeaders.length).setValues(finalRows);
  }

  // --------------------------------------------------
  // STEP 4 — Delete moved rows from Skip Trace Finished
  // --------------------------------------------------
  rowsToDelete.sort((a, b) => b - a).forEach(r => stf.deleteRow(r));

  need.setFrozenRows(1);
  need.getRange(1, 1, 1, stfHeaders.length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    'Need Manual DM built.\n' +
    `Moved from Skip Trace Finished: ${rowsFromSTF.length}\n` +
    `Appended from Not In Direct: ${rowsFromNID.length}\n` +
    `Deleted from Skip Trace Finished: ${rowsToDelete.length}\n` +
    `Total unique rows in Need Manual DM: ${finalRows.length}`
  );
}

// --------------------------------------------------
// Helpers
// --------------------------------------------------

function findHeaderIdx_(headers, names) {
  const wanted = names.map(normalizeHeader_);
  for (let i = 0; i < headers.length; i++) {
    if (wanted.includes(normalizeHeader_(headers[i]))) return i;
  }
  return -1;
}

function normalizeHeader_(h) {
  return String(h || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeZip_(z) {
  const m = String(z || '').match(/(\d{5})/);
  return m ? m[1] : '';
}

function setMappedValue_(outRow, targetHeaders, srcRow, srcNormMap, sourceNames, targetNames) {
  let value = '';
  for (const srcName of sourceNames) {
    const srcIdx = srcNormMap.get(normalizeHeader_(srcName));
    if (srcIdx !== undefined) {
      const v = srcRow[srcIdx];
      if (String(v || '').trim() !== '') {
        value = v;
        break;
      }
    }
  }
  if (String(value || '').trim() === '') return;

  for (const targetName of targetNames) {
    const tIdx = findHeaderIdx_(targetHeaders, [targetName]);
    if (tIdx !== -1 && String(outRow[tIdx] || '').trim() === '') {
      outRow[tIdx] = value;
      return;
    }
  }
}

function normalizeOutZip_(outRow, headers, headerNames) {
  const idx = findHeaderIdx_(headers, headerNames);
  if (idx !== -1) outRow[idx] = normalizeZip_(outRow[idx]);
}

function buildRowKey_(row) {
  return row.map(v => String(v || '').trim()).join('||');
}
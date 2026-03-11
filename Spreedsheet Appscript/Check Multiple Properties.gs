function ST_buildMultipleProperties() {

  const ss = SpreadsheetApp.getActive();
  const stf = ss.getSheetByName('Skip Trace Finished');
  const fin = ss.getSheetByName('For Import');
  const out = ss.getSheetByName('Multiple Properties') || ss.insertSheet('Multiple Properties');

  if (!stf || stf.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished missing or empty');
    return;
  }

  if (!fin || fin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('For Import missing or empty');
    return;
  }

  const stfHeaders = stf.getRange(1,1,1,stf.getLastColumn()).getValues()[0];
  const finHeaders = fin.getRange(1,1,1,fin.getLastColumn()).getValues()[0];

  const stfData = stf.getRange(2,1,stf.getLastRow()-1,stf.getLastColumn()).getValues();
  const finData = fin.getRange(2,1,fin.getLastRow()-1,fin.getLastColumn()).getValues();

  const STF_MAIL_ADDR = 12; // M
  const STF_MAIL_ZIP  = 15; // P

  const FIN_REP_ADDR = 13; // N
  const FIN_REP_ZIP  = 16; // Q

  const targetHeaders = stfHeaders.slice();

  out.clearContents();
  out.getRange(1,1,1,targetHeaders.length).setValues([targetHeaders]);
  out.setFrozenRows(1);

  function normAddr(s){
    return String(s||'')
      .toLowerCase()
      .replace(/[^\w\s]/g,' ')
      .replace(/\s+/g,' ')
      .trim();
  }

  function normZip(z){
    const m = String(z||'').match(/(\d{5})/);
    return m ? m[1] : '';
  }

  function headerIndex(headers){
    const map = {};
    headers.forEach((h,i)=>{
      map[String(h||'').trim().toLowerCase()] = i;
    });
    return map;
  }

  const finIdx = headerIndex(finHeaders);
  const tgtIdx = headerIndex(targetHeaders);

  const finMap = new Map();

  for (const r of finData){

    const addr = normAddr(r[FIN_REP_ADDR]);
    const zip  = normZip(r[FIN_REP_ZIP]);

    if(!addr || !zip) continue;

    const key = addr + "|" + zip;

    if(!finMap.has(key)) finMap.set(key,[]);
    finMap.get(key).push(r);
  }

  const output = [];
  const written = new Set();

  let groups = 0;

  for(const stfRow of stfData){

    const addr = normAddr(stfRow[STF_MAIL_ADDR]);
    const zip  = normZip(stfRow[STF_MAIL_ZIP]);

    if(!addr || !zip) continue;

    const key = addr + "|" + zip;

    const matches = finMap.get(key) || [];

    // MORE THAN ONCE
    if(matches.length > 1){

      groups++;

      for(const finRow of matches){

        const row = new Array(targetHeaders.length).fill('');

        // direct header matches
        for(let i=0;i<targetHeaders.length;i++){

          const h = String(targetHeaders[i]||'').toLowerCase().trim();
          const src = finIdx[h];

          if(src !== undefined){
            row[i] = finRow[src];
          }
        }

        // representative → mailing mapping
        if(tgtIdx['mailing address'] !== undefined)
          row[tgtIdx['mailing address']] = finRow[finIdx['representative address']] || '';

        if(tgtIdx['mailing city'] !== undefined)
          row[tgtIdx['mailing city']] = finRow[finIdx['representative city']] || '';

        if(tgtIdx['mailing state'] !== undefined)
          row[tgtIdx['mailing state']] = finRow[finIdx['representative state']] || '';

        if(tgtIdx['mailing zipcode'] !== undefined)
          row[tgtIdx['mailing zipcode']] = finRow[finIdx['representative zip']] || '';

        const sig = JSON.stringify(row);

        if(!written.has(sig)){
          written.add(sig);
          output.push(row);
        }

      }

    }

  }

  if(output.length){
    out.getRange(2,1,output.length,targetHeaders.length).setValues(output);
  }

  SpreadsheetApp.getUi().alert(
    "Multiple Properties rebuilt\n" +
    "STF rows scanned: " + stfData.length + "\n" +
    "Matching duplicate groups: " + groups + "\n" +
    "Rows written: " + output.length
  );

}
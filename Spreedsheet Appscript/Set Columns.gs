function ST_setColumns_DeleteExtras() {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Skip Trace Finished');
  const fin = ss.getSheetByName('For Import');

  if (!sh) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished not found.');
    return;
  }

  if (!fin || fin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('For Import not found or empty.');
    return;
  }

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  if (lastCol < 1) {
    SpreadsheetApp.getUi().alert('Skip Trace Finished is empty.');
    return;
  }

  const currentHeaders = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const data = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues() : [];

  const finLastCol = fin.getLastColumn();
  const finHeaders = fin.getRange(1, 1, 1, finLastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const finData = fin.getLastRow() > 1 ? fin.getRange(2, 1, fin.getLastRow() - 1, finLastCol).getDisplayValues() : [];

  const FIRST_NAME_COL = 250; // IP
  const LAST_NAME_COL  = 251; // IQ

  const targetHeaders = [
    '1st Date Contact Added',
    'Marketing Stage of Contact',
    'Contact Type',
    'Prospect Type',
    'Property County',

    'Property Address',
    'Property City',
    'Property State',
    'Property Zip',

    'First Name',
    'Last Name',
    'Second POC Name',

    'Mailing Address',
    'Mailing City',
    'Mailing State',
    'Mailing Zipcode',

    'Contact Source',
    'Phone',
    'Additional Phone',
    'Landline',
    'Additional Landlines',
    'Email',
    'Additional Emails',

    'Owner Type',
    'Last Sales Date',
    'Last Sales Price',
    'Square Footage',
    'Property Type',

    'Baths',
    'Beds',
    'House Style',
    'Year Built',
    'Pool',

    'AVM',
    'ROS Offer',

    'Rental Estimate Low',
    'Rental Estimate High',

    'Wholesale Value',
    'Market Value',

    'Assessed Total',
    'Total Loans',
    'LTV',

    'Recording Date',
    'Maturity Date',

    'Estimated Mortgage Balance',
    'Estimated Mortgage Payment',
    'Mortgage Interest Rate',

    'Absentee Owner',
    'Active Investor Owned',
    'Active Listing',
    'Bored Investor',
    'Cash Buyer',

    'Delinquent Tax Activity',
    'Flipped',
    'Foreclosures',

    'Free And Clear',
    'High Equity',
    'Long Term Owner',
    'Low Equity',

    'Potentially Inherited',
    'Pre-Foreclosure',
    'Upside Down',
    'Vacancy',
    'Zombie Property',
    'Deceased Probate',

    'Retail Score',
    'Rental Score',
    'Wholesale Score',

    'Auction Date',
    'Date Of Death',

    'PR File Date',
    'Appraised Value',

    'Type Of Phone Number',
    'Tags'
  ];

  const headerMap = {
    'Street Address': 'Property Address',
    'City': 'Property City',
    'State': 'Property State',
    'Zip': 'Property Zip',

    'Mailing Add': 'Mailing Address',
    'Mailing Cit': 'Mailing City',
    'Mailing ST': 'Mailing State',
    'Mailing Zip': 'Mailing Zipcode',

    'OwnerType': 'Owner Type',
    'LastSalesDate': 'Last Sales Date',
    'LastSalesPrice': 'Last Sales Price',
    'SquareFootage': 'Square Footage',
    'PropertyType': 'Property Type',
    'HouseStyle': 'House Style',
    'YearBuilt': 'Year Built',

    'RentalEstimateLow': 'Rental Estimate Low',
    'RentalEstimateHigh': 'Rental Estimate High',

    'WholesaleValue': 'Wholesale Value',
    'MarketValue': 'Market Value',

    'AssessedTotal': 'Assessed Total',
    'TotalLoans': 'Total Loans',

    'RecordingDate': 'Recording Date',
    'MaturityDate': 'Maturity Date',

    'EstimatedMortgageBalance': 'Estimated Mortgage Balance',
    'EstimatedMortgagePayment': 'Estimated Mortgage Payment',
    'MortgageInterestRate': 'Mortgage Interest Rate',

    'AbsenteeOwner': 'Absentee Owner',
    'ActiveInvestorOwned': 'Active Investor Owned',
    'ActiveListing': 'Active Listing',
    'BoredInvestor': 'Bored Investor',
    'CashBuyer': 'Cash Buyer',
    'DelinquentTaxActivity': 'Delinquent Tax Activity',

    'FreeAndClear': 'Free And Clear',
    'HighEquity': 'High Equity',
    'LongTermOwner': 'Long Term Owner',
    'LowEquity': 'Low Equity',

    'PotentiallyInherited': 'Potentially Inherited',
    'PreForeclosure': 'Pre-Foreclosure',
    'UpsideDown': 'Upside Down',
    'ZombieProperty': 'Zombie Property',

    'RetailScore': 'Retail Score',
    'RentalScore': 'Rental Score',
    'WholesaleScore': 'Wholesale Score',

    'AuctionDate': 'Auction Date',

    'PropertyAddress': 'Property Address',
    'PropertyCity': 'Property City',
    'PropertyState': 'Property State',
    'PropertyPostalCode': 'Property Zip',

    'FirstN': 'First Name',
    'LastN': 'Last Name',
    'FirstName': 'First Name',
    'LastName': 'Last Name',

    'Additional Landline': 'Additional Landlines'
  };

  function norm(v) {
    return String(v || '').trim().toLowerCase();
  }

  function normAddr(v) {
    return String(v || '')
      .toLowerCase()
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function zip5(v) {
    const m = String(v || '').match(/(\d{5})/);
    return m ? m[1] : '';
  }

  function cleanText(v) {
    return String(v == null ? '' : v).trim();
  }

  function isBlank(v) {
    return cleanText(v) === '';
  }

  function cleanLTV(v) {
    let s = String(v == null ? '' : v).trim();
    if (!s) return '';
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s)) return '';
    const numLike = s.replace(/[^0-9.\-]/g, '');
    if (!numLike || numLike === '.' || numLike === '-' || numLike === '-.') return '';
    return numLike;
  }

  function buildMappedIndex(headers) {
    const map = new Map();
    headers.forEach((h, i) => {
      const mapped = headerMap[h] || h;
      const key = norm(mapped);
      if (!map.has(key)) map.set(key, i);
    });
    return map;
  }

  function buildPlainIndex(headers) {
    const map = new Map();
    headers.forEach((h, i) => {
      const key = norm(h);
      if (!map.has(key)) map.set(key, i);
    });
    return map;
  }

  const stfMappedIndex = buildMappedIndex(currentHeaders);
  const finIndex = buildPlainIndex(finHeaders);

  function getSTF(row, names) {
    for (const name of names) {
      const idx = stfMappedIndex.get(norm(name));
      if (idx !== undefined) return row[idx];
    }
    return '';
  }

  function getFinIdx(names) {
    for (const name of names) {
      const idx = finIndex.get(norm(name));
      if (idx !== undefined) return idx;
    }
    return -1;
  }

  const finRepAddrIdx = getFinIdx(['Representative Address']);
  const finRepZipIdx = getFinIdx(['Representative Zip', 'Representative Zipcode']);

  const finPropertyCountyIdx = getFinIdx(['Property County', 'County']);
  const finPropertyAddrIdx = getFinIdx(['Property Address', 'Street Address']);
  const finPropertyCityIdx = getFinIdx(['Property City', 'City']);
  const finPropertyStateIdx = getFinIdx(['Property State', 'State']);
  const finPropertyZipIdx = getFinIdx(['Property Zip', 'Property Zip Code', 'Zip', 'Zip Code', 'Zipcode']);
  const finFirstNameIdx = getFinIdx(['First Name']);
  const finLastNameIdx = getFinIdx(['Last Name']);
  const finSecondPOCIdx = getFinIdx(['Second POC Name']);

  const finMap = new Map();
  if (finRepAddrIdx !== -1 && finRepZipIdx !== -1) {
    for (const row of finData) {
      const key = normAddr(row[finRepAddrIdx]) + '|' + zip5(row[finRepZipIdx]);
      if (key === '|') continue;
      if (!finMap.has(key)) finMap.set(key, row);
    }
  }

  const output = data.map(row => {
    const newRow = new Array(targetHeaders.length).fill('');

    for (let c = 0; c < targetHeaders.length; c++) {
      const header = targetHeaders[c];

      if (header === 'First Name') {
        newRow[c] = cleanText(row[FIRST_NAME_COL - 1] || '');
      } else if (header === 'Last Name') {
        newRow[c] = cleanText(row[LAST_NAME_COL - 1] || '');
      } else if (header === 'Property Zip' || header === 'Mailing Zipcode') {
        newRow[c] = zip5(getSTF(row, [header]));
      } else if (header === 'LTV') {
        newRow[c] = cleanLTV(getSTF(row, [header]));
      } else {
        newRow[c] = cleanText(getSTF(row, [header]));
      }
    }

    // Build match from ORIGINAL STF row, not rebuilt row
    const srcMailAddr = getSTF(row, ['Mailing Address', 'Mailing Add']);
    const srcMailZip  = getSTF(row, ['Mailing Zipcode', 'Mailing Zip']);
    const finKey = normAddr(srcMailAddr) + '|' + zip5(srcMailZip);
    const finRow = finMap.get(finKey);

    if (finRow) {
      const backfills = [
        ['Property County', finPropertyCountyIdx, false],
        ['Property Address', finPropertyAddrIdx, false],
        ['Property City', finPropertyCityIdx, false],
        ['Property State', finPropertyStateIdx, false],
        ['Property Zip', finPropertyZipIdx, true],
        ['Second POC Name', finSecondPOCIdx, false]
      ];

      for (const [targetHeader, finIdx, isZip] of backfills) {
        if (finIdx === -1) continue;
        const outIdx = targetHeaders.indexOf(targetHeader);
        if (outIdx === -1) continue;
        if (isBlank(newRow[outIdx])) {
          newRow[outIdx] = isZip ? zip5(finRow[finIdx]) : cleanText(finRow[finIdx]);
        }
      }

      // Only backfill names from For Import if hard-coded IP/IQ did not provide them
      const firstOutIdx = targetHeaders.indexOf('First Name');
      const lastOutIdx = targetHeaders.indexOf('Last Name');

      if (firstOutIdx !== -1 && isBlank(newRow[firstOutIdx]) && finFirstNameIdx !== -1) {
        newRow[firstOutIdx] = cleanText(finRow[finFirstNameIdx]);
      }
      if (lastOutIdx !== -1 && isBlank(newRow[lastOutIdx]) && finLastNameIdx !== -1) {
        newRow[lastOutIdx] = cleanText(finRow[finLastNameIdx]);
      }
    }

    return newRow;
  });

  sh.clearContents();
  sh.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);

  if (output.length) {
    sh.getRange(2, 1, output.length, targetHeaders.length).setValues(output);
  }

  sh.setFrozenRows(1);

  function forcePlainText_(header) {
    const col = targetHeaders.indexOf(header) + 1;
    if (col <= 0 || output.length === 0) return;
    const rng = sh.getRange(2, col, output.length, 1);
    rng.setNumberFormat('@');
    const vals = rng.getDisplayValues();
    rng.setValues(vals.map(r => [String(r[0] || '').trim()]));
  }

  function forceZip5_(header) {
    const col = targetHeaders.indexOf(header) + 1;
    if (col <= 0 || output.length === 0) return;
    const rng = sh.getRange(2, col, output.length, 1);
    rng.setNumberFormat('@');
    const vals = rng.getDisplayValues();
    rng.setValues(vals.map(r => [zip5(r[0])]));
  }

  forceZip5_('Property Zip');
  forceZip5_('Mailing Zipcode');

  forcePlainText_('Phone');
  forcePlainText_('Additional Phone');
  forcePlainText_('Landline');
  forcePlainText_('Additional Landlines');
  forcePlainText_('LTV');

  SpreadsheetApp.getUi().alert('Columns standardized, renamed, and backfilled successfully.');
}
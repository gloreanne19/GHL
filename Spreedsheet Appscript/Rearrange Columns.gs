function reorderColumnsToRequiredOrder() {
  const SHEET_NAME = ''; // optional: set sheet name or leave blank for active sheet
  const HEADER_ROW = 1;
  const CREATE_MISSING_COLUMNS = true;

  const REQUIRED_ORDER = [
    "Created",
    "Contact Id",
    "Prospect Type",
    "Marketing Stage of Contact",
    "Contact Type",              // was "Type"
    "Full Name of POC",           // was "Name"
    "First Name",
    "Last Name",
    "Phone",
    "Email",
    "Subject Property Address",
    "Street Address",             // fixed casing
    "City",
    "State",
    "Postal Code",
    "Property County",
    "Mailing Street",
    "Mailing City",
    "Mailing State",
    "Mailing Zipcode",
    "Source",
    "Auction Date",
    "Property Type",
    "House Style",
    "Year Built",
    "Square Footage",
    "Beds",
    "Baths",
    "Pool",
    "Last Sales Date",
    "Deed Date",
    "-- Record - Believed Bad OR Correct OR Deceased --",
    "4. Bad Call, Kill or NOT A Fit",
    "Owner Type",
    "Retail Score",
    "Rental Score",
    "Loan To Value",
    "Loan Balance +15K",
    "Original Loan",
    "Second POC Name",
    "Deceased Owner",
    "PR File Date",
    "1st Date Contact Added",
    "Equity of Property",
    "Date of Death",
    "Loan Date",
    "Foreclosing Lien",
    "Complaint Type",
    "ROS OFFER",
    "Foreclosures",
    "Absentee Owner",
    "High Equity",
    "Pre-Foreclosure",
    "Deceased Probate",
    "Last Sales Price",
    "Recording Date",
    "AVM",
    "Estimated Value",
    "Market Value",
    "Assessed Total",
    "Rental Estimate Low",
    "Rental Estimate High",
    "Total Loans",
    "Estimated Mortgage Balance",
    "Estimated Mortgage Payment",
    "Mortgage Interest Rate",
    "LTV",
    "Maturity Date",
    "Free And Clear",
    "Equity Percent",
    "Additional Phones",
    "Additional Emails"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sh) throw new Error('Sheet not found. Check SHEET_NAME.');

  const lastCol = sh.getLastColumn();
  if (lastCol === 0) throw new Error('Sheet has no columns.');

  // Read headers
  let headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(h => (h == null ? '' : String(h).trim()));

  // Map header → column index (1-based)
  const headerToCol = {};
  headers.forEach((h, i) => {
    if (h && headerToCol[h] === undefined) headerToCol[h] = i + 1;
  });

  // Create missing required columns
  if (CREATE_MISSING_COLUMNS) {
    const missing = REQUIRED_ORDER.filter(h => headerToCol[h] === undefined);
    if (missing.length) {
      const startCol = sh.getLastColumn() + 1;
      sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
      sh.getRange(HEADER_ROW, startCol, 1, missing.length).setValues([missing]);
    }
  }

  // Re-read headers after insertions
  const finalHeaders = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => (h == null ? '' : String(h).trim()));

  const requiredSet = {};
  REQUIRED_ORDER.forEach(h => requiredSet[h] = true);

  // Identify extra columns
  const extras = finalHeaders.filter(h => !requiredSet[h]);

  // Desired final order
  const existingRequired = REQUIRED_ORDER.filter(h => finalHeaders.includes(h));
  const desired = existingRequired.concat(extras);

  // Move columns into place
  for (let targetIndex = 1; targetIndex <= desired.length; targetIndex++) {
    const desiredHeader = desired[targetIndex - 1];

    const currentHeaders = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(h => (h == null ? '' : String(h).trim()));

    const currentIndex = currentHeaders.indexOf(desiredHeader) + 1;
    if (currentIndex > 0 && currentIndex !== targetIndex) {
      sh.moveColumns(
        sh.getRange(1, currentIndex, sh.getMaxRows(), 1),
        targetIndex
      );
    }
  }

  SpreadsheetApp.flush();
}function myFunction() {
  
}

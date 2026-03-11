function addOutOfStateTag_HighEquity_LongTermOwner_EstateZ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Skip Trace Finished");

  if (!sheet) {
    Logger.log("Sheet 'Skip Trace Finished' not found.");
    return;
  }

  const range = sheet.getDataRange();
  let data = range.getValues();

  if (data.length <= 1) {
    Logger.log("No data rows found.");
    return;
  }

  // -----------------------------
  // HEADER NORMALIZATION
  // -----------------------------
  let headers = data[0].map(h => String(h || "").trim());

  const renameMap = {
    "Property Address": "Street Address",
    "Property City": "City",
    "Property State": "State",
    "Property Zip": "Zip"
  };

  let headerChanged = false;

  headers = headers.map(h => {
    if (renameMap[h]) {
      headerChanged = true;
      return renameMap[h];
    }
    return h;
  });

  if (headerChanged) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    data[0] = headers;
  }

  // -----------------------------
  // Helper: normalize header
  // -----------------------------
  const normH = (h) =>
    String(h || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");

  function findCol_(variants) {
    const wants = variants.map(normH);
    for (let c = 0; c < headers.length; c++) {
      const hc = normH(headers[c]);
      if (wants.includes(hc)) return c;
    }
    return -1;
  }

  // -----------------------------
  // REQUIRED columns
  // -----------------------------
  const idxProspectType =
    findCol_(["Prospect Type"]) !== -1 ? findCol_(["Prospect Type"]) : 3;

  const idxPropertyAddr =
    findCol_(["Street Address", "Property Address"]) !== -1
      ? findCol_(["Street Address", "Property Address"])
      : 5;

  const idxEstateIndicator =
    findCol_(["Estate", "Estate indicator", "Deceased Probate", "Deceased", "EstateZ"]) !== -1
      ? findCol_(["Estate", "Estate indicator", "Deceased Probate", "Deceased", "EstateZ"])
      : 9;

  const idxMailingStreet =
    findCol_(["Mailing Street", "Mailing Address", "Mailing Street Address"]) !== -1
      ? findCol_(["Mailing Street", "Mailing Address", "Mailing Street Address"])
      : 12;

  const idxMailingState =
    findCol_(["Mailing State"]) !== -1 ? findCol_(["Mailing State"]) : 14;

  const idxHighEquity =
    findCol_(["High Equity"]) !== -1 ? findCol_(["High Equity"]) : 56;

  const idxLongTermOwner =
    findCol_(["Long Term Owner"]) !== -1 ? findCol_(["Long Term Owner"]) : 57;

  const idxPrFileDate =
    findCol_(["PR File Date"]) !== -1 ? findCol_(["PR File Date"]) : 0;

  // -----------------------------
  // OPTIONAL columns
  // -----------------------------
  const idxTags = findCol_(["Tags", "Tag"]);
  const idxMktStage = findCol_(["Marketing Stage of Contact", "Marketing Stage"]);
  const idxContactType = findCol_(["Contact Type"]);
  const idxContactSource = findCol_(["Contact Source", "Source"]);

  // -----------------------------
  // Indicator columns to force to 1 or blank
  // -----------------------------
  const indicatorCols = [
    findCol_(["Absentee Owner"]),
    findCol_(["Active Investor Owned"]),
    findCol_(["Active Listing"]),
    findCol_(["Bored Investor"]),
    findCol_(["Cash Buyer"]),
    findCol_(["Delinquent Tax Activity"]),
    findCol_(["Flipped"]),
    findCol_(["Foreclosures"]),
    findCol_(["Free And Clear"]),
    findCol_(["High Equity"]),
    findCol_(["Long Term Owner"]),
    findCol_(["Low Equity"]),
    findCol_(["Potentially Inherited"]),
    findCol_(["Pre-Foreclosure"]),
    findCol_(["Upside Down"]),
    findCol_(["Vacancy"]),
    findCol_(["Zombie Property"]),
    findCol_(["Deceased Probate"])
  ].filter(idx => idx !== -1);

  // -----------------------------
  // Utilities
  // -----------------------------
  const splitTags = (s) =>
    String(s || "")
      .split(/\s*\/\s*/g)
      .map(t => t.trim())
      .filter(Boolean);

  const uniq = (arr) => {
    const seen = new Set();
    const out = [];

    for (const x of arr) {
      const v = String(x || "").trim();
      if (!v) continue;
      if (seen.has(v)) continue;
      seen.add(v);
      out.push(v);
    }

    return out;
  };

  const isEstate = (value) => {
    const txt = String(value || "").trim();
    if (!txt) return false;

    return /\bEST\b/i.test(txt) ||
           /\bEstate\b/i.test(txt) ||
           /\bEstate of\b/i.test(txt);
  };

  function isOne_(v) {
    const s = String(v || "").trim().toUpperCase();
    return s === "1" || s === "TRUE";
  }

  function formatAsMDY_(value) {
    let d;

    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
      d = value;
    } else {
      const s = String(value || "").trim();
      if (!s) return "";
      d = new Date(s);
      if (isNaN(d)) return s;
    }

    const m = d.getMonth() + 1;
    const day = d.getDate();
    const y = d.getFullYear();
    return `${m}/${day}/${y}`;
  }

  function todayMDY_() {
    const d = new Date();
    const m = d.getMonth() + 1;
    const day = d.getDate();
    const y = d.getFullYear();
    return `${m}/${day}/${y}`;
  }

  // -----------------------------
  // MAIN LOOP
  // -----------------------------
  for (let i = 1; i < data.length; i++) {
    const mailingState = String(data[i][idxMailingState] || "").trim().toUpperCase();
    const propertyAddr = String(data[i][idxPropertyAddr] || "").trim();
    const mailingStreet = String(data[i][idxMailingStreet] || "").trim();

    const highEquity = data[i][idxHighEquity];
    const longTermOwner = data[i][idxLongTermOwner];
    const estateCol = data[i][idxEstateIndicator];

    let ptTokens = splitTags(data[i][idxProspectType]).map(t => t.toUpperCase());

    if (!ptTokens.includes("PO")) ptTokens.unshift("PO");

    // OS / AO
    if (mailingState) {
      if (mailingState !== "TX") {
        if (!ptTokens.includes("OS")) ptTokens.push("OS");
        if (!ptTokens.includes("AO")) ptTokens.push("AO");
      } else {
        if (propertyAddr && mailingStreet && propertyAddr !== mailingStreet) {
          if (!ptTokens.includes("AO")) ptTokens.push("AO");
        }
      }
    }

    // HQ
    if (isOne_(highEquity)) {
      if (!ptTokens.includes("HQ")) ptTokens.push("HQ");
    }

    // SD
    if (isOne_(longTermOwner)) {
      if (!ptTokens.includes("SD")) ptTokens.push("SD");
    }

    // Z
    if (isEstate(estateCol) || isOne_(estateCol)) {
      if (!ptTokens.includes("Z")) ptTokens.push("Z");
    }

    data[i][idxProspectType] = uniq(ptTokens).join(" / ");

    // -----------------------------
    // Force indicator fields to 1 or blank
    // -----------------------------
    for (const idx of indicatorCols) {
      data[i][idx] = isOne_(data[i][idx]) ? 1 : "";
    }

    // -----------------------------
    // Make sure PR File Date is not empty
    // Format: 2/19/2026
    // -----------------------------
    if (idxPrFileDate !== -1) {
      const cur = data[i][idxPrFileDate];
      if (!String(cur || "").trim()) {
        data[i][idxPrFileDate] = todayMDY_();
      } else {
        data[i][idxPrFileDate] = formatAsMDY_(cur);
      }
    }

    // -----------------------------
    // OPTIONAL FIELD UPDATES
    // -----------------------------
    if (idxTags !== -1) {
      const cur = String(data[i][idxTags] || "").trim();
      const parts = cur ? cur.split(/[,;\n]+/).map(x => x.trim()).filter(Boolean) : [];

      const cleaned = parts.filter(x => x.toLowerCase() !== "dfw");
      const lowerCleaned = cleaned.map(x => x.toLowerCase());

      if (!lowerCleaned.includes("dfw_area")) {
        cleaned.push("dfw_area");
      }

      data[i][idxTags] = cleaned.join(", ");
    }

    if (idxMktStage !== -1) {
      if (!String(data[i][idxMktStage] || "").trim()) {
        data[i][idxMktStage] = "Prospect - Not Contacted";
      }
    }

    if (idxContactType !== -1) {
      if (!String(data[i][idxContactType] || "").trim()) {
        data[i][idxContactType] = "Property Owner - POC";
      }
    }

    if (idxContactSource !== -1) {
      if (!String(data[i][idxContactSource] || "").trim()) {
        data[i][idxContactSource] = "Propelio";
      }
    }

    // NOTE: Type Of Phone Number is intentionally NOT updated
  }

  // -----------------------------
  // Write updated data back
  // -----------------------------
  range.setValues(data);

  // -----------------------------
  // Force PR File Date display format
  // -----------------------------
  if (idxPrFileDate !== -1 && sheet.getLastRow() > 1) {
    sheet.getRange(2, idxPrFileDate + 1, sheet.getLastRow() - 1, 1).setNumberFormat("m/d/yyyy");
  }

  // -----------------------------
  // Delete columns after all changes
  // -----------------------------
  const columnsToDelete = [
    "Owner Type",
    "Last Sales Date",
    "Last Sales Price",
    "Square Footage",
    "Property Type",
    "Baths",
    "Beds",
    "House Style",
    "Year Built",
    "Pool",
    "AVM",
    "ROS Offer",
    "Rental Estimate Low",
    "Rental Estimate High",
    "Wholesale Value",
    "Market Value",
    "Assessed Total",
    "Total Loans",
    "LTV",
    "Recording Date",
    "Maturity Date",
    "Estimated Mortgage Balance",
    "Estimated Mortgage Payment",
    "Mortgage Interest Rate",
    "Absentee Owner",
    "Active Investor Owned",
    "Active Listing",
    "Bored Investor",
    "Cash Buyer",
    "Delinquent Tax Activity",
    "Flipped",
    "Foreclosures",
    "Free And Clear",
    "High Equity",
    "Long Term Owner",
    "Low Equity",
    "Potentially Inherited",
    "Pre-Foreclosure",
    "Upside Down",
    "Vacancy",
    "Zombie Property",
    "Deceased Probate",
    "Retail Score",
    "Rental Score",
    "Wholesale Score",
    "Auction Date"
  ];

  let currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h || "").trim());

  const deleteIdxs = [];
  for (let c = 0; c < currentHeaders.length; c++) {
    if (columnsToDelete.includes(currentHeaders[c])) {
      deleteIdxs.push(c + 1); // 1-based
    }
  }

  // Delete from right to left
  deleteIdxs.sort((a, b) => b - a);
  for (const col of deleteIdxs) {
    sheet.deleteColumn(col);
  }

  Logger.log("Updated headers, Prospect Type, PR File Date, Tags, marketing fields, and deleted selected columns.");
}
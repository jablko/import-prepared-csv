function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu().addItem("Start", "start").addToUi();
}

function start() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile("Index");
  ui.showModalDialog(html, "Import prepared CSV");
}

/**
 * Do the things ... parse CSV, compare it to the existing spreadsheet and add
 * and/or modify transactions.
 *
 * @param {string} csv
 * @param {boolean} preview
 */
function doImport(csv, preview) {
  const sheet = /** @type {GoogleAppsScript.Spreadsheet.Sheet} */ (
    ss.getSheetByName("Transactions")
  );
  if (sheet === null) {
    throw new Error("Transactions sheet not found");
  }
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  Logger.log(`Get headers from range ${headerRange.getA1Notation()}`);
  /**
   * Header names -> (single last) zero-based spreadsheet column numbers. The
   * first header is normally a GoogleAppsScript.Spreadsheet.CellImage.
   *
   * @type {ReadonlyMap<unknown, number>}
   */
  const tHeader = new Map(
    headerRange
      .getValues()
      .flat()
      .map((name, i) => [name, i]),
  );
  const idRange = sheet.getRange(
    2,
    /** @type {number} */ (tHeader.get("Transaction ID")) + 1,
    sheet.getLastRow() - 1,
  );
  Logger.log(`Get transaction IDs from range ${idRange.getA1Notation()}`);
  /**
   * Transaction IDs -> (multiple) zero-based spreadsheet row numbers.
   *
   * @type {ReadonlyMap<unknown, readonly number[]>}
   */
  const byID = invertArray(idRange.getValues().flat());
  /**
   * @type {[
   *   (readonly string[])?,
   *   ...(readonly (string | number | Date)[][]),
   * ]}
   */
  const [headerData, ...sData] = Utilities.parseCsv(csv);
  if (headerData === undefined) {
    throw new Error("Empty CSV");
  }
  /**
   * Header names -> (single) CSV column numbers. We exploit it being the same
   * size as the number of CSV headers/columns.
   */
  const sHeader = new Map(headerData.map((name, i) => [name, i]));
  if (sHeader.size !== headerData.length) {
    throw new Error("Duplicate CSV headers");
  }
  // Prepare CSV
  prepareDate(sHeader, sData);
  prepareAmount(sHeader, sData);
  prepareDescription(sHeader, sData);
  prepareID(sHeader, sData, tHeader);
  /**
   * For duplicate IDs, map CSV transactions -> spreadsheet transactions by ID
   * occurrence number.
   *
   * @type {Record<string, number>}
   */
  const counter = {};
  /**
   * CSV -> spreadsheet --- row numbers -> row numbers.
   *
   * @type {readonly (number | undefined)[]}
   */
  const rMap = sData.map((sRow) => {
    const value = sRow[/** @type {number} */ (sHeader.get("Transaction ID"))];
    //return byID.get(value)?.[(counter[/** @type {never} */ (value)] ??= 0)++];
    return byID.get(value)?.[
      (counter[/** @type {never} */ (value)] =
        (counter[/** @type {never} */ (value)] ?? -1) + 1)
    ];
  });
  /**
   * CSV -> spreadsheet --- column numbers -> column numbers.
   *
   * @type {readonly (number | undefined)[]}
   */
  const cMap = [...sHeader.keys().map((name) => tHeader.get(name))];
  // The smallest rectangle that covers the imported rows and columns. The
  // Spreadsheet service operates on one rectangle at a time, it is not
  // asynchronous and doing an unbounded number of operations is slow, so
  // operate on one, minimum rectangle for theoretical performance? The Sheets
  // API batchGet()/batchUpdate() are overkill?
  const rMin = Math.min(...rMap.filter((i) => i !== undefined));
  const rMax = Math.max(...rMap.filter((i) => i !== undefined));
  const cMin = Math.min(...cMap.filter((j) => j !== undefined));
  const cMax = Math.max(...cMap.filter((j) => j !== undefined));
  const addCMin = Math.min(cMin, tHeader.get("Date Added") ?? Math.min());
  const addCMax = Math.max(cMax, tHeader.get("Date Added") ?? Math.max());
  // Separate transactions to add and ones to modify
  /** @type {(string | number | Date)[][]} */
  const addData = [];
  /** @type {(string | number | Date)[][]} */
  let modifyData = [];
  for (const [i, sRow] of sData.entries()) {
    const newRow = [];
    for (const [j, value] of sRow.entries()) {
      newRow[/** @type {never} */ (cMap[j])] = value;
    }
    if (rMap[i] === undefined) {
      newRow[/** @type {never} */ (tHeader.get("Date Added"))] = dateAdded;
      addData.push(newRow.slice(addCMin));
    } else {
      modifyData[rMap[i] - rMin] = newRow.slice(cMin);
    }
  }
  /**
   * Existing spreadsheet transactions for comparison.
   *
   * @type {(string | number | Date)[][]}
   */
  let tData;
  if (modifyData.length > 0) {
    const range = sheet.getRange(
      1 + rMin + 1,
      cMin + 1,
      rMax - rMin + 1,
      cMax - cMin + 1,
    );
    Logger.log(`Get range ${range.getA1Notation()} for comparison`);
    tData = range.getValues();
  }
  if (preview) {
    /** Compare to the existing spreadsheet. */
    const nModify = modifyData.reduce(
      (n, newRow, i) =>
        n +
        /** @type {never} */ (
          newRow.some((value, j) => !equals(value, tData[i][j]))
        ),
      0,
    );
    return {
      nAdd: addData.length,
      nModify,
      nTotal: sData.length,
      headers: headerData.map((name) => [
        name,
        tHeader.get(name) !== undefined,
      ]),
    };
  }
  add();
  modify();
  sheet.sort(/** @type {number} */ (tHeader.get("Date")) + 1, false);

  function add() {
    // No transactions to add or disjoint headers
    if (addData.length < 1 || cMax < cMin) {
      return;
    }
    const nGrow = sheet.getLastRow() + addData.length - sheet.getMaxRows();
    if (nGrow > 0) {
      Logger.log(`Grow spreadsheet by ${nGrow} rows`);
      sheet.insertRowsAfter(sheet.getMaxRows(), nGrow);
    }
    const range = sheet.getRange(
      sheet.getLastRow() + 1,
      addCMin + 1,
      addData.length,
      addCMax - addCMin + 1,
    );
    Logger.log(`Add transactions to range ${range.getA1Notation()}`);
    noDataValidationSetValues(range, addData);
  }

  function modify() {
    // Compare to the existing spreadsheet and shrink the rectangle that we
    // modify to cover just the rows and columns that differ
    const modifyRMin = Number(
      Object.keys(modifyData).find((i) =>
        modifyData[i].some((value, j) => !equals(value, tData[i][j])),
      ),
    );
    // All are already identical
    if (Number.isNaN(modifyRMin)) {
      return;
    }
    const modifyRMax = Number(
      Object.keys(modifyData).findLast((i) =>
        modifyData[i].some((value, j) => !equals(value, tData[i][j])),
      ),
    );
    modifyData = modifyData.slice(modifyRMin, modifyRMax + 1);
    tData = tData.slice(modifyRMin, modifyRMax + 1);
    const modifyCMin = Math.min(
      ...modifyData
        .map((newRow, i) =>
          Object.keys(newRow).find((j) => !equals(newRow[j], tData[i][j])),
        )
        .filter((j) => j !== undefined),
    );
    const modifyCMax = Math.max(
      ...modifyData
        .map((newRow, i) =>
          Object.keys(newRow).findLast((j) => !equals(newRow[j], tData[i][j])),
        )
        .filter((j) => j !== undefined),
    );
    tData = tData.map((tRow) => tRow.slice(modifyCMin, modifyCMax + 1));
    // Merge CSV and the existing spreadsheet. `modifyData` is sparse.
    for (const [i, newRow] of Object.entries(modifyData)) {
      Object.assign(tData[i], newRow.slice(modifyCMin, modifyCMax + 1));
    }
    const range = sheet.getRange(
      1 + rMin + modifyRMin + 1,
      cMin + modifyCMin + 1,
      modifyRMax - modifyRMin + 1,
      modifyCMax - modifyCMin + 1,
    );
    Logger.log(`Modify transactions in range ${range.getA1Notation()}`);
    noDataValidationSetValues(range, tData);
  }
}

const ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Values -> (multiple) indexes.
 *
 * @template T
 * @param {readonly T[]} array
 */
function invertArray(array) {
  /** @type {Map<T, number[]>} */
  const groups = new Map();
  for (const [i, value] of array.entries()) {
    if (groups.get(value)?.push(i) === undefined) {
      groups.set(value, [i]);
    }
  }
  return groups;
}

/**
 * Parse the date and derive the month and week columns.
 *
 * @param {Map<string, number>} sHeader
 * @param {readonly (string | number | Date)[][]} sData
 */
function prepareDate(sHeader, sData) {
  if (sHeader.get("Date") === undefined) {
    return;
  }
  for (const sRow of sData) {
    sRow[/** @type {number} */ (sHeader.get("Date"))] = new Date(
      sRow[/** @type {number} */ (sHeader.get("Date"))],
    );
  }
  /**
   * JS interprets YYYY-MM-DD strings using UTC and other strings without
   * explicit time zones using the script time zone, which may or may not equal
   * the spreadsheet time zone. We do not detect the case it encountered
   * (explicit, YYYY-MM-DD or other) but if dates are all midnight, local time
   * or UTC, then guess they were all date-only and reinterpret them using the
   * spreadsheet time zone. The ultimate solution is probably Temporal objects
   * when Apps Script implements that, or a third-party library until then?
   */
  const from = sData.every((sRow) =>
    isMidnight(
      /** @type {never} */ (sRow[/** @type {number} */ (sHeader.get("Date"))]),
    ),
  )
    ? scriptTimeZone
    : sData.every((sRow) =>
          isMidnightUTC(
            /** @type {never} */ (
              sRow[/** @type {number} */ (sHeader.get("Date"))]
            ),
          ),
        )
      ? "GMT"
      : null;
  if (from !== null) {
    Logger.log(
      `Interpret date-only strings using the spreadsheet time zone ${ssTimeZone}`,
    );
    for (const sRow of sData) {
      sRow[/** @type {number} */ (sHeader.get("Date"))] = addTimeZoneOffset(
        /** @type {never} */ (
          sRow[/** @type {number} */ (sHeader.get("Date"))]
        ),
        ssTimeZone,
        from,
      );
    }
  }
  sHeader.set("Month", sHeader.get("Month") ?? sHeader.size);
  sHeader.set("Week", sHeader.get("Week") ?? sHeader.size);
  for (const sRow of sData) {
    /** Spreadsheet time zone calendar day. */
    const local = addTimeZoneOffset(
      /** @type {never} */ (sRow[/** @type {number} */ (sHeader.get("Date"))]),
      scriptTimeZone,
      ssTimeZone,
    );
    sRow[/** @type {number} */ (sHeader.get("Month"))] = addTimeZoneOffset(
      new Date(local.getFullYear(), local.getMonth()),
      ssTimeZone,
      scriptTimeZone,
    );
    sRow[/** @type {number} */ (sHeader.get("Week"))] = addTimeZoneOffset(
      new Date(
        local.getFullYear(),
        local.getMonth(),
        local.getDate() - local.getDay(),
      ),
      ssTimeZone,
      scriptTimeZone,
    );
  }
}

/** @param {Date} date */
function isMidnight(date) {
  return (
    date.getHours() === 0 &&
    date.getMinutes() === 0 &&
    date.getSeconds() === 0 &&
    date.getMilliseconds() === 0
  );
}

/** @param {Date} date */
function isMidnightUTC(date) {
  return (
    date.getUTCHours() === 0 &&
    date.getUTCMinutes() === 0 &&
    date.getUTCSeconds() === 0 &&
    date.getUTCMilliseconds() === 0
  );
}

const scriptTimeZone = Session.getScriptTimeZone();
const ssTimeZone = ss.getSpreadsheetTimeZone();

/**
 * Add to `date` the difference between the `to` and `from` time zones (`to`
 * minus `from`). `from` defaults to GMT, i.e. zero.
 *
 * The implementation formats `date` using the `from` time zone, then
 * reinterprets that representation using the `to` time zone. Using the JS date
 * time string format is arbitrary, equivalent formats make no difference.
 *
 * @param {Date} date
 * @param {string} to
 */
function addTimeZoneOffset(date, to, from = "GMT") {
  return Utilities.parseDate(
    Utilities.formatDate(date, from, dateTimeStringFormat),
    to,
    dateTimeStringFormat,
  );
}

/** https://tc39.es/ecma262/multipage/numbers-and-dates.html#sec-date-time-string-format */
const dateTimeStringFormat = "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'";

/**
 * Normalize the amount like the Spreadsheet service does, for `equals()` and
 * `prepareID()`.
 *
 * @param {ReadonlyMap<string, number>} sHeader
 * @param {readonly (string | number | Date)[][]} sData
 */
function prepareAmount(sHeader, sData) {
  if (sHeader.get("Amount") === undefined) {
    return;
  }
  for (const sRow of sData) {
    sRow[/** @type {number} */ (sHeader.get("Amount"))] = Number(
      /** @type {string} */ (
        sRow[/** @type {number} */ (sHeader.get("Amount"))]
      ).replace(/[$,]/g, ""),
    );
  }
}

/**
 * Derive the description and full description from either column.
 *
 * @param {Map<string, number>} sHeader
 * @param {readonly (string | number | Date)[][]} sData
 */
function prepareDescription(sHeader, sData) {
  const j = sHeader.get("Full Description") ?? sHeader.get("Description");
  if (j === undefined) {
    return;
  }
  sHeader.set("Description", sHeader.get("Description") ?? sHeader.size);
  sHeader.set(
    "Full Description",
    sHeader.get("Full Description") ?? sHeader.size,
  );
  for (const sRow of sData) {
    sRow[/** @type {number} */ (sHeader.get("Description"))] = toDescription(
      /** @type {never} */ (
        sRow[/** @type {number} */ (sHeader.get("Full Description"))] = sRow[j]
      ),
    );
  }
}

/**
 * ?Yodlee redacts numbers and Tiller title cases the description.
 *
 * @param {string} fullDescription
 */
function toDescription(fullDescription) {
  return fullDescription
    .replace(/ {2,}/g, " ")
    .replace(/[0-9](?=[- 0-9]+[0-9]{3})/g, "X")
    .replace(/[^- ]{2}[^ ]*/g, (match, start) => {
      switch (match.toUpperCase()) {
        case "A":
        case "AN":
        case "AND":
        case "AT":
        case "BUT":
        case "BY":
        case "FOR":
        case "IN":
        case "NOR":
        case "OF":
        case "OFF":
        case "ON":
        case "OR":
        case "OUT":
        case "SO":
        case "THE":
        case "TO":
        case "UP":
        case "VIA":
        case "YET":
          if (start === 0) {
            break;
          }
          return match.toLowerCase();
        case "AB":
        case "AK":
        case "AL":
        case "AR":
        case "AZ":
        case "BC":
        case "CA":
        case "CO":
        case "CT":
        case "DC":
        case "DE":
        case "FL":
        case "GA":
        case "HI":
        case "IA":
        case "ID":
        case "IL":
        case "IN":
        case "KS":
        case "KY":
        case "LA":
        case "MA":
        case "MB":
        case "MD":
        case "ME":
        case "MI":
        case "MN":
        case "MO":
        case "MS":
        case "MT":
        case "NB":
        case "NC":
        case "ND":
        case "NE":
        case "NH":
        case "NJ":
        case "NL":
        case "NM":
        case "NS":
        case "NT":
        case "NU":
        case "NV":
        case "NY":
        //case "OH":
        case "OK":
        case "ON":
        case "OR":
        case "PA":
        case "PEI":
        case "QC":
        case "RI":
        case "SC":
        case "SD":
        case "SK":
        case "TN":
        case "TX":
        case "UT":
        case "VA":
        case "VT":
        case "WA":
        case "WI":
        case "WV":
        case "WY":
        case "YT":
          return match;
      }
      const [first, ...rest] = match;
      return first + rest.join("").toLowerCase();
    })
    .replace(/X{3,}/gi, "x");
}

/**
 * Give transactions content IDs so we do not import them again, at least.
 *
 * @param {Map<string, number>} sHeader
 * @param {readonly (string | number | Date)[][]} sData
 * @param {ReadonlyMap<unknown, number>} tHeader
 */
function prepareID(sHeader, sData, tHeader) {
  if (sHeader.get("Transaction ID") !== undefined) {
    return;
  }
  Logger.log("Give transactions content IDs");
  sHeader.set("Transaction ID", sHeader.size);
  /**
   * Exclude the derived month and week and include only the normalized
   * description.
   */
  const canonical = [
    ...sHeader.keys().filter((name) => {
      switch (name) {
        case "Month":
        case "Week":
        case "Full Description":
          return false;
      }
      return tHeader.get(name) !== undefined;
    }),
  ];
  canonical.sort();
  /** Do not give identical transactions duplicate IDs. */
  const groups = /** @type {Record<string, (string | number | Date)[][]>} */ (
    Object.groupBy(sData, (sRow) => {
      /** @type {Record<string, string | number | Date>} */
      const o = {};
      for (const name of canonical) {
        const value = sRow[/** @type {number} */ (sHeader.get(name))];
        if (value !== "") {
          o[name] = value;
        }
      }
      return JSON.stringify(o).toUpperCase();
    })
  );
  for (const [s, g] of Object.entries(groups)) {
    for (const [i, sRow] of g.entries()) {
      const digest = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        s + i,
      );
      // RFC 6920 Naming Things with Hashes
      sRow[/** @type {number} */ (sHeader.get("Transaction ID"))] =
        `ni:///sha-256;${Utilities.base64EncodeWebSafe(digest).replaceAll("=", "")}`;
    }
  }
}

const dateAdded = new Date();

/**
 * Compare two Date objects as numbers (of milliseconds since the epoch) and
 * compare strings and numbers loosely.
 *
 * @param {string | number | Date} a
 * @param {string | number | Date} b
 * @returns {boolean}
 */
function equals(a, b) {
  return a instanceof Date && b instanceof Date
    ? equals(Number(a), Number(b))
    : a == b;
}

/**
 * Suppress data validation errors.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {(string | number | Date)[][]} values
 */
function noDataValidationSetValues(range, values) {
  const rules = range.getDataValidations();
  range.clearDataValidations();
  try {
    range.setValues(values);
  } finally {
    range.setDataValidations(rules);
  }
}

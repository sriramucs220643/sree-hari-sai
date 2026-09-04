/**
 * BILL TRACKER — Google Apps Script Backend
 * ------------------------------------------
 * Database: Google Sheets (one tab per day, named DD-MM-YYYY)
 * Columns:  S.No | Date | Time | Payment Method | Amount
 *
 * SETUP:
 * 1. Paste this file into a Google Apps Script project bound to (or
 *    referencing) your spreadsheet.
 * 2. Set SPREADSHEET_ID below to your spreadsheet's ID.
 * 3. Deploy as a Web App (see README.md for full steps).
 */

// ⚠️ REPLACE with your actual Google Spreadsheet ID
const SPREADSHEET_ID = 'PUT_YOUR_SPREADSHEET_ID_HERE';

const HEADERS = ['S.No', 'Date', 'Time', 'Payment Method', 'Amount'];
const VALID_METHODS = ['Cash', 'GPay'];

// ---------------------------------------------------------------------
// ENTRY POINTS
// ---------------------------------------------------------------------

function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'getBills') {
      return jsonResponse(getBills(e.parameter.date));
    }

    if (action === 'getDates') {
      return jsonResponse(getAllDates());
    }

    return jsonResponse({ success: false, message: 'Unknown GET action' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function doPost(e) {
  try {
    // Frontend sends JSON as text/plain to avoid CORS preflight.
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'addBill') {
      return jsonResponse(addBill(body.paymentMethod, body.amount));
    }

    if (action === 'updateBill') {
      return jsonResponse(updateBill(body.date, body.sno, body.paymentMethod, body.amount));
    }

    if (action === 'deleteBill') {
      return jsonResponse(deleteBill(body.date, body.sno));
    }

    return jsonResponse({ success: false, message: 'Unknown POST action' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

// ---------------------------------------------------------------------
// CORE ACTIONS
// ---------------------------------------------------------------------

/**
 * Adds a bill to TODAY's sheet. Auto-generates S.No, Date, Time.
 * Uses LockService so concurrent requests never collide on S.No.
 */
function addBill(paymentMethod, amountRaw) {
  const validation = validateInput(paymentMethod, amountRaw);
  if (!validation.valid) {
    return { success: false, message: validation.message };
  }
  const amount = validation.amount;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // wait up to 30s for the lock

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const todayStr = formatDate(new Date());
    const sheet = getOrCreateDailySheet(ss, todayStr);

    const nextSno = getNextSerial(sheet);
    const timeStr = formatTime(new Date());

    sheet.appendRow([nextSno, todayStr, timeStr, paymentMethod, amount]);

    return {
      success: true,
      message: 'Bill added successfully',
      bill: { sno: nextSno, date: todayStr, time: timeStr, paymentMethod, amount }
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns all bills for a given date (DD-MM-YYYY), newest first,
 * plus computed summary totals.
 */
function getBills(dateStr) {
  if (!dateStr) {
    return { success: false, message: 'Date is required' };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(dateStr);

  if (!sheet) {
    return {
      success: true,
      date: dateStr,
      bills: [],
      summary: { totalSales: 0, cash: 0, gpay: 0, totalBills: 0 }
    };
  }

  const lastRow = sheet.getLastRow();
  let bills = [];

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    bills = data
      .filter(row => row[0] !== '' && row[0] !== null) // skip blank rows
      .map(row => ({
        sno: row[0],
        date: row[1],
        time: row[2],
        paymentMethod: row[3],
        amount: Number(row[4])
      }));
  }

  // Newest first
  bills.sort((a, b) => b.sno - a.sno);

  const summary = bills.reduce(
    (acc, b) => {
      acc.totalSales += b.amount;
      if (b.paymentMethod === 'Cash') acc.cash += b.amount;
      if (b.paymentMethod === 'GPay') acc.gpay += b.amount;
      acc.totalBills += 1;
      return acc;
    },
    { totalSales: 0, cash: 0, gpay: 0, totalBills: 0 }
  );

  return { success: true, date: dateStr, bills, summary };
}

/**
 * Updates payment method and/or amount for a specific S.No on a specific date.
 */
function updateBill(dateStr, sno, paymentMethod, amountRaw) {
  const validation = validateInput(paymentMethod, amountRaw);
  if (!validation.valid) {
    return { success: false, message: validation.message };
  }
  const amount = validation.amount;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(dateStr);
    if (!sheet) {
      return { success: false, message: 'No sheet found for that date' };
    }

    const rowIndex = findRowBySno(sheet, sno);
    if (rowIndex === -1) {
      return { success: false, message: 'Bill not found' };
    }

    sheet.getRange(rowIndex, 4).setValue(paymentMethod); // Payment Method
    sheet.getRange(rowIndex, 5).setValue(amount);        // Amount

    return { success: true, message: 'Bill updated successfully' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a bill row by S.No on a given date. Does not renumber others.
 */
function deleteBill(dateStr, sno) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(dateStr);
    if (!sheet) {
      return { success: false, message: 'No sheet found for that date' };
    }

    const rowIndex = findRowBySno(sheet, sno);
    if (rowIndex === -1) {
      return { success: false, message: 'Bill not found' };
    }

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'Bill deleted successfully' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns all existing daily sheet names, sorted newest first,
 * so the frontend can populate the "Select Date" dropdown.
 */
function getAllDates() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const names = ss.getSheets()
    .map(s => s.getName())
    .filter(name => /^\d{2}-\d{2}-\d{4}$/.test(name))
    .sort((a, b) => parseDDMMYYYY(b) - parseDDMMYYYY(a));

  return { success: true, dates: names };
}

// ---------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------

function getOrCreateDailySheet(ss, dateStr) {
  let sheet = ss.getSheetByName(dateStr);
  if (!sheet) {
    sheet = ss.insertSheet(dateStr);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getNextSerial(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const snoValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => Number(r[0]))
    .filter(n => !isNaN(n));

  if (snoValues.length === 0) return 1;
  return Math.max(...snoValues) + 1;
}

function findRowBySno(sheet, sno) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const snoValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < snoValues.length; i++) {
    if (Number(snoValues[i][0]) === Number(sno)) {
      return i + 2; // +2 because data starts at row 2 and arrays are 0-indexed
    }
  }
  return -1;
}

function validateInput(paymentMethod, amountRaw) {
  if (VALID_METHODS.indexOf(paymentMethod) === -1) {
    return { valid: false, message: 'Invalid payment method' };
  }
  const amount = Number(amountRaw);
  if (isNaN(amount) || amount <= 0) {
    return { valid: false, message: 'Invalid amount' };
  }
  return { valid: true, amount: amount };
}

function formatDate(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'hh:mm a');
}

function parseDDMMYYYY(str) {
  const [dd, mm, yyyy] = str.split('-').map(Number);
  return new Date(yyyy, mm - 1, dd).getTime();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

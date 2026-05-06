// ============================================================
// Daily Sales Tracker — Google Apps Script
// 
// SETUP INSTRUCTIONS:
// 1. Open Google Sheet > Extensions > Apps Script
// 2. Delete ALL existing code
// 3. Paste this entire file
// 4. Press Ctrl+S to save
// 5. Click "Deploy" > "New deployment"
// 6. Click the gear icon > select "Web app"
// 7. Description: Sales Tracker
// 8. Execute as: Me
// 9. Who has access: Anyone
// 10. Click "Deploy"
// 11. Click "Authorise access" > Choose your Google account > Allow
// 12. COPY the Web App URL (looks like: https://script.google.com/macros/s/ABC123.../exec)
// 13. Paste it into the Settings tab of your Sales Tracker app
//
// IMPORTANT: Every time you edit this script, you must:
// Deploy > Manage deployments > Edit (pencil icon) > Version: New version > Deploy
// ============================================================

var SHEET_NAME = 'Sales Data';
var HEADERS    = ['Date','Cash','Card','Web Sales','Just Eat','Uber Eats','Deliveroo','Foodhub','Total'];

function doGet(e) {
  var cb     = (e.parameter && e.parameter.callback) ? e.parameter.callback : 'cb';
  var action = (e.parameter && e.parameter.action)   ? e.parameter.action   : 'getAll';
  var result;

  try {
    if (action === 'save') {
      var rowStr = e.parameter.row ? decodeURIComponent(e.parameter.row) : '{}';
      var row    = JSON.parse(rowStr);
      result     = saveRow(row);
    } else if (action === 'getAll') {
      result = getAllRows();
    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  var json   = JSON.stringify(result);
  var output = ContentService.createTextOutput(cb + '(' + json + ');');
  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  return output;
}

// Also handle POST (some browsers send preflight)
function doPost(e) {
  return doGet(e);
}

function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var hRange = sheet.getRange(1, 1, 1, HEADERS.length);
    hRange.setValues([HEADERS]);
    hRange.setBackground('#1a1a2e');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    hRange.setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    for (var i = 2; i <= HEADERS.length; i++) {
      sheet.setColumnWidth(i, 105);
    }
  }
  return sheet;
}

function cellDateToStr(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(cell).trim();
}

function saveRow(row) {
  if (!row || !row.date) {
    return { success: false, error: 'No date provided' };
  }

  var sheet  = getOrCreateSheet();
  var date   = String(row.date).trim();
  var values = [
    date,
    Number(row.cash)      || 0,
    Number(row.card)      || 0,
    Number(row.web)       || 0,
    Number(row.justeat)   || 0,
    Number(row.ubereats)  || 0,
    Number(row.deliveroo) || 0,
    Number(row.foodhub)   || 0,
    Number(row.total)     || 0
  ];

  var lastRow  = sheet.getLastRow();
  var foundRow = -1;

  if (lastRow >= 2) {
    var dateVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < dateVals.length; i++) {
      if (cellDateToStr(dateVals[i][0]) === date) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
    var newLast = sheet.getLastRow();
    if (newLast > 2) {
      sheet.getRange(2, 1, newLast - 1, HEADERS.length).sort({ column: 1, ascending: true });
    }
  }

  applyStyles(sheet);
  return { success: true, message: 'Saved: ' + date };
}

function getAllRows() {
  var sheet   = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, rows: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (r[0] === '' || r[0] === null || r[0] === undefined) continue;
    rows.push({
      date:      cellDateToStr(r[0]),
      cash:      Number(r[1]) || 0,
      card:      Number(r[2]) || 0,
      web:       Number(r[3]) || 0,
      justeat:   Number(r[4]) || 0,
      ubereats:  Number(r[5]) || 0,
      deliveroo: Number(r[6]) || 0,
      foodhub:   Number(r[7]) || 0,
      total:     Number(r[8]) || 0
    });
  }

  return { success: true, rows: rows };
}

function applyStyles(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (var i = 2; i <= lastRow; i++) {
    sheet.getRange(i, 1, 1, HEADERS.length)
         .setBackground(i % 2 === 0 ? '#f0f4ff' : '#ffffff');
  }
  sheet.getRange(2, 2, lastRow - 1, HEADERS.length - 1)
       .setNumberFormat('"£"#,##0.00');
}

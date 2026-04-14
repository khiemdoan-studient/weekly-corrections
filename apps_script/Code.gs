/**
 * Automated Weekly Corrections — Apps Script
 *
 * Two behaviors:
 *   1. Checkbox (Sheet 1, col A, row 4+): copies row to Sheet 3 with date stamp
 *   2. Dropdown filter (any sheet, row 2): hides/shows data rows to match selection
 *
 * Installation:
 *   1. Open the "Automated Weekly Corrections" spreadsheet
 *   2. Extensions > Apps Script
 *   3. Paste this code (replace everything), then Save (Ctrl+S)
 *   4. Do NOT click Run — triggers fire automatically on edits
 */

function onEdit(e) {
  // Guard: e is undefined when run manually from the script editor
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();

  // ── Dropdown filter (row 2 on any sheet) ─────────────────────────────
  if (row === 2) {
    applyDropdownFilter_(sheet);
    return;
  }

  // ── Checkbox approval (Sheet 1, col A, row 4+) ──────────────────────
  if (sheet.getName() !== "Corrected Roster Info") return;
  if (col !== 1) return;
  if (row <= 3) return; // rows 1-2 = filters, row 3 = header

  var newValue = range.getValue();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  if (newValue === true) {
    // Read columns B through M (12 data columns)
    var data = sheet.getRange(row, 2, 1, 12).getValues()[0];

    var targetSheet = ss.getSheetByName("Automated Correction List");
    if (!targetSheet) return;

    // Append with date stamp
    var approvalDate = new Date();
    targetSheet.appendRow([approvalDate].concat(data));

    // Format the date cell
    var lastRow = targetSheet.getLastRow();
    targetSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");

    // Grey out the approved row
    sheet.getRange(row, 1, 1, 14).setBackground("#E8ECF1");

  } else if (newValue === false) {
    // Unchecked — remove grey background
    sheet.getRange(row, 1, 1, 14).setBackground(null);
  }
}


/**
 * Apply dropdown filter: hide data rows that don't match any active dropdown.
 * "All" or empty = show everything for that column.
 */
function applyDropdownFilter_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 3) return; // no data rows

  var sheetName = sheet.getName();
  var dataStartRow = 4; // row 4 is first data row (1=labels, 2=dropdowns, 3=headers)

  // Read all dropdown values from row 2
  var lastCol = sheet.getLastColumn();
  var filterValues = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // Read column headers from row 3 to know which column each filter maps to
  var headers = sheet.getRange(3, 1, 1, lastCol).getValues()[0];

  // Build filter map: column index → required value (skip "All" and empty)
  var filters = {};
  for (var c = 0; c < filterValues.length; c++) {
    var fv = String(filterValues[c]).trim();
    if (fv && fv !== "All") {
      filters[c] = fv.toLowerCase();
    }
  }

  // If no active filters, show all rows
  if (Object.keys(filters).length === 0) {
    showAllRows_(sheet, dataStartRow, lastRow);
    return;
  }

  // Read all data rows at once for performance
  var numDataRows = lastRow - dataStartRow + 1;
  var data = sheet.getRange(dataStartRow, 1, numDataRows, lastCol).getValues();

  // Determine which rows to hide/show
  var rowsToHide = [];
  var rowsToShow = [];

  for (var r = 0; r < data.length; r++) {
    var rowData = data[r];
    var match = true;

    for (var c in filters) {
      var cellVal = String(rowData[c]).trim().toLowerCase();
      if (cellVal !== filters[c]) {
        match = false;
        break;
      }
    }

    if (match) {
      rowsToShow.push(dataStartRow + r);
    } else {
      rowsToHide.push(dataStartRow + r);
    }
  }

  // Batch hide/show for performance
  if (rowsToHide.length > 0) {
    batchSetRowVisibility_(sheet, rowsToHide, true);
  }
  if (rowsToShow.length > 0) {
    batchSetRowVisibility_(sheet, rowsToShow, false);
  }
}


/**
 * Show all data rows (reset filter).
 */
function showAllRows_(sheet, startRow, endRow) {
  sheet.showRows(startRow, endRow - startRow + 1);
}


/**
 * Hide or show rows in batches of consecutive ranges for performance.
 */
function batchSetRowVisibility_(sheet, rows, hide) {
  if (rows.length === 0) return;

  rows.sort(function(a, b) { return a - b; });

  var start = rows[0];
  var count = 1;

  for (var i = 1; i < rows.length; i++) {
    if (rows[i] === rows[i - 1] + 1) {
      count++;
    } else {
      if (hide) {
        sheet.hideRows(start, count);
      } else {
        sheet.showRows(start, count);
      }
      start = rows[i];
      count = 1;
    }
  }

  // Final batch
  if (hide) {
    sheet.hideRows(start, count);
  } else {
    sheet.showRows(start, count);
  }
}

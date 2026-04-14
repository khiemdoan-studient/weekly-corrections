/**
 * Automated Weekly Corrections — Apps Script
 *
 * Checkbox (Sheet 1 "Corrected Roster Info", col A, row 7+):
 *   Copies the student row to hidden _ApprovedData tab with a date stamp.
 *   The visible "Automated Correction List" sheet reads from _ApprovedData
 *   via a QUERY formula, so it updates automatically.
 *
 * Dropdown filter (row 5): clears stale checkboxes so they align with
 *   the new SORT(QUERY()) output after a filter/sort change.
 *
 * Installation:
 *   1. Open the "Automated Weekly Corrections" spreadsheet
 *   2. Extensions > Apps Script
 *   3. Paste this code (replace everything), then Save (Ctrl+S)
 *   4. Do NOT click Run — triggers fire automatically on edits
 */

function onEdit(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var row = range.getRow();
  var col = range.getColumn();

  // ── Dropdown filter or Sort By changed (row 5) → clear stale checkboxes
  if (row === 5 && sheetName === "Corrected Roster Info") {
    clearCheckboxes_(sheet);
    return;
  }

  // ── Checkbox approval (Sheet 1, col A, row 7+) ──────────────────
  if (sheetName !== "Corrected Roster Info") return;
  if (col !== 1) return;
  if (row <= 6) return;

  var newValue = range.getValue();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  if (newValue === true) {
    // Read columns B through M (12 data columns from QUERY output)
    var data = sheet.getRange(row, 2, 1, 12).getValues()[0];

    // Skip empty rows (QUERY may leave gaps)
    if (!data[0] && !data[10]) return;

    // Append to hidden _ApprovedData (not the visible sheet)
    var targetSheet = ss.getSheetByName("_ApprovedData");
    if (!targetSheet) {
      // Fallback: create _ApprovedData if it doesn't exist
      targetSheet = ss.insertSheet("_ApprovedData");
      targetSheet.hideSheet();
    }

    var approvalDate = new Date();
    targetSheet.appendRow([approvalDate].concat(data));

    // Format the date cell
    var lastRow = targetSheet.getLastRow();
    targetSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");

    // Grey out the approved row on Sheet 1
    sheet.getRange(row, 1, 1, 14).setBackground("#E8ECF1");

  } else if (newValue === false) {
    sheet.getRange(row, 1, 1, 14).setBackground(null);
  }
}

function clearCheckboxes_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 6) return;
  var checkRange = sheet.getRange(7, 1, lastRow - 6, 1);
  checkRange.setValue(false);
  sheet.getRange(7, 1, lastRow - 6, 14).setBackground(null);
}

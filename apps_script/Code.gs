/**
 * Automated Weekly Corrections — Apps Script
 *
 * Two behaviors:
 *   1. Checkbox (Sheet 1 "Corrected Roster Info", col A, row 7+):
 *      copies the student row to Sheet 3 with a date stamp
 *   2. Dropdown filter (row 5): clears stale checkboxes so they
 *      align with the new QUERY output
 *
 * Filtering is handled by QUERY formulas on the visible sheets —
 * NOT by this script. The QUERY references dropdown cells in row 5
 * and auto-recalculates when a dropdown changes.
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
  var sheetName = sheet.getName();
  var row = range.getRow();
  var col = range.getColumn();

  // ── Dropdown filter changed (row 5) → clear stale checkboxes ─────
  if (row === 5 && sheetName === "Corrected Roster Info") {
    clearCheckboxes_(sheet);
    return;
  }

  // ── Checkbox approval (Sheet 1, col A, row 7+) ──────────────────
  if (sheetName !== "Corrected Roster Info") return;
  if (col !== 1) return;
  if (row <= 6) return; // rows 1-2=title, 3=spacer, 4-5=filters, 6=header

  var newValue = range.getValue();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  if (newValue === true) {
    // Read columns B through M (12 data columns from QUERY output)
    var data = sheet.getRange(row, 2, 1, 12).getValues()[0];

    // Skip empty rows (QUERY may leave gaps)
    if (!data[0] && !data[10]) return; // no Campus and no Student_ID = empty

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
 * Clear all checkboxes in column A when a filter dropdown changes.
 * The QUERY formula recalculates and shows different rows, so
 * existing checkboxes are stale and must be reset.
 */
function clearCheckboxes_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 6) return;

  var checkRange = sheet.getRange(7, 1, lastRow - 6, 1);
  checkRange.setValue(false);
  // Also clear any grey backgrounds from previous approvals
  sheet.getRange(7, 1, lastRow - 6, 14).setBackground(null);
}

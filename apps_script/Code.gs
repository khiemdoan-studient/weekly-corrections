/**
 * Automated Weekly Corrections — Apps Script
 *
 * Sheet 1 "Corrected Roster Info" has two checkbox columns:
 *   Col A = Accept Changes (light green) — routes to approval sheets by mismatch type
 *   Col B = Reject Changes (light red)  — routes all rejections to _RejectedData
 *
 * Accept routing (col A checkbox = true):
 *   Reads Mismatch Summary from col O (column 15) to determine target:
 *     "Roster Addition"  → _AdditionsData → "Roster Additions" sheet
 *     "Unenrolling"      → _UnenrollData  → "Roster Unenrollments" sheet
 *     field mismatches   → _ApprovedData  → "Automated Correction List" sheet
 *
 * Reject routing (col B checkbox = true):
 *   All rejected rows → _RejectedData → "Rejected Changes" sheet
 *   "Reason for Rejection" column on Rejected Changes is blank for manual entry.
 *
 * Mutual exclusion: checking Accept unchecks Reject and vice versa.
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

  // ── Checkbox action (Sheet 1, col A or B, row 7+) ──────────────────
  if (sheetName !== "Corrected Roster Info") return;
  if (col !== 1 && col !== 2) return;
  if (row <= 6) return;

  var newValue = range.getValue();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  if (newValue === true) {
    // Read columns C through N (12 data columns from QUERY output)
    var data = sheet.getRange(row, 3, 1, 12).getValues()[0];

    // Skip empty rows (QUERY may leave gaps)
    if (!data[0] && !data[10]) return;

    // Mutual exclusion: uncheck the other column
    var otherCol = (col === 1) ? 2 : 1;
    sheet.getRange(row, otherCol).setValue(false);

    // Read Mismatch Summary (col O = column 15) — used by both accept and reject
    var mismatchSummary = sheet.getRange(row, 15).getValue();

    if (col === 1) {
      // ── ACCEPT: route by mismatch type ──────────────────────────
      var targetTabName;
      if (mismatchSummary === "Roster Addition") {
        targetTabName = "_AdditionsData";
      } else if (mismatchSummary === "Unenrolling") {
        targetTabName = "_UnenrollData";
      } else {
        targetTabName = "_ApprovedData";
      }

      var targetSheet = ss.getSheetByName(targetTabName);
      if (!targetSheet) {
        targetSheet = ss.insertSheet(targetTabName);
        targetSheet.hideSheet();
      }

      var approvalDate = new Date();
      targetSheet.appendRow([approvalDate, mismatchSummary].concat(data));

      var lastRow = targetSheet.getLastRow();
      targetSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");

    } else {
      // ── REJECT: all rejections go to _RejectedData ──────────────
      var targetSheet = ss.getSheetByName("_RejectedData");
      if (!targetSheet) {
        targetSheet = ss.insertSheet("_RejectedData");
        targetSheet.hideSheet();
      }

      var rejectionDate = new Date();
      targetSheet.appendRow([rejectionDate, mismatchSummary].concat(data));

      var lastRow = targetSheet.getLastRow();
      targetSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
    }

    // Grey out the row on Sheet 1
    sheet.getRange(row, 1, 1, 15).setBackground("#E8ECF1");

  } else if (newValue === false) {
    sheet.getRange(row, 1, 1, 15).setBackground(null);
  }
}

function clearCheckboxes_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 6) return;
  var numRows = lastRow - 6;
  // Clear both Accept (col A) and Reject (col B) checkboxes
  sheet.getRange(7, 1, numRows, 2).setValue(false);
  sheet.getRange(7, 1, numRows, 15).setBackground(null);
}

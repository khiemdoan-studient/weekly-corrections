/**
 * Automated Weekly Corrections — Apps Script
 *
 * onEdit trigger: when a manager checks a checkbox in Sheet 1 ("Corrected Roster Info")
 * column A, the student's row is copied to Sheet 3 ("Automated Correction List")
 * with today's date as the approval timestamp.
 *
 * Installation:
 *   1. Open the "Automated Weekly Corrections" spreadsheet
 *   2. Extensions > Apps Script
 *   3. Paste this code, replacing any existing code
 *   4. Save (Ctrl+S)
 *   5. No need to deploy — onEdit triggers run automatically
 */

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Only act on "Corrected Roster Info" sheet, column A (checkbox)
  if (sheet.getName() !== "Corrected Roster Info") return;
  if (range.getColumn() !== 1) return;
  if (range.getRow() <= 1) return; // skip header

  var newValue = range.getValue();

  if (newValue === true) {
    // Checkbox was CHECKED — copy row to Automated Correction List
    var row = range.getRow();

    // Read columns B through M (12 data columns: Campus through External Student ID)
    var data = sheet.getRange(row, 2, 1, 12).getValues()[0];

    // Get target sheet
    var targetSheet = e.source.getSheetByName("Automated Correction List");
    if (!targetSheet) {
      SpreadsheetApp.getUi().alert(
        "Error: 'Automated Correction List' tab not found."
      );
      return;
    }

    // Append row with date stamp
    var approvalDate = new Date();
    var outputRow = [approvalDate].concat(data);
    targetSheet.appendRow(outputRow);

    // Format the date cell in the new row
    var lastRow = targetSheet.getLastRow();
    targetSheet
      .getRange(lastRow, 1)
      .setNumberFormat("yyyy-MM-dd HH:mm:ss");

    // Visual feedback: grey out the approved row in Sheet 1
    sheet.getRange(row, 1, 1, 14).setBackground("#E8ECF1");

  } else if (newValue === false) {
    // Checkbox was UNCHECKED — remove grey background (allow re-review)
    sheet.getRange(range.getRow(), 1, 1, 14).setBackground(null);
  }
}

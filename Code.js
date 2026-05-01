/**
 * Automated Weekly Corrections — Apps Script (v2.6.0)
 *
 * Source of truth: this file in the repo. Deployed to the Apps Script
 * project bound to the corrections spreadsheet via clasp:
 *
 *   npm run deploy   # local: runs node --check Code.js && clasp push
 *
 * GHA also auto-deploys on every push to main that touches Code.js or
 * appsscript.json (.github/workflows/deploy-apps-script.yml). The live
 * Apps Script ALWAYS matches HEAD on main — no manual paste step.
 *
 * This file combines two features that share the same Apps Script project.
 * Each feature activates only when its relevant sheet tab is present, so
 * pushing this into either the Weekly Corrections spreadsheet OR any ISR
 * (which has Student Cards) is safe.
 *
 * FEATURE 1: Accept/Reject checkbox handler (onEdit)
 *   - Active only when current sheet is "Corrected Roster Info"
 *   - Accept (col A, green) routes to _ApprovedData / _AdditionsData / _UnenrollData by mismatch type
 *   - Reject (col B, red) routes to _RejectedData
 *   - IDEMPOTENT: toggling a checkbox (accept → uncheck → reject) removes stale entries;
 *     student only appears in one cumulative tab at a time (the latest state)
 *   - PRE-FORMATTED dates: eliminates race condition on concurrent checkbox edits
 *   - PRESERVES col A (green) + col B (red) backgrounds — only greys/clears data cols C:O
 *
 * FEATURE 2: Student Cards generator (onOpen menu)
 *   - Active only when current sheet has a "Copy of MAP Roster" tab (i.e. ISRs)
 *   - Adds "Student Cards" menu → generates Letter portrait PDF cards
 */

// ═════════════════════════════════════════════════════════════════════════
// SHARED: onOpen — conditionally shows Student Cards menu
// ═════════════════════════════════════════════════════════════════════════

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Student Cards menu: only on spreadsheets that actually have the source tab
  if (ss.getSheetByName(SHEET_NAME)) {
    ui.createMenu("Student Cards")
      .addItem("Generate Letter Slides + PDF", "generateStudentCardsFromTemplate")
      .addToUi();
  }
}

// ═════════════════════════════════════════════════════════════════════════
// FEATURE 1: ACCEPT/REJECT CHECKBOX HANDLER
// ═════════════════════════════════════════════════════════════════════════

function onEdit(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();

  // Only handle edits on the corrections sheet; no-op everywhere else
  if (sheetName !== "Corrected Roster Info") return;

  var row = range.getRow();
  var col = range.getColumn();

  // Dropdown filter or Sort By changed (row 5) → clear stale checkboxes
  if (row === 5) {
    clearCheckboxes_(sheet);
    return;
  }

  // Only respond to col A (Accept) or col B (Reject), row 7+
  if (col !== 1 && col !== 2) return;
  if (row <= 6) return;

  var newValue = range.getValue();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  // Read the student record once (needed for both check and uncheck paths)
  var data = sheet.getRange(row, 3, 1, 12).getValues()[0];

  // Skip empty rows (QUERY may leave gaps)
  if (!data[0] && !data[10]) return;

  var studentId = String(data[10] || "").trim();

  // IDEMPOTENCY — on ANY checkbox state change, first remove all existing
  // rows for this student from every cumulative tab. Then, if newValue=true,
  // append one row to the target tab. This handles: accept → uncheck → reject,
  // reject → uncheck → accept, and repeated accepts (no duplicates).
  removeStudentFromCumulativeTabs_(ss, studentId);

  if (newValue === true) {
    // Mutual exclusion: uncheck the other column (fires another onEdit but
    // that one sees newValue=false and just removes — which we already did)
    var otherCol = col === 1 ? 2 : 1;
    sheet.getRange(row, otherCol).setValue(false);

    // Read the Mismatch Summary (col O) for routing
    var mismatchSummary = sheet.getRange(row, 15).getValue();

    // Pre-format the timestamp to eliminate the setNumberFormat race condition
    var tz = Session.getScriptTimeZone();
    var dateString = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    var targetTabName;
    if (col === 1) {
      // ACCEPT: route by mismatch type
      if (mismatchSummary === "Roster Addition") {
        targetTabName = "_AdditionsData";
      } else if (mismatchSummary === "Unenrolling") {
        targetTabName = "_UnenrollData";
      } else {
        targetTabName = "_ApprovedData";
      }
    } else {
      // REJECT: always _RejectedData
      targetTabName = "_RejectedData";
    }

    var targetSheet = ss.getSheetByName(targetTabName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(targetTabName);
      targetSheet.hideSheet();
    }
    targetSheet.appendRow([dateString, mismatchSummary].concat(data));

    // Grey-out DATA cols only (C:O, cols 3–15). Preserve col A (green) and col B (red) base colors.
    sheet.getRange(row, 3, 1, 13).setBackground("#E8ECF1");
  } else {
    // Uncheck — clear data col backgrounds only; A/B keep their permanent colors
    sheet.getRange(row, 3, 1, 13).setBackground(null);
  }
}

function clearCheckboxes_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 6) return;
  var numRows = lastRow - 6;
  // Clear both Accept (col A) and Reject (col B) checkboxes
  sheet.getRange(7, 1, numRows, 2).setValue(false);
  // Reset only DATA col backgrounds (cols 3–15). Preserve A/B column colors.
  sheet.getRange(7, 3, numRows, 13).setBackground(null);
}

/**
 * Remove every row for `studentId` from all cumulative tabs. Called before
 * any append so the student only exists in one tab (the latest choice).
 * Student_ID is at col M (col 13) in the 14-col layout:
 * [Date, MismatchSummary, Campus, Grade, Level, First, Last, Email,
 *  StudentGroup, GuideFirst, GuideLast, GuideEmail, StudentID, ExtStudentID]
 */
function removeStudentFromCumulativeTabs_(ss, studentId) {
  if (!studentId) return;
  var tabs = ["_ApprovedData", "_AdditionsData", "_UnenrollData", "_RejectedData"];
  var SID_COL = 13; // col M
  for (var t = 0; t < tabs.length; t++) {
    var tab = ss.getSheetByName(tabs[t]);
    if (!tab) continue;
    var lastRow = tab.getLastRow();
    if (lastRow < 1) continue;
    var values = tab.getRange(1, SID_COL, lastRow, 1).getValues();
    // Delete from bottom up so row indices don't shift while iterating
    for (var i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0] || "").trim() === studentId) {
        tab.deleteRow(i + 1);
      }
    }
  }
}

// ═════════════════════════════════════════════════════════════════════════
// FEATURE 2: STUDENT CARDS (unchanged from original Student Cards script)
// ═════════════════════════════════════════════════════════════════════════

// ---------- CONFIG ----------
const SHEET_NAME = "Copy of MAP Roster";
const OUTPUT_FOLDER_ID = "1jmWYiBj-jdgEB6YwHRSdfyHH9K9vNayA";

// Template (Letter Portrait)
const TEMPLATE_PRESENTATION_URL_OR_ID =
  "https://docs.google.com/presentation/d/1H-ovHPaTuIpmIZm3_WeCLKqkLMK-sAEcRgH-xrkGddw/edit";

// Logo file (Drive)
const LOGO_URL_OR_ID =
  "https://drive.google.com/file/d/1v33iHmEPPgeK7Y2INBCX3s-__wwdn-zY/view?usp=sharing";

// Password line (same for every card)
const PW_LINE = "PW: iloveschool";

// Exact headers in your sheet
const COL_FIRST  = "First Name";
const COL_LAST   = "Last Name";
const COL_EMAIL  = "Student Email";
const COL_CAMPUS = "Campus";
const COL_PERIOD = "Period";

// Filters
const SKIP_IF_NO_EMAIL = true;
const SKIP_IF_NO_NAME  = true;

// Layout
const GRID_ROWS = 3;
const GRID_COLS = 2;
const CARDS_PER_PAGE = GRID_ROWS * GRID_COLS;

// Letter portrait points
const PAGE_WIDTH_PT  = 612;
const PAGE_HEIGHT_PT = 792;

// Margins and spacing
const MARGIN_PT = 36;
const GUTTER_PT = 12;

// Typography
const NAME_FONT   = 20;
const EMAIL_FONT  = 12;
const PW_FONT     = 12;
const EXTRA_FONT  = 10;

// Logo placement
const LOGO_WIDTH_PT = 80;
const LOGO_MARGIN_PT = 10;
// ----------------------------

function generateStudentCardsFromTemplate() {
  // 1) Read data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("No data rows found.");

  const headers = values[0].map(function (h) { return String(h).trim(); });
  const idx = headerIndex_(headers);

  [COL_FIRST, COL_LAST, COL_EMAIL, COL_CAMPUS, COL_PERIOD].forEach(function (col) {
    if (!(col in idx)) throw new Error("Missing required header: \"" + col + "\"");
  });

  const students = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const first  = clean_(row[idx[COL_FIRST]]);
    const last   = clean_(row[idx[COL_LAST]]);
    const email  = clean_(row[idx[COL_EMAIL]]);
    const campus = clean_(row[idx[COL_CAMPUS]]);
    const period = clean_(row[idx[COL_PERIOD]]);

    if (SKIP_IF_NO_NAME && !(first || last)) continue;
    if (SKIP_IF_NO_EMAIL && !email) continue;

    students.push({ first: first, last: last, email: email, campus: campus, period: period });
  }
  if (!students.length) throw new Error("No eligible students to print.");

  // 2) Prepare logo blob
  const logoId = extractId_(LOGO_URL_OR_ID);
  const logoBlob = DriveApp.getFileById(logoId).getBlob();

  // 3) Copy template
  const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const templateId = extractId_(TEMPLATE_PRESENTATION_URL_OR_ID);
  const templateFile = DriveApp.getFileById(templateId);

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const outputName = "JHES - Hardeeville Elementary School Student Cards (Letter) v2- " + stamp;
  const copyFile = templateFile.makeCopy(outputName, folder);
  const presId = copyFile.getId();

  // 4) Clear slides
  const pres = SlidesApp.openById(presId);
  const existingSlides = pres.getSlides();
  for (let i = existingSlides.length - 1; i >= 0; i--) {
    existingSlides[i].remove();
  }

  // 5) Geometry
  const usableW = PAGE_WIDTH_PT - (2 * MARGIN_PT);
  const usableH = PAGE_HEIGHT_PT - (2 * MARGIN_PT);
  const cardW = (usableW - (GUTTER_PT * (GRID_COLS - 1))) / GRID_COLS;
  const cardH = (usableH - (GUTTER_PT * (GRID_ROWS - 1))) / GRID_ROWS;

  // 6) Build slides
  for (let i = 0; i < students.length; i += CARDS_PER_PAGE) {
    const pageStudents = students.slice(i, i + CARDS_PER_PAGE);
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    for (let s = 0; s < CARDS_PER_PAGE; s++) {
      const rr = Math.floor(s / GRID_COLS);
      const cc = s % GRID_COLS;
      const x = MARGIN_PT + cc * (cardW + GUTTER_PT);
      const y = MARGIN_PT + rr * (cardH + GUTTER_PT);

      const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, cardW, cardH);
      shape.getFill().setTransparent();
      const border = shape.getBorder();
      if (border) border.setWeight(1);
      shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

      const logoX = x + cardW - LOGO_WIDTH_PT - LOGO_MARGIN_PT;
      const logoY = y + LOGO_MARGIN_PT;
      slide.insertImage(logoBlob, logoX, logoY, LOGO_WIDTH_PT, LOGO_WIDTH_PT);

      const tf = shape.getText();
      tf.clear();

      const student = pageStudents[s];
      if (!student) {
        tf.setText(" ");
        continue;
      }

      const fullName = [student.first, student.last].filter(Boolean).join(" ").trim();
      const email = student.email || "";
      const campus = student.campus || "";
      const period = student.period || "";

      const combined = fullName + "\n" + email + "\n" + PW_LINE + "\n" + campus + "\n" + period;
      tf.setText(combined);

      tf.getRange(0, combined.length)
        .getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      tf.getTextStyle().setFontSize(EXTRA_FONT).setBold(false);

      const nameLen = fullName.length;
      const emailLen = email.length;
      const pwLen = PW_LINE.length;
      const campusLen = campus.length;
      const periodLen = period.length;

      const startName = 0;
      const endName = startName + nameLen;
      const startEmail = endName + 1;
      const endEmail = startEmail + emailLen;
      const startPW = endEmail + 1;
      const endPW = startPW + pwLen;
      const startCampus = endPW + 1;
      const endCampus = startCampus + campusLen;
      const startPeriod = endCampus + 1;
      const endPeriod = startPeriod + periodLen;

      if (nameLen > 0) tf.getRange(startName, endName).getTextStyle().setFontSize(NAME_FONT).setBold(true);
      if (emailLen > 0) tf.getRange(startEmail, endEmail).getTextStyle().setFontSize(EMAIL_FONT).setBold(false);
      if (pwLen > 0) tf.getRange(startPW, endPW).getTextStyle().setFontSize(PW_FONT).setBold(false);
      if (campusLen > 0) tf.getRange(startCampus, endCampus).getTextStyle().setFontSize(EXTRA_FONT).setBold(false);
      if (periodLen > 0) tf.getRange(startPeriod, endPeriod).getTextStyle().setFontSize(EXTRA_FONT).setBold(false);
    }
  }

  pres.saveAndClose();

  // 7) Export PDF
  const pdfBlob = DriveApp.getFileById(presId).getAs(MimeType.PDF).setName(outputName + ".pdf");
  folder.createFile(pdfBlob);

  SpreadsheetApp.getUi().alert("Done!\n\nSlides:\n" + copyFile.getUrl());
}

function extractId_(urlOrId) {
  const s = String(urlOrId || "").trim();
  const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s;
}

function headerIndex_(headers) {
  const map = {};
  headers.forEach(function (h, i) {
    if (h) map[h] = i;
  });
  return map;
}

function clean_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

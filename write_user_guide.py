"""write_user_guide.py — Write formatted user guide to the Google Doc.

Follows the same structure/quality as the Student Performance Dashboard
user guide (write_claude_guide.py).

Target: https://docs.google.com/document/d/1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4
"""

import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import SERVICE_ACCOUNT_KEY

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_KEY, scopes=["https://www.googleapis.com/auth/documents"]
)
docs = build("docs", "v1", credentials=creds)
DOC_ID = "1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4"

# ── Clear existing content ────────────────────────────────────────────────
doc = docs.documents().get(documentId=DOC_ID).execute()
body_content = doc.get("body", {}).get("content", [])
end_index = body_content[-1]["endIndex"] if body_content else 1
requests = []
if end_index > 2:
    requests.append(
        {"deleteContentRange": {"range": {"startIndex": 1, "endIndex": end_index - 1}}}
    )

# ── Document text ─────────────────────────────────────────────────────────
text = (
    "Automated Weekly Corrections\n"
    "User Guide \u2014 Implementation Manager Reference\n"
    "\n"
    "Overview\n"
    "This spreadsheet identifies students whose enrollment information in the SIS (Student "
    "Information System) does not match the MAP roster. Implementation Managers review these "
    "mismatches weekly, check off corrections, and the data team submits the changes every Friday.\n"
    "\n"
    "The spreadsheet has five sheets:\n"
    "\u2022 Corrected Roster Info \u2014 What the data should be (from the MAP roster)\n"
    "\u2022 Current Roster Info in SIS \u2014 What the SIS currently has (for comparison)\n"
    "\u2022 Automated Correction List \u2014 Running history of approved field-mismatch corrections\n"
    "\u2022 Roster Additions \u2014 Running history of approved new student enrollments\n"
    "\u2022 Roster Unenrollments \u2014 Running history of approved student unenrollments\n"
    "\u2022 Rejected Changes \u2014 Running history of rejected corrections with reason column\n"
    "\n"
    "PART 1: HOW TO REVIEW CORRECTIONS\n"
    "\n"
    "Step 1: Open the spreadsheet and set your filters\n"
    "At the top of each sheet, there are dropdown filters for Campus, Grade, Level, Student "
    "Group, and Guide Email. Select your campus from the Campus dropdown to see only your students. "
    "You can also use the Sort By dropdown to change the order (e.g., sort by Grade or by "
    "Mismatch Summary to see specific types of issues).\n"
    "\n"
    "Tip: Set all your filters before checking any boxes. Changing a filter resets all checkboxes.\n"
    "\n"
    "Step 2: Compare the two sheets side by side\n"
    "Open Corrected Roster Info (Sheet 1) and Current Roster Info in SIS (Sheet 2) side by side. "
    "Both sheets show the same students in the same order. The last column in Sheet 1, "
    "\u201cMismatch Summary,\u201d tells you exactly which fields are different (e.g., \u201cCampus, Grade\u201d "
    "or \u201cGuide Email, Guide Name\u201d).\n"
    "\n"
    "Step 3: Accept or reject corrections\n"
    "Each row has two checkboxes: Accept Changes (column A, green) and Reject Changes (column B, red).\n"
    "\u2022 Accept \u2014 Check column A if the MAP roster data is correct and should replace the SIS data. "
    "The student automatically appears in the appropriate approval sheet with today\u2019s date.\n"
    "\u2022 Reject \u2014 Check column B if the correction should NOT be applied. "
    "The student appears in the Rejected Changes sheet where you can add a reason.\n"
    "Checking one automatically unchecks the other.\n"
    "\n"
    "Important: Do not check boxes for students outside your campus. Only check corrections for "
    "schools you manage.\n"
    "\n"
    "PART 2: WHAT EACH SHEET DOES\n"
    "\n"
    "Corrected Roster Info\n"
    "What it shows: MAP roster data (source of truth) for students whose information doesn\u2019t "
    "match the SIS. The Mismatch Summary column is color-coded:\n"
    "\u2022 Green \u2014 Roster Addition (student is enrolled in MAP but not yet in the SIS)\n"
    "\u2022 Yellow \u2014 Field mismatch (student exists in both but specific fields differ)\n"
    "\u2022 Light yellow \u2014 Unenrolling (student is no longer enrolled in MAP but still enrolled in SIS)\n"
    "What you do: Review the data. Check Accept Changes (column A, green) to approve, or "
    "Reject Changes (column B, red) to reject. Your choice is automatically routed to the correct sheet.\n"
    "\n"
    "Current Roster Info in SIS\n"
    "What it shows: What the SIS pipeline currently has for the same students.\n"
    "What you do: Compare this with Sheet 1 to see what\u2019s wrong. Do not edit this sheet.\n"
    "Note: Both sheets show the same students in the same row order for easy comparison.\n"
    "\n"
    "Automated Correction List\n"
    "What it shows: A running list of approved field-mismatch corrections with the date each was checked off.\n"
    "What you do: This sheet is read-only for IMs. The data team uses it every Friday.\n"
    "Note: This list is cumulative \u2014 it keeps a full history and is never cleared automatically.\n"
    "\n"
    "Roster Additions\n"
    "What it shows: A running list of approved new student enrollments (students in MAP not yet in SIS).\n"
    "What you do: Read-only for IMs. The data team processes new enrollments every Friday.\n"
    "\n"
    "Roster Unenrollments\n"
    "What it shows: A running list of approved student unenrollments (students no longer enrolled in MAP but still in SIS).\n"
    "What you do: Read-only for IMs. The data team processes unenrollments every Friday.\n"
    "\n"
    "Rejected Changes\n"
    "What it shows: A running list of rejected corrections with the date and a Reason for Rejection column.\n"
    "What you do: After rejecting a row on Sheet 1, go to this sheet and optionally add a reason in the last column.\n"
    "\n"
    "PART 3: DROPDOWN FILTERS AND SORTING\n"
    "\n"
    "Filter Dropdowns\n"
    "Each sheet has 5 filter dropdowns at the top:\n"
    "\u2022 Campus \u2014 Filter by school (e.g., AFMS, JHES, Reading CCSD)\n"
    "\u2022 Grade \u2014 Filter by grade level\n"
    "\u2022 Level \u2014 Filter by school level (Elementary, Middle, High School)\n"
    "\u2022 Student Group \u2014 Filter by student group assignment\n"
    "\u2022 Guide Email \u2014 Filter by teacher/guide email address\n"
    "\n"
    "How they work: Select a value to show only matching students. Select \u201cAll\u201d to remove "
    "that filter. Filters work together \u2014 selecting a Campus and a Grade shows only students "
    "matching both.\n"
    "\n"
    "Sort By Dropdown\n"
    "Each sheet also has a Sort By dropdown (to the right of the filters).\n"
    "\u2022 Choose any column name to sort the data by that field\n"
    "\u2022 Text fields sort A\u2192Z (ascending)\n"
    "\u2022 Default sort: Campus for Sheets 1 and 2, Date Approved for Sheets 3, 4, and 5\n"
    "\n"
    "Important: Changing a filter or sort option on Sheet 1 automatically clears all checkboxes, "
    "because the visible rows change. Always set your filters first, then check boxes.\n"
    "\n"
    "PART 4: WEEKLY WORKFLOW\n"
    "\n"
    "Monday through Thursday\n"
    "Implementation Managers review the Corrected Roster Info sheet. Filter by your campus, "
    "compare with the SIS sheet, and check off corrections that need to be made.\n"
    "\n"
    "Friday\n"
    "The data team reviews all three approval sheets (Automated Correction List, Roster Additions, "
    "Roster Unenrollments) and sends them to the support team to update the SIS.\n"
    "\n"
    "Automatic refresh\n"
    "The mismatch data refreshes automatically when the dashboard pipeline runs. New mismatches "
    "appear as student data changes in either the MAP roster or SIS.\n"
    "\n"
    "PART 5: UNENROLLING A STUDENT (NEW)\n"
    "\n"
    "If a student has left your campus and needs to be unenrolled from the SIS, you can now flag "
    "them directly from your campus\u2019s Student Roster spreadsheet.\n"
    "\n"
    "How it works\n"
    "\u2022 Open your campus\u2019s Student Roster spreadsheet (the one that feeds the MAP roster).\n"
    "\u2022 Find the student\u2019s row on the \u201cStudent Roster\u201d tab.\n"
    "\u2022 Check the \u201cUnenroll\u201d checkbox in the last column.\n"
    "\u2022 Within a minute, the checkmark propagates through your MAP Roster tab into the main MAP roster sheet.\n"
    "\u2022 The next time the weekly correction pipeline runs, the student will appear in the \u201cRoster Unenrollments\u201d sheet for the data team to process.\n"
    "\n"
    "Important rules\n"
    "\u2022 Leave the checkbox CHECKED after unenrollment \u2014 it\u2019s a permanent historical record.\n"
    "\u2022 Do not delete the row. Unchecking or deleting will NOT reverse an unenrollment that\u2019s already been processed.\n"
    "\u2022 If the student is both unenrolled by you AND has field mismatches (e.g. wrong grade), the Unenroll takes priority \u2014 we process the unenrollment, not the field changes.\n"
    "\n"
    "When your checkbox won\u2019t trigger an Unenrolling entry\n"
    "\u2022 If the SIS already has the student as not-enrolled, there\u2019s nothing to correct \u2014 the system stays quiet.\n"
    "\u2022 If you check Unenroll on a student that\u2019s still enrolled everywhere, expect them to show up in \u201cRoster Unenrollments\u201d on the next pipeline run.\n"
    "\n"
    "PART 6: FIELDS COMPARED\n"
    "\n"
    "The system compares these 10 fields between MAP roster and SIS:\n"
    "\u2022 Campus\n"
    "\u2022 Grade\n"
    "\u2022 Level\n"
    "\u2022 First Name\n"
    "\u2022 Last Name\n"
    "\u2022 Email\n"
    "\u2022 Student Group\n"
    "\u2022 Guide Name (Teacher first + last name)\n"
    "\u2022 Guide Email\n"
    "\u2022 External Student ID (SUNS Number for SC campuses)\n"
    "\n"
    "If any of these fields differ between the MAP roster and SIS, the student appears in the "
    "corrections sheet.\n"
    "\n"
    "PART 7: TROUBLESHOOTING\n"
    "\n"
    "\u201cI don\u2019t see my campus in the dropdown\u201d\n"
    "Your campus may not have any mismatched students this week. If all data matches between "
    "MAP roster and SIS, the campus won\u2019t appear in the filter options.\n"
    "\n"
    "\u201cCheckboxes disappeared after changing a filter\u201d\n"
    "This is expected behavior. Checkboxes reset when you change a dropdown because the visible "
    "students change. Set your filters first, then check boxes.\n"
    "\n"
    "\u201cThe data looks stale\u201d\n"
    "The corrections sheet refreshes when the data pipeline runs. Contact Khiem Doan if data "
    "appears outdated.\n"
    "\n"
    "\u201cMy approved correction isn\u2019t in the right sheet\u201d\n"
    "Corrections are automatically routed based on the Mismatch Summary type: field mismatches go "
    "to Automated Correction List, Roster Additions go to the Roster Additions sheet, and "
    "Unenrolling goes to the Roster Unenrollments sheet. Make sure you checked the actual checkbox "
    "(not just clicked the cell). Also verify the Apps Script is installed (Extensions > Apps Script "
    "should show the onEdit code).\n"
    "\n"
    "\u201cI see a student in Sheet 1 but not Sheet 2\u201d\n"
    "The student exists in the MAP roster but was not found in the SIS at all. Sheet 2 will "
    "show \u201cNOT FOUND IN SIS\u201d for that student. The Mismatch Summary will say \u201cRoster Addition\u201d "
    "(highlighted green).\n"
    "\n"
    "\u201cMy Unenroll checkbox isn\u2019t showing in the corrections list\u201d\n"
    "Three things to check: (1) IMPORTRANGE can take up to a minute to refresh \u2014 wait and reload. "
    "(2) Make sure the SIS still has the student as Enrolled \u2014 if it already matches, no correction "
    "is needed. (3) The pipeline only checks on runs \u2014 ask Khiem to run `python generate_corrections.py` "
    "or wait for the next scheduled run.\n"
)

requests.append({"insertText": {"location": {"index": 1}, "text": text}})
docs.documents().batchUpdate(documentId=DOC_ID, body={"requests": requests}).execute()
print("Text inserted.")

# ── Formatting ────────────────────────────────────────────────────────────
fmt = []

# Title (H1)
idx = 1
end = 1 + text.index("\n")
fmt.append(
    {
        "updateParagraphStyle": {
            "range": {"startIndex": idx, "endIndex": end},
            "paragraphStyle": {"namedStyleType": "HEADING_1"},
            "fields": "namedStyleType",
        }
    }
)

# Subtitle
sub = "User Guide \u2014 Implementation Manager Reference"
sub_start = 1 + text.index(sub)
sub_end = sub_start + len(sub)
fmt.append(
    {
        "updateParagraphStyle": {
            "range": {"startIndex": sub_start, "endIndex": sub_end},
            "paragraphStyle": {"namedStyleType": "SUBTITLE"},
            "fields": "namedStyleType",
        }
    }
)

# H2 headings (PART sections + Overview)
for title in [
    "Overview",
    "PART 1: HOW TO REVIEW CORRECTIONS",
    "PART 2: WHAT EACH SHEET DOES",
    "PART 3: DROPDOWN FILTERS AND SORTING",
    "PART 4: WEEKLY WORKFLOW",
    "PART 5: UNENROLLING A STUDENT (NEW)",
    "PART 6: FIELDS COMPARED",
    "PART 7: TROUBLESHOOTING",
]:
    i = text.find(title)
    if i >= 0:
        fmt.append(
            {
                "updateParagraphStyle": {
                    "range": {"startIndex": 1 + i, "endIndex": 1 + i + len(title)},
                    "paragraphStyle": {"namedStyleType": "HEADING_2"},
                    "fields": "namedStyleType",
                }
            }
        )

# H3 headings
for title in [
    "Step 1: Open the spreadsheet and set your filters",
    "Step 2: Compare the two sheets side by side",
    "Step 3: Accept or reject corrections",
    "Corrected Roster Info",
    "Current Roster Info in SIS",
    "Automated Correction List",
    "Roster Additions",
    "Roster Unenrollments",
    "Rejected Changes",
    "Filter Dropdowns",
    "Sort By Dropdown",
    "Monday through Thursday",
    "Friday",
    "Automatic refresh",
    "How it works",
    "Important rules",
    "When your checkbox won\u2019t trigger an Unenrolling entry",
]:
    # Find within appropriate section to avoid duplicates
    i = text.find(title + "\n")
    if i >= 0:
        fmt.append(
            {
                "updateParagraphStyle": {
                    "range": {"startIndex": 1 + i, "endIndex": 1 + i + len(title)},
                    "paragraphStyle": {"namedStyleType": "HEADING_3"},
                    "fields": "namedStyleType",
                }
            }
        )

# Bold labels (What it shows:, What you do:, etc.)
for label in [
    "What it shows:",
    "What you do:",
    "Mismatch Summary:",
    "Note:",
    "How they work:",
    "Tip:",
    "Important:",
    "Automatic refresh",
    "Default sort:",
]:
    start = 0
    while True:
        i = text.find(label, start)
        if i < 0:
            break
        fmt.append(
            {
                "updateTextStyle": {
                    "range": {"startIndex": 1 + i, "endIndex": 1 + i + len(label)},
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            }
        )
        start = i + len(label)

# Bold troubleshooting issue titles (quoted)
for issue in [
    "\u201cI don\u2019t see my campus in the dropdown\u201d",
    "\u201cCheckboxes disappeared after changing a filter\u201d",
    "\u201cThe data looks stale\u201d",
    "\u201cMy approved correction isn\u2019t in the right sheet\u201d",
    "\u201cI see a student in Sheet 1 but not Sheet 2\u201d",
    "\u201cMy Unenroll checkbox isn\u2019t showing in the corrections list\u201d",
]:
    i = text.find(issue)
    if i >= 0:
        fmt.append(
            {
                "updateTextStyle": {
                    "range": {"startIndex": 1 + i, "endIndex": 1 + i + len(issue)},
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            }
        )

docs.documents().batchUpdate(documentId=DOC_ID, body={"requests": fmt}).execute()
print("Formatting applied.")
print(f"Done! https://docs.google.com/document/d/{DOC_ID}/edit")

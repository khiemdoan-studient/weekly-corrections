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
    "The spreadsheet has three sheets:\n"
    "\u2022 Corrected Roster Info \u2014 What the data should be (from the MAP roster)\n"
    "\u2022 Current Roster Info in SIS \u2014 What the SIS currently has (for comparison)\n"
    "\u2022 Automated Correction List \u2014 Running history of approved corrections\n"
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
    "Step 3: Check off corrections\n"
    "If the MAP roster data (Sheet 1) is correct and should replace the SIS data (Sheet 2), "
    "check the checkbox in column A. The student automatically appears in the Automated "
    "Correction List (Sheet 3) with today\u2019s date.\n"
    "\n"
    "Important: Do not check boxes for students outside your campus. Only check corrections for "
    "schools you manage.\n"
    "\n"
    "PART 2: WHAT EACH SHEET DOES\n"
    "\n"
    "Corrected Roster Info\n"
    "What it shows: MAP roster data (source of truth) for students whose information doesn\u2019t "
    "match the SIS.\n"
    "What you do: Review the data, check the checkbox in column A if the correction is valid.\n"
    "Mismatch Summary: The last column lists which specific fields differ between MAP and SIS.\n"
    "\n"
    "Current Roster Info in SIS\n"
    "What it shows: What the SIS pipeline currently has for the same students.\n"
    "What you do: Compare this with Sheet 1 to see what\u2019s wrong. Do not edit this sheet.\n"
    "Note: Both sheets show the same students in the same row order for easy comparison.\n"
    "\n"
    "Automated Correction List\n"
    "What it shows: A running list of all approved corrections with the date each was checked off.\n"
    "What you do: This sheet is read-only for IMs. The data team uses it every Friday.\n"
    "Note: This list is cumulative \u2014 it keeps a full history and is never cleared automatically.\n"
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
    "\u2022 Default sort: Campus for Sheets 1 and 2, Date Approved for Sheet 3\n"
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
    "The data team copies the Automated Correction List (Sheet 3) and sends it to the support "
    "team to update the SIS. The corrections are then applied in the student information system.\n"
    "\n"
    "Automatic refresh\n"
    "The mismatch data refreshes automatically when the dashboard pipeline runs. New mismatches "
    "appear as student data changes in either the MAP roster or SIS.\n"
    "\n"
    "PART 5: FIELDS COMPARED\n"
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
    "PART 6: TROUBLESHOOTING\n"
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
    "\u201cMy approved correction isn\u2019t in Sheet 3\u201d\n"
    "Make sure you checked the actual checkbox (not just clicked the cell). The checkbox must "
    "show a checkmark. Also verify the Apps Script is installed (Extensions > Apps Script should "
    "show the onEdit code).\n"
    "\n"
    "\u201cI see a student in Sheet 1 but not Sheet 2\u201d\n"
    "The student exists in the MAP roster but was not found in the SIS at all. Sheet 2 will "
    "show \u201cNOT FOUND IN SIS\u201d for that student.\n"
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
    "PART 5: FIELDS COMPARED",
    "PART 6: TROUBLESHOOTING",
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
    "Step 3: Check off corrections",
    "Corrected Roster Info",
    "Current Roster Info in SIS",
    "Automated Correction List",
    "Filter Dropdowns",
    "Sort By Dropdown",
    "Monday through Thursday",
    "Friday",
    "Automatic refresh",
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
    "\u201cMy approved correction isn\u2019t in Sheet 3\u201d",
    "\u201cI see a student in Sheet 1 but not Sheet 2\u201d",
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

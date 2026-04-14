"""write_user_guide.py — Write formatted user guide to the Google Doc.

Target doc: https://docs.google.com/document/d/1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4
Pattern follows: email-automation/write_doc.py
"""

import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import SERVICE_ACCOUNT_KEY

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_KEY,
    scopes=["https://www.googleapis.com/auth/documents"],
)
docs = build("docs", "v1", credentials=creds)
DOC_ID = "1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4"

# ── Clear existing content ─────────────────────────────────────────────────
doc = docs.documents().get(documentId=DOC_ID).execute()
body_content = doc.get("body", {}).get("content", [])
end_index = body_content[-1]["endIndex"] if body_content else 1
requests = []
if end_index > 2:
    requests.append(
        {"deleteContentRange": {"range": {"startIndex": 1, "endIndex": end_index - 1}}}
    )

# ── Document text ──────────────────────────────────────────────────────────
text = (
    "Automated Weekly Corrections\n"
    "User Guide\n"
    "\n"
    "What Is This?\n"
    "This spreadsheet compares student enrollment data between two sources: "
    "the MAP roster (our source of truth) and the SIS pipeline (the system of record). "
    "When a student\u2019s information doesn\u2019t match between the two, it shows up here "
    "so Implementation Managers can review and approve corrections.\n"
    "\n"
    "How to Review Corrections (3 Steps)\n"
    "\n"
    "Step 1: Open the spreadsheet\n"
    "Open the Automated Weekly Corrections spreadsheet. You\u2019ll see three sheets "
    "at the bottom: Corrected Roster Info, Current Roster Info in SIS, and "
    "Automated Correction List.\n"
    "\n"
    "Step 2: Use dropdown filters to find your campus\n"
    "At the top of each sheet, there are dropdown filters for Campus, Grade, Level, "
    "Student Group, and Guide Email. Select your campus from the Campus dropdown to "
    "see only students at your school. Use \u201cAll\u201d to see all campuses.\n"
    "\n"
    "Step 3: Check off corrections\n"
    "On the Corrected Roster Info sheet, compare each student\u2019s data with the "
    "Current Roster Info in SIS sheet (same students, same order). "
    "The \u201cMismatch Summary\u201d column (last column) tells you which fields are different. "
    "If the MAP roster data is correct and should replace the SIS data, check the "
    "checkbox in column A. The student will automatically appear in the "
    "Automated Correction List with today\u2019s date.\n"
    "\n"
    "Important: When you change a dropdown filter, all checkboxes are automatically "
    "cleared because the filtered data changes. Check boxes after setting your filters.\n"
    "\n"
    "What Each Sheet Does\n"
    "\n"
    "Corrected Roster Info \u2014 Shows the MAP roster data (source of truth) for students "
    "whose information doesn\u2019t match the SIS. This is the \u201ccorrect\u201d version. "
    "Use the checkbox in column A to approve corrections.\n"
    "\n"
    "Current Roster Info in SIS \u2014 Shows what the SIS pipeline currently has for the "
    "same students. Compare this side-by-side with Sheet 1 to see what\u2019s wrong.\n"
    "\n"
    "Automated Correction List \u2014 A running list of all approved corrections with "
    "the date each was checked off. This list is sent to the support team every Friday. "
    "This sheet is never cleared by the system \u2014 it keeps a full history.\n"
    "\n"
    "Dropdown Filters\n"
    "\n"
    "Each sheet has 5 dropdown filters at the top:\n"
    "\u2022 Campus \u2014 Filter by school (e.g., AFMS, JHES, Reading CCSD)\n"
    "\u2022 Grade \u2014 Filter by grade level\n"
    "\u2022 Level \u2014 Filter by school level (Elementary, Middle, High School)\n"
    "\u2022 Student Group \u2014 Filter by student group assignment\n"
    "\u2022 Guide Email \u2014 Filter by teacher/guide email\n"
    "\n"
    "Select \u201cAll\u201d in any dropdown to show all values for that field. "
    "Filters work together \u2014 if you select a Campus and a Grade, you\u2019ll only see "
    "students matching both.\n"
    "\n"
    "Weekly Workflow\n"
    "\n"
    "Monday through Thursday: Implementation Managers review the Corrected Roster Info "
    "sheet, filter by their campus, and check off corrections that need to be made.\n"
    "\n"
    "Friday: The data team copies the Automated Correction List (Sheet 3) and sends it "
    "to the support team to update the SIS. The corrections are then applied in the system.\n"
    "\n"
    "The data refreshes automatically when the main dashboard pipeline runs, so new "
    "mismatches will appear as student data changes.\n"
    "\n"
    "Fields Compared\n"
    "\n"
    "The system compares these 10 fields between MAP roster and SIS:\n"
    "\u2022 Campus\n"
    "\u2022 Grade\n"
    "\u2022 Level\n"
    "\u2022 First Name\n"
    "\u2022 Last Name\n"
    "\u2022 Email\n"
    "\u2022 Student Group\n"
    "\u2022 Guide Name (Teacher)\n"
    "\u2022 Guide Email\n"
    "\u2022 External Student ID (SUNS Number)\n"
    "\n"
    "If any of these fields differ, the student appears in the corrections sheet.\n"
    "\n"
    "Troubleshooting\n"
    "\n"
    "\u201cI don\u2019t see my campus in the dropdown\u201d \u2014 Your campus may not have any "
    "mismatched students this week. If all data matches, the campus won\u2019t appear.\n"
    "\n"
    "\u201cCheckboxes disappeared after changing a filter\u201d \u2014 This is expected. "
    "Checkboxes reset when you change a dropdown because the visible students change. "
    "Set your filters first, then check boxes.\n"
    "\n"
    "\u201cThe data looks stale\u201d \u2014 The corrections sheet refreshes when the data pipeline "
    "runs. Contact the data team if data appears outdated.\n"
    "\n"
    "\u201cMy approved correction isn\u2019t in Sheet 3\u201d \u2014 Make sure you checked the box "
    "(not just clicked the cell). The checkbox must show a checkmark.\n"
)

requests.append({"insertText": {"location": {"index": 1}, "text": text}})
docs.documents().batchUpdate(documentId=DOC_ID, body={"requests": requests}).execute()
print("Text inserted.")

# ── Formatting ─────────────────────────────────────────────────────────────
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
sub_start = end + 1
sub_end = sub_start + len("User Guide")
fmt.append(
    {
        "updateParagraphStyle": {
            "range": {"startIndex": sub_start, "endIndex": sub_end},
            "paragraphStyle": {"namedStyleType": "SUBTITLE"},
            "fields": "namedStyleType",
        }
    }
)

# H2 headings
for title in [
    "What Is This?",
    "How to Review Corrections (3 Steps)",
    "What Each Sheet Does",
    "Dropdown Filters",
    "Weekly Workflow",
    "Fields Compared",
    "Troubleshooting",
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
    "Step 1: Open the spreadsheet",
    "Step 2: Use dropdown filters to find your campus",
    "Step 3: Check off corrections",
]:
    i = text.find(title)
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

# Bold important text
for bold_text in [
    "Important: When you change a dropdown filter, all checkboxes are automatically "
    "cleared because the filtered data changes. Check boxes after setting your filters.",
    "Corrected Roster Info",
    "Current Roster Info in SIS",
    "Automated Correction List",
]:
    # Find within "What Each Sheet Does" section for tab names
    i = text.find(bold_text)
    if i >= 0:
        fmt.append(
            {
                "updateTextStyle": {
                    "range": {"startIndex": 1 + i, "endIndex": 1 + i + len(bold_text)},
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            }
        )

# Bold sheet tab names at the start of their descriptions
for tab in [
    "Corrected Roster Info \u2014",
    "Current Roster Info in SIS \u2014",
    "Automated Correction List \u2014",
]:
    start = text.find("What Each Sheet Does")
    i = text.find(tab, start)
    if i >= 0:
        tab_name = tab.split(" \u2014")[0]
        fmt.append(
            {
                "updateTextStyle": {
                    "range": {"startIndex": 1 + i, "endIndex": 1 + i + len(tab_name)},
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            }
        )

docs.documents().batchUpdate(documentId=DOC_ID, body={"requests": fmt}).execute()
print("Formatting applied.")
print(f"Done! https://docs.google.com/document/d/{DOC_ID}/edit")

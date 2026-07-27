"""
File: generate_migration_report.py

Author:
    Lin, Hua
    M365 Copilot

Purpose:
    Scan exported VBA modules and generate
    a WPS migration report.

Output:
    output/migration_report.txt
"""

import os
from collections import defaultdict

# -------------------------------------------------
# Paths
# -------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

VBA_FOLDER = os.path.join(
    SCRIPT_DIR,
    "..",
    "vba_export"
)

OUTPUT_FILE = os.path.join(
    SCRIPT_DIR,
    "..",
    "output",
    "migration_report.txt"
)

# -------------------------------------------------
# Rules
# -------------------------------------------------

RULES = {

    "UserForm": {
        "severity": "HIGH",
        "migration":
        "Convert UserForm into worksheet form or Python Tkinter GUI."
    },

    "CreateObject": {
        "severity": "HIGH",
        "migration":
        "Replace COM automation with Python module."
    },

    "GetObject": {
        "severity": "HIGH",
        "migration":
        "Replace COM automation with Python module."
    },

    "MSForms": {
        "severity": "HIGH",
        "migration":
        "Replace MSForms controls."
    },

    "ActiveX": {
        "severity": "HIGH",
        "migration":
        "Replace ActiveX controls with worksheet controls."
    },

    "Workbook_Open": {
        "severity": "HIGH",
        "migration":
        "Move startup logic into Python launcher."
    },

    "Application.OnTime": {
        "severity": "HIGH",
        "migration":
        "Replace scheduler with Python timer."
    },

    "FileDialog": {
        "severity": "MEDIUM",
        "migration":
        "Replace with tkinter.filedialog."
    },

    "Shell": {
        "severity": "MEDIUM",
        "migration":
        "Replace with subprocess module."
    },

    "PrintOut": {
        "severity": "MEDIUM",
        "migration":
        "Replace with ReportLab PDF generation."
    },

    "ExportAsFixedFormat": {
        "severity": "MEDIUM",
        "migration":
        "Generate PDF using Python."
    },

    "Worksheet_Change": {
        "severity": "LOW",
        "migration":
        "Convert to validation routine."
    },

    "SelectionChange": {
        "severity": "LOW",
        "migration":
        "Convert to worksheet formulas if possible."
    }
}


# -------------------------------------------------
# Scan VBA modules
# -------------------------------------------------

report_lines = []

report_lines.append("=" * 80)
report_lines.append("VBA TO WPS MIGRATION REPORT")
report_lines.append("=" * 80)
report_lines.append("")

summary = defaultdict(int)

files_processed = 0

for file_name in os.listdir(VBA_FOLDER):

    if not file_name.lower().endswith(
        (".bas", ".frm", ".cls")
    ):
        continue

    files_processed += 1

    full_path = os.path.join(
        VBA_FOLDER,
        file_name
    )

    with open(
        full_path,
        "r",
        encoding="utf-8",
        errors="ignore"
    ) as f:

        text = f.read()

    report_lines.append("")
    report_lines.append("-" * 80)
    report_lines.append(f"FILE : {file_name}")
    report_lines.append("-" * 80)

    found = False

    for keyword, info in RULES.items():

        count = text.lower().count(
            keyword.lower()
        )

        if count > 0:

            found = True

            summary[keyword] += count

            report_lines.append(
                f"\nKeyword : {keyword}"
            )

            report_lines.append(
                f"Occurrences : {count}"
            )

            report_lines.append(
                f"Severity : {info['severity']}"
            )

            report_lines.append(
                f"Migration : {info['migration']}"
            )

    if not found:

        report_lines.append(
            "No obvious WPS migration issues found."
        )

# -------------------------------------------------
# Summary
# -------------------------------------------------

report_lines.append("\n")
report_lines.append("=" * 80)
report_lines.append("SUMMARY")
report_lines.append("=" * 80)

report_lines.append(
    f"\nModules Processed : {files_processed}"
)

for keyword, count in sorted(
        summary.items()):

    report_lines.append(
        f"{keyword:<25} {count}"
    )

# -------------------------------------------------
# Recommendations
# -------------------------------------------------

report_lines.append("\n")
report_lines.append("=" * 80)
report_lines.append("RECOMMENDED NEXT STEPS")
report_lines.append("=" * 80)

report_lines.append(
    "\n1. Remove UserForms."
)

report_lines.append(
    "2. Replace ActiveX controls."
)

report_lines.append(
    "3. Convert CheckMe routines to Python."
)

report_lines.append(
    "4. Convert Export/Print PDF logic to Python."
)

report_lines.append(
    "5. Load CL_InfoDB.csv using pandas."
)

report_lines.append(
    "6. Save final workbook as XLSX."
)

# -------------------------------------------------
# Save Report
# -------------------------------------------------

os.makedirs(
    os.path.dirname(OUTPUT_FILE),
    exist_ok=True
)

with open(
    OUTPUT_FILE,
    "w",
    encoding="utf-8"
) as f:

    f.write(
        "\n".join(report_lines)
    )

print("")
print("=" * 80)
print("Migration report generated")
print(OUTPUT_FILE)
print("=" * 80)
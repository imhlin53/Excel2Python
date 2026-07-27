"""
scan_vba.py

Purpose:
    Scan all exported VBA files and identify Excel features
    that are not well supported by WPS Office.
"""

import os

VBA_FOLDER = r"../vba_export"

KEYWORDS = [
    "UserForm",
    "CreateObject",
    "GetObject",
    "Shell",
    "FileDialog",
    "ActiveX",
    "MSForms",
    "Workbook_Open",
    "Application.OnTime"
]

for file_name in os.listdir(VBA_FOLDER):

    full_path = os.path.join(VBA_FOLDER, file_name)

    if not file_name.lower().endswith(
        (".bas", ".cls", ".frm")
    ):
        continue

    with open(
        full_path,
        "r",
        encoding="utf-8",
        errors="ignore"
    ) as f:

        text = f.read()

    for keyword in KEYWORDS:

        if keyword.lower() in text.lower():

            print(
                f"[FOUND] {keyword:<20} "
                f"in {file_name}"
            )
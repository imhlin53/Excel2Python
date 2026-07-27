"""
extract_named_ranges.py
"""

import openpyxl
import csv
from pathlib import Path

SOURCE_FOLDER = Path("../source")

excel_files = list(
    SOURCE_FOLDER.glob("*.xlsm")
)

if not excel_files:
    raise Exception(
        "No XLSM workbook found in source folder"
    )

WORKBOOK = excel_files[0]

print(f"Workbook = {WORKBOOK}")

wb = openpyxl.load_workbook(
    WORKBOOK,
    keep_vba=True,
    data_only=False
)

OUTPUT = "../output/named_ranges.csv"

with open(
    OUTPUT,
    "w",
    newline="",
    encoding="utf-8"
) as f:

    writer = csv.writer(f)

    writer.writerow([
        "Name",
        "Reference"
    ])

    for name in wb.defined_names.values():

        writer.writerow([
            name.name,
            name.attr_text
        ])

print()
print("Done")
print(OUTPUT)
import os
import re
import csv

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
    "vba_function_inventory_v2.csv"
)

rows = []

pattern = re.compile(
    r'^\s*(?:Public|Private|Friend)?\s*'
    r'(Sub|Function)\s+'
    r'([A-Za-z0-9_]+)',
    re.IGNORECASE
)

for file_name in os.listdir(VBA_FOLDER):

    if not file_name.lower().endswith(
        (".bas", ".cls", ".frm")
    ):
        continue

    full_path = os.path.join(
        VBA_FOLDER,
        file_name
    )

    with open(
        full_path,
        encoding="utf-8",
        errors="ignore"
    ) as f:

        for line_number, line in enumerate(f, start=1):

            match = pattern.search(line)

            if match:

                rows.append([
                    file_name,
                    line_number,
                    match.group(1),
                    match.group(2)
                ])

with open(
    OUTPUT_FILE,
    "w",
    newline="",
    encoding="utf-8"
) as csvfile:

    writer = csv.writer(csvfile)

    writer.writerow([
        "File",
        "Line",
        "Type",
        "Procedure"
    ])

    writer.writerows(rows)

print()
print("Functions found:", len(rows))
print("Output:", OUTPUT_FILE)
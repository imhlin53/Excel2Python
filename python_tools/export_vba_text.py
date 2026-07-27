# export_vba_text.py

import os

source_dir = r"..\vba_export"

with open(
    r"..\output\all_vba.txt",
    "w",
    encoding="utf-8"
) as out:

    for f in os.listdir(source_dir):

        if not f.lower().endswith(
            (".bas", ".frm", ".cls")
        ):
            continue

        out.write("\n")
        out.write("=" * 80)
        out.write("\n")
        out.write(f"FILE: {f}\n")
        out.write("=" * 80)
        out.write("\n\n")

        with open(
            os.path.join(source_dir, f),
            encoding="utf-8",
            errors="ignore"
        ) as inp:

            out.write(inp.read())
            out.write("\n\n")
# extract_vba.py

from oletools.olevba import VBA_Parser

file_path = r"source\H23-VOBES_FB_Komponenten_Epsilon_ASV_702J_20260714.xls"

vbaparser = VBA_Parser(file_path)

for (_, _, module_name, code) in vbaparser.extract_macros():

    out_file = f"vba_export/{module_name}.bas"

    with open(out_file, "w", encoding="utf-8", errors="ignore") as f:
        f.write(code)

    print(module_name)
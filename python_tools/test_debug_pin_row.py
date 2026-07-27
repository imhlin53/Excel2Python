from pathlib import Path
from vobes_workbook import VobesWorkbook

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

wb = VobesWorkbook(xlsm)

wb.debug_pin_row(17)

wb.debug_pin_row(20)
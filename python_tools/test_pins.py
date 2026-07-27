from pathlib import Path
from vobes_workbook import VobesWorkbook

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

wb = VobesWorkbook(xlsm)

print()

for pin in wb.get_pins():

    print(pin)
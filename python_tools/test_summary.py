from pathlib import Path
from vobes_workbook import VobesWorkbook

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

wb = VobesWorkbook(xlsm)

wb.print_summary()

print()

print("Slots")
print("-----")

for slot in wb.get_slots():

    print(slot)

print()

print("Pins")
print("----")

for pin in wb.get_pins():

    print(pin)
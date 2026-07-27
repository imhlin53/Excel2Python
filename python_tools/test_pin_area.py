from pathlib import Path
from vobes_workbook import VobesWorkbook

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

wb = VobesWorkbook(xlsm)

pin_start, pin_end = wb.get_pin_bounds()

slot_start, slot_end = wb.get_slot_bounds()

print()
print("Pin Area")
print("----------------")
print(pin_start, pin_end)

print()

print("Slot Area")
print("----------------")
print(slot_start, slot_end)
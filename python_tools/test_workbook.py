from pathlib import Path
from vobes_workbook import VobesWorkbook

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

wb = VobesWorkbook(xlsm)

print()
print("Part Number :", wb.get_part_number())
print("Pin Count   :", wb.get_pin_count())
print("Slot Count  :", wb.get_slot_count())
print("State       :", wb.get_status())
print("Checked     :", wb.get_data_checked())
from pathlib import Path
from vobes_workbook import VobesWorkbook

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

wb = VobesWorkbook(xlsm)

names = [
    "Teilenummer",
    "SumPins",
    "SumSlots",
    "State",
    "DataChecked",
    "StartPin",
    "EndPin",
    "GER"
]

for name in names:

    try:

       print(
            f"{name:15}",
            wb.get_named_value(name)
        )

    except Exception as e:

        print(name,  e)

from pathlib import Path

from vobes_service import VobesService
from vobes_pin_manager import VobesPinManager


xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

csv_file = Path(
    "../source/CL_InfoDB.csv"
)

svc = VobesService(
    xlsm,
    csv_file
)

pm = VobesPinManager(
    svc
)

pm.print_pin_summary()

print()
print("CheckMe")
print("--------")

issues = pm.check_me()

if not issues:

    print("No issues found")

else:

    for issue in issues:

        print(issue)

print()
print("Search Description: KL30")
print("------------------------")

for pin in pm.find_pins_by_description("KL30"):

    print(pin)
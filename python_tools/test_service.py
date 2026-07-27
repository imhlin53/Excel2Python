from pathlib import Path

from vobes_service import VobesService

xlsm = list(
    Path("../source").glob("*.xlsm")
)[0]

csv_file = Path(
    "../source/CL_InfoDB.csv")

svc = VobesService(
    xlsm,
    csv_file
)

svc.print_summary()

print()
print("Pins")
print("----")
for pin in svc.get_all_pins():

   print(pin)

print()
print("Searc* CL Clamp: 30")
print("-----------*-------")

results = svc.search_clamp("30")

for row in results[:5]:
    print(row)
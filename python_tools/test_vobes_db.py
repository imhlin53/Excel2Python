from pathlib import Path

from vobes_db import VobesDB

csv_file = Path(
    "../source/CL_InfoDB.csv"
)

db = VobesDB()

db.load_csv(csv_file)

db.print_summary()

print()
print("Functions containing CAN")
print("------------------------")

matches = db.search_function(
    "CAN"
)

for row in matches[:10]:

    print(row)
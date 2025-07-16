import smartsheet
import csv
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

# ————————————————
# Setup und Config
# ————————————————
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
client = smartsheet.Smartsheet(token)

sheet_ids = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

phase_fields = [
    ("Kontrolle",      "K von",        1),
    ("BE am",          "BE von",       2),
    ("K am",           "K2 von",       3),
    ("C am",           "C von",        4),
    ("Reopen C2 am",   "Reopen C2 von",5),
]

# ————————————————
# Zeitfenster letzte 7 Tage
# ————————————————
today  = datetime.utcnow().date()
cutoff = today - timedelta(days=7)

# ————————————————
# Speicherort: weekly/weekly_changes.csv (immer überschreiben)
# ————————————————
weekly_dir  = "weekly"
os.makedirs(weekly_dir, exist_ok=True)
weekly_file = os.path.join(weekly_dir, "weekly_changes.csv")

# ————————————————
# CSV schreiben
# ————————————————
with open(weekly_file, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Erzeugt am","Produktgruppe","Phase","Datumsfeld","Datum","Mitarbeiter"])

    for group, sid in sheet_ids.items():
        sheet = client.Sheets.get_sheet(sid)
        col_map = {col.title: col.id for col in sheet.columns}

        for row in sheet.rows:
            for date_col, user_col, phase_no in phase_fields:
                date_val = None
                user_val = ""
                for cell in row.cells:
                    if cell.column_id == col_map.get(date_col):
                        date_val = cell.value
                    if cell.column_id == col_map.get(user_col):
                        user_val = cell.display_value or ""
                if date_val:
                    try:
                        dt = datetime.fromisoformat(date_val).date()
                    except ValueError:
                        continue
                    if dt >= cutoff:
                        w.writerow([
                            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                            group,
                            phase_no,
                            date_col,
                            dt.isoformat(),
                            user_val
                        ])

print(f"✅ Weekly report overwritten at {weekly_file}")

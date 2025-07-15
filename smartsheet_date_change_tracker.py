import smartsheet
import csv
import os
from datetime import datetime
from dotenv import load_dotenv

# --- Setup ---
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
sm = smartsheet.Smartsheet(token)

# --- Konfiguration ---
sheet_ids = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# Mapping: Datumsfeld → Phase‑Nummer und Mitarbeiter‑Feld
phase_fields = [
    ("Kontrolle",      "K von",        1),
    ("BE am",          "BE von",       2),
    ("K am",           "K2 von",       3),
    ("C am",           "C von",        4),
    ("Reopen C2 am",   "Reopen C2 von",5),
]

# --- Backup laden ---
backup_file = "date_backup.csv"
try:
    with open(backup_file, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        headers = next(reader)
        prev = {row[0]: row for row in reader}
except FileNotFoundError:
    prev = {}
    headers = []

# Neues Log sammeln
now = datetime.utcnow()
month_year = now.strftime("%m.%Y")
log_file = f"date_changes_log_{month_year}.csv"
timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

# Sammle alle neuen Änderungen
changes = []
for group, sid in sheet_ids.items():
    sheet = sm.Sheets.get_sheet(sid)
    for row in sheet.rows:
        key = f"{group}:{row.id}"
        for date_col, user_col, phase_no in phase_fields:
            date_val = ""
            user_val = ""
            for cell in row.cells:
                # identify column titles
                for col in sheet.columns:
                    if col.id == cell.column_id:
                        if col.title == date_col:
                            date_val = cell.display_value or ""
                        if col.title == user_col:
                            user_val = cell.display_value or ""
            if date_val:
                prev_row = prev.get(key, [])
                if date_val not in prev_row:
                    changes.append([
                        timestamp,
                        group,
                        phase_no,
                        date_col,
                        date_val,
                        user_val
                    ])
                    prev.setdefault(key, []).append(date_val)

# Logfile schreiben
if changes:
    write_header = not os.path.exists(log_file)
    with open(log_file, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(["Änderung am","Produktgruppe","Phase","Feld","Datum","Mitarbeiter"])
        w.writerows(changes)
    print(f"✅ {len(changes)} Änderung(en) protokolliert in {log_file}")
else:
    print("🔍 Keine Änderungen gefunden.")

# Backup aktualisieren
with open(backup_file, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["RowKey","LoggedDates..."])
    for k, vals in prev.items():
        w.writerow([k] + vals)

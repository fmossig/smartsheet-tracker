import smartsheet
import csv
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

# → Token laden
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
client = smartsheet.Smartsheet(token)

# → Sheet‑IDs
sheet_ids = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# → Zu überwachende Felder (Datumsfeld, Mitarbeiterspalte, Phase‑Nummer)
phase_fields = [
    ("Kontrolle",      "K von",        1),
    ("BE am",          "BE von",       2),
    ("K am",           "K2 von",       3),
    ("C am",           "C von",        4),
    ("Reopen C2 am",   "Reopen C2 von",5),
]

# → heute und Cutoffs
today  = datetime.utcnow().date()
cutoff30  = today - timedelta(days=30)

# → Output‑Ordner anlegen
status_dir = "status"
os.makedirs(status_dir, exist_ok=True)

# → Dateinamen in diesem Ordner
date_str      = today.isoformat()
snapshot_file = os.path.join(status_dir, f"status_snapshot_{date_str}.csv")
summary_file  = os.path.join(status_dir, f"status_summary_{date_str}.csv")

# --- 1) Detaillierter 30‑Tage‑Snapshot ---
with open(snapshot_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Produktgruppe","RowID","Phase","Datumsfeld","Datum","Mitarbeiter"])
    for group, sid in sheet_ids.items():
        sheet = client.Sheets.get_sheet(sid)
        col_map = {col.title: col.id for col in sheet.columns}
        for row in sheet.rows:
            for date_col, user_col, phase_no in phase_fields:
                val = next((cell.value for cell in row.cells if cell.column_id == col_map.get(date_col)), None)
                user = next((cell.display_value for cell in row.cells if cell.column_id == col_map.get(user_col)), "")
                if not val:
                    continue
                try:
                    dt = datetime.fromisoformat(val).date()
                except ValueError:
                    continue
                if dt >= cutoff30:
                    writer.writerow([group, row.id, phase_no, date_col, dt.isoformat(), user])

print(f"✅ 30‑Tage Snapshot geschrieben: {snapshot_file}")

# --- 2) Perioden‑Vergleich über alle Phasen-Events ---
# Wir sammeln alle Events (unabhängig vom 30‑Tage‑Filter) und jeder Row das jüngste Datum
cutoff60 = today - timedelta(days=60)
cutoff90 = today - timedelta(days=90)

events = []
latest = {}  # (Gruppe,RowID) → jüngstes Datum

for group, sid in sheet_ids.items():
    sheet = client.Sheets.get_sheet(sid)
    col_map = {col.title: col.id for col in sheet.columns}
    for row in sheet.rows:
        key = (group, row.id)
        for date_col, user_col, phase_no in phase_fields:
            val = next((cell.value for cell in row.cells if cell.column_id == col_map.get(date_col)), None)
            if not val:
                continue
            try:
                dt = datetime.fromisoformat(val).date()
            except ValueError:
                continue
            events.append(dt)
            prev = latest.get(key)
            if prev is None or dt > prev:
                latest[key] = dt

# Zähle Phase‑Events je Zeitraum
cnt_evt = {"0-30":0, "31-60":0, "61-90":0}
for dt in events:
    if   dt >= cutoff30:  cnt_evt["0-30"]  += 1
    elif dt >= cutoff60:  cnt_evt["31-60"] += 1
    elif dt >= cutoff90:  cnt_evt["61-90"] += 1

# Zähle Rows nach jüngster Phase
cnt_row = {"0-30":0, "31-60":0, "61-90":0, ">90":0}
for dt in latest.values():
    if   dt >= cutoff30:  cnt_row["0-30"]  += 1
    elif dt >= cutoff60:  cnt_row["31-60"] += 1
    elif dt >= cutoff90:  cnt_row["61-90"] += 1
    else:                  cnt_row[">90"]   += 1

# Summary‑CSV schreiben
with open(summary_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Metrik","0-30 Tage","31-60 Tage","61-90 Tage",">90 Tage"])
    writer.writerow([
        "Anzahl Phase‑Events",
        cnt_evt["0-30"], cnt_evt["31-60"], cnt_evt["61-90"], ""
    ])
    writer.writerow([
        "Distinct Rows mit jüngster Phase",
        cnt_row["0-30"], cnt_row["31-60"], cnt_row["61-90"], cnt_row[">90"]
    ])

print(f"✅ Perioden‑Vergleich geschrieben: {summary_file}")

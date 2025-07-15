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

# → Zeitfenster: letzte 30 Tage
today = datetime.utcnow().date()
cutoff = today - timedelta(days=30)

# → Output‑Dateiname
snapshot_file = f"status_snapshot_{today.isoformat()}.csv"

# → CSV schreiben
with open(snapshot_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Produktgruppe", "RowID", "Phase", "Datumsfeld", "Datum", "Mitarbeiter"])

    for group, sid in sheet_ids.items():
        sheet = client.Sheets.get_sheet(sid)

        # Spalten‑Mapping (Titel → ID)
        col_map = {col.title: col.id for col in sheet.columns}

        for row in sheet.rows:
            for date_col, user_col, phase_no in phase_fields:
                # Werte suchen
                date_val = None
                user_val = ""
                for cell in row.cells:
                    if cell.column_id == col_map.get(date_col):
                        date_val = cell.value  # ISO‑String z.B. "2025-07-15"
                    if cell.column_id == col_map.get(user_col):
                        user_val = cell.display_value or ""
                # Falls Datum gesetzt und im Fenster
                if date_val:
                    try:
                        dt = datetime.fromisoformat(date_val).date()
                    except:
                        continue
                    if dt >= cutoff:
                        writer.writerow([group, row.id, phase_no, date_col, dt.isoformat(), user_val])

print(f"✅ Snapshot geschrieben in {snapshot_file}")

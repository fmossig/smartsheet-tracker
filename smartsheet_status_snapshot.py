# smartsheet_status_snapshot.py

import smartsheet
import csv
import os
from datetime import datetime
from dotenv import load_dotenv

# Lade Token
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
client = smartsheet.Smartsheet(token)

# Deine Sheet‑IDs
sheet_ids = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# Felder, die wir ausgeben wollen
fields = [
    ("Kontrolle",    "K von"),
    ("BE am",        "BE von"),
    ("K am",         "K2 von"),
    ("C am",         "C von"),
    ("Reopen C2 am", "Reopen C2 von"),
]

# Dateiname mit Datum (YYYY‑MM‑DD)
today = datetime.utcnow().strftime("%Y-%m-%d")
filename = f"status_snapshot_{today}.csv"

# Header
headers = ["Produktgruppe", "RowID"]
for date_col, emp_col in fields:
    headers += [date_col, emp_col]

# Schreibe CSV
with open(filename, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(headers)

    for group, sid in sheet_ids.items():
        sheet = client.Sheets.get_sheet(sid)
        for row in sheet.rows:
            # Zellen in dict packen
            row_vals = {col.title: "" for col in sheet.columns}
            for cell in row.cells:
                # Spaltenbezeichnung finden
                for col in sheet.columns:
                    if col.id == cell.column_id:
                        row_vals[col.title] = cell.display_value or ""
                        break

            # Eine Zeile ergeben
            line = [group, row.id]
            for date_col, emp_col in fields:
                line.append(row_vals.get(date_col, ""))
                line.append(row_vals.get(emp_col, ""))
            writer.writerow(line)

print(f"✅ Snapshot geschrieben als {filename}")

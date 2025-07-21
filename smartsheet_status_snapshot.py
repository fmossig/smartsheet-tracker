import smartsheet, csv, os
from datetime import datetime, timedelta
from dotenv import load_dotenv

# — Setup —
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
    ("Kontrolle","K von",1),
    ("BE am","BE von",2),
    ("K am","K2 von",3),
    ("C am","C von",4),
    ("Reopen C2 am","Reopen C2 von",5),
]

today = datetime.utcnow().date()
cut30 = today - timedelta(days=30)
cut60 = today - timedelta(days=60)
cut90 = today - timedelta(days=90)

# Stelle Output‑Ordner sicher
os.makedirs("status", exist_ok=True)
date_str = today.isoformat()
file_snap  = f"status/status_snapshot_{date_str}.csv"
file_sum   = f"status/status_summary_{date_str}.csv"

# 1) Snapshot (0–30 Tage)
with open(file_snap, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Produktgruppe","RowID","Phase","Datumsfeld","Datum","Mitarbeiter"])
    for grp,sid in sheet_ids.items():
        sheet = client.Sheets.get_sheet(sid)
        cols = {c.title:c.id for c in sheet.columns}
        for row in sheet.rows:
            for dcol,ucol,ph in phase_fields:
                val = next((c.value for c in row.cells if c.column_id==cols[dcol]), None)
                usr = next((c.display_value for c in row.cells if c.column_id==cols[ucol]), "")
                if not val: continue
                try:
                    dt = datetime.fromisoformat(val).date()
                except ValueError:
                    continue
                if dt>=cut30:
                    w.writerow([grp,row.id,ph,dcol,dt.isoformat(),usr])

# 2) Summary über alle Events
events = []
latest = {}
for grp,sid in sheet_ids.items():
    sheet = client.Sheets.get_sheet(sid)
    cols = {c.title:c.id for c in sheet.columns}
    for row in sheet.rows:
        key = (grp,row.id)
        for dcol,_,_ in phase_fields:
            val = next((c.value for c in row.cells if c.column_id==cols[dcol]), None)
            if not val: continue
            try:
                dt = datetime.fromisoformat(val).date()
            except ValueError:
                continue
            events.append(dt)
            if key not in latest or dt>latest[key]:
                latest[key]=dt

cnt_evt = {"0-30":0,"31-60":0,"61-90":0}
for dt in events:
    if   dt>=cut30:  cnt_evt["0-30"]+=1
    elif dt>=cut60:  cnt_evt["31-60"]+=1
    elif dt>=cut90:  cnt_evt["61-90"]+=1

cnt_row={"0-30":0,"31-60":0,"61-90":0,">90":0}
for dt in latest.values():
    if   dt>=cut30:  cnt_row["0-30"]+=1
    elif dt>=cut60:  cnt_row["31-60"]+=1
    elif dt>=cut90:  cnt_row["61-90"]+=1
    else:             cnt_row[">90"]+=1

with open(file_sum,"w",newline="",encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Metrik","0-30 Tage","31-60 Tage","61-90 Tage",">90 Tage"])
    w.writerow(["Anzahl Phase‑Events",
                cnt_evt["0-30"],cnt_evt["31-60"],cnt_evt["61-90"],""])
    w.writerow(["Distinct Rows mit jüngster Phase",
                cnt_row["0-30"],cnt_row["31-60"],cnt_row["61-90"],cnt_row[">90"]])

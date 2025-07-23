import os
import csv
import argparse
from datetime import datetime, timedelta
from collections import defaultdict

import smartsheet
from dotenv import load_dotenv

# -------------------- KONFIG --------------------
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}
SEED_SHEET_ID = 6879355327172484  # einmalige Bootstrap-Quelle

PHASE_FIELDS = [
    ("Kontrolle",    "K von",        1),
    ("BE am",        "BE von",       2),
    ("K am",         "K2 von",       3),
    ("C am",         "C von",        4),
    ("Reopen C2 am", "Reopen C2 von",5),
]

AMAZON_COL = "Amazon"

# Logs in Unterordner
LOG_DIR = "tracker_logs"
os.makedirs(LOG_DIR, exist_ok=True)

def month_file_from_date(d):
    return os.path.join(LOG_DIR, f"date_changes_log_{d.strftime('%m.%Y')}.csv")

BACKUP_FILE = os.path.join(LOG_DIR, "date_backup.csv")

# -------------------- Writer Cache --------------------
_writer_cache = {}  # month_file -> (file_handle, csv_writer)

def get_writer_for_date(dt):
    """Return csv.writer for the month of dt, open/cached."""
    path = month_file_from_date(dt)
    if path not in _writer_cache:
        write_header = not os.path.exists(path)
        f = open(path, "a", newline="", encoding="utf-8")
        w = csv.writer(f)
        if write_header:
            w.writerow(["Änderung am", "Produktgruppe", "Land/Marketplace",
                        "Phase", "Feld", "Datum", "Mitarbeiter"])
        _writer_cache[path] = (f, w)
    return _writer_cache[path][1]

def close_writers():
    for f, _ in _writer_cache.values():
        f.close()
    _writer_cache.clear()

# -------------------- Backup Handling --------------------
def load_backup(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            r = csv.reader(f)
            next(r)
            return {row[0]: row[1:] for row in r}
    except FileNotFoundError:
        return {}

def save_backup(path, prev):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["RowKey", "LoggedDates..."])
        for k, vals in prev.items():
            w.writerow([k] + vals)

# -------------------- Parsing helpers --------------------
def parse_date_fuzzy(s):
    """Try to parse various date formats; return datetime.date or None."""
    if not s:
        return None
    # try ISO
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
    # try common European / fallback
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return None

# -------------------- Core append --------------------
def append_change(writer, when_ts, group, land, phase_no, field, date_val, user):
    writer.writerow([when_ts, group, land, phase_no, field, date_val, user])

# -------------------- Nightly tracker --------------------
def track_incremental(sm, prev):
    now = datetime.utcnow()
    ts_str = now.strftime("%Y-%m-%d %H:%M:%S")
    total = 0

    for group, sid in SHEET_IDS.items():
        sheet = sm.Sheets.get_sheet(sid)
        col_map = {c.id: c.title for c in sheet.columns}

        for row in sheet.rows:
            row_key = f"{group}:{row.id}"
            seen = prev.get(row_key, [])

            # Land/Marketplace
            land = ""
            for cell in row.cells:
                if col_map.get(cell.column_id) == AMAZON_COL:
                    land = (cell.display_value or "").strip()
                    break

            for date_col, user_col, phase_no in PHASE_FIELDS:
                date_val = ""
                user_val = ""
                for cell in row.cells:
                    title = col_map.get(cell.column_id, "")
                    if title == date_col:
                        date_val = (cell.display_value or "").strip()
                    elif title == user_col:
                        user_val = (cell.display_value or "").strip()

                if not date_val:
                    continue

                if date_val not in seen:
                    dt_obj = parse_date_fuzzy(date_val) or now.date()
                    w = get_writer_for_date(dt_obj)
                    append_change(w, ts_str, group, land, phase_no, date_col, date_val, user_val)
                    seen.append(date_val)
                    total += 1

            prev[row_key] = seen
    return total

# -------------------- Bootstrap --------------------
def seed_from_sheet(sm, prev):
    """
    Liest ALLE Zeilen aus dem Seed-Sheet (keine 90-Tage-Filterung mehr)
    und schreibt sie in die monatlichen CSVs, sofern noch nicht im Backup.
    Erwartete Spaltennamen im Seed-Sheet:
      Änderung am, Produktgruppe, Land/Marketplace, Phase, Feld, Datum, Mitarbeiter
    """
    sheet = sm.Sheets.get_sheet(SEED_SHEET_ID)
    col_map = {c.id: c.title for c in sheet.columns}

    # Wir prüfen nicht mehr streng – holen nur, was wir finden.
    needed = ["Änderung am","Produktgruppe","Land/Marketplace","Phase","Feld","Datum","Mitarbeiter"]

    added = 0
    for row in sheet.rows:
        rec = {}
        for cell in row.cells:
            title = col_map.get(cell.column_id, "")
            if title in needed:
                rec[title] = (cell.display_value or "").strip()

        # Mindestens Datum + Produktgruppe müssen da sein
        dt = parse_date_fuzzy(rec.get("Datum"))
        if not dt:
            continue

        # Fallbacks
        ts_str   = rec.get("Änderung am", datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"))
        grp      = rec.get("Produktgruppe", "")
        land     = rec.get("Land/Marketplace", "")
        phase    = rec.get("Phase", "")
        field    = rec.get("Feld", "")
        date_val = rec.get("Datum", "")
        user     = rec.get("Mitarbeiter", "")

        # Backup-Key: pro Row einzigartig genug gestalten
        key = f"{grp}:{land}:{field}:{date_val}"
        seen = prev.get(key, [])

        if date_val in seen:
            continue

        w = get_writer_for_date(dt)
        append_change(w, ts_str, grp, land, phase, field, date_val, user)
        seen.append(date_val)
        prev[key] = seen
        added += 1

    return added

# -------------------- MAIN --------------------
def main():
    load_dotenv()
    token = os.getenv("SMARTSHEET_TOKEN")
    sm = smartsheet.Smartsheet(token)

    parser = argparse.ArgumentParser()
    parser.add_argument("--bootstrap", action="store_true",
                        help="Einmalige Initial-Füllung aus Seed-Sheet.")
    # days bleibt für Abwärtskompatibilität, ist aber egal
    parser.add_argument("--days", type=int, default=-1,
                        help="(Ignoriert beim Bootstrap) Zeitraum in Tagen.")
    args = parser.parse_args()

    prev = load_backup(BACKUP_FILE)

    if args.bootstrap:
        added = seed_from_sheet(sm, prev)
        print(f"✅ Bootstrap: {added} Zeilen übernommen.")
    else:
        added = track_incremental(sm, prev)
        print(f"✅ Nightly: {added} neue Änderung(en) protokolliert.")

    close_writers()
    save_backup(BACKUP_FILE, prev)

if __name__ == "__main__":
    main()

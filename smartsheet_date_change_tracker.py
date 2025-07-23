import os
import csv
import argparse
from datetime import datetime, timedelta, timezone
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

SEED_SHEET_ID = 6879355327172484          # einmalige Bootstrap-Quelle
TRACK_DIR     = "tracker_logs"             # Ablageordner für CSVs
BACKUP_FILE   = os.path.join(TRACK_DIR, "date_backup.csv")

AMAZON_COL = "Amazon"
PHASE_FIELDS = [
    ("Kontrolle",    "K von",        1),
    ("BE am",        "BE von",       2),
    ("K am",         "K2 von",       3),
    ("C am",         "C von",        4),
    ("Reopen C2 am", "Reopen C2 von",5),
]

# -------------------- Writer Cache --------------------
_writer_cache = {}  # month_file -> (file_handle, csv_writer)

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def month_file_from_date(d):
    return os.path.join(TRACK_DIR, f"date_changes_log_{d.strftime('%m.%Y')}.csv")

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
    ensure_dir(os.path.dirname(path))
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
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
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
    now = datetime.now(timezone.utc)
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

# -------------------- Bootstrap (scan Seed-Sheet wie eine Gruppe) --------------------
def seed_from_sheet(sm, prev, days_back):
    cutoff = datetime.now(timezone.utc).date() - timedelta(days=days_back)
    sheet = sm.Sheets.get_sheet(SEED_SHEET_ID)
    col_map = {c.id: c.title for c in sheet.columns}

    added = 0
    now_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    # Wir behandeln jede Zeile wie in track_incremental:
    # Produktgruppe + Amazon + die PHASE_FIELDS
    for row in sheet.rows:
        # Produktgruppe
        group = ""
        land  = ""
        for cell in row.cells:
            title = col_map.get(cell.column_id, "")
            if title == "Produktgruppe":
                group = (cell.display_value or "").strip()
            elif title == AMAZON_COL:
                land = (cell.display_value or "").strip()

        # Fallback, falls keine Produktgruppe im Seed-Sheet => überspringen
        if not group:
            continue

        row_key = f"{group}:{row.id}"
        seen = prev.get(row_key, [])

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

            dt_obj = parse_date_fuzzy(date_val)
            if not dt_obj or dt_obj < cutoff:
                continue

            if date_val in seen:
                continue

            w = get_writer_for_date(dt_obj)
            append_change(w, now_ts, group, land, phase_no, date_col, date_val, user_val)
            seen.append(date_val)
            prev[row_key] = seen
            added += 1

    return added

# -------------------- MAIN --------------------
def main():
    ensure_dir(TRACK_DIR)

    load_dotenv()
    token = os.getenv("SMARTSHEET_TOKEN")
    sm = smartsheet.Smartsheet(token)

    parser = argparse.ArgumentParser()
    parser.add_argument("--bootstrap", action="store_true",
                        help="Einmalige Initial-Füllung aus Seed-Sheet.")
    parser.add_argument("--days", type=int, default=90,
                        help="Zeitraum für Bootstrap (Tage).")
    args = parser.parse_args()

    prev = load_backup(BACKUP_FILE)

    if args.bootstrap:
        added = seed_from_sheet(sm, prev, args.days)
        print(f"✅ Bootstrap: {added} Zeilen übernommen.")
    else:
        added = track_incremental(sm, prev)
        print(f"✅ Nightly: {added} neue Änderung(en) protokolliert.")

    close_writers()
    save_backup(BACKUP_FILE, prev)

if __name__ == "__main__":
    main()

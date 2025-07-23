import os
import re
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
    ("Kontrolle",    "K von",         1),
    ("BE am",        "BE von",        2),
    ("K am",         "K2 von",        3),
    ("C am",         "C von",         4),
    ("Reopen C2 am", "Reopen C2 von", 5),
]

AMAZON_COL = "Amazon"

# Logs in Unterordner
LOG_DIR = "tracker_logs"
os.makedirs(LOG_DIR, exist_ok=True)

def month_file_from_date(d):
    return os.path.join(LOG_DIR, f"date_changes_log_{d.strftime('%m.%Y')}.csv")

BACKUP_FILE = os.path.join(LOG_DIR, "date_backup.csv")

# -------------------- Writer Cache --------------------
_writer_cache = {}  # path -> (file_handle, csv_writer)

def get_writer_for_date(dt):
    """Return csv.writer for the month of dt, open/cached."""
    path = month_file_from_date(dt)
    if path not in _writer_cache:
        write_header = not os.path.exists(path)
        f = open(path, "a", newline="", encoding="utf-8")
        w = csv.writer(f)
        if write_header:
            w.writerow([
                "Änderung am",
                "Produktgruppe",
                "Land/Marketplace",
                "Phase",
                "Feld",
                "Datum",
                "Mitarbeiter",
            ])
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
    # ISO (handles 'YYYY-MM-DD' and 'YYYY-MM-DD HH:MM:SS')
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

# -------------------- Helpers --------------------
def get_cell_val(cell):
    return (cell.display_value or cell.value or "") if cell else ""

def row_dict(row, col_map):
    """title -> value dict für eine Row."""
    d = {}
    for c in row.cells:
        t = col_map.get(c.column_id, "")
        if t:
            d[t] = get_cell_val(c).strip()
    return d

def infer_group_from_primary(text, group_keys):
    if not text:
        return ""
    txt = text.strip().upper()
    pref = txt[:2]
    if pref in [g.upper() for g in group_keys]:
        return pref
    for g in group_keys:
        if txt.startswith(g.upper()):
            return g
    return ""

# -------------------- Bootstrap --------------------
def seed_from_sheet(sm, prev, dedupe=False, verbose=True):
    """
    Bootstrap: ALLE Phase-Daten aus dem Seed-Sheet schreiben.
    - Kein 90-Tage-Filter
    - Produktgruppe aus Primärspalte, falls leer
    - dedupe=False: ignoriert backup (schreibt alles)
    """
    sheet = sm.Sheets.get_sheet(SEED_SHEET_ID)
    col_map = {c.id: c.title for c in sheet.columns}
    primary_col_id = next((c.id for c in sheet.columns if getattr(c, "primary", False)), None)

    now_ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    added = 0
    skipped_no_date = 0
    skipped_dup = 0

    # Statistik je Phase
    per_phase_found = {p[0]: 0 for p in PHASE_FIELDS}
    per_phase_written = {p[0]: 0 for p in PHASE_FIELDS}

    for row in sheet.rows:
        rdict = row_dict(row, col_map)

        # Gruppe
        grp = rdict.get("Produktgruppe", "")
        if not grp and primary_col_id:
            prim_cell = next((c for c in row.cells if c.column_id == primary_col_id), None)
            grp = infer_group_from_primary(get_cell_val(prim_cell), SHEET_IDS.keys())

        land = rdict.get("Land/Marketplace", "") or rdict.get(AMAZON_COL, "")
        change_ts = rdict.get("Änderung am", "") or now_ts
        row_key_part = str(row.id)

        for date_col, user_col, phase_no in PHASE_FIELDS:
            date_val = rdict.get(date_col, "")
            if not date_val:
                continue
            per_phase_found[date_col] += 1

            dt_obj = parse_date_fuzzy(date_val)
            if not dt_obj:
                skipped_no_date += 1
                continue

            user_val = rdict.get(user_col, "") or rdict.get("Mitarbeiter", "")

            if dedupe:
                # Backup-Key inkl. Row-ID damit identische Daten in anderer Zeile nicht wegfallen
                key = f"{grp}:{land}:{phase_no}:{date_col}:{row_key_part}"
                seen = set(prev.get(key, []))
                pair = f"{date_val}"
                if pair in seen:
                    skipped_dup += 1
                    continue
                seen.add(pair)
                prev[key] = list(seen)
            # Wenn dedupe=False, wir schreiben einfach.

            w = get_writer_for_date(dt_obj)
            append_change(w, change_ts, grp, land, phase_no, date_col, date_val, user_val)
            per_phase_written[date_col] += 1
            added += 1

    if verbose:
        print("Bootstrap-Statistik:")
        print(f"  +{added} Zeilen geschrieben")
        print(f"  {skipped_no_date} ohne/ungültiges Datum")
        print(f"  {skipped_dup} Duplikate (nur wenn dedupe=True)")
        for col in PHASE_FIELDS:
            name = col[0]
            print(f"  {name}: gefunden {per_phase_found[name]}, geschrieben {per_phase_written[name]}")
    return added

# -------------------- MAIN --------------------
def main():
    load_dotenv()
    token = os.getenv("SMARTSHEET_TOKEN")
    sm = smartsheet.Smartsheet(token)

    parser = argparse.ArgumentParser()
    parser.add_argument("--bootstrap", action="store_true",
                        help="Einmalige Initial-Füllung aus Seed-Sheet.")
    parser.add_argument("--days", type=int, default=-1,
                        help="(Ignoriert beim Bootstrap) Zeitraum in Tagen für alte Variante.")
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

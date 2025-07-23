#!/usr/bin/env python3
import os
import csv
import argparse
from datetime import datetime, timedelta, timezone, date
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
SEED_SHEET_ID = 6879355327172484  # Tracker-/Seed-Sheet

PHASE_FIELDS = [
    ("Kontrolle",    "K von",        1),
    ("BE am",        "BE von",       2),
    ("K am",         "K2 von",       3),
    ("C am",         "C von",        4),
    ("Reopen C2 am", "Reopen C2 von",5),
]

AMAZON_COL   = "Amazon"
GROUP_COL_SEED = "Produktgruppe"  # in Seed sheet vorhanden

# Ordner für Logs
LOG_DIR = "tracker_logs"
os.makedirs(LOG_DIR, exist_ok=True)

def month_file_from_date(d: date) -> str:
    return os.path.join(LOG_DIR, f"date_changes_log_{d.strftime('%m.%Y')}.csv")

BACKUP_FILE = os.path.join(LOG_DIR, "date_backup.csv")

# -------------------- Writer Cache --------------------
_writer_cache = {}  # month_path -> (file_handle, writer)

def _open_writer_for_path(path: str):
    write_header = not os.path.exists(path)
    f = open(path, "a", newline="", encoding="utf-8")
    w = csv.writer(f)
    if write_header:
        w.writerow(["Änderung am", "Produktgruppe", "Land/Marketplace",
                    "Phase", "Feld", "Datum", "Mitarbeiter"])
    return f, w

def get_writer_for_date(dt: date):
    path = month_file_from_date(dt)
    if path not in _writer_cache:
        _writer_cache[path] = _open_writer_for_path(path)
    return _writer_cache[path][1]

def ensure_month_file_exists(dt: date):
    path = month_file_from_date(dt)
    if not os.path.exists(path):
        f, w = _open_writer_for_path(path)
        f.close()

def close_writers():
    for f, _w in _writer_cache.values():
        f.close()
    _writer_cache.clear()

# -------------------- Backup Handling --------------------
def load_backup(path: str):
    try:
        with open(path, "r", encoding="utf-8") as f:
            r = csv.reader(f)
            next(r)
            return {row[0]: row[1:] for row in r}
    except FileNotFoundError:
        return {}

def save_backup(path: str, prev: dict):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["RowKey", "LoggedDates..."])
        for k, vals in prev.items():
            w.writerow([k] + vals)

# -------------------- Date Parsing --------------------
def parse_date_fuzzy(s: str):
    if not s:
        return None
    s = s.strip()
    # ISO
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
    # Common forms
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return None

# -------------------- CSV Append --------------------
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

            # Land
            land = ""
            for cell in row.cells:
                if col_map.get(cell.column_id) == AMAZON_COL:
                    land = (cell.display_value or "").strip()
                    break

            for date_col, user_col, phase_no in PHASE_FIELDS:
                date_val = ""
                user_val = ""
                for cell in row.cells:
                    t = col_map.get(cell.column_id, "")
                    if t == date_col:
                        date_val = (cell.display_value or "").strip()
                    elif t == user_col:
                        user_val = (cell.display_value or "").strip()

                if not date_val:
                    continue
                if date_val in seen:
                    continue

                dt_obj = parse_date_fuzzy(date_val) or now.date()
                w = get_writer_for_date(dt_obj)
                append_change(w, ts_str, group, land, phase_no, date_col, date_val, user_val)
                seen.append(date_val)
                total += 1

            prev[row_key] = seen

    return total

# -------------------- Bootstrap from Seed Sheet --------------------
def seed_from_sheet(sm, prev, days_back: int):
    """
    Lies die Tracker-Tabelle (SEED_SHEET_ID):
    - Hole für jede Zeile die 5 Phasen-Datumsfelder
    - Ermittle das neueste Datum
    - Wenn neuestes Datum <= days_back liegt: logge ALLE Phasen-Daten in diesem Zeitraum.
    """
    cutoff = datetime.now(timezone.utc).date() - timedelta(days=days_back)

    sheet = sm.Sheets.get_sheet(SEED_SHEET_ID)
    col_map_id2title = {c.id: c.title for c in sheet.columns}
    title2id = {c.title: c.id for c in sheet.columns}

    added = 0
    now_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    for row in sheet.rows:
        # Produktgruppe und Land
        group = ""
        land  = ""
        # Sammle alle Phase-Daten
        phase_values = []  # [(phase_no, field_name, date_str, user_str), ...]

        for date_col, user_col, phase_no in PHASE_FIELDS:
            d_str = ""
            u_str = ""
            # hole display_values
            dcid = title2id.get(date_col)
            ucid = title2id.get(user_col)
            if dcid:
                cell = next((c for c in row.cells if c.column_id == dcid), None)
                if cell:
                    d_str = (cell.display_value or "").strip()
            if ucid:
                cell = next((c for c in row.cells if c.column_id == ucid), None)
                if cell:
                    u_str = (cell.display_value or "").strip()

            if d_str:
                phase_values.append((phase_no, date_col, d_str, u_str))

        # group & land/marketplace
        # (wenn vorhanden – sonst leer)
        gid = title2id.get(GROUP_COL_SEED)
        if gid:
            c = next((c for c in row.cells if c.column_id == gid), None)
            if c:
                group = (c.display_value or "").strip()

        amid = title2id.get(AMAZON_COL)
        if amid:
            c = next((c for c in row.cells if c.column_id == amid), None)
            if c:
                land = (c.display_value or "").strip()

        if not phase_values:
            continue

        # Latest date of the 5 phases
        parsed_dates = [parse_date_fuzzy(x[2]) for x in phase_values if parse_date_fuzzy(x[2])]
        if not parsed_dates:
            continue
        latest = max(parsed_dates)
        if latest < cutoff:
            # alles zu alt
            continue

        # row_key für Backup (seed-spezifisch)
        row_key = f"seed:{row.id}"
        seen = prev.get(row_key, [])

        # Logge alle Phase-Daten, die >= cutoff sind & noch nicht logged
        for phase_no, field_name, d_str, u_str in phase_values:
            d_parsed = parse_date_fuzzy(d_str)
            if not d_parsed or d_parsed < cutoff:
                continue
            if d_str in seen:
                continue
            w = get_writer_for_date(d_parsed)
            append_change(w, now_ts, group, land, phase_no, field_name, d_str, u_str)
            seen.append(d_str)
            added += 1

        prev[row_key] = seen

    return added

# -------------------- MAIN --------------------
def main():
    load_dotenv()
    token = os.getenv("SMARTSHEET_TOKEN")
    if not token:
        raise RuntimeError("SMARTSHEET_TOKEN fehlt!")

    sm = smartsheet.Smartsheet(token)

    parser = argparse.ArgumentParser()
    parser.add_argument("--bootstrap", action="store_true",
                        help="Einmalige Initial-Füllung aus Seed-Sheet.")
    parser.add_argument("--days", type=int, default=90,
                        help="Zeitraum (Tage) für Bootstrap-Filter.")
    args = parser.parse_args()

    # Stelle sicher, dass für aktuellen Monat eine leere CSV da ist
    ensure_month_file_exists(datetime.now(timezone.utc).date())

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

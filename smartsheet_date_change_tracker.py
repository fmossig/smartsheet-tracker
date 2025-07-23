#!/usr/bin/env python3
import os
import csv
import argparse
from datetime import datetime, timedelta, timezone
from collections import defaultdict

import smartsheet
from dotenv import load_dotenv

# -------------------- KONFIG --------------------
TRACKER_DIR   = "tracker_logs"
os.makedirs(TRACKER_DIR, exist_ok=True)

SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}
SEED_SHEET_ID = 6879355327172484  # Bootstrap-Quelle (dein Tracking-Sheet)

PHASE_FIELDS = [
    ("Kontrolle",    "K von",        1),
    ("BE am",        "BE von",       2),
    ("K am",         "K2 von",       3),
    ("C am",         "C von",        4),
    ("Reopen C2 am", "Reopen C2 von",5),
]

AMAZON_COL   = "Amazon"
GROUP_COL_SEED = "Produktgruppe"

BACKUP_FILE  = os.path.join(TRACKER_DIR, "date_backup.csv")

# -------------------- Writer Cache --------------------
_writer_cache = {}  # month_path -> (file_handle, csv_writer)

def month_file_from_date(d):
    return os.path.join(TRACKER_DIR, f"date_changes_log_{d.strftime('%m.%Y')}.csv")

def get_writer_for_date(dt):
    path = month_file_from_date(dt)
    if path not in _writer_cache:
        write_header = not os.path.exists(path)
        f = open(path, "a", newline="", encoding="utf-8")
        w = csv.writer(f)
        if write_header:
            w.writerow([
                "Änderung am", "Produktgruppe", "Land/Marketplace",
                "Phase", "Feld", "Datum", "Mitarbeiter"
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
                if date_val in seen:
                    continue

                d_parsed = parse_date_fuzzy(date_val) or now.date()
                w = get_writer_for_date(d_parsed)
                append_change(w, ts_str, group, land, phase_no, date_col, date_val, user_val)
                seen.append(date_val)
                total += 1

            prev[row_key] = seen
    return total

# -------------------- Bootstrap --------------------
def seed_from_sheet(sm, prev, days_back: int | None, no_cutoff: bool):
    """
    Liest SEED_SHEET_ID und füllt CSVs.
    - Wenn Log-Spalten vorhanden: direkt übernehmen (optional Cutoff)
    - Sonst: Phasen-Felder auslesen (wie Nightly)
    """
    if no_cutoff:
        cutoff = datetime(1970, 1, 1).date()
    else:
        if days_back is None:
            days_back = 90
        cutoff = datetime.now(timezone.utc).date() - timedelta(days=days_back)

    sheet = sm.Sheets.get_sheet(SEED_SHEET_ID)
    col_map_id2title = {c.id: c.title for c in sheet.columns}
    title2id = {c.title: c.id for c in sheet.columns}

    needed_cols = {"Änderung am", "Produktgruppe", "Land/Marketplace",
                   "Phase", "Feld", "Datum", "Mitarbeiter"}
    have_cols = set(title2id.keys())

    added = 0
    now_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    # MODE 1: fertiges Log
    if needed_cols.issubset(have_cols):
        for row in sheet.rows:
            rec = {}
            for cell in row.cells:
                t = col_map_id2title[cell.column_id]
                if t in needed_cols:
                    rec[t] = (cell.display_value or "").strip()

            if not rec.get("Datum"):
                continue
            d_parsed = parse_date_fuzzy(rec["Datum"])
            if not d_parsed or d_parsed < cutoff:
                continue

            key = f"seed:{row.id}:{rec['Datum']}:{rec['Feld']}"
            seen = prev.get(key, [])
            if rec["Datum"] in seen:
                continue

            w = get_writer_for_date(d_parsed)
            append_change(w,
                          rec.get("Änderung am", now_ts),
                          rec.get("Produktgruppe", ""),
                          rec.get("Land/Marketplace", ""),
                          rec.get("Phase", ""),
                          rec.get("Feld", ""),
                          rec.get("Datum", ""),
                          rec.get("Mitarbeiter", ""))
            seen.append(rec["Datum"])
            prev[key] = seen
            added += 1
        return added

    # MODE 2: wie Produktgruppen-Sheets
    for row in sheet.rows:
        # Gruppe / Land
        group = ""
        land  = ""

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

        phase_values = []
        for date_col, user_col, phase_no in PHASE_FIELDS:
            d_str = ""
            u_str = ""
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

        if not phase_values:
            continue

        # jüngstes Datum prüfen
        parsed_dates = [parse_date_fuzzy(x[2]) for x in phase_values if parse_date_fuzzy(x[2])]
        if not parsed_dates:
            continue
        latest = max(parsed_dates)
        if latest < cutoff:
            continue

        row_key = f"seed:{row.id}"
        seen = prev.get(row_key, [])

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
    sm = smartsheet.Smartsheet(token)

    parser = argparse.ArgumentParser()
    parser.add_argument("--bootstrap", action="store_true",
                        help="Initiale Füllung aus Seed-Sheet.")
    parser.add_argument("--days", type=int, default=90,
                        help="Zeitraum für Bootstrap (Tage).")
    parser.add_argument("--no-cutoff", action="store_true",
                        help="Bootstrap ohne Datumsgrenze.")
    args = parser.parse_args()

    prev = load_backup(BACKUP_FILE)

    if args.bootstrap:
        added = seed_from_sheet(sm, prev, args.days, no_cutoff=args.no_cutoff)
        print(f"✅ Bootstrap: {added} Zeile(n) übernommen.")
    else:
        added = track_incremental(sm, prev)
        print(f"✅ Nightly: {added} neue Änderung(en) protokolliert.")

    close_writers()
    save_backup(BACKUP_FILE, prev)

if __name__ == "__main__":
    main()

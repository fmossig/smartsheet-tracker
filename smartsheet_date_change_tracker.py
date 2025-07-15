import smartsheet
import csv
import os
from dotenv import load_dotenv

load_dotenv()
access_token = os.getenv("SMARTSHEET_TOKEN")
smartsheet_client = smartsheet.Smartsheet(access_token)

sheet_ids = {
    "Amazon Cases (NA)": 6141179298008964,
    "Amazon Cases (NF)": 615755411312516,
    "Amazon Cases (NH)": 123340632051588,
    "Amazon Cases (NP)": 3009924800925572,
    "Amazon Cases (NT)": 2199739350077316,
    "Amazon Cases (NV)": 8955413669040004,
    "Amazon Cases (NM)": 4275419734822788,
}

date_columns = ["Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"]

try:
    with open('date_backup.csv', 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        headers = next(reader)
        previous_rows = {row[0]: row for row in reader}
except FileNotFoundError:
    previous_rows = {}
    headers = []

changes = []
log_headers_written = os.path.exists("date_changes_log.csv")

for sheet_name, sheet_id in sheet_ids.items():
    try:
        sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    except Exception as e:
        print(f"❌ Fehler beim Laden von {sheet_name}: {e}")
        continue

    col_map = {col.title.strip(): col.id for col in sheet.columns}
    header_order = [col.title for col in sheet.columns]

    for row in sheet.rows:
        row_data = {col.title: "" for col in sheet.columns}
        for cell in row.cells:
            col_title = next((col.title for col in sheet.columns if col.id == cell.column_id), "")
            row_data[col_title] = cell.display_value or ""

        key = f"{sheet_name}:{row.id}"
        current_values = [row_data.get(col, "") for col in date_columns]
        old_values = previous_rows.get(key)

        if old_values:
            old_vals = [old_values[header_order.index(col)] for col in date_columns if col in header_order]
            if current_values != old_vals:
                changes.append([key] + [row_data.get(col, "") for col in header_order])
        else:
            changes.append([key] + [row_data.get(col, "") for col in header_order])

        previous_rows[key] = [key] + [row_data.get(col, "") for col in header_order]
        headers = ["RowKey"] + header_order

if changes:
    with open("date_changes_log.csv", "a", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not log_headers_written:
            writer.writerow(headers)
        for row in changes:
            writer.writerow(row)
    print(f"✅ {len(changes)} Änderung(en) protokolliert.")
else:
    print("🔍 Keine Änderungen erkannt.")

with open("date_backup.csv", "w", newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    for row in previous_rows.values():
        writer.writerow(row)

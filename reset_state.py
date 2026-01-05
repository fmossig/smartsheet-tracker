import json
import smartsheet
import os
import subprocess
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
if not token:
    print("Error: SMARTSHEET_TOKEN not found in environment or .env file")
    exit(1)

# Sheets to process
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# Fields to track
PHASE_FIELDS = [
    ("Kontrolle", "K von", 1),
    ("BE am", "BE von", 2),
    ("K am", "K2 von", 3),
    ("C am", "C von", 4),
    ("Reopen C2 am", "Reopen C2 von", 5),
]

# First, reset or create the change history file
history_file = "tracking_data/change_history.csv"
os.makedirs("tracking_data", exist_ok=True)

# Try to reset using git if the file is tracked
try:
    print("Attempting to reset change history file using git...")
    subprocess.run(["git", "checkout", "--", history_file], check=False)
except Exception as e:
    print(f"Note: Git command failed: {e}")

# Create a fresh change history file with headers
print("Creating fresh change history file...")
with open(history_file, "w", newline="", encoding="utf-8") as f:
    f.write("Timestamp,Group,RowID,Phase,DateField,Date,User,Marketplace\n")

print("Connecting to Smartsheet...")
# Create a fresh state file that marks all current data as processed
client = smartsheet.Smartsheet(token)
client.errors_as_exceptions(True)
state = {"last_run": "2025-10-18 16:39:14", "processed": {}}

# Process each sheet
for group, sid in SHEET_IDS.items():
    print(f"Processing sheet {group}...")
    try:
        sheet = client.Sheets.get_sheet(sid)
        
        # Map column titles to IDs
        col_map = {col.title: col.id for col in sheet.columns}
        
        # Process each row
        for row in sheet.rows:
            for date_col, _, _ in PHASE_FIELDS:
                col_id = col_map.get(date_col)
                if not col_id:
                    continue
                    
                # Find cell with this column ID
                for cell in row.cells:
                    if cell.column_id == col_id and cell.value:
                        # Add to processed state with normalized date (YYYY-MM-DD)
                        field_key = f"{group}:{row.id}:{date_col}"
                        # Normalize to YYYY-MM-DD format
                        val = cell.value
                        if hasattr(val, 'date'):
                            val = val.date().isoformat()
                        elif hasattr(val, 'isoformat'):
                            val = val.isoformat()
                        else:
                            val = str(val).strip()[:10]  # Take just YYYY-MM-DD part
                        state["processed"][field_key] = val
                        break
    except Exception as e:
        print(f"Error processing sheet {group}: {e}")
        continue

# Save state
with open("tracking_data/tracker_state.json", "w") as f:
    json.dump(state, f)

print(f"Created state file with {len(state['processed'])} processed items")
print("System reset complete - tracking will now only capture new changes")

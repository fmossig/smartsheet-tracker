import json
import os
import smartsheet
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
if not token:
    print("ERROR: No Smartsheet token found")
    exit(1)

# State file
STATE_FILE = "tracking_data/tracker_state.json"

# Check if state file exists
if not os.path.exists(STATE_FILE):
    print(f"ERROR: State file not found: {STATE_FILE}")
    exit(1)

# Load state
with open(STATE_FILE, 'r') as f:
    state = json.load(f)
    processed = state.get("processed", {})
    print(f"Loaded state file with {len(processed)} entries")

# Connect to Smartsheet
client = smartsheet.Smartsheet(token)
print("Connected to Smartsheet")

# Example fields we're checking
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

PHASE_FIELDS = [
    "Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"
]

# Collect current values for fields in state
print("Checking for differences...")
print("=" * 50)

differences = []
for key, stored_value in processed.items():
    try:
        # Parse the key (format: "GROUP:ROW_ID:FIELD")
        parts = key.split(":")
        if len(parts) != 3:
            continue
            
        group, row_id, field = parts
        
        if group not in SHEET_IDS:
            continue
            
        if field not in PHASE_FIELDS:
            continue
            
        # Get current value from Smartsheet
        sheet = client.Sheets.get_sheet(SHEET_IDS[group])
        
        # Map column titles to IDs
        col_map = {col.title: col.id for col in sheet.columns}
        if field not in col_map:
            continue
            
        # Find the row
        row = next((r for r in sheet.rows if str(r.id) == row_id), None)
        if not row:
            print(f"Row not found: {row_id} in {group}")
            continue
            
        # Get the cell value
        cell = next((c for c in row.cells if c.column_id == col_map[field]), None)
        current_value = cell.value if cell else None
        
        # Compare with stored value
        if stored_value != current_value:
            differences.append({
                "key": key,
                "stored": stored_value,
                "current": current_value
            })
            print(f"DIFFERENCE FOUND: {key}")
            print(f"  Stored: {stored_value}")
            print(f"  Current: {current_value}")
            
    except Exception as e:
        print(f"Error checking {key}: {e}")

print("=" * 50)
print(f"Total differences found: {len(differences)}")

if len(differences) == 0:
    # Check a sample of entries
    print("\nChecking random sample of 5 entries:")
    sample_count = 0
    for key, stored_value in list(processed.items())[:5]:
        print(f"Sample {sample_count+1}: {key} = {stored_value}")
        sample_count += 1
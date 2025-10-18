import os
import json
import csv
from datetime import datetime
import smartsheet
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
if not token:
    print("ERROR: SMARTSHEET_TOKEN not found in environment or .env file")
    exit(1)

# Constants from your tracking script
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
    ("Kontrolle", "K von", 1),
    ("BE am", "BE von", 2),
    ("K am", "K2 von", 3),
    ("C am", "C von", 4),
    ("Reopen C2 am", "Reopen C2 von", 5),
]

DATA_DIR = "tracking_data"
STATE_FILE = os.path.join(DATA_DIR, "tracker_state.json")
CHANGES_FILE = os.path.join(DATA_DIR, "change_history.csv")

def load_state():
    """Load state file and return the data."""
    if not os.path.exists(STATE_FILE):
        print(f"State file not found: {STATE_FILE}")
        return {"last_run": None, "processed": {}}
    
    try:
        with open(STATE_FILE, 'r') as f:
            state = json.load(f)
            print(f"Loaded state file with {len(state.get('processed', {}))} processed items")
            return state
    except Exception as e:
        print(f"Error loading state: {e}")
        return {"last_run": None, "processed": {}}

def save_state(state):
    """Save state to file."""
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(state, f)
            print(f"Saved state with {len(state.get('processed', {}))} processed items")
    except Exception as e:
        print(f"Error saving state: {e}")

def find_differences():
    """Find differences between stored state and current Smartsheet values."""
    state = load_state()
    processed = state.get("processed", {})
    
    # Connect to Smartsheet
    client = smartsheet.Smartsheet(token)
    client.errors_as_exceptions(True)
    print("Connected to Smartsheet API")
    
    # Track differences
    differences = []
    current_values = {}
    
    # Process each sheet
    for group, sheet_id in SHEET_IDS.items():
        print(f"\nProcessing sheet {group} (ID: {sheet_id})")
        
        try:
            sheet = client.Sheets.get_sheet(sheet_id)
            print(f"Sheet {group} has {len(sheet.rows)} rows")
            
            col_map = {col.title: col.id for col in sheet.columns}
            
            # Check which tracked fields exist in this sheet
            found_fields = [f for f, _, _ in PHASE_FIELDS if f in col_map]
            print(f"Found tracked fields: {found_fields}")
            
            # Process each row
            for row in sheet.rows:
                for date_col, user_col, phase_no in PHASE_FIELDS:
                    if date_col not in col_map:
                        continue
                    
                    # Get current value from Smartsheet
                    date_val = None
                    user_val = ""
                    
                    for cell in row.cells:
                        if cell.column_id == col_map.get(date_col):
                            date_val = cell.value
                        if cell.column_id == col_map.get(user_col):
                            user_val = cell.display_value or ""
                    
                    if not date_val:
                        continue
                    
                    # Create field key
                    field_key = f"{group}:{row.id}:{date_col}"
                    
                    # Store current value
                    current_values[field_key] = date_val
                    
                    # Compare with stored state
                    prev_val = processed.get(field_key)
                    
                    if prev_val != date_val:
                        differences.append({
                            "field_key": field_key,
                            "prev_value": prev_val,
                            "current_value": date_val,
                            "user": user_val
                        })
                        
        except Exception as e:
            print(f"Error processing sheet {group}: {e}")
    
    # Print results
    print("\n" + "="*50)
    print(f"FOUND {len(differences)} DIFFERENCES")
    print("="*50)
    
    if differences:
        print("\nTop 10 differences:")
        for i, diff in enumerate(differences[:10]):
            print(f"{i+1}. {diff['field_key']}")
            print(f"   Old: {diff['prev_value']}")
            print(f"   New: {diff['current_value']}")
            print(f"   User: {diff['user']}")
    
    return differences, current_values

def force_track_changes(differences):
    """Force tracking of detected differences."""
    if not differences:
        print("No differences to track.")
        return False
    
    print(f"\nForcing tracking of {len(differences)} changes...")
    
    # Load current state
    state = load_state()
    
    # Get timestamp
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    
    # Ensure tracking_data directory exists
    os.makedirs(DATA_DIR, exist_ok=True)
    
    # Append to changes file
    try:
        # Create file with headers if it doesn't exist
        if not os.path.exists(CHANGES_FILE):
            with open(CHANGES_FILE, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([
                    "Timestamp", 
                    "Group", 
                    "RowID",
                    "Phase", 
                    "DateField", 
                    "Date", 
                    "User",
                    "Marketplace"
                ])
                print(f"Created new changes file: {CHANGES_FILE}")
        
        with open(CHANGES_FILE, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            for diff in differences:
                # Parse field key
                parts = diff['field_key'].split(":")
                if len(parts) != 3:
                    print(f"Invalid field key format: {diff['field_key']}")
                    continue
                
                group, row_id, date_col = parts
                
                # Find phase number
                phase_no = 0
                for dc, _, p in PHASE_FIELDS:
                    if dc == date_col:
                        phase_no = p
                        break
                
                # Parse date
                date_val = diff['current_value']
                try:
                    # Try ISO format first
                    dt = datetime.fromisoformat(date_val).date()
                except ValueError:
                    try:
                        # Try other formats
                        for fmt in ('%Y-%m-%d', '%d.%m.%Y'):
                            try:
                                dt = datetime.strptime(date_val, fmt).date()
                                break
                            except ValueError:
                                continue
                    except Exception:
                        print(f"Could not parse date: {date_val}")
                        continue
                
                # Write change record
                writer.writerow([
                    timestamp,
                    group,
                    row_id,
                    phase_no,
                    date_col,
                    dt.isoformat(),
                    diff['user'],
                    ""  # Marketplace (empty)
                ])
            
            print(f"Added {len(differences)} changes to {CHANGES_FILE}")
            
            # Update state
            for diff in differences:
                state["processed"][diff['field_key']] = diff['current_value']
            
            state["last_run"] = timestamp
            save_state(state)
            
            return True
    except Exception as e:
        print(f"Error writing changes: {e}")
        return False

def update_state_from_current(current_values):
    """Update state file with current Smartsheet values."""
    print("\nUpdating state file with current Smartsheet values...")
    
    # Load current state
    state = load_state()
    
    # Update with current values
    state["processed"] = current_values
    state["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Save updated state
    save_state(state)
    
    print("State file updated with current values from Smartsheet")
    return True

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Diagnose Smartsheet tracking issues")
    parser.add_argument("--check", action="store_true", help="Check for differences between state and current values")
    parser.add_argument("--force-track", action="store_true", help="Force tracking of any differences found")
    parser.add_argument("--update-state", action="store_true", help="Update state file with current Smartsheet values")
    
    args = parser.parse_args()
    
    if args.check or args.force_track:
        differences, current_values = find_differences()
        
        if args.force_track and differences:
            force_track_changes(differences)
    
    if args.update_state:
        _, current_values = find_differences()
        update_state_from_current(current_values)
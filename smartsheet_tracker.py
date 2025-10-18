import os
import csv
import json
from datetime import datetime, timedelta
import logging

import smartsheet
from dotenv import load_dotenv

# Set up logging - INCREASED VERBOSITY
logging.basicConfig(
    level=logging.DEBUG,  # Changed from INFO to DEBUG
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("smartsheet_tracker.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
if not token:
    logger.error("SMARTSHEET_TOKEN not found in environment or .env file")
    exit(1)

# Smartsheet IDs and columns to track
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# Fields to track - (date_column, user_column, phase_number)
PHASE_FIELDS = [
    ("Kontrolle", "K von", 1),
    ("BE am", "BE von", 2),
    ("K am", "K2 von", 3),
    ("C am", "C von", 4),
    ("Reopen C2 am", "Reopen C2 von", 5),
]

# Directory to store data
DATA_DIR = "tracking_data"
os.makedirs(DATA_DIR, exist_ok=True)

# State file to track what we've already processed
STATE_FILE = os.path.join(DATA_DIR, "tracker_state.json")
CHANGES_FILE = os.path.join(DATA_DIR, "change_history.csv")

def load_state():
    """Load previously saved state or create empty state."""
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, 'r') as f:
                state = json.load(f)
                logger.info(f"Loaded state file with {len(state.get('processed', {}))} processed items")
                return state
        else:
            logger.warning(f"State file not found: {STATE_FILE}")
            return {"last_run": None, "processed": {}}
    except Exception as e:
        logger.error(f"Error loading state: {e}")
        return {"last_run": None, "processed": {}}

def save_state(state):
    """Save current state to file."""
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(state, f)
            logger.info(f"Saved state with {len(state.get('processed', {}))} processed items")
    except Exception as e:
        logger.error(f"Error saving state: {e}")

def ensure_changes_file():
    """Create changes file with headers if it doesn't exist."""
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
            logger.info(f"Created new changes file: {CHANGES_FILE}")
    else:
        logger.info(f"Changes file exists: {CHANGES_FILE}")

def parse_date(date_str):
    """Parse date from string, supporting multiple formats and handling typos."""
    if not date_str:
        return None
        
    # Clean up common typos
    cleaned = date_str.strip()
    if cleaned and not cleaned[-1].isdigit():
        # Remove any trailing non-digit characters
        cleaned = ''.join([c for i, c in enumerate(cleaned) 
                          if i < len(cleaned)-1 or c.isdigit()])
        
    # Try various formats
    for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%d', '%d.%m.%Y', '%m/%d/%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
            
    try:
        # Try ISO format (catches many variations)
        return datetime.fromisoformat(cleaned).date()
    except:
        return None

def track_changes():
    """Main function to track changes in Smartsheet tables."""
    logger.info("Starting Smartsheet change tracking")
    
    # Initialize
    state = load_state()
    ensure_changes_file()
    
    # Connect to Smartsheet
    try:
        client = smartsheet.Smartsheet(token)
        client.errors_as_exceptions(True)
        logger.info("Connected to Smartsheet API")
    except Exception as e:
        logger.error(f"Failed to connect to Smartsheet: {e}")
        return False
    
    # Current timestamp
    now = datetime.now()
    today = now.date()
    
    # Track new changes
    changes_found = 0
    fields_checked = 0
    
    try:
        # Open file in append mode
        with open(CHANGES_FILE, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # Process each sheet
            for group, sheet_id in SHEET_IDS.items():
                logger.info(f"Processing sheet {group} (ID: {sheet_id})")
                
                try:
                    # Get sheet with columns and rows
                    sheet = client.Sheets.get_sheet(sheet_id)
                    logger.info(f"Sheet {group} has {len(sheet.rows)} rows")
                    
                    # Map column titles to IDs
                    col_map = {col.title: col.id for col in sheet.columns}
                    amazon_col_id = col_map.get("Amazon")
                    
                    # Log which tracked fields are found
                    found_fields = [f for f, _, _ in PHASE_FIELDS if f in col_map]
                    logger.info(f"Found tracked fields in sheet {group}: {found_fields}")
                    
                    # Process each row
                    for row in sheet.rows:
                        row_key = f"{group}:{row.id}"
                        
                        # Get marketplace if available
                        marketplace = ""
                        if amazon_col_id:
                            for cell in row.cells:
                                if cell.column_id == amazon_col_id:
                                    marketplace = (cell.display_value or "").strip()
                                    break
                        
                        # Check each phase field
                        for date_col, user_col, phase_no in PHASE_FIELDS:
                            date_val = None
                            user_val = ""
                            
                            # Get date and user values
                            for cell in row.cells:
                                if cell.column_id == col_map.get(date_col):
                                    date_val = cell.value
                                if cell.column_id == col_map.get(user_col):
                                    user_val = cell.display_value or ""
                            
                            if not date_val:
                                continue
                                
                            fields_checked += 1
                            
                            # Create unique key for this specific field
                            field_key = f"{row_key}:{date_col}"
                            
                            # Check if this is a new or changed value
                            prev_val = state["processed"].get(field_key)
                            
                            # Debug log for comparison
                            logger.debug(f"Comparing field {field_key}: old='{prev_val}', new='{date_val}'")
                            
                            if prev_val == date_val:
                                continue
                                
                            # Parse date
                            parsed_date = parse_date(date_val)
                            if not parsed_date:
                                logger.warning(f"Could not parse date: {date_val} for {field_key}")
                                continue
                                
                            # Found a change!
                            logger.info(f"CHANGE DETECTED in {field_key}: old='{prev_val}', new='{date_val}'")
                                
                            # Record the change
                            writer.writerow([
                                now.strftime("%Y-%m-%d %H:%M:%S"),
                                group,
                                row.id,
                                phase_no,
                                date_col,
                                parsed_date.isoformat(),
                                user_val,
                                marketplace
                            ])
                            
                            # Update state
                            state["processed"][field_key] = date_val
                            changes_found += 1
                            
                except Exception as e:
                    logger.error(f"Error processing sheet {group}: {e}")
                    continue
    
    except Exception as e:
        logger.error(f"Error writing to changes file: {e}")
        return False
    
    # Update state
    state["last_run"] = now.strftime("%Y-%m-%d %H:%M:%S")
    save_state(state)
    
    logger.info(f"Change tracking completed. Checked {fields_checked} fields, found {changes_found} changes.")
    return True

def bootstrap_tracking(days_back=0):
    """Initialize tracking for new data only."""
    logger.info(f"Starting bootstrap (tracking new data only)")
    
    # Reset state to force reprocessing
    state = {"last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "processed": {}}
    save_state(state)
    
    # Ensure we have a new changes file
    if os.path.exists(CHANGES_FILE):
        # Create backup
        backup_file = f"{CHANGES_FILE}.bak.{datetime.now().strftime('%Y%m%d%H%M%S')}"
        try:
            os.rename(CHANGES_FILE, backup_file)
            logger.info(f"Previous changes file backed up to {backup_file}")
        except Exception as e:
            logger.error(f"Failed to backup changes file: {e}")
    
    # Create new file and ensure report directories exist
    ensure_changes_file()
    
    # Create report directories to avoid git errors
    os.makedirs("reports/weekly", exist_ok=True)
    os.makedirs("reports/monthly", exist_ok=True)
    
    logger.info("Bootstrap complete. System is ready to track new changes.")
    return True

def reset_tracking_state():
    """Reset the tracking state to current Smartsheet data."""
    logger.info("Resetting tracking state...")
    
    # Connect to Smartsheet
    client = smartsheet.Smartsheet(token)
    state = {"last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "processed": {}}
    
    # Process each sheet to build state
    for group, sid in SHEET_IDS.items():
        logger.info(f"Processing sheet {group}...")
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
                        # Add to processed state
                        field_key = f"{group}:{row.id}:{date_col}"
                        state["processed"][field_key] = cell.value
                        break
    
    # Save state
    with open(STATE_FILE, "w") as f:
        json.dump(state, f)
    
    # Reset change history file
    with open(CHANGES_FILE, "w", newline="", encoding="utf-8") as f:
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
    
    logger.info(f"Reset complete: Marked {len(state['processed'])} items as processed")
    return True

def force_sample_change():
    """Debugging function to force detecting a sample change."""
    logger.info("Forcing a sample change detection for testing...")
    
    # Load current state
    state = load_state()
    
    # Find a key to modify
    if not state["processed"]:
        logger.error("No processed items found in state file")
        return False
    
    # Take first key and modify it
    sample_key = next(iter(state["processed"].keys()))
    old_val = state["processed"][sample_key]
    logger.info(f"Selected key for forced change: {sample_key}, current value: {old_val}")
    
    # Remove this key from state to force detection
    del state["processed"][sample_key]
    save_state(state)
    
    logger.info(f"Removed key {sample_key} from state - next tracking run will detect it as a change")
    return True

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Track changes in Smartsheet tables")
    parser.add_argument("--bootstrap", action="store_true", help="Initialize tracking for all data")
    parser.add_argument("--reset", action="store_true", help="Reset tracking state to current data")
    parser.add_argument("--force-change", action="store_true", help="Force detection of a sample change (testing)")
    args = parser.parse_args()
    
    if args.reset:
        success = reset_tracking_state()
    elif args.bootstrap:
        success = bootstrap_tracking()
    elif args.force_change:
        success = force_sample_change()
    else:
        success = track_changes()
        
    exit(0 if success else 1)

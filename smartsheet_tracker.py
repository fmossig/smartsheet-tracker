import os
import csv
import json
from datetime import datetime, timedelta
import logging

import smartsheet
from dotenv import load_dotenv

# Set up logging
logging.basicConfig(
    level=logging.INFO,
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
                return json.load(f)
        else:
            return {"last_run": None, "processed": {}}
    except Exception as e:
        logger.error(f"Error loading state: {e}")
        return {"last_run": None, "processed": {}}

def save_state(state):
    """Save current state to file."""
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(state, f)
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

def parse_date(date_str):
    """Parse date from string, supporting multiple formats."""
    if not date_str:
        return None
        
    # Try various formats
    for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%d', '%d.%m.%Y'):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
            
    try:
        # Try ISO format (catches many variations)
        return datetime.fromisoformat(date_str).date()
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
                    
                    # Map column titles to IDs
                    col_map = {col.title: col.id for col in sheet.columns}
                    amazon_col_id = col_map.get("Amazon")
                    
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
                                
                            # Create unique key for this specific field
                            field_key = f"{row_key}:{date_col}"
                            
                            # Check if this is a new or changed value
                            if field_key in state["processed"] and state["processed"][field_key] == date_val:
                                continue
                                
                            # Parse date
                            parsed_date = parse_date(date_val)
                            if not parsed_date:
                                logger.warning(f"Could not parse date: {date_val} for {field_key}")
                                continue
                                
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
    
    logger.info(f"Change tracking completed. Found {changes_found} changes.")
    return True

# The bootstrap function will initialize the system without historical data
def bootstrap_tracking():
    """Initialize tracking without historical data."""
    logger.info("Starting fresh bootstrap (no historical data)")
    
    # Reset state to start fresh
    state = {"last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "processed": {}}
    save_state(state)
    
    # Create empty changes file
    ensure_changes_file()
    
    logger.info("Bootstrap complete. System ready to track new changes.")
    return True
    
    # Create new file
    ensure_changes_file()
    
    # Run tracking
    return track_changes()

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Track changes in Smartsheet tables")
    parser.add_argument("--bootstrap", action="store_true", help="Initialize tracking for all data")
    args = parser.parse_args()
    
    if args.bootstrap:
        success = bootstrap_tracking()
    else:
        success = track_changes()
        
    exit(0 if success else 1)

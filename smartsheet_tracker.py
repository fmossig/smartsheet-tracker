import os
import csv
import json
from datetime import datetime, timedelta, date
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

def parse_date(value):
    """Parse date from Smartsheet cell values (string/date/datetime).

    Smartsheet may return date columns as datetime.date / datetime.datetime objects
    or as strings depending on column configuration and SDK behavior.
    """
    if not value:
        return None

    # Accept native date/datetime objects directly
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    # Fall back to string parsing
    cleaned = str(value).strip()

    # Clean up common trailing characters (e.g., accidental suffixes)
    if cleaned and not cleaned[-1].isdigit():
        cleaned = cleaned.rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ')

    # Try various formats
    for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%d', '%d.%m.%Y', '%m/%d/%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue

    try:
        # Try ISO format (catches many variations)
        return datetime.fromisoformat(cleaned).date()
    except Exception:
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

    # Verify state format and structure
    processed = state.get("processed", {})
    if not processed:
        logger.warning("State file has empty or invalid 'processed' dict - may detect ALL changes as new")

    # Force detect one change for testing
    test_mode = False  # Set to True for testing
    if test_mode:
        logger.info("TEST MODE: Will detect at least one change")
        if processed:
            # Remove one random key from processed
            import random
            key_to_remove = random.choice(list(processed.keys()))
            logger.info(f"Removing key {key_to_remove} for test")
            processed.pop(key_to_remove, None)

    # Current timestamp
    now = datetime.now()

    # Track new changes
    changes_found = 0

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

                # Check which phase fields exist in this sheet
                found_fields = [date_col for date_col, user_col, _ in PHASE_FIELDS if date_col in col_map and user_col in col_map]
                logger.info(f"Found {len(found_fields)} phase fields in {group}: {found_fields}")

                # Process each row
                for row in sheet.rows:
                    # Get marketplace if available
                    marketplace = ""
                    if amazon_col_id:
                        for cell in row.cells:
                            if cell.column_id == amazon_col_id:
                                marketplace = (cell.display_value or "").strip()
                                break

                    # Check each phase field
                    for date_col, user_col, phase_no in PHASE_FIELDS:
                        if date_col not in col_map or user_col not in col_map:
                            continue

                        date_cell = None
                        user_cell = None

                        # Get date and user values
                        for cell in row.cells:
                            if cell.column_id == col_map[date_col]:
                                date_cell = cell
                            if cell.column_id == col_map[user_col]:
                                user_cell = cell

                        # Skip if no date value
                        if not date_cell or not date_cell.value:
                            continue

                        date_val = date_cell.value
                        user_val = user_cell.display_value if user_cell else ""

                        # Create unique key for this field
                        field_key = f"{group}:{row.id}:{date_col}"

                        # Check if changed - EXPLICITLY CAST BOTH TO STRING FOR COMPARISON
                        str_date_val = str(date_val).strip()
                        prev_val = state["processed"].get(field_key)
                        str_prev_val = str(prev_val).strip() if prev_val is not None else None

                        if str_prev_val == str_date_val:
                            continue

                        # Only log detailed info for changes
                        logger.info(f"Change detected in {field_key}")
                        logger.info(f"  Previous: '{prev_val}'")
                        logger.info(f"  Current:  '{date_val}'")

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

                        # Update state for this field (JSON-safe)
                        state["processed"][field_key] = str_date_val

                        changes_found += 1

            except Exception as e:
                logger.error(f"Error processing sheet {group}: {e}")
                continue

    # Update state
    state["last_run"] = now.strftime("%Y-%m-%d %H:%M:%S")
    save_state(state)

    logger.info(f"Change tracking completed. Found {changes_found} changes.")
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
                        state["processed"][field_key] = str(cell.value).strip()
                        break

    # Save state
    save_state(state)

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

def bootstrap_tracking(days_back=0):
    """Initialize tracking for new data only."""
    logger.info(f"Starting bootstrap (tracking new data only)")

    # Reset state to force reprocessing
    state = {"last_run": None, "processed": {}}
    save_state(state)

    # Run tracking
    return track_changes()

def test_changes():
    """Test function to force detect at least one change."""
    logger.info("Testing change detection...")

    # Load state
    state = load_state()
    processed = state.get("processed", {})

    if not processed:
        logger.error("No processed items found in state file. Run reset first.")
        return False

    # Remove one item to force detection
    import random
    key_to_remove = random.choice(list(processed.keys()))
    logger.info(f"Removing key {key_to_remove} to force change detection")
    processed.pop(key_to_remove, None)

    # Save modified state
    save_state(state)
    logger.info("Test modification saved. Now run tracking to detect the forced change.")
    return True

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Smartsheet Change Tracker")
    parser.add_argument("--bootstrap", action="store_true", help="Initialize tracking for all data")
    parser.add_argument("--reset", action="store_true", help="Reset tracking state to current data")
    parser.add_argument("--test", action="store_true", help="Test change detection by forcing changes")
    args = parser.parse_args()

    if args.reset:
        success = reset_tracking_state()
    elif args.bootstrap:
        success = bootstrap_tracking()
    elif args.test:
        success = test_changes()
    else:
        success = track_changes()

    exit(0 if success else 1)

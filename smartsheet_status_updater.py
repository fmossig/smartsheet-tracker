#!/usr/bin/env python3
"""
Smartsheet Status Updater

Pushes automated status updates and weekly statistics to Smartsheet.
Creates and maintains two sheets:
1. Status Updates - Real-time tracking of system runs
2. Weekly Stats - Aggregated metrics per week

Usage:
    python smartsheet_status_updater.py --status          # Push status update
    python smartsheet_status_updater.py --weekly-stats    # Push weekly statistics
    python smartsheet_status_updater.py --setup           # Create/setup sheets
"""

import os
import csv
import json
import argparse
import logging
from datetime import datetime, date, timedelta
from collections import defaultdict

import smartsheet
from dotenv import load_dotenv

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("smartsheet_status.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")

# Configuration - Sheet IDs for status and stats tracking
STATUS_SHEET_ID = 1175618067582852
WEEKLY_STATS_SHEET_ID = 5679217694953348

# Data files
DATA_DIR = "tracking_data"
CHANGES_FILE = os.path.join(DATA_DIR, "change_history.csv")
STATE_FILE = os.path.join(DATA_DIR, "tracker_state.json")

# Sheet configurations
STATUS_SHEET_NAME = "Amazon Content Management - System Status"
WEEKLY_STATS_SHEET_NAME = "Amazon Content Management - Weekly Stats"

STATUS_COLUMNS = [
    {"title": "Timestamp", "type": "TEXT_NUMBER", "width": 150},
    {"title": "Run Type", "type": "TEXT_NUMBER", "width": 100},
    {"title": "Status", "type": "TEXT_NUMBER", "width": 80},
    {"title": "Changes Detected", "type": "TEXT_NUMBER", "width": 120},
    {"title": "Sheets Processed", "type": "TEXT_NUMBER", "width": 120},
    {"title": "Errors", "type": "TEXT_NUMBER", "width": 80},
    {"title": "Duration (sec)", "type": "TEXT_NUMBER", "width": 100},
    {"title": "Details", "type": "TEXT_NUMBER", "width": 300},
]

WEEKLY_STATS_COLUMNS = [
    {"title": "Week", "type": "TEXT_NUMBER", "width": 100},
    {"title": "Start Date", "type": "DATE", "width": 100},
    {"title": "End Date", "type": "DATE", "width": 100},
    {"title": "Total Changes", "type": "TEXT_NUMBER", "width": 100},
    {"title": "Active Users", "type": "TEXT_NUMBER", "width": 100},
    {"title": "Active Groups", "type": "TEXT_NUMBER", "width": 100},
    # Per-user columns (will be added dynamically)
    {"title": "DM", "type": "TEXT_NUMBER", "width": 60},
    {"title": "EK", "type": "TEXT_NUMBER", "width": 60},
    {"title": "HI", "type": "TEXT_NUMBER", "width": 60},
    {"title": "JHU", "type": "TEXT_NUMBER", "width": 60},
    {"title": "LK", "type": "TEXT_NUMBER", "width": 60},
    {"title": "SM", "type": "TEXT_NUMBER", "width": 60},
    # Per-group columns
    {"title": "NA", "type": "TEXT_NUMBER", "width": 60},
    {"title": "NF", "type": "TEXT_NUMBER", "width": 60},
    {"title": "NH", "type": "TEXT_NUMBER", "width": 60},
    {"title": "NM", "type": "TEXT_NUMBER", "width": 60},
    {"title": "NP", "type": "TEXT_NUMBER", "width": 60},
    {"title": "NT", "type": "TEXT_NUMBER", "width": 60},
    {"title": "NV", "type": "TEXT_NUMBER", "width": 60},
    # Per-phase columns
    {"title": "Phase 1", "type": "TEXT_NUMBER", "width": 70},
    {"title": "Phase 2", "type": "TEXT_NUMBER", "width": 70},
    {"title": "Phase 3", "type": "TEXT_NUMBER", "width": 70},
    {"title": "Phase 4", "type": "TEXT_NUMBER", "width": 70},
    {"title": "Phase 5", "type": "TEXT_NUMBER", "width": 70},
    {"title": "Report Generated", "type": "CHECKBOX", "width": 120},
    {"title": "Notes", "type": "TEXT_NUMBER", "width": 200},
]


def get_client():
    """Get authenticated Smartsheet client."""
    if not token:
        logger.error("SMARTSHEET_TOKEN not found in environment")
        return None
    
    try:
        client = smartsheet.Smartsheet(token)
        client.errors_as_exceptions(True)
        return client
    except Exception as e:
        logger.error(f"Failed to connect to Smartsheet: {e}")
        return None


def create_sheet(client, name, columns):
    """Create a new sheet with specified columns."""
    try:
        # Build column specifications
        col_specs = []
        for i, col in enumerate(columns):
            spec = smartsheet.models.Column({
                'title': col['title'],
                'type': col['type'],
                'width': col.get('width', 100),
                'primary': (i == 0)  # First column is primary
            })
            col_specs.append(spec)
        
        # Create sheet
        sheet_spec = smartsheet.models.Sheet({
            'name': name,
            'columns': col_specs
        })
        
        response = client.Home.create_sheet(sheet_spec)
        sheet = response.result
        
        logger.info(f"Created sheet '{name}' with ID: {sheet.id}")
        return sheet.id
        
    except Exception as e:
        logger.error(f"Failed to create sheet '{name}': {e}")
        return None


def get_column_map(client, sheet_id):
    """Get mapping of column titles to IDs."""
    try:
        sheet = client.Sheets.get_sheet(sheet_id)
        return {col.title: col.id for col in sheet.columns}
    except Exception as e:
        logger.error(f"Failed to get column map: {e}")
        return {}


def setup_sheets():
    """Create status and weekly stats sheets if they don't exist."""
    client = get_client()
    if not client:
        return False
    
    created = {}
    
    # Create Status Sheet
    if not STATUS_SHEET_ID:
        sheet_id = create_sheet(client, STATUS_SHEET_NAME, STATUS_COLUMNS)
        if sheet_id:
            created['STATUS_SHEET_ID'] = sheet_id
    else:
        logger.info(f"Status sheet already configured: {STATUS_SHEET_ID}")
    
    # Create Weekly Stats Sheet
    if not WEEKLY_STATS_SHEET_ID:
        sheet_id = create_sheet(client, WEEKLY_STATS_SHEET_NAME, WEEKLY_STATS_COLUMNS)
        if sheet_id:
            created['WEEKLY_STATS_SHEET_ID'] = sheet_id
    else:
        logger.info(f"Weekly stats sheet already configured: {WEEKLY_STATS_SHEET_ID}")
    
    if created:
        print("\n" + "=" * 60)
        print("SHEETS CREATED - Add these to your .env or GitHub Secrets:")
        print("=" * 60)
        for key, value in created.items():
            print(f"{key}={value}")
        print("=" * 60 + "\n")
    
    return True


def push_status_update(run_type="tracking", status="success", changes_detected=0,
                       sheets_processed=0, errors=0, duration=0, details=""):
    """Push a status update row to the Status sheet."""
    if not STATUS_SHEET_ID:
        logger.warning("STATUS_SHEET_ID not configured. Run --setup first.")
        return False
    
    client = get_client()
    if not client:
        return False
    
    try:
        col_map = get_column_map(client, int(STATUS_SHEET_ID))
        if not col_map:
            return False
        
        # Build row
        new_row = smartsheet.models.Row()
        new_row.to_top = True  # Add to top of sheet
        
        cells = [
            {"column_id": col_map.get("Timestamp"), "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            {"column_id": col_map.get("Run Type"), "value": run_type},
            {"column_id": col_map.get("Status"), "value": status},
            {"column_id": col_map.get("Changes Detected"), "value": changes_detected},
            {"column_id": col_map.get("Sheets Processed"), "value": sheets_processed},
            {"column_id": col_map.get("Errors"), "value": errors},
            {"column_id": col_map.get("Duration (sec)"), "value": round(duration, 2)},
            {"column_id": col_map.get("Details"), "value": details[:500] if details else ""},
        ]
        
        for cell in cells:
            if cell["column_id"]:
                new_row.cells.append(cell)
        
        response = client.Sheets.add_rows(int(STATUS_SHEET_ID), [new_row])
        logger.info(f"Status update pushed: {run_type} - {status}")
        return True
        
    except Exception as e:
        logger.error(f"Failed to push status update: {e}")
        return False


def load_changes(start_date, end_date):
    """Load changes from CSV within date range."""
    changes = []
    
    if not os.path.exists(CHANGES_FILE):
        return changes
    
    try:
        with open(CHANGES_FILE, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    # Parse the date from the Date field
                    change_date = datetime.strptime(row['Date'], '%Y-%m-%d').date()
                    if start_date <= change_date <= end_date:
                        row['ParsedDate'] = change_date
                        changes.append(row)
                except (ValueError, KeyError):
                    continue
    except Exception as e:
        logger.error(f"Error loading changes: {e}")
    
    return changes


def calculate_weekly_stats(start_date, end_date):
    """Calculate statistics for a week."""
    changes = load_changes(start_date, end_date)
    
    stats = {
        "total_changes": len(changes),
        "users": defaultdict(int),
        "groups": defaultdict(int),
        "phases": defaultdict(int),
    }
    
    for change in changes:
        user = change.get('User', '').strip()
        group = change.get('Group', '').strip()
        phase = change.get('Phase', '').strip()
        
        if user:
            stats["users"][user] += 1
        if group:
            stats["groups"][group] += 1
        if phase:
            stats["phases"][f"Phase {phase}"] += 1
    
    stats["active_users"] = len(stats["users"])
    stats["active_groups"] = len(stats["groups"])
    
    return stats


def push_weekly_stats(year=None, week=None):
    """Push weekly statistics to the Weekly Stats sheet."""
    if not WEEKLY_STATS_SHEET_ID:
        logger.warning("WEEKLY_STATS_SHEET_ID not configured. Run --setup first.")
        return False
    
    client = get_client()
    if not client:
        return False
    
    # Calculate week dates
    if year is None or week is None:
        # Default to previous week
        today = date.today()
        # Get Monday of current week, then go back one week
        current_monday = today - timedelta(days=today.weekday())
        start_date = current_monday - timedelta(weeks=1)
        end_date = start_date + timedelta(days=6)
        year = start_date.isocalendar()[0]
        week = start_date.isocalendar()[1]
    else:
        # Calculate dates from year/week
        start_date = date.fromisocalendar(year, week, 1)
        end_date = start_date + timedelta(days=6)
    
    week_str = f"{year}-W{week:02d}"
    logger.info(f"Calculating stats for {week_str} ({start_date} to {end_date})")
    
    # Get statistics
    stats = calculate_weekly_stats(start_date, end_date)
    
    try:
        col_map = get_column_map(client, int(WEEKLY_STATS_SHEET_ID))
        if not col_map:
            return False
        
        # Check if row for this week already exists
        sheet = client.Sheets.get_sheet(int(WEEKLY_STATS_SHEET_ID))
        week_col_id = col_map.get("Week")
        existing_row_id = None
        
        if week_col_id:
            for row in sheet.rows:
                for cell in row.cells:
                    if cell.column_id == week_col_id and cell.value == week_str:
                        existing_row_id = row.id
                        break
                if existing_row_id:
                    break
        
        # Build cells
        cells = [
            {"column_id": col_map.get("Week"), "value": week_str},
            {"column_id": col_map.get("Start Date"), "value": start_date.isoformat()},
            {"column_id": col_map.get("End Date"), "value": end_date.isoformat()},
            {"column_id": col_map.get("Total Changes"), "value": stats["total_changes"]},
            {"column_id": col_map.get("Active Users"), "value": stats["active_users"]},
            {"column_id": col_map.get("Active Groups"), "value": stats["active_groups"]},
        ]
        
        # Add per-user stats
        for user in ["DM", "EK", "HI", "JHU", "LK", "SM"]:
            col_id = col_map.get(user)
            if col_id:
                cells.append({"column_id": col_id, "value": stats["users"].get(user, 0)})
        
        # Add per-group stats
        for group in ["NA", "NF", "NH", "NM", "NP", "NT", "NV"]:
            col_id = col_map.get(group)
            if col_id:
                cells.append({"column_id": col_id, "value": stats["groups"].get(group, 0)})
        
        # Add per-phase stats
        for phase in ["Phase 1", "Phase 2", "Phase 3", "Phase 4", "Phase 5"]:
            col_id = col_map.get(phase)
            if col_id:
                cells.append({"column_id": col_id, "value": stats["phases"].get(phase, 0)})
        
        # Filter out cells with None column_id
        cells = [c for c in cells if c["column_id"]]
        
        if existing_row_id:
            # Update existing row
            update_row = smartsheet.models.Row()
            update_row.id = existing_row_id
            update_row.cells = cells
            
            response = client.Sheets.update_rows(int(WEEKLY_STATS_SHEET_ID), [update_row])
            logger.info(f"Updated weekly stats for {week_str}")
        else:
            # Create new row
            new_row = smartsheet.models.Row()
            new_row.to_top = True
            new_row.cells = cells
            
            response = client.Sheets.add_rows(int(WEEKLY_STATS_SHEET_ID), [new_row])
            logger.info(f"Added weekly stats for {week_str}")
        
        # Also push a status update
        push_status_update(
            run_type="weekly_stats",
            status="success",
            changes_detected=stats["total_changes"],
            details=f"Week {week_str}: {stats['active_users']} users, {stats['active_groups']} groups"
        )
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to push weekly stats: {e}")
        return False


def mark_report_generated(year, week):
    """Mark a week's report as generated in the Weekly Stats sheet."""
    if not WEEKLY_STATS_SHEET_ID:
        return False
    
    client = get_client()
    if not client:
        return False
    
    week_str = f"{year}-W{week:02d}"
    
    try:
        col_map = get_column_map(client, int(WEEKLY_STATS_SHEET_ID))
        sheet = client.Sheets.get_sheet(int(WEEKLY_STATS_SHEET_ID))
        
        week_col_id = col_map.get("Week")
        report_col_id = col_map.get("Report Generated")
        
        if not week_col_id or not report_col_id:
            return False
        
        # Find the row for this week
        for row in sheet.rows:
            for cell in row.cells:
                if cell.column_id == week_col_id and cell.value == week_str:
                    # Update the Report Generated checkbox
                    update_row = smartsheet.models.Row()
                    update_row.id = row.id
                    update_row.cells = [{"column_id": report_col_id, "value": True}]
                    
                    client.Sheets.update_rows(int(WEEKLY_STATS_SHEET_ID), [update_row])
                    logger.info(f"Marked report as generated for {week_str}")
                    return True
        
        logger.warning(f"No row found for {week_str}")
        return False
        
    except Exception as e:
        logger.error(f"Failed to mark report generated: {e}")
        return False


def get_tracking_summary():
    """Get a summary of current tracking state."""
    summary = {
        "last_run": None,
        "total_tracked_items": 0,
        "total_changes_recorded": 0,
    }
    
    # Load state
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                state = json.load(f)
                summary["last_run"] = state.get("last_run")
                summary["total_tracked_items"] = len(state.get("processed", {}))
        except:
            pass
    
    # Count changes
    if os.path.exists(CHANGES_FILE):
        try:
            with open(CHANGES_FILE, 'r', encoding='utf-8') as f:
                summary["total_changes_recorded"] = sum(1 for _ in f) - 1  # Subtract header
        except:
            pass
    
    return summary


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Smartsheet Status Updater")
    parser.add_argument("--setup", action="store_true", help="Create status and stats sheets")
    parser.add_argument("--status", action="store_true", help="Push a status update")
    parser.add_argument("--weekly-stats", action="store_true", help="Push weekly statistics")
    parser.add_argument("--year", type=int, help="Year for weekly stats")
    parser.add_argument("--week", type=int, help="Week number for weekly stats")
    parser.add_argument("--run-type", default="manual", help="Run type for status update")
    parser.add_argument("--changes", type=int, default=0, help="Number of changes detected")
    parser.add_argument("--mark-report", action="store_true", help="Mark report as generated")
    
    args = parser.parse_args()
    
    if args.setup:
        setup_sheets()
    elif args.status:
        summary = get_tracking_summary()
        push_status_update(
            run_type=args.run_type,
            status="success",
            changes_detected=args.changes,
            sheets_processed=7,
            details=f"Tracked items: {summary['total_tracked_items']}, Total changes: {summary['total_changes_recorded']}"
        )
    elif args.weekly_stats:
        push_weekly_stats(args.year, args.week)
    elif args.mark_report:
        if args.year and args.week:
            mark_report_generated(args.year, args.week)
        else:
            # Default to previous week
            today = date.today()
            current_monday = today - timedelta(days=today.weekday())
            prev_monday = current_monday - timedelta(weeks=1)
            mark_report_generated(prev_monday.year, prev_monday.isocalendar()[1])
    else:
        parser.print_help()

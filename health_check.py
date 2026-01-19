#!/usr/bin/env python3
"""
Health Check Script for Smartsheet Tracker

Verifies system health by checking:
1. Configuration validity
2. Smartsheet API connectivity
3. State file integrity
4. Change history file status
5. Required directories

Usage:
    python health_check.py           # Run all checks
    python health_check.py --verbose # Verbose output
    python health_check.py --json    # JSON output for automation
"""

import os
import sys
import json
import argparse
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional

import smartsheet
from dotenv import load_dotenv

# Import centralized configuration
from config import (
    SHEET_IDS,
    PHASE_FIELDS,
    DATA_DIR,
    STATE_FILE,
    CHANGES_FILE,
    REPORTS_DIR,
    STATUS_SHEET_ID,
    WEEKLY_STATS_SHEET_ID,
    DAILY_STATS_SHEET_ID,
    USERS,
    get_product_groups,
    ensure_directories,
)

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class HealthCheck:
    """Health check runner for the Smartsheet Tracker system."""

    def __init__(self, verbose: bool = False):
        self.verbose = verbose
        self.results: Dict[str, Any] = {
            "timestamp": datetime.now().isoformat(),
            "checks": {},
            "overall_status": "unknown",
            "warnings": [],
            "errors": [],
        }

    def log(self, message: str, level: str = "info"):
        """Log a message if verbose mode is enabled."""
        if self.verbose:
            if level == "info":
                logger.info(message)
            elif level == "warning":
                logger.warning(message)
            elif level == "error":
                logger.error(message)

    def check_environment(self) -> bool:
        """Check environment variables."""
        self.log("Checking environment variables...")

        token = os.getenv("SMARTSHEET_TOKEN")
        if not token:
            self.results["errors"].append("SMARTSHEET_TOKEN not found in environment")
            self.results["checks"]["environment"] = {
                "status": "failed",
                "message": "SMARTSHEET_TOKEN not found"
            }
            return False

        self.results["checks"]["environment"] = {
            "status": "passed",
            "message": "Environment variables configured"
        }
        return True

    def check_directories(self) -> bool:
        """Check required directories exist."""
        self.log("Checking directories...")

        directories = [DATA_DIR, REPORTS_DIR]
        missing = []

        for directory in directories:
            if not os.path.exists(directory):
                missing.append(directory)

        if missing:
            self.results["warnings"].append(f"Missing directories: {missing}")
            self.results["checks"]["directories"] = {
                "status": "warning",
                "message": f"Missing directories: {missing}",
                "missing": missing
            }
            # Try to create them
            try:
                ensure_directories()
                self.log(f"Created missing directories: {missing}")
            except Exception as e:
                self.results["errors"].append(f"Failed to create directories: {e}")
                return False
        else:
            self.results["checks"]["directories"] = {
                "status": "passed",
                "message": "All directories exist"
            }

        return True

    def check_state_file(self) -> bool:
        """Check state file integrity."""
        self.log("Checking state file...")

        if not os.path.exists(STATE_FILE):
            self.results["warnings"].append("State file not found - tracking may not be initialized")
            self.results["checks"]["state_file"] = {
                "status": "warning",
                "message": "State file not found"
            }
            return True  # Not a failure, just needs initialization

        try:
            with open(STATE_FILE, 'r') as f:
                state = json.load(f)

            last_run = state.get("last_run")
            processed_count = len(state.get("processed", {}))

            # Check if last run was recent (within 25 hours for daily runs)
            is_recent = False
            if last_run:
                try:
                    last_run_dt = datetime.strptime(last_run, "%Y-%m-%d %H:%M:%S")
                    hours_ago = (datetime.now() - last_run_dt).total_seconds() / 3600
                    is_recent = hours_ago < 25
                except ValueError:
                    pass

            self.results["checks"]["state_file"] = {
                "status": "passed",
                "message": "State file is valid",
                "last_run": last_run,
                "processed_items": processed_count,
                "is_recent": is_recent
            }

            if not is_recent and last_run:
                self.results["warnings"].append(f"Last tracking run was more than 25 hours ago: {last_run}")

            return True

        except json.JSONDecodeError as e:
            self.results["errors"].append(f"State file is corrupted: {e}")
            self.results["checks"]["state_file"] = {
                "status": "failed",
                "message": f"State file is corrupted: {e}"
            }
            return False
        except Exception as e:
            self.results["errors"].append(f"Error reading state file: {e}")
            self.results["checks"]["state_file"] = {
                "status": "failed",
                "message": f"Error reading state file: {e}"
            }
            return False

    def check_changes_file(self) -> bool:
        """Check change history file."""
        self.log("Checking change history file...")

        if not os.path.exists(CHANGES_FILE):
            self.results["warnings"].append("Change history file not found")
            self.results["checks"]["changes_file"] = {
                "status": "warning",
                "message": "Change history file not found"
            }
            return True

        try:
            with open(CHANGES_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            total_records = len(lines) - 1  # Subtract header

            # Get recent changes (last 7 days)
            recent_count = 0
            cutoff = datetime.now() - timedelta(days=7)

            for line in lines[1:]:  # Skip header
                parts = line.strip().split(',')
                if len(parts) >= 6:
                    try:
                        change_date = datetime.strptime(parts[5], '%Y-%m-%d')
                        if change_date >= cutoff:
                            recent_count += 1
                    except ValueError:
                        pass

            self.results["checks"]["changes_file"] = {
                "status": "passed",
                "message": "Change history file is valid",
                "total_records": total_records,
                "recent_changes_7d": recent_count
            }
            return True

        except Exception as e:
            self.results["errors"].append(f"Error reading change history: {e}")
            self.results["checks"]["changes_file"] = {
                "status": "failed",
                "message": f"Error reading change history: {e}"
            }
            return False

    def check_smartsheet_api(self) -> bool:
        """Check Smartsheet API connectivity."""
        self.log("Checking Smartsheet API connectivity...")

        token = os.getenv("SMARTSHEET_TOKEN")
        if not token:
            self.results["checks"]["smartsheet_api"] = {
                "status": "skipped",
                "message": "No token available"
            }
            return False

        try:
            client = smartsheet.Smartsheet(token)
            client.errors_as_exceptions(True)

            # Try to get current user info (lightweight API call)
            user = client.Users.get_current_user()

            self.results["checks"]["smartsheet_api"] = {
                "status": "passed",
                "message": "API connection successful",
                "user_email": user.email if hasattr(user, 'email') else "unknown"
            }
            return True

        except Exception as e:
            self.results["errors"].append(f"Smartsheet API error: {e}")
            self.results["checks"]["smartsheet_api"] = {
                "status": "failed",
                "message": f"API connection failed: {e}"
            }
            return False

    def check_sheet_access(self) -> bool:
        """Check access to all configured sheets."""
        self.log("Checking sheet access...")

        token = os.getenv("SMARTSHEET_TOKEN")
        if not token:
            self.results["checks"]["sheet_access"] = {
                "status": "skipped",
                "message": "No token available"
            }
            return False

        try:
            client = smartsheet.Smartsheet(token)
            client.errors_as_exceptions(True)

            accessible = []
            inaccessible = []

            # Check product group sheets
            for group, sheet_id in SHEET_IDS.items():
                try:
                    sheet = client.Sheets.get_sheet(sheet_id, page_size=1)
                    accessible.append({
                        "name": group,
                        "id": sheet_id,
                        "rows": sheet.total_row_count if hasattr(sheet, 'total_row_count') else len(sheet.rows)
                    })
                except Exception as e:
                    inaccessible.append({
                        "name": group,
                        "id": sheet_id,
                        "error": str(e)
                    })

            if inaccessible:
                self.results["warnings"].append(f"Cannot access {len(inaccessible)} sheets")
                self.results["checks"]["sheet_access"] = {
                    "status": "warning",
                    "message": f"Cannot access {len(inaccessible)} of {len(SHEET_IDS)} sheets",
                    "accessible": accessible,
                    "inaccessible": inaccessible
                }
            else:
                self.results["checks"]["sheet_access"] = {
                    "status": "passed",
                    "message": f"All {len(SHEET_IDS)} sheets accessible",
                    "sheets": accessible
                }

            return len(inaccessible) == 0

        except Exception as e:
            self.results["errors"].append(f"Sheet access check failed: {e}")
            self.results["checks"]["sheet_access"] = {
                "status": "failed",
                "message": f"Check failed: {e}"
            }
            return False

    def check_configuration(self) -> bool:
        """Check configuration validity."""
        self.log("Checking configuration...")

        issues = []

        # Check SHEET_IDS
        if not SHEET_IDS:
            issues.append("No sheet IDs configured")
        elif len(SHEET_IDS) < 7:
            issues.append(f"Only {len(SHEET_IDS)} sheet IDs configured (expected 7)")

        # Check PHASE_FIELDS
        if not PHASE_FIELDS:
            issues.append("No phase fields configured")

        # Check USERS
        if not USERS:
            issues.append("No users configured")

        if issues:
            self.results["warnings"].extend(issues)
            self.results["checks"]["configuration"] = {
                "status": "warning",
                "message": "Configuration has issues",
                "issues": issues
            }
        else:
            self.results["checks"]["configuration"] = {
                "status": "passed",
                "message": "Configuration is valid",
                "sheet_count": len(SHEET_IDS),
                "phase_count": len(PHASE_FIELDS),
                "user_count": len(USERS),
                "product_groups": get_product_groups()
            }

        return len(issues) == 0

    def run_all_checks(self) -> Dict[str, Any]:
        """Run all health checks and return results."""
        self.log("Starting health checks...")

        checks = [
            ("environment", self.check_environment),
            ("directories", self.check_directories),
            ("configuration", self.check_configuration),
            ("state_file", self.check_state_file),
            ("changes_file", self.check_changes_file),
            ("smartsheet_api", self.check_smartsheet_api),
            ("sheet_access", self.check_sheet_access),
        ]

        all_passed = True
        for name, check_func in checks:
            try:
                result = check_func()
                if not result and name not in ["sheet_access"]:  # sheet_access can be partial
                    all_passed = False
            except Exception as e:
                self.results["errors"].append(f"Check '{name}' failed with exception: {e}")
                self.results["checks"][name] = {
                    "status": "failed",
                    "message": f"Exception: {e}"
                }
                all_passed = False

        # Determine overall status
        if self.results["errors"]:
            self.results["overall_status"] = "unhealthy"
        elif self.results["warnings"]:
            self.results["overall_status"] = "degraded"
        else:
            self.results["overall_status"] = "healthy"

        self.log(f"Health check complete: {self.results['overall_status']}")

        return self.results


def print_results(results: Dict[str, Any], verbose: bool = False):
    """Print health check results in human-readable format."""
    print("\n" + "=" * 60)
    print("SMARTSHEET TRACKER HEALTH CHECK")
    print("=" * 60)
    print(f"Timestamp: {results['timestamp']}")
    print(f"Overall Status: {results['overall_status'].upper()}")
    print("-" * 60)

    for check_name, check_result in results["checks"].items():
        status = check_result.get("status", "unknown")
        message = check_result.get("message", "")

        status_icon = {
            "passed": "[OK]",
            "warning": "[!]",
            "failed": "[X]",
            "skipped": "[-]"
        }.get(status, "[?]")

        print(f"{status_icon} {check_name}: {message}")

        if verbose and isinstance(check_result, dict):
            for key, value in check_result.items():
                if key not in ["status", "message"]:
                    print(f"      {key}: {value}")

    if results["warnings"]:
        print("-" * 60)
        print("WARNINGS:")
        for warning in results["warnings"]:
            print(f"  ! {warning}")

    if results["errors"]:
        print("-" * 60)
        print("ERRORS:")
        for error in results["errors"]:
            print(f"  X {error}")

    print("=" * 60)


def main():
    """Main entry point for health check."""
    parser = argparse.ArgumentParser(description="Smartsheet Tracker Health Check")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--json", "-j", action="store_true", help="Output as JSON")
    args = parser.parse_args()

    checker = HealthCheck(verbose=args.verbose)
    results = checker.run_all_checks()

    if args.json:
        print(json.dumps(results, indent=2))
    else:
        print_results(results, verbose=args.verbose)

    # Exit with appropriate code
    if results["overall_status"] == "healthy":
        sys.exit(0)
    elif results["overall_status"] == "degraded":
        sys.exit(1)
    else:
        sys.exit(2)


if __name__ == "__main__":
    main()

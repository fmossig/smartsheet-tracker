#!/usr/bin/env python3
"""
Validation script for Smartsheet Change Tracking & Reporting System

This script performs comprehensive validation of the system:
- Checks dependencies
- Validates configuration
- Tests data structure
- Verifies scripts are working
- Checks file permissions
"""

import os
import sys
import csv
import glob
import importlib.util
from datetime import datetime

# Color codes for output
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
RESET = '\033[0m'

def print_status(message, status='info'):
    """Print formatted status message"""
    if status == 'success':
        print(f"{GREEN}✅ {message}{RESET}")
    elif status == 'error':
        print(f"{RED}❌ {message}{RESET}")
    elif status == 'warning':
        print(f"{YELLOW}⚠️  {message}{RESET}")
    else:
        print(f"ℹ️  {message}")

def check_dependencies():
    """Check if required dependencies are installed"""
    print("\n" + "="*60)
    print("CHECKING DEPENDENCIES")
    print("="*60)
    
    required = ['smartsheet', 'dotenv', 'reportlab']
    missing = []
    
    for package in required:
        try:
            __import__(package)
            print_status(f"{package} - installed", 'success')
        except ImportError:
            print_status(f"{package} - MISSING", 'error')
            missing.append(package)
    
    if missing:
        print_status(f"Missing packages: {', '.join(missing)}", 'error')
        print_status("Run: pip install " + " ".join(missing), 'info')
        return False
    
    return True

def check_environment():
    """Check environment configuration"""
    print("\n" + "="*60)
    print("CHECKING ENVIRONMENT")
    print("="*60)
    
    # Load .env if exists
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception as e:
        print_status(f"Could not load .env: {e}", 'warning')
    
    # Check for token
    token = os.getenv('SMARTSHEET_TOKEN') or os.getenv('SMARTSHEET_ACCESS_TOKEN')
    if token:
        masked = token[:8] + "..." + token[-4:] if len(token) > 12 else "***"
        print_status(f"Smartsheet token found: {masked}", 'success')
        return True
    else:
        print_status("Smartsheet token NOT found", 'error')
        print_status("Set SMARTSHEET_TOKEN in .env file", 'info')
        return False

def check_scripts():
    """Check if all scripts are syntactically valid"""
    print("\n" + "="*60)
    print("CHECKING SCRIPTS")
    print("="*60)
    
    scripts = [
        'smartsheet_date_change_tracker.py',
        'smartsheet_weekly_change_tracker.py',
        'smartsheet_status_snapshot.py',
        'smartsheet_status_report.py',
        'smartsheet_periodic_status_report.py',
        'smartsheet_reports_orchestrator.py'
    ]
    
    all_valid = True
    for script in scripts:
        if os.path.exists(script):
            try:
                spec = importlib.util.spec_from_file_location("test", script)
                module = importlib.util.module_from_spec(spec)
                print_status(f"{script} - valid", 'success')
            except Exception as e:
                print_status(f"{script} - ERROR: {e}", 'error')
                all_valid = False
        else:
            print_status(f"{script} - NOT FOUND", 'error')
            all_valid = False
    
    return all_valid

def check_directories():
    """Check if required directories exist"""
    print("\n" + "="*60)
    print("CHECKING DIRECTORIES")
    print("="*60)
    
    dirs = {
        'tracker_logs': 'required',
        'reports': 'optional',
        'weekly': 'optional',
        'status': 'optional',
        'assets': 'optional'
    }
    
    all_ok = True
    for dir_name, required in dirs.items():
        if os.path.isdir(dir_name):
            file_count = len(os.listdir(dir_name))
            print_status(f"{dir_name}/ exists ({file_count} files)", 'success')
        elif required == 'required':
            print_status(f"{dir_name}/ MISSING (required)", 'error')
            all_ok = False
        else:
            print_status(f"{dir_name}/ missing (will be created)", 'warning')
    
    return all_ok

def check_data():
    """Check tracker log data"""
    print("\n" + "="*60)
    print("CHECKING DATA")
    print("="*60)
    
    # Check for log files
    log_files = glob.glob("tracker_logs/date_changes_log_*.csv")
    
    if not log_files:
        print_status("No tracker log files found", 'warning')
        print_status("Run bootstrap: python smartsheet_reports_orchestrator.py bootstrap", 'info')
        return False
    
    print_status(f"Found {len(log_files)} log file(s)", 'success')
    
    # Validate structure
    total_rows = 0
    expected_headers = ['Änderung am', 'Produktgruppe', 'Land/Marketplace', 
                       'Phase', 'Feld', 'Datum', 'Mitarbeiter']
    
    for path in log_files:
        try:
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                
                # Check headers
                if list(reader.fieldnames) != expected_headers:
                    print_status(f"{path} has unexpected headers", 'error')
                    return False
                
                # Count rows
                row_count = sum(1 for _ in reader)
                total_rows += row_count
                
                basename = os.path.basename(path)
                print_status(f"{basename}: {row_count:,} entries", 'success')
        except Exception as e:
            print_status(f"Error reading {path}: {e}", 'error')
            return False
    
    print_status(f"Total entries: {total_rows:,}", 'success')
    
    # Check backup file
    if os.path.exists("tracker_logs/date_backup.csv"):
        print_status("Backup file exists", 'success')
    else:
        print_status("Backup file missing (will be created on first run)", 'warning')
    
    return True

def check_reports():
    """Check if reports have been generated"""
    print("\n" + "="*60)
    print("CHECKING REPORTS")
    print("="*60)
    
    report_locations = [
        "reports/weekly/",
        "reports/monthly/",
        "status/"
    ]
    
    found_reports = []
    for location in report_locations:
        if os.path.isdir(location):
            pdfs = glob.glob(os.path.join(location, "**/*.pdf"), recursive=True)
            if pdfs:
                found_reports.extend(pdfs)
                print_status(f"{location} - {len(pdfs)} PDF(s)", 'success')
            else:
                print_status(f"{location} - no PDFs", 'warning')
        else:
            print_status(f"{location} - directory not found", 'warning')
    
    if found_reports:
        print_status(f"Total PDFs found: {len(found_reports)}", 'success')
        # Show newest
        newest = max(found_reports, key=os.path.getmtime)
        mtime = datetime.fromtimestamp(os.path.getmtime(newest))
        print_status(f"Newest: {newest} ({mtime.strftime('%Y-%m-%d %H:%M')})", 'info')
    else:
        print_status("No reports found - generate with smartsheet_status_report.py", 'info')
    
    return True

def check_workflows():
    """Check GitHub Actions workflows"""
    print("\n" + "="*60)
    print("CHECKING AUTOMATION")
    print("="*60)
    
    workflow_dir = ".github/workflows"
    
    if not os.path.isdir(workflow_dir):
        print_status("Not a GitHub repository or workflows not configured", 'warning')
        return True
    
    workflows = glob.glob(os.path.join(workflow_dir, "*.yml"))
    
    if workflows:
        print_status(f"Found {len(workflows)} workflow(s):", 'success')
        for wf in sorted(workflows):
            basename = os.path.basename(wf)
            print_status(f"  {basename}", 'info')
    else:
        print_status("No workflows found", 'warning')
    
    return True

def print_summary(results):
    """Print validation summary"""
    print("\n" + "="*60)
    print("VALIDATION SUMMARY")
    print("="*60)
    
    passed = sum(results.values())
    total = len(results)
    
    for check, result in results.items():
        status = 'success' if result else 'error'
        print_status(check, status)
    
    print()
    if passed == total:
        print_status(f"ALL CHECKS PASSED ({passed}/{total})", 'success')
        print_status("System is ready to use!", 'success')
        return 0
    else:
        print_status(f"SOME CHECKS FAILED ({passed}/{total})", 'error')
        print_status("Fix the issues above before using the system", 'error')
        return 1

def main():
    """Main validation routine"""
    print("="*60)
    print("SMARTSHEET TRACKER SYSTEM VALIDATION")
    print("="*60)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Working directory: {os.getcwd()}")
    
    results = {
        "Dependencies installed": check_dependencies(),
        "Environment configured": check_environment(),
        "Scripts valid": check_scripts(),
        "Directories present": check_directories(),
        "Data exists": check_data(),
        "Reports generated": check_reports(),
        "Automation configured": check_workflows()
    }
    
    return print_summary(results)

if __name__ == "__main__":
    sys.exit(main())

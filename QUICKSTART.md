# Quick Start Guide

Get the Smartsheet Change Tracking & Reporting System up and running in 5 minutes.

## Prerequisites

- Python 3.11 or higher
- Smartsheet API access token
- Access to target Smartsheet sheets

## Step 1: Clone and Install

```bash
# Navigate to repository
cd smartsheet-tracker

# Install dependencies
pip install -r requirements.txt
```

## Step 2: Configure

Create a `.env` file in the project root:

```bash
SMARTSHEET_TOKEN=your_access_token_here
```

To get a Smartsheet API token:
1. Log in to Smartsheet
2. Go to Account > Apps & Integrations > API Access
3. Generate new access token
4. Copy token to `.env` file

## Step 3: Bootstrap (First Time Only)

Initialize the system with historical data:

```bash
python smartsheet_reports_orchestrator.py bootstrap
```

This will:
- Load all existing data from Smartsheet
- Create tracker log files for the last 3 months
- Generate initial PDF reports

**Expected output:**
```
Bootstrap: 1234 Zeilen übernommen.
✅ PDF Report erstellt: reports/monthly/2025/status_report_2025-08.pdf
```

## Step 4: Daily Tracking

Set up daily incremental tracking:

```bash
# Run manually for testing
python smartsheet_date_change_tracker.py
```

**Expected output:**
```
✅ Nightly: 42 neue Änderung(en) protokolliert.
```

## Step 5: Generate Reports

### Quick Status Report (Last 30 Days)
```bash
python smartsheet_status_report.py
```

Output: `status/status_report_YYYY-MM-DD.pdf`

### Weekly Report
```bash
python smartsheet_reports_orchestrator.py weekly
```

Output: `reports/weekly/YYYY/YYYY-Wxx/status_report_YYYY-Wxx.pdf`

### Monthly Report
```bash
python smartsheet_reports_orchestrator.py monthly
```

Output: `reports/monthly/YYYY/status_report_YYYY-MM.pdf`

## Common Tasks

### View Latest Changes
```bash
# Check what was tracked today
tail -20 tracker_logs/date_changes_log_$(date +%m.%Y).csv
```

### Generate Custom Period Report
```bash
# Report for August 2025
REPORT_SINCE=2025-08-01 \
REPORT_UNTIL=2025-08-31 \
REPORT_LABEL=2025-08 \
REPORT_OUTDIR=reports/custom \
python smartsheet_status_report.py
```

### Weekly Snapshot
```bash
# Create snapshot of last 7 days
python smartsheet_weekly_change_tracker.py

# View results
cat weekly/weekly_changes.csv | head -10
```

### Status Summary
```bash
# Generate current status
python smartsheet_status_snapshot.py

# View summary
cat status/status_summary_$(date +%Y-%m-%d).csv
```

## Automation (Optional)

### GitHub Actions

If your repository is on GitHub, the workflows are already configured:

1. **Add Secret**: Go to Repository Settings > Secrets and add `SMARTSHEET_TOKEN`

2. **Enable Workflows**: Workflows will run automatically:
   - **Nightly**: Every day at 00:00 UTC (tracks changes)
   - **Weekly**: Every Monday at 01:15 UTC (generates report)
   - **Monthly**: 1st of month at 02:00 UTC (generates report)

3. **Manual Trigger**: Go to Actions tab, select workflow, click "Run workflow"

### Local Cron (Linux/Mac)

Add to crontab (`crontab -e`):

```bash
# Daily tracking at midnight
0 0 * * * cd /path/to/smartsheet-tracker && python smartsheet_date_change_tracker.py

# Weekly report every Monday at 1 AM
0 1 * * 1 cd /path/to/smartsheet-tracker && python smartsheet_reports_orchestrator.py weekly

# Monthly report on 1st at 2 AM
0 2 1 * * cd /path/to/smartsheet-tracker && python smartsheet_reports_orchestrator.py monthly
```

### Windows Task Scheduler

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (daily/weekly/monthly)
4. Action: Start a program
   - Program: `python.exe`
   - Arguments: `C:\path\to\smartsheet_date_change_tracker.py`
   - Start in: `C:\path\to\smartsheet-tracker`

## Verify Installation

Run this test to verify everything works:

```bash
# Test 1: Check dependencies
python -c "import smartsheet, dotenv; print('✅ Dependencies OK')"

# Test 2: Check environment
python -c "from dotenv import load_dotenv; import os; load_dotenv(); print('✅ Token found' if os.getenv('SMARTSHEET_TOKEN') else '❌ Token missing')"

# Test 3: Check Smartsheet connection
python -c "
from dotenv import load_dotenv
import os, smartsheet
load_dotenv()
client = smartsheet.Smartsheet(os.getenv('SMARTSHEET_TOKEN'))
client.Sheets.list_sheets()
print('✅ Smartsheet connection OK')
"

# Test 4: Run tracker (dry run - won't commit)
python smartsheet_date_change_tracker.py
```

All tests passing? You're ready to go! 🚀

## Next Steps

1. **Customize** - Edit sheet IDs and colors to match your setup
2. **Review Reports** - Check the generated PDFs in `reports/` and `status/`
3. **Set Up Automation** - Configure GitHub Actions or cron jobs
4. **Monitor** - Check `tracker_logs/` daily to ensure tracking works

## Troubleshooting

### "ModuleNotFoundError: No module named 'smartsheet'"
```bash
pip install smartsheet-python-sdk python-dotenv
```

### "Token not found" or API errors
- Check `.env` file exists and contains valid token
- Verify token has not expired
- Ensure you have read access to the sheets

### "No data in reports"
- Run bootstrap first: `python smartsheet_reports_orchestrator.py bootstrap`
- Check tracker logs exist: `ls -l tracker_logs/`
- Verify date range covers available data

### "Permission denied" when writing files
- Check write permissions on directories
- Create directories manually: `mkdir -p reports/weekly reports/monthly status tracker_logs`

### Reports look empty or incomplete
- Wait until after bootstrap completes
- Check that sheets have data in the tracked date fields
- Verify sheet IDs are correct in scripts

## Getting Help

1. Check main [README.md](README.md) for detailed documentation
2. Review error messages carefully - they usually indicate the issue
3. Check GitHub Actions logs if using automation
4. Verify Smartsheet API status: https://status.smartsheet.com/

## Tips

- **Start small**: Test with one product group first
- **Check data**: Review CSV logs before generating reports
- **Backup regularly**: Commit `tracker_logs/` to version control
- **Monitor size**: Archive old logs if they grow too large
- **Stay updated**: Pull latest changes regularly if using shared repository

Happy tracking! 📊

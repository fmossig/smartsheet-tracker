# Smartsheet Change Tracking & Reporting System

A comprehensive reporting system that tracks changes on Smartsheet tables and generates detailed PDF reports with visualizations, KPIs, and analytics.

## Overview

This system monitors multiple Smartsheet sheets for content changes, logs them to CSV files, and generates professional PDF reports with:
- Activity charts and graphs
- Phase distribution analytics
- Employee contribution tracking
- Country/marketplace statistics
- KPI metrics and rankings
- Historical trend analysis

## Components

### 1. Data Collection Scripts

#### `smartsheet_date_change_tracker.py`
**Main tracker** - Monitors date field changes across all product groups and logs them incrementally.

**Features:**
- Tracks changes in phase date fields (Kontrolle, BE am, K am, C am, Reopen C2 am)
- Stores data in monthly CSV files (`tracker_logs/date_changes_log_MM.YYYY.csv`)
- Supports bootstrap mode for initial data seeding
- Maintains backup file to avoid duplicate logging
- Tracks employee assignments and marketplace information

**Usage:**
```bash
# Nightly incremental tracking
python smartsheet_date_change_tracker.py

# Bootstrap from seed sheet (initial setup)
python smartsheet_date_change_tracker.py --bootstrap

# Bootstrap with specific time range
python smartsheet_date_change_tracker.py --bootstrap --days 90
```

#### `smartsheet_weekly_change_tracker.py`
Creates a weekly snapshot of changes from the last 7 days.

**Output:** `weekly/weekly_changes.csv`

**Usage:**
```bash
python smartsheet_weekly_change_tracker.py
```

#### `smartsheet_status_snapshot.py`
Generates status snapshots for different time periods (0-30, 31-60, 61-90 days).

**Output:**
- `status/status_snapshot_YYYY-MM-DD.csv` - Detailed snapshot
- `status/status_summary_YYYY-MM-DD.csv` - Summary statistics

**Usage:**
```bash
python smartsheet_status_snapshot.py
```

### 2. Report Generation Scripts

#### `smartsheet_status_report.py`
**Main report generator** - Creates comprehensive PDF reports with visualizations.

**Features:**
- Cover page with metadata
- Product group bar charts with phase distribution pie charts
- Per-group detailed pages with:
  - Employee contribution stacked bar charts
  - KPI boxes (article counts, marketplace articles, % edited)
  - Country/marketplace rankings (most/least active)
- Customizable time periods
- Professional styling with color coding

**Usage:**
```bash
# Generate report for last 30 days (default)
python smartsheet_status_report.py

# Generate report for specific period
REPORT_SINCE=2025-08-01 REPORT_UNTIL=2025-08-31 \
REPORT_LABEL=2025-08 REPORT_OUTDIR=reports/monthly \
python smartsheet_status_report.py
```

**Environment Variables:**
- `REPORT_SINCE` - Start date (YYYY-MM-DD)
- `REPORT_UNTIL` - End date (YYYY-MM-DD)
- `REPORT_LABEL` - Label for filename (e.g., "2025-W36" or "2025-08")
- `REPORT_OUTDIR` - Output directory (default: "status")
- `SMARTSHEET_TOKEN` - Smartsheet API access token

#### `smartsheet_periodic_status_report.py`
Alternative report generator using tracker logs for weekly/monthly reports.

**Usage:**
```bash
# Weekly report for specific ISO week
python smartsheet_periodic_status_report.py --week 2025-W37

# Monthly report
python smartsheet_periodic_status_report.py --month 2025-08

# Custom date range
python smartsheet_periodic_status_report.py --from 2025-08-01 --to 2025-08-31
```

### 3. Orchestration Script

#### `smartsheet_reports_orchestrator.py`
Master script that orchestrates different reporting modes.

**Modes:**
- `nightly` - Run incremental tracker
- `weekly` - Generate weekly report for previous week
- `monthly` - Generate monthly report for previous month
- `bootstrap` - Initialize system with last 3 months of data

**Usage:**
```bash
# Nightly tracking (runs daily)
python smartsheet_reports_orchestrator.py nightly

# Weekly report (runs on Mondays)
python smartsheet_reports_orchestrator.py weekly

# Monthly report (runs on 1st of month)
python smartsheet_reports_orchestrator.py monthly

# Bootstrap system
python smartsheet_reports_orchestrator.py bootstrap --months 3
```

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

**Required packages:**
- `smartsheet-python-sdk` - Smartsheet API client
- `python-dotenv` - Environment variable management
- `reportlab` - PDF generation (installed by orchestrator when needed)

### 2. Configure Environment

Create a `.env` file with your Smartsheet API token:

```bash
SMARTSHEET_TOKEN=your_smartsheet_access_token_here
```

Or set it as an environment variable:
```bash
export SMARTSHEET_TOKEN=your_token_here
```

### 3. Configure Product Groups

Edit the `SHEET_IDS` dictionary in the scripts to match your Smartsheet IDs:

```python
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}
```

### 4. Initial Bootstrap

Run the bootstrap to populate historical data:

```bash
python smartsheet_reports_orchestrator.py bootstrap
```

This will:
1. Load all data from the seed sheet
2. Create monthly tracker log files
3. Trim to last 3 months
4. Generate PDF reports for all available months

## Automation with GitHub Actions

The system includes GitHub Actions workflows for automated execution:

### Nightly Tracking
**File:** `.github/workflows/smartsheet-tracker-nightly.yml`
**Schedule:** Daily at 00:00 UTC
**Purpose:** Tracks new changes incrementally

### Weekly Reports
**File:** `.github/workflows/ci-weekly-reports.yml`
**Schedule:** Mondays at 01:15 UTC
**Purpose:** Generates weekly PDF report for previous week

### Monthly Reports
**File:** `.github/workflows/ci-monthly-reports.yml`
**Schedule:** 1st of month at 02:00 UTC
**Purpose:** Generates monthly PDF report for previous month

### Manual Triggers
All workflows support manual triggering via `workflow_dispatch`.

## Directory Structure

```
.
├── tracker_logs/                    # Tracked changes (CSV)
│   ├── date_changes_log_08.2025.csv
│   ├── date_changes_log_09.2025.csv
│   └── date_backup.csv             # Deduplication state
├── reports/                         # Generated PDF reports
│   ├── weekly/
│   │   └── 2025/
│   │       └── 2025-W37/
│   └── monthly/
│       └── 2025/
│           └── report_monthly_2025-08.pdf
├── weekly/                          # Weekly snapshots
│   └── weekly_changes.csv
├── status/                          # Status snapshots & reports
│   ├── status_report_2025-10-18.pdf
│   ├── status_snapshot_2025-10-18.csv
│   └── status_summary_2025-10-18.csv
└── assets/                          # Logo files for PDFs
    ├── amazon_logo.png
    └── noctua_logo.png
```

## Data Schema

### Tracker Log CSV Format
```csv
Änderung am,Produktgruppe,Land/Marketplace,Phase,Feld,Datum,Mitarbeiter
2025-09-14 16:35:03,NF,it,5,Reopen C2 am,2025-08-01,DM
```

**Columns:**
- `Änderung am` - When change was logged (timestamp)
- `Produktgruppe` - Product group code (NA, NF, NH, NM, NP, NT, NV)
- `Land/Marketplace` - Country/marketplace (de, com, co.uk, etc.)
- `Phase` - Phase number (1-5)
- `Feld` - Field name that changed
- `Datum` - Date value in the field
- `Mitarbeiter` - Employee who made the change

### Phase Fields
1. **Phase 1** - Kontrolle (K von)
2. **Phase 2** - BE am (BE von)
3. **Phase 3** - K am (K2 von)
4. **Phase 4** - C am (C von)
5. **Phase 5** - Reopen C2 am (Reopen C2 von)

## Report Features

### Product Group Overview
- Bar chart showing total phase openings per group
- Pie charts showing phase distribution per group
- Color-coded by product group

### Per-Group Details
- Stacked bar charts of employee contributions per phase
- KPI boxes:
  - Total articles
  - Marketplace articles
  - Edited articles (in period)
  - Percentage edited
- Country rankings:
  - Most inactive countries (by average age)
  - Most active countries (by average age)

### Customization
- Configurable color schemes for groups, phases, and employees
- Flexible time periods via environment variables
- PDF footer with company logos

## Color Schemes

### Product Groups
- NA: #E63946 (Red)
- NF: #457B9D (Blue)
- NH: #2A9D8F (Teal)
- NM: #E9C46A (Yellow)
- NP: #F4A261 (Orange)
- NT: #9D4EDD (Purple)
- NV: #00B4D8 (Cyan)

### Phases
- Phase 1: #1F77B4
- Phase 2: #FF7F0E
- Phase 3: #2CA02C
- Phase 4: #9467BD
- Phase 5: #D62728

### Employees
Customizable in `EMP_COLORS` dictionary.

## System Validation

Use the validation script to check if everything is configured correctly:

```bash
python validate_system.py
```

This will check:
- ✅ Dependencies installed
- ✅ Environment configured (token present)
- ✅ All scripts are syntactically valid
- ✅ Required directories exist
- ✅ Data files are present and valid
- ✅ Reports have been generated
- ✅ Automation workflows are configured

**Expected output when properly configured:**
```
ALL CHECKS PASSED (7/7)
System is ready to use!
```

## Troubleshooting

### Issue: "Smartsheet API error"
- Verify your `SMARTSHEET_TOKEN` is valid
- Check network connectivity
- Ensure sheet IDs are correct

### Issue: "No data in reports"
- Run bootstrap first: `python smartsheet_reports_orchestrator.py bootstrap`
- Check that tracker logs exist in `tracker_logs/`
- Verify date ranges match available data

### Issue: "Missing dependencies"
- Run: `pip install -r requirements.txt`
- For PDF generation: `pip install reportlab`

### Issue: "Permission denied on GitHub Actions"
- Ensure `SMARTSHEET_TOKEN` secret is set in repository settings
- Check workflow permissions (needs `contents: write`)

### Issue: Validation script fails
Run `python validate_system.py` to identify specific issues and follow the recommendations provided.

## Development

### Adding New Product Groups
1. Add to `SHEET_IDS` dictionary
2. Add color to `GROUP_COLORS` dictionary
3. Update documentation

### Modifying Report Layout
- Edit `smartsheet_status_report.py`
- Adjust sizing constants (mm units)
- Modify chart parameters

### Extending Tracking
- Add fields to `PHASE_FIELDS` list
- Update column mappings
- Maintain CSV schema compatibility

## License

Internal use only. Amazon/Noctua confidential.

## Version

Current version: 2025-07-23_tracker

## Support

For issues or questions, contact the development team.

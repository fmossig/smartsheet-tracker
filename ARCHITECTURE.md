# System Architecture

## Overview

The Smartsheet Change Tracking & Reporting System is a data pipeline that monitors changes in Smartsheet tables, logs them incrementally, and generates visual reports. The architecture follows a three-tier approach: **Collection → Storage → Presentation**.

```
┌─────────────────────────────────────────────────────────────────┐
│                        Smartsheet API                            │
│  (Multiple product group sheets: NA, NF, NH, NM, NP, NT, NV)   │
└─────────────────────┬───────────────────────────────────────────┘
                      │
                      │ Daily Polling
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Data Collection Layer                         │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ smartsheet_date_change_tracker.py                        │   │
│  │ - Incremental tracking                                   │   │
│  │ - Deduplication via backup CSV                          │   │
│  │ - Monthly log rotation                                   │   │
│  └──────────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ smartsheet_weekly_change_tracker.py                      │   │
│  │ - Weekly snapshots (last 7 days)                        │   │
│  └──────────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ smartsheet_status_snapshot.py                            │   │
│  │ - Status snapshots (0-30, 31-60, 61-90, >90 days)      │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────┬───────────────────────────────────────────┘
                      │
                      │ Write
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Storage Layer                               │
│  tracker_logs/                                                   │
│  ├── date_changes_log_08.2025.csv  (monthly rotation)          │
│  ├── date_changes_log_09.2025.csv                               │
│  └── date_backup.csv                (deduplication state)       │
│                                                                  │
│  weekly/                                                         │
│  └── weekly_changes.csv             (overwritten weekly)        │
│                                                                  │
│  status/                                                         │
│  ├── status_snapshot_YYYY-MM-DD.csv                            │
│  └── status_summary_YYYY-MM-DD.csv                             │
└─────────────────────┬───────────────────────────────────────────┘
                      │
                      │ Read & Aggregate
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│                   Report Generation Layer                        │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ smartsheet_status_report.py                              │   │
│  │ - Main report with full visualizations                   │   │
│  │ - Reads: tracker_logs + live Smartsheet data           │   │
│  │ - Outputs: Professional PDF with charts                 │   │
│  └──────────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ smartsheet_periodic_status_report.py                     │   │
│  │ - Alternative report from tracker logs only              │   │
│  │ - Simpler charts, faster generation                     │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────┬───────────────────────────────────────────┘
                      │
                      │ Generate
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│                        Output Layer                              │
│  reports/                                                        │
│  ├── weekly/                                                    │
│  │   └── 2025/                                                 │
│  │       └── 2025-W37/                                         │
│  │           └── status_report_2025-W37.pdf                    │
│  └── monthly/                                                   │
│      └── 2025/                                                  │
│          └── status_report_2025-08.pdf                          │
│                                                                  │
│  status/                                                         │
│  └── status_report_YYYY-MM-DD.pdf                              │
└─────────────────────────────────────────────────────────────────┘
```

## Components

### 1. Data Collection Layer

#### Incremental Tracker (`smartsheet_date_change_tracker.py`)

**Purpose**: Track changes without reprocessing historical data.

**Algorithm**:
1. Load backup state (which rows/dates already logged)
2. Fetch current state from all sheets
3. Compare with backup to find new changes
4. Write new changes to monthly CSV
5. Update backup state

**Key Design Decisions**:
- **Monthly log rotation**: Files named `date_changes_log_MM.YYYY.csv` based on the date value, not timestamp
- **Deduplication**: Uses composite key `{group}:{row_id}` → `[logged_dates]`
- **Graceful handling**: Parses multiple date formats, handles missing fields
- **Writer caching**: Opens each monthly file once per run for efficiency

**Data Flow**:
```
Smartsheet Sheet → Row Iteration → Cell Parsing → Date Detection →
Backup Check → New? → Writer Selection (by month) → CSV Append →
Backup Update
```

#### Weekly Snapshot (`smartsheet_weekly_change_tracker.py`)

**Purpose**: Provide a simple, always-current view of recent activity.

**Characteristics**:
- Overwrites previous file each run
- No deduplication (fresh query)
- Filters to last 7 days client-side
- Useful for quick checks without parsing monthly logs

#### Status Snapshot (`smartsheet_status_snapshot.py`)

**Purpose**: Capture current state for historical comparison.

**Outputs**:
1. **Snapshot CSV**: All phase events in last 30 days
2. **Summary CSV**: Aggregated counts by time buckets (0-30, 31-60, 61-90, >90 days)

**Use Case**: Trend analysis, comparing activity levels across weeks/months.

### 2. Storage Layer

#### CSV Format Design

**Tracker Log Schema**:
```csv
Änderung am,Produktgruppe,Land/Marketplace,Phase,Feld,Datum,Mitarbeiter
```

**Rationale**:
- **Änderung am**: When change was logged (not when it occurred in Smartsheet)
- **Datum**: The actual date value from the field
- **Phase**: Numeric (1-5) for easy filtering and sorting
- **Land/Marketplace**: Enables geographic analysis
- **Mitarbeiter**: Tracks individual contributions

**File Organization**:
- Monthly rotation prevents files from growing too large
- Easy to archive/delete old months
- Parallel reading possible (one reader per month)

**Backup Format**:
```csv
RowKey,LoggedDates...
NA:7037475763654532,2025-08-01,2025-08-15
```

Simple append-based format for fast lookups and updates.

### 3. Report Generation Layer

#### Main Report Generator (`smartsheet_status_report.py`)

**Architecture**:
- **Two-source design**: Combines tracker logs (historical activity) with live Smartsheet data (current KPIs)
- **Flexible periods**: Supports 30-day default or custom ranges via environment variables
- **Multi-page layout**: Cover + overview + per-group details

**Report Structure**:

1. **Cover Page**
   - Title, timestamp, period, version

2. **Overview Page**
   - Bar chart: Phase openings per group
   - Pie charts: Phase distribution per group (on gray banner)
   - Phase legend

3. **Per-Group Pages** (one per product group)
   - Header chip with group code
   - Employee legend on gray banner
   - Stacked bar chart: Employee contributions per phase
   - KPI boxes: Articles, marketplace articles, edited %, etc.
   - Country rankings: Most/least active (by average age since last activity)

**Visual Design Principles**:
- **Color consistency**: Each group, phase, and employee has a fixed color
- **Information density**: Maximize insight per page without clutter
- **Professional styling**: Clean, corporate look with logos

**Technology Stack**:
- ReportLab for PDF generation
- Custom Drawing objects for charts
- TableStyle for layout control

#### Periodic Report Generator (`smartsheet_periodic_status_report.py`)

**Differences from Main Report**:
- Reads only from tracker logs (no live Smartsheet calls)
- Simpler charts (bar and pie via ReportLab charts API)
- Faster generation
- Better for historical period analysis

**Use Case**: Archival reports, performance reviews, trend analysis.

### 4. Orchestration Layer

#### Master Orchestrator (`smartsheet_reports_orchestrator.py`)

**Purpose**: Simplify complex multi-step operations.

**Modes**:

1. **Nightly** (`--mode nightly`)
   - Runs incremental tracker
   - Updates backup
   - Commits logs to version control (via GitHub Actions)

2. **Weekly** (`--mode weekly`)
   - Calculates previous week's date range (Monday-Sunday)
   - Sets environment variables for period
   - Generates report to `reports/weekly/YYYY/YYYY-Wxx/`

3. **Monthly** (`--mode monthly`)
   - Calculates previous month's date range
   - Generates report to `reports/monthly/YYYY/`

4. **Bootstrap** (`--mode bootstrap`)
   - Runs tracker in bootstrap mode (loads all historical data)
   - Trims old logs to last N months
   - Generates PDFs for all available months
   - **Use Case**: Initial setup or recovery after data loss

**Script Resolution**:
- Handles multiple filename variants (e.g., `(1)` suffixes from downloads)
- Searches for wildcards as fallback
- Raises clear error if script not found

**Log Trimming**:
- Parses `MM.YYYY` from filenames
- Sorts by date (newest first)
- Keeps N newest, deletes rest
- Prevents unbounded storage growth

### 5. Automation Layer (GitHub Actions)

#### Workflow Design

**Nightly Workflow** (`smartsheet-tracker-nightly.yml`):
```yaml
Trigger: cron '0 0 * * *' (daily)
Steps:
  1. Checkout code
  2. Setup Python 3.11
  3. Install dependencies
  4. Run tracker
  5. Commit & push CSVs
  6. Upload artifacts
```

**Concurrency**: `cancel-in-progress: false` to prevent data loss from interrupted runs.

**Weekly Workflow** (`ci-weekly-reports.yml`):
```yaml
Trigger: cron '15 1 * * MON'
Steps:
  1. Checkout
  2. Setup Python
  3. Install deps + reportlab
  4. Run orchestrator weekly
  5. Commit PDFs
```

**Monthly Workflow** (`ci-monthly-reports.yml`):
Similar to weekly, runs on 1st of month.

**Bootstrap Workflow** (`ci-bootstrap-3m.yml`):
- Manual trigger only
- Initializes system
- Used for recovery or new setup

#### Artifact Strategy

- Nightly: Uploads CSVs as artifacts (backup)
- Reports: Commits PDFs directly to repo (for easy access)
- Logs: Commits to repo daily (for version history)

## Data Flow Examples

### Example 1: Daily Tracking

```
00:00 UTC - GitHub Action triggers

1. Clone repo with existing tracker_logs/
2. Load date_backup.csv into memory
3. Query Smartsheet API (7 sheets × ~100-500 rows each)
4. For each row:
   - Check 5 phase date fields
   - Parse dates
   - Look up in backup
   - If new: append to monthly CSV, add to backup
5. Close CSV writers
6. Save updated backup
7. Git commit + push
8. Upload artifacts

Total time: ~2-5 minutes
API calls: ~7 (one per sheet)
New entries: ~10-50 per night (typical)
```

### Example 2: Weekly Report Generation

```
Monday 01:15 UTC - GitHub Action triggers

1. Orchestrator calculates previous week:
   - Today: 2025-09-15 (Monday)
   - Previous Monday: 2025-09-08
   - Previous Sunday: 2025-09-14
   - ISO week: 2025-W37

2. Set environment variables:
   REPORT_SINCE=2025-09-08
   REPORT_UNTIL=2025-09-14
   REPORT_LABEL=2025-W37
   REPORT_OUTDIR=reports/weekly/2025/2025-W37

3. Status report generator:
   a. Load tracker logs from 08.2025 and 09.2025
   b. Filter rows to date range
   c. Aggregate by group, phase, employee
   d. Query Smartsheet for KPIs
   e. Generate PDF with all charts
   
4. Commit PDF to repo

Total time: ~5-10 minutes
API calls: ~7 (KPIs for each group)
Output size: ~500KB-2MB PDF
```

### Example 3: Bootstrap

```
Manual trigger (initial setup or recovery)

1. Orchestrator runs tracker with --bootstrap
2. Tracker queries SEED_SHEET_ID (one large sheet with all data)
3. For each row:
   - Extract all 5 phase dates
   - Infer product group from primary column
   - Parse date → determine month → route to correct CSV
   - Write without deduplication check (dedupe=False)
4. Result: Multiple monthly CSVs created

5. Orchestrator trims to last N months:
   - Parse filenames for dates
   - Sort chronologically
   - Keep 3 newest, delete rest

6. For each remaining month:
   - Set period to full month
   - Generate PDF
   
7. Final state: Clean logs + complete report history

Total time: ~30-60 minutes (depends on data volume)
Entries written: ~10,000-50,000 (typical)
```

## Design Patterns

### 1. Incremental Processing

**Problem**: Reprocessing all Smartsheet data is slow and wasteful.

**Solution**: Maintain state (backup CSV) to track what's been logged.

**Benefits**:
- Fast daily runs (only process new changes)
- Reduced API load
- Lower cost (if rate-limited)

### 2. Monthly Log Rotation

**Problem**: Single CSV grows unbounded.

**Solution**: Route entries to monthly files based on date value.

**Benefits**:
- Bounded file sizes
- Easy to archive old data
- Parallel processing possible

### 3. Deferred Deduplication

**Problem**: Same date might appear in multiple tracker runs (Smartsheet data doesn't change).

**Solution**: Track seen dates per row in backup.

**Benefits**:
- No duplicates in logs
- Fast lookup (dict-based)
- Survives restarts

### 4. Two-Source Reporting

**Problem**: Some metrics need current state (Smartsheet), others need history (logs).

**Solution**: Hybrid approach: read both sources.

**Benefits**:
- Rich reports combining activity + current state
- Accurate KPIs
- Historical trend analysis

### 5. Environment-Based Configuration

**Problem**: Same script needs to generate reports for different periods.

**Solution**: Accept period via environment variables.

**Benefits**:
- Single codebase
- Easy automation (GitHub Actions just sets env vars)
- Flexible for ad-hoc reports

## Scalability Considerations

### Current Scale
- **Sheets**: 7 product groups
- **Rows per sheet**: ~100-500
- **Changes per day**: ~10-50
- **Monthly log size**: ~100KB-500KB
- **Report generation**: ~5-10 minutes

### Bottlenecks
1. **Smartsheet API rate limits**: Mitigated by incremental tracking
2. **PDF generation time**: Grows with data volume; consider pagination for very large datasets
3. **CSV parsing**: Linear scan; acceptable up to ~1M rows per file

### Optimization Strategies (if needed)
- **Parallel sheet fetching**: Use threading for Smartsheet API calls
- **Incremental aggregation**: Pre-compute aggregates, update daily
- **Database backend**: Replace CSV with SQLite for faster queries
- **Caching**: Cache Smartsheet metadata (column mappings)

## Security & Privacy

### Sensitive Data
- **SMARTSHEET_TOKEN**: Stored as GitHub Secret, never in code
- **Sheet data**: Contains internal product/employee information

### Access Control
- Repository: Private
- Workflows: Require authentication
- Artifacts: Available only to repo members

### Data Retention
- Tracker logs: Trimmed to N months
- Reports: Kept indefinitely (small size)
- Backups: Overwritten daily

## Error Handling

### Tracker
- **API errors**: Script fails, GitHub Actions alerts (no partial commits)
- **Parse errors**: Skip row, log to stderr, continue
- **Write errors**: Fail fast (prevent data loss)

### Report Generator
- **Missing logs**: Warning, generate partial report
- **API errors**: Fatal (need live data for KPIs)
- **PDF errors**: Fatal (output is essential)

### Recovery
- **Backup corruption**: Re-run bootstrap
- **Missing logs**: Re-run bootstrap
- **Wrong data**: Manual CSV edit + re-run tracker

## Testing Strategy

### Current State
- No automated tests (manual validation)

### Recommended Tests (future)
1. **Unit tests**:
   - Date parsing functions
   - CSV writing
   - Backup state management

2. **Integration tests**:
   - End-to-end tracker run (mock Smartsheet API)
   - Report generation (sample logs)

3. **Smoke tests**:
   - GitHub Actions includes basic validation
   - Check output file existence

## Future Enhancements

### Potential Improvements
1. **Dashboard**: Web-based live dashboard instead of PDFs
2. **Alerts**: Notify on anomalies (sudden activity drop, etc.)
3. **Drill-down**: Interactive reports with filtering
4. **Predictions**: ML-based forecasting of completion times
5. **Mobile**: Mobile-friendly report viewer
6. **API**: REST API for programmatic access to metrics

### Requested Features
- Custom report templates
- Export to Excel
- Email delivery of reports
- Slack/Teams integration

## Maintenance

### Regular Tasks
- **Weekly**: Review GitHub Actions logs for errors
- **Monthly**: Check storage usage (logs + reports)
- **Quarterly**: Audit sheet IDs, add new groups if needed
- **Yearly**: Archive old reports, review retention policy

### Updates
- **Dependencies**: Run `pip list --outdated` monthly
- **Scripts**: Version tracked in `CODE_VERSION` constant
- **Workflows**: Update actions (e.g., `actions/checkout@v4` → `@v5`)

## Glossary

- **Product Group**: Category of products (NA, NF, etc.)
- **Phase**: Stage in content lifecycle (1-5)
- **Marketplace**: Amazon country site (de, com, co.uk, etc.)
- **Tracker**: Data collection script
- **Bootstrap**: Initial data load
- **Snapshot**: Point-in-time export
- **Orchestrator**: Master control script

## References

- [Smartsheet API Documentation](https://smartsheet-platform.github.io/api-docs/)
- [ReportLab User Guide](https://www.reportlab.com/docs/reportlab-userguide.pdf)
- [GitHub Actions Documentation](https://docs.github.com/en/actions)

# Usage Examples

Practical examples for common use cases of the Smartsheet Change Tracking & Reporting System.

## Table of Contents
- [First-Time Setup](#first-time-setup)
- [Daily Operations](#daily-operations)
- [Report Generation](#report-generation)
- [Data Analysis](#data-analysis)
- [Custom Reports](#custom-reports)
- [Troubleshooting](#troubleshooting)
- [Advanced Usage](#advanced-usage)

## First-Time Setup

### Example 1: Complete Initial Setup

```bash
# Clone repository
git clone https://github.com/your-org/smartsheet-tracker.git
cd smartsheet-tracker

# Install dependencies
pip install -r requirements.txt

# Create .env file
cat > .env << EOF
SMARTSHEET_TOKEN=your_actual_token_here
EOF

# Bootstrap system with last 3 months of data
python smartsheet_reports_orchestrator.py bootstrap --months 3

# Verify setup
ls -l tracker_logs/
ls -l reports/monthly/2025/
```

**Expected output:**
```
Bootstrap-Statistik:
  +2847 Zeilen geschrieben
  143 ohne/ungültiges Datum
  0 Duplikate (nur wenn dedupe=True)
  Kontrolle: gefunden 589, geschrieben 589
  BE am: gefunden 612, geschrieben 612
  K am: gefunden 534, geschrieben 534
  C am: gefunden 587, geschrieben 587
  Reopen C2 am: gefunden 525, geschrieben 525
✅ Bootstrap: 2847 Zeilen übernommen.
✅ PDF Report erstellt: reports/monthly/2025/status_report_2025-08.pdf
✅ PDF Report erstellt: reports/monthly/2025/status_report_2025-09.pdf
```

### Example 2: Minimal Setup for Testing

```bash
# Quick setup without bootstrap
pip install smartsheet-python-sdk python-dotenv

# Test connection only
python -c "
from dotenv import load_dotenv
import os, smartsheet
load_dotenv()
client = smartsheet.Smartsheet(os.getenv('SMARTSHEET_TOKEN'))
sheets = client.Sheets.list_sheets()
print(f'✅ Connected! Found {len(sheets.data)} sheets')
"

# Run single tracking cycle
python smartsheet_date_change_tracker.py

# Generate quick report
python smartsheet_status_report.py
```

## Daily Operations

### Example 3: Manual Daily Tracking

```bash
# Run tracker
python smartsheet_date_change_tracker.py

# Check what was tracked
tail -20 tracker_logs/date_changes_log_$(date +%m.%Y).csv

# Review changes for specific group
grep "^.*,NF," tracker_logs/date_changes_log_$(date +%m.%Y).csv | tail -10

# Count today's changes
grep "$(date +%Y-%m-%d)" tracker_logs/date_changes_log_$(date +%m.%Y).csv | wc -l
```

**Sample output:**
```
✅ Nightly: 23 neue Änderung(en) protokolliert.

2025-10-18 09:23:45,NF,de,2,BE am,2025-10-18,SM
2025-10-18 09:23:45,NF,it,2,BE am,2025-10-18,SM
2025-10-18 09:23:45,NH,com,3,K am,2025-10-17,HI
...

23
```

### Example 4: Weekly Snapshot

```bash
# Generate weekly snapshot
python smartsheet_weekly_change_tracker.py

# View summary by product group
awk -F',' 'NR>1 {count[$2]++} END {for (g in count) print g, count[g]}' \
  weekly/weekly_changes.csv | sort

# View summary by employee
awk -F',' 'NR>1 {count[$6]++} END {for (e in count) print e, count[e]}' \
  weekly/weekly_changes.csv | sort -rn -k2
```

**Sample output:**
```
NA 45
NF 67
NH 23
NM 12
NP 34
NT 28
NV 56

SM 89
HI 67
DM 45
EK 34
JHU 23
LK 7
```

## Report Generation

### Example 5: Standard Monthly Report

```bash
# Generate report for current month to date
python smartsheet_status_report.py

# Output location
ls -lh status/status_report_*.pdf

# View with system PDF viewer
xdg-open status/status_report_$(date +%Y-%m-%d).pdf  # Linux
open status/status_report_$(date +%Y-%m-%d).pdf      # Mac
start status/status_report_$(date +%Y-%m-%d).pdf     # Windows
```

### Example 6: Weekly Report via Orchestrator

```bash
# Generate last week's report
python smartsheet_reports_orchestrator.py weekly

# Find the report
find reports/weekly -name "*.pdf" -type f -mtime -1

# Copy to shared location
cp reports/weekly/2025/2025-W*/status_report_*.pdf /shared/reports/
```

### Example 7: Monthly Report via Orchestrator

```bash
# Generate previous month's report
python smartsheet_reports_orchestrator.py monthly

# List all monthly reports
ls -lh reports/monthly/2025/

# Create archive
tar -czf monthly_reports_2025.tar.gz reports/monthly/2025/
```

## Data Analysis

### Example 8: Find Most Active Employee

```bash
# Count changes per employee in October
awk -F',' '$1 ~ /^2025-10/ {count[$7]++} END {
  for (emp in count) print count[emp], emp
}' tracker_logs/date_changes_log_10.2025.csv | sort -rn | head -5
```

**Sample output:**
```
234 SM
189 HI
156 DM
123 EK
98 JHU
```

### Example 9: Track Phase 2 Completion Rate

```bash
# Count Phase 2 events by week
awk -F',' '$4 == "2" {
  split($6, date, "-")
  week = strftime("%Y-W%V", mktime(date[1]" "date[2]" "date[3]" 0 0 0"))
  count[week]++
}
END {
  for (w in count) print w, count[w]
}' tracker_logs/date_changes_log_*.csv | sort
```

**Sample output:**
```
2025-W34 45
2025-W35 52
2025-W36 48
2025-W37 61
2025-W38 55
```

### Example 10: Marketplace Activity Comparison

```bash
# Compare activity by marketplace
awk -F',' 'NR>1 {
  marketplace[$3]++
}
END {
  for (mp in marketplace) 
    printf "%-10s %5d\n", mp, marketplace[mp]
}' tracker_logs/date_changes_log_*.csv | sort -k2 -rn | head -10
```

**Sample output:**
```
de           1234
com           987
co.uk         876
fr            654
it            543
es            432
nl            321
se            234
pl            198
ca            156
```

### Example 11: Identify Inactive Rows

```bash
# Find rows not updated in 60+ days
python << 'EOF'
import csv, glob
from datetime import datetime, timedelta
from collections import defaultdict

cutoff = datetime.now().date() - timedelta(days=60)
latest = defaultdict(lambda: None)

for path in glob.glob("tracker_logs/date_changes_log_*.csv"):
    with open(path) as f:
        for row in csv.DictReader(f):
            key = f"{row['Produktgruppe']}:{row['Land/Marketplace']}"
            date = datetime.fromisoformat(row['Datum']).date()
            if latest[key] is None or date > latest[key]:
                latest[key] = date

inactive = [(k, d) for k, d in latest.items() if d < cutoff]
inactive.sort(key=lambda x: x[1])

print(f"Found {len(inactive)} inactive items (no updates in 60+ days):\n")
for key, last_date in inactive[:20]:
    days = (datetime.now().date() - last_date).days
    print(f"{key:30s} - {days:3d} days ago ({last_date})")
EOF
```

## Custom Reports

### Example 12: Quarter Report

```bash
# Q3 2025 (July-September)
REPORT_SINCE=2025-07-01 \
REPORT_UNTIL=2025-09-30 \
REPORT_LABEL=2025-Q3 \
REPORT_OUTDIR=reports/quarterly \
python smartsheet_status_report.py

ls -lh reports/quarterly/status_report_2025-Q3.pdf
```

### Example 13: Sprint Report (2 weeks)

```bash
# Two-week sprint ending yesterday
END_DATE=$(date -d "yesterday" +%Y-%m-%d)
START_DATE=$(date -d "yesterday -14 days" +%Y-%m-%d)

REPORT_SINCE=$START_DATE \
REPORT_UNTIL=$END_DATE \
REPORT_LABEL="Sprint-$(date -d "$END_DATE" +%Y-W%V)" \
REPORT_OUTDIR=reports/sprints \
python smartsheet_status_report.py

echo "Report created: reports/sprints/status_report_Sprint-*.pdf"
```

### Example 14: Single Product Group Report

```bash
# Filter logs for one group, generate custom report
mkdir -p /tmp/nf_only/tracker_logs

# Extract NF data only
for f in tracker_logs/date_changes_log_*.csv; do
    head -1 "$f" > "/tmp/nf_only/$f"
    grep "^[^,]*,NF," "$f" >> "/tmp/nf_only/$f"
done

# Generate report using filtered data
cd /tmp/nf_only
python /path/to/smartsheet_periodic_status_report.py \
  --month 2025-10 \
  --tracker-logs . \
  --out-dir reports

ls -lh reports/
```

### Example 15: Comparison Report (Month-over-Month)

```bash
# Generate two months and compare
mkdir -p reports/comparison

# August
REPORT_SINCE=2025-08-01 REPORT_UNTIL=2025-08-31 \
REPORT_LABEL=2025-08 REPORT_OUTDIR=reports/comparison \
python smartsheet_status_report.py

# September
REPORT_SINCE=2025-09-01 REPORT_UNTIL=2025-09-30 \
REPORT_LABEL=2025-09 REPORT_OUTDIR=reports/comparison \
python smartsheet_status_report.py

# Compare side-by-side
echo "Reports generated:"
ls -lh reports/comparison/
```

## Troubleshooting

### Example 16: Validate Data Integrity

```bash
# Check for duplicates in logs
python << 'EOF'
import csv, glob
from collections import defaultdict

seen = defaultdict(list)
duplicates = 0

for path in glob.glob("tracker_logs/date_changes_log_*.csv"):
    with open(path) as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=2):
            key = f"{row['Produktgruppe']}:{row['Phase']}:{row['Datum']}:{row['Mitarbeiter']}"
            seen[key].append((path, i))
            if len(seen[key]) > 1:
                duplicates += 1

print(f"Found {duplicates} potential duplicates")
if duplicates > 0:
    print("\nSample duplicates:")
    for key, locs in list(seen.items())[:5]:
        if len(locs) > 1:
            print(f"\n{key}:")
            for loc in locs[:3]:
                print(f"  {loc[0]}:{loc[1]}")
EOF
```

### Example 17: Recover from Corrupted Backup

```bash
# Backup is corrupted, regenerate from logs
echo "Backing up corrupted file..."
mv tracker_logs/date_backup.csv tracker_logs/date_backup.csv.corrupted

# Rebuild backup from logs
python << 'EOF'
import csv, glob
from collections import defaultdict

backup = defaultdict(list)

for path in glob.glob("tracker_logs/date_changes_log_*.csv"):
    with open(path) as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Simplified key (just group:date:phase)
            key = f"{row['Produktgruppe']}:{row['Datum']}:{row['Phase']}"
            backup[key].append(row['Datum'])

# Write new backup
with open("tracker_logs/date_backup.csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["RowKey", "LoggedDates..."])
    for key, dates in backup.items():
        writer.writerow([key] + list(set(dates)))

print(f"✅ Rebuilt backup with {len(backup)} entries")
EOF
```

### Example 18: Verify API Connection

```bash
# Test all sheet connections
python << 'EOF'
from dotenv import load_dotenv
import os, smartsheet

load_dotenv()
client = smartsheet.Smartsheet(os.getenv('SMARTSHEET_TOKEN'))

sheet_ids = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

for name, sid in sheet_ids.items():
    try:
        sheet = client.Sheets.get_sheet(sid)
        print(f"✅ {name}: {sheet.name} ({len(sheet.rows)} rows)")
    except Exception as e:
        print(f"❌ {name}: Error - {e}")
EOF
```

## Advanced Usage

### Example 19: Automated Email Delivery

```bash
# Generate report and email it
python smartsheet_reports_orchestrator.py weekly

LATEST_REPORT=$(find reports/weekly -name "*.pdf" -type f -mtime -1 -print -quit)

# Using mail command (Linux)
echo "Weekly report attached" | mail -s "Smartsheet Weekly Report" \
  -A "$LATEST_REPORT" team@example.com

# Or using Python
python << EOF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

msg = MIMEMultipart()
msg['Subject'] = 'Smartsheet Weekly Report'
msg['From'] = 'reports@example.com'
msg['To'] = 'team@example.com'

with open('$LATEST_REPORT', 'rb') as f:
    pdf = MIMEApplication(f.read(), _subtype='pdf')
    pdf.add_header('Content-Disposition', 'attachment', 
                   filename='weekly_report.pdf')
    msg.attach(pdf)

# Send via SMTP
with smtplib.SMTP('smtp.example.com', 587) as s:
    s.starttls()
    s.login('user', 'pass')
    s.send_message(msg)
    print('✅ Email sent')
EOF
```

### Example 20: Custom Aggregation Script

```bash
# Create custom aggregation for executive summary
python << 'EOF'
import csv, glob
from datetime import datetime, timedelta
from collections import defaultdict

# Load last 30 days
cutoff = datetime.now().date() - timedelta(days=30)
metrics = defaultdict(lambda: defaultdict(int))

for path in glob.glob("tracker_logs/date_changes_log_*.csv"):
    with open(path) as f:
        reader = csv.DictReader(f)
        for row in reader:
            date = datetime.fromisoformat(row['Datum']).date()
            if date >= cutoff:
                group = row['Produktgruppe']
                phase = row['Phase']
                metrics[group][f'phase_{phase}'] += 1
                metrics[group]['total'] += 1

# Generate executive summary
print("Executive Summary - Last 30 Days")
print("=" * 60)
print(f"{'Group':<10} {'Total':<10} {'P1':<8} {'P2':<8} {'P3':<8} {'P4':<8} {'P5':<8}")
print("-" * 60)

total_all = 0
for group in sorted(metrics.keys()):
    m = metrics[group]
    total = m['total']
    total_all += total
    print(f"{group:<10} {total:<10} {m['phase_1']:<8} {m['phase_2']:<8} "
          f"{m['phase_3']:<8} {m['phase_4']:<8} {m['phase_5']:<8}")

print("-" * 60)
print(f"{'TOTAL':<10} {total_all:<10}")
print()

# Top contributors
from collections import Counter
contributors = Counter()
for path in glob.glob("tracker_logs/date_changes_log_*.csv"):
    with open(path) as f:
        reader = csv.DictReader(f)
        for row in reader:
            date = datetime.fromisoformat(row['Datum']).date()
            if date >= cutoff:
                contributors[row['Mitarbeiter']] += 1

print("Top Contributors:")
for emp, count in contributors.most_common(5):
    print(f"  {emp:<10} {count:>5} changes")
EOF
```

### Example 21: Data Export to Excel

```bash
# Install openpyxl
pip install openpyxl

# Export tracker logs to Excel
python << 'EOF'
import csv, glob
from openpyxl import Workbook
from datetime import datetime

wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

# Create sheet per month
for path in sorted(glob.glob("tracker_logs/date_changes_log_*.csv")):
    month = path.split('_')[-1].replace('.csv', '')
    ws = wb.create_sheet(title=month)
    
    with open(path) as f:
        reader = csv.reader(f)
        for row in reader:
            ws.append(row)

# Save
output = f"tracker_export_{datetime.now().strftime('%Y%m%d')}.xlsx"
wb.save(output)
print(f"✅ Exported to {output}")
EOF
```

### Example 22: Integration with Slack

```bash
# Post report summary to Slack
python << 'EOF'
import json
import urllib.request
from datetime import datetime, timedelta

# Generate summary
# ... (use code from Example 20)

# Post to Slack
webhook_url = "https://hooks.slack.com/services/YOUR/WEBHOOK/URL"
message = {
    "text": "📊 Weekly Smartsheet Report",
    "blocks": [
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": "*Weekly Report Summary*\n" + 
                        f"Total changes: {total_all}\n" +
                        f"Top contributor: {top_contributor}"
            }
        }
    ]
}

req = urllib.request.Request(
    webhook_url,
    data=json.dumps(message).encode(),
    headers={'Content-Type': 'application/json'}
)
urllib.request.urlopen(req)
print("✅ Posted to Slack")
EOF
```

## Tips and Best Practices

### Performance Tips
```bash
# Parallel report generation for multiple periods
for month in 08 09 10; do
    REPORT_SINCE=2025-${month}-01 \
    REPORT_UNTIL=2025-${month}-31 \
    REPORT_LABEL=2025-${month} \
    REPORT_OUTDIR=reports/batch \
    python smartsheet_status_report.py &
done
wait
echo "All reports generated"
```

### Backup Strategy
```bash
# Daily backup script
#!/bin/bash
DATE=$(date +%Y%m%d)
tar -czf backups/tracker_backup_$DATE.tar.gz tracker_logs/
find backups/ -name "*.tar.gz" -mtime +30 -delete
echo "✅ Backup created: backups/tracker_backup_$DATE.tar.gz"
```

### Monitoring Script
```bash
# Check for tracking failures
#!/bin/bash
LATEST_LOG=$(ls -t tracker_logs/date_changes_log_*.csv | head -1)
LATEST_ENTRY=$(tail -1 "$LATEST_LOG" | cut -d',' -f1)
LATEST_DATE=$(echo "$LATEST_ENTRY" | cut -d' ' -f1)

if [[ "$LATEST_DATE" != $(date +%Y-%m-%d) ]]; then
    echo "⚠️  WARNING: No entries today in tracker logs!"
    echo "Last entry: $LATEST_ENTRY"
    # Send alert
else
    echo "✅ Tracking up to date"
fi
```

These examples should cover most use cases. For more specific needs, combine and modify these patterns as needed.

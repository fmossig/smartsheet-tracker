import os
import csv
import json
from datetime import datetime, timedelta, date
from collections import defaultdict, Counter
import logging

import smartsheet
from dotenv import load_dotenv
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("smartsheet_report.log"),
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

# Constants
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

GROUP_COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8",
}

PHASE_COLORS = {
    1: "#1F77B4",
    2: "#FF7F0E",
    3: "#2CA02C",
    4: "#9467BD",
    5: "#D62728",
}

# Directories
DATA_DIR = "tracking_data"
REPORTS_DIR = "reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

# File paths
CHANGES_FILE = os.path.join(DATA_DIR, "change_history.csv")

def parse_date(date_str):
    """Parse date from string, supporting multiple formats."""
    if not date_str:
        return None
        
    # Try various formats
    for fmt in ('%Y-%m-%d', '%d.%m.%Y'):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
            
    try:
        # Try ISO format (catches many variations)
        return datetime.fromisoformat(date_str).date()
    except:
        return None

def load_changes(start_date=None, end_date=None):
    """Load changes from the CSV file within the given date range."""
    if not os.path.exists(CHANGES_FILE):
        logger.error(f"Changes file not found: {CHANGES_FILE}")
        return []

    changes = []
    try:
        with open(CHANGES_FILE, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Parse the timestamp
                try:
                    ts = datetime.strptime(row['Timestamp'], "%Y-%m-%d %H:%M:%S").date()
                    
                    # Apply date filter if specified
                    if start_date and end_date:
                        if start_date <= ts <= end_date:
                            # Also parse the date field for later use
                            row['ParsedDate'] = parse_date(row['Date'])
                            changes.append(row)
                    else:
                        # No date filter, include all changes
                        row['ParsedDate'] = parse_date(row['Date'])
                        changes.append(row)
                except (ValueError, KeyError) as e:
                    logger.warning(f"Error parsing row: {row} - {e}")
                    continue
    except Exception as e:
        logger.error(f"Error reading changes file: {e}")
    
    if start_date and end_date:
        logger.info(f"Loaded {len(changes)} changes between {start_date} and {end_date}")
    else:
        logger.info(f"Loaded {len(changes)} total changes")
    return changes

def collect_metrics(changes):
    """Collect metrics from the changes data."""
    metrics = {
        "total_changes": len(changes),
        "groups": defaultdict(int),
        "phases": defaultdict(int),
        "users": defaultdict(int),
        "marketplaces": defaultdict(int),
        "daily": defaultdict(int),
        "group_phase": defaultdict(lambda: defaultdict(int))
    }
    
    # Add sample data if no changes
    if not changes:
        # Sample data for empty report
        metrics["groups"] = {"NA": 5, "NF": 3, "NH": 2}
        metrics["phases"] = {"1": 3, "2": 4, "3": 2, "4": 1}
        metrics["users"] = {"Sample User 1": 5, "Sample User 2": 3, "Sample User 3": 2}
        metrics["marketplaces"] = {"DE": 4, "FR": 3, "UK": 2, "ES": 1}
        return metrics
    
    for change in changes:
        group = change.get('Group', '')
        phase = change.get('Phase', '')
        user = change.get('User', '')
        marketplace = change.get('Marketplace', '')
        date_str = change.get('Date', '')
        
        metrics["groups"][group] += 1
        metrics["phases"][phase] += 1
        metrics["users"][user] += 1
        metrics["marketplaces"][marketplace] += 1
        
        if change.get('ParsedDate'):
            date_key = change['ParsedDate'].isoformat()
            metrics["daily"][date_key] += 1
            
        if group and phase:
            metrics["group_phase"][group][phase] += 1
            
    return metrics

def make_pie_chart(data_dict, title, width=400, height=200):
    """Create a pie chart from the given data."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-15, title,
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # Create the pie chart
    pie = Pie()
    pie.x = 100
    pie.y = 25
    pie.width = 150
    pie.height = 150
    
    # Prepare data
    labels = []
    data = []
    
    # If data is empty, add sample data
    if not data_dict:
        data_dict = {"Sample": 1}
        
    for key, value in data_dict.items():
        if value > 0:  # Only include non-zero values
            labels.append(str(key))
            data.append(value)
    
    # If still no data, add a placeholder
    if not data:
        labels = ['No Data']
        data = [1]
    
    pie.data = data
    pie.labels = [f"{labels[i]}: {data[i]}" for i in range(len(labels))]
    
    # Set slice properties
    for i in range(len(data)):
        pie.slices[i].strokeWidth = 0.5
    
    drawing.add(pie)
    return drawing

def make_bar_chart(data_dict, title, width=500, height=250):
    """Create a bar chart from the given data."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-15, title,
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # If data is empty, add sample data
    if not data_dict:
        data_dict = {"Sample": 1}
    
    # Create the bar chart
    chart = VerticalBarChart()
    chart.x = 50
    chart.y = 50
    chart.height = 150
    chart.width = 400
    chart.categoryAxis.categoryNames = list(data_dict.keys())
    chart.data = [list(data_dict.values())]
    chart.bars[0].fillColor = colors.steelblue
    
    drawing.add(chart)
    return drawing

def create_sample_image(title, message, width=500, height=200):
    """Create a placeholder image with text."""
    from reportlab.lib.colors import lightgrey, black
    from reportlab.graphics.shapes import Rect
    
    drawing = Drawing(width, height)
    
    # Add a background rectangle
    drawing.add(Rect(0, 0, width, height, fillColor=lightgrey))
    
    # Add title
    drawing.add(String(width/2, height-30, title,
                     fontName='Helvetica-Bold', fontSize=14, textAnchor='middle'))
    
    # Add message
    drawing.add(String(width/2, height/2, message,
                     fontName='Helvetica', fontSize=12, textAnchor='middle'))
    
    return drawing

def create_weekly_report(start_date, end_date, force=False):
    """Create a weekly PDF report."""
    # Generate output filename
    week_str = f"{start_date.isocalendar()[0]}-W{start_date.isocalendar()[1]:02d}"
    out_dir = os.path.join(REPORTS_DIR, "weekly")
    os.makedirs(out_dir, exist_ok=True)
    filename = os.path.join(out_dir, f"weekly_report_{week_str}.pdf")
    
    # Load changes for the week
    changes = load_changes(start_date, end_date)
    
    # Check if we have data
    has_data = len(changes) > 0
    
    # If no changes and not forcing, return None
    if not changes and not force:
        logger.warning(f"No changes found for week {week_str}")
        return None
    
    # Collect metrics (will use sample data if no changes)
    metrics = collect_metrics(changes)
    
    # Create PDF document
    doc = SimpleDocTemplate(filename, pagesize=A4,
                          leftMargin=25*mm, rightMargin=25*mm,
                          topMargin=20*mm, bottomMargin=20*mm)
    
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    heading_style = styles['Heading1']
    normal_style = styles['Normal']
    
    # Build the PDF content
    story = []
    
    # Title
    story.append(Paragraph(f"Weekly Smartsheet Changes Report", title_style))
    
    # Period information
    if not has_data:
        story.append(Paragraph(f"Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}", normal_style))
        story.append(Paragraph(f"<i>Sample report - no data available for this period</i>", normal_style))
    else:
        story.append(Paragraph(f"Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}", normal_style))
    
    story.append(Spacer(1, 10*mm))
    
    # Summary
    story.append(Paragraph("Summary", heading_style))
    summary_data = [
        ["Total Changes", str(metrics["total_changes"])],
        ["Groups with Activity", str(len(metrics["groups"]))],
        ["Users Active", str(len(metrics["users"]))],
        ["Marketplaces Affected", str(len(metrics["marketplaces"]))]
    ]
    summary_table = Table(summary_data, colWidths=[100*mm, 50*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (1,-1), 'RIGHT')
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10*mm))
    
    # Group activity
    story.append(Paragraph("Activity by Product Group", heading_style))
    if not has_data:
        story.append(create_sample_image("Product Group Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["groups"], "Changes by Group"))
    story.append(Spacer(1, 5*mm))
    
    # Phase activity
    story.append(Paragraph("Activity by Phase", heading_style))
    if not has_data:
        story.append(create_sample_image("Phase Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["phases"], "Changes by Phase"))
    story.append(Spacer(1, 5*mm))
    
    # Most active users
    story.append(Paragraph("Most Active Users", heading_style))
    if not has_data:
        story.append(create_sample_image("User Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_users = dict(sorted(metrics["users"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_users, "Top 10 Users by Activity"))
    story.append(Spacer(1, 5*mm))
    
    # Most active marketplaces
    story.append(Paragraph("Most Active Marketplaces", heading_style))
    if not has_data:
        story.append(create_sample_image("Marketplace Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_markets = dict(sorted(metrics["marketplaces"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_markets, "Top 10 Marketplaces by Activity"))
    
    # Build the PDF
    doc.build(story)
    logger.info(f"Weekly report created: {filename}")
    return filename

def create_monthly_report(year, month, force=False):
    """Create a monthly PDF report."""
    # Determine the month's date range
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(year, month + 1, 1) - timedelta(days=1)
    
    # Generate output filename
    month_str = f"{year}-{month:02d}"
    out_dir = os.path.join(REPORTS_DIR, "monthly")
    os.makedirs(out_dir, exist_ok=True)
    filename = os.path.join(out_dir, f"monthly_report_{month_str}.pdf")
    
    # Load changes for the month
    changes = load_changes(start_date, end_date)
    
    # Check if we have data
    has_data = len(changes) > 0
    
    # If no changes and not forcing, return None
    if not changes and not force:
        logger.warning(f"No changes found for month {month_str}")
        return None
    
    # Collect metrics (will use sample data if no changes)
    metrics = collect_metrics(changes)
    
    # Create PDF document
    doc = SimpleDocTemplate(filename, pagesize=A4,
                          leftMargin=25*mm, rightMargin=25*mm,
                          topMargin=20*mm, bottomMargin=20*mm)
    
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    heading_style = styles['Heading1']
    subheading_style = styles['Heading2']
    normal_style = styles['Normal']
    
    # Build the PDF content
    story = []
    
    # Title
    story.append(Paragraph(f"Monthly Smartsheet Changes Report", title_style))
    
    # Period information
    if not has_data:
        story.append(Paragraph(f"Period: {start_date.strftime('%B %Y')}", normal_style))
        story.append(Paragraph(f"<i>Sample report - no data available for this period</i>", normal_style))
    else:
        story.append(Paragraph(f"Period: {start_date.strftime('%B %Y')}", normal_style))
    
    story.append(Spacer(1, 10*mm))
    
    # Summary
    story.append(Paragraph("Monthly Summary", heading_style))
    summary_data = [
        ["Total Changes", str(metrics["total_changes"])],
        ["Groups with Activity", str(len(metrics["groups"]))],
        ["Users Active", str(len(metrics["users"]))],
        ["Marketplaces Affected", str(len(metrics["marketplaces"]))]
    ]
    summary_table = Table(summary_data, colWidths=[100*mm, 50*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (1,-1), 'RIGHT')
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10*mm))
    
    # Group activity
    story.append(Paragraph("Activity by Product Group", heading_style))
    if not has_data:
        story.append(create_sample_image("Product Group Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["groups"], "Changes by Group"))
    story.append(Spacer(1, 5*mm))
    
    # Phase activity
    story.append(Paragraph("Activity by Phase", heading_style))
    if not has_data:
        story.append(create_sample_image("Phase Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["phases"], "Changes by Phase"))
    story.append(Spacer(1, 5*mm))
    
    # Most active users
    story.append(Paragraph("Most Active Users", heading_style))
    if not has_data:
        story.append(create_sample_image("User Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_users = dict(sorted(metrics["users"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_users, "Top 10 Users by Activity"))
    story.append(Spacer(1, 5*mm))
    
    # Most active marketplaces
    story.append(PageBreak())
    story.append(Paragraph("Most Active Marketplaces", heading_style))
    if not has_data:
        story.append(create_sample_image("Marketplace Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_markets = dict(sorted(metrics["marketplaces"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_markets, "Top 10 Marketplaces by Activity"))
    story.append(Spacer(1, 10*mm))
    
    # Group detail pages
    story.append(PageBreak())
    story.append(Paragraph("Product Group Details", heading_style))
    
    if not has_data:
        # Add sample group details
        for group in ["NA", "NF", "NH"]:
            story.append(Spacer(1, 5*mm))
            story.append(Paragraph(f"Group {group}", subheading_style))
            story.append(create_sample_image(f"Group {group} Activity", "Sample data shown - no actual changes in this period"))
            story.append(Spacer(1, 5*mm))
    else:
        # Add real group details
        for group, count in sorted(metrics["groups"].items(), key=lambda x: x[1], reverse=True):
            if not group:
                continue
                
            story.append(Spacer(1, 5*mm))
            story.append(Paragraph(f"Group {group}", subheading_style))
            
            # Phase distribution for this group
            phase_data = metrics["group_phase"].get(group, {})
            if phase_data:
                story.append(make_pie_chart(phase_data, f"Phase Distribution for {group}"))
                story.append(Spacer(1, 5*mm))
    
    # Build the PDF
    doc.build(story)
    logger.info(f"Monthly report created: {filename}")
    return filename

def get_previous_week():
    """Get the start and end dates for the previous week (Monday to Sunday)."""
    today = date.today()
    previous_week_end = today - timedelta(days=today.weekday() + 1)
    previous_week_start = previous_week_end - timedelta(days=6)
    return previous_week_start, previous_week_end

def get_previous_month():
    """Get the year and month for the previous month."""
    today = date.today()
    if today.month == 1:
        return today.year - 1, 12
    else:
        return today.year, today.month - 1

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Generate Smartsheet change reports")
    report_type = parser.add_mutually_exclusive_group(required=True)
    report_type.add_argument("--weekly", action="store_true", help="Generate weekly report")
    report_type.add_argument("--monthly", action="store_true", help="Generate monthly report")
    
    parser.add_argument("--year", type=int, help="Year for report (defaults to current year)")
    parser.add_argument("--month", type=int, help="Month number for monthly report")
    parser.add_argument("--week", type=int, help="ISO week number for weekly report")
    parser.add_argument("--previous", action="store_true", help="Generate report for previous week/month")
    parser.add_argument("--force", action="store_true", help="Force report generation even with no data")
    
    args = parser.parse_args()
    
    try:
        if args.weekly:
            if args.previous:
                start_date, end_date = get_previous_week()
                filename = create_weekly_report(start_date, end_date, force=args.force)
            elif args.year and args.week:
                # Calculate date from ISO week
                start_date = datetime.fromisocalendar(args.year, args.week, 1).date()
                end_date = start_date + timedelta(days=6)
                filename = create_weekly_report(start_date, end_date, force=args.force)
            else:
                logger.error("For weekly reports, specify --previous OR (--year and --week)")
                exit(1)
                
        elif args.monthly:
            if args.previous:
                year, month = get_previous_month()
                filename = create_monthly_report(year, month, force=args.force)
            elif args.year and args.month:
                filename = create_monthly_report(args.year, args.month, force=args.force)
            else:
                logger.error("For monthly reports, specify --previous OR (--year and --month)")
                exit(1)
                
        if filename:
            logger.info(f"Report successfully generated: {filename}")
            exit(0)
        else:
            logger.warning("Report generation completed but no file was created")
            exit(1)
            
    except Exception as e:
        logger.error(f"Error generating report: {e}", exc_info=True)
        exit(1)

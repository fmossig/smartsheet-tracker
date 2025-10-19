import os
import csv
import json
from datetime import datetime, timedelta, date
from collections import defaultdict, Counter
import logging

import smartsheet
from dotenv import load_dotenv
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart, HorizontalBarChart
from reportlab.graphics.charts.legends import Legend
from reportlab.graphics.widgets.markers import makeMarker

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

# Color scheme
GROUP_COLORS = {
    "NA": colors.HexColor("#E63946"),
    "NF": colors.HexColor("#457B9D"),
    "NH": colors.HexColor("#2A9D8F"),
    "NM": colors.HexColor("#E9C46A"),
    "NP": colors.HexColor("#F4A261"),
    "NT": colors.HexColor("#9D4EDD"),
    "NV": colors.HexColor("#00B4D8"),
}

# Phase names for better readability
PHASE_NAMES = {
    "1": "Kontrolle",
    "2": "BE",
    "3": "K2",
    "4": "C",
    "5": "Reopen C2",
}

# User colors (will be generated dynamically)
USER_COLORS = {}

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
        "group_phase_user": defaultdict(lambda: defaultdict(lambda: defaultdict(int))),
    }
    
    # Add sample data if no changes
    if not changes:
        # Sample data for empty report
        metrics["groups"] = {"NA": 5, "NF": 3, "NH": 2, "NP": 1, "NT": 4, "NV": 2, "NM": 3}
        metrics["phases"] = {"1": 3, "2": 4, "3": 2, "4": 1, "5": 3}
        
        # Sample user data for each group
        sample_users = ["User A", "User B", "User C", "User D"]
        for group in metrics["groups"]:
            for phase in metrics["phases"]:
                for user in sample_users:
                    # Create random distribution of users across phases
                    if (ord(group[-1]) + ord(phase) + ord(user[-1])) % 3 == 0:
                        metrics["group_phase_user"][group][phase][user] = (ord(group[-1]) + ord(phase) + ord(user[-1])) % 5 + 1
            
        return metrics
    
    # Process real data
    for change in changes:
        group = change.get('Group', '')
        phase = change.get('Phase', '')
        user = change.get('User', '')
        
        metrics["groups"][group] += 1
        metrics["phases"][phase] += 1
        metrics["users"][user] += 1
        
        # Detailed metrics for group-phase-user breakdown
        if group and phase and user:
            metrics["group_phase_user"][group][phase][user] += 1
            
    return metrics

def generate_user_colors(users):
    """Generate consistent colors for users."""
    # Base colors
    base_colors = [
        colors.HexColor("#1f77b4"),  # Blue
        colors.HexColor("#ff7f0e"),  # Orange
        colors.HexColor("#2ca02c"),  # Green
        colors.HexColor("#d62728"),  # Red
        colors.HexColor("#9467bd"),  # Purple
        colors.HexColor("#8c564b"),  # Brown
        colors.HexColor("#e377c2"),  # Pink
        colors.HexColor("#7f7f7f"),  # Gray
        colors.HexColor("#bcbd22"),  # Yellow-green
        colors.HexColor("#17becf"),  # Cyan
    ]
    
    # Clear existing colors
    USER_COLORS.clear()
    
    # Assign colors to users
    sorted_users = sorted(users.keys())
    for i, user in enumerate(sorted_users):
        if user:  # Only assign colors to non-empty user names
            USER_COLORS[user] = base_colors[i % len(base_colors)]
    
    return USER_COLORS

def make_group_bar_chart(data_dict, title, width=500, height=300):
    """Create a bar chart showing counts by group."""
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
    chart.height = 200
    chart.width = 400
    
    # Sort groups alphabetically
    sorted_keys = sorted(data_dict.keys())
    chart.categoryAxis.categoryNames = sorted_keys
    chart.data = [list(data_dict[k] for k in sorted_keys)]
    
    # Add colors for each group
    for i, key in enumerate(sorted_keys):
        if key in GROUP_COLORS:
            chart.bars[0].fillColor = GROUP_COLORS[key]
            
    # Adjust labels
    chart.categoryAxis.labels.fontSize = 10
    chart.categoryAxis.labels.boxAnchor = 'n'
    chart.categoryAxis.labels.angle = 0
    chart.categoryAxis.labels.dy = -10
    
    # Value axis
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max(data_dict.values()) * 1.1 if data_dict else 10
    chart.valueAxis.valueStep = max(1, int(max(data_dict.values()) / 5)) if data_dict else 2
    
    # Add group colors
    for i, group in enumerate(sorted_keys):
        if group in GROUP_COLORS:
            chart.bars[(0, i)].fillColor = GROUP_COLORS.get(group, colors.steelblue)
    
    drawing.add(chart)
    return drawing

def make_phase_bar_chart(data_dict, title, width=500, height=300):
    """Create a bar chart showing counts by phase."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-15, title,
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # If data is empty, add sample data
    if not data_dict:
        data_dict = {"1": 3, "2": 4, "3": 2, "4": 1, "5": 3}
    
    # Create the bar chart
    chart = VerticalBarChart()
    chart.x = 50
    chart.y = 50
    chart.height = 200
    chart.width = 400
    
    # Sort phases numerically
    sorted_keys = sorted(data_dict.keys(), key=lambda x: int(x) if x.isdigit() else 999)
    
    # Use phase names for display
    chart.categoryAxis.categoryNames = [f"{k}: {PHASE_NAMES.get(k, '')}" for k in sorted_keys]
    chart.data = [list(data_dict[k] for k in sorted_keys)]
    
    # Adjust labels
    chart.categoryAxis.labels.fontSize = 10
    chart.categoryAxis.labels.boxAnchor = 'n'
    chart.categoryAxis.labels.angle = 0
    chart.categoryAxis.labels.dy = -10
    
    # Value axis
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max(data_dict.values()) * 1.1 if data_dict else 10
    chart.valueAxis.valueStep = max(1, int(max(data_dict.values()) / 5)) if data_dict else 2
    
    # Add phase colors (blue to red spectrum)
    phase_colors = {
        "1": colors.HexColor("#1f77b4"),  # Blue
        "2": colors.HexColor("#ff7f0e"),  # Orange
        "3": colors.HexColor("#2ca02c"),  # Green
        "4": colors.HexColor("#9467bd"),  # Purple
        "5": colors.HexColor("#d62728"),  # Red
    }
    
    for i, phase in enumerate(sorted_keys):
        chart.bars[(0, i)].fillColor = phase_colors.get(phase, colors.steelblue)
    
    drawing.add(chart)
    return drawing

def make_stacked_bar_chart_by_user(group, phase_user_data, title, width=500, height=400):
    """Create a horizontal stacked bar chart showing user contributions per phase."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-15, title,
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # Create the bar chart
    chart = HorizontalBarChart()
    chart.x = 100  # More space for phase labels
    chart.y = 50
    chart.height = 250
    chart.width = 350
    
    # Sort phases numerically
    phases = sorted(phase_user_data.keys(), key=lambda x: int(x) if x.isdigit() else 999)
    
    # Get all users across all phases
    all_users = set()
    for phase_data in phase_user_data.values():
        all_users.update(phase_data.keys())
    all_users = sorted(all_users)
    
    # Generate consistent colors for users
    user_colors = generate_user_colors({user: 1 for user in all_users})
    
    # Prepare data for stacked bars
    data = []
    for user in all_users:
        user_data = []
        for phase in phases:
            user_data.append(phase_user_data.get(phase, {}).get(user, 0))
        data.append(user_data)
    
    # Set the data
    chart.data = data
    
    # Use phase names for categories
    chart.categoryAxis.categoryNames = [f"{p}: {PHASE_NAMES.get(p, '')}" for p in phases]
    chart.categoryAxis.labels.fontSize = 10
    
    # Value axis
    max_total = 0
    for phase in phases:
        phase_total = sum(phase_user_data.get(phase, {}).values())
        if phase_total > max_total:
            max_total = phase_total
    
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max_total * 1.1 if max_total else 10
    chart.valueAxis.valueStep = max(1, int(max_total / 5)) if max_total else 2
    
    # Set colors for each user's bars
    for i, user in enumerate(all_users):
        chart.bars[i].fillColor = user_colors.get(user, colors.steelblue)
    
    drawing.add(chart)
    
    # Add a legend for users
    if all_users:
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 50
        legend.y = height - 50
        legend.columnMaximum = 8
        legend.dxTextSpace = 5
        legend.fontSize = 8
        
        legend.colorNamePairs = [(user_colors.get(user, colors.steelblue), user) for user in all_users]
        drawing.add(legend)
    
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
    
    # Try to load all changes if no data for this period
    all_changes = changes if has_data else load_changes()
    
    # Collect metrics
    metrics = collect_metrics(changes if has_data else all_changes)
    
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
    story.append(Paragraph(f"Weekly Smartsheet Changes Report", title_style))
    
    # Period information
    if not has_data:
        story.append(Paragraph(f"Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}", normal_style))
        if all_changes:
            story.append(Paragraph(f"<i>No data for this period. Showing sample with data from all available history.</i>", normal_style))
        else:
            story.append(Paragraph(f"<i>Sample report - no data available yet</i>", normal_style))
    else:
        story.append(Paragraph(f"Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}", normal_style))
    
    story.append(Spacer(1, 10*mm))
    
    # Summary
    story.append(Paragraph("Summary", heading_style))
    summary_data = [
        ["Total Changes", str(metrics["total_changes"])],
        ["Groups with Activity", str(len(metrics["groups"]))],
        ["Users Active", str(len(metrics["users"]))],
    ]
    summary_table = Table(summary_data, colWidths=[100*mm, 50*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (1,-1), 'RIGHT')
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10*mm))
    
    # Main page charts
    
    # Chart 1: Group activity bar chart
    story.append(Paragraph("Activity by Product Group", heading_style))
    group_chart = make_group_bar_chart(metrics["groups"], "Changes by Group")
    story.append(group_chart)
    story.append(Spacer(1, 10*mm))
    
    # Chart 2: Phase activity bar chart
    story.append(Paragraph("Activity by Phase", heading_style))
    phase_chart = make_phase_bar_chart(metrics["phases"], "Changes by Phase")
    story.append(phase_chart)
    
    # Group detail pages with stacked bar charts
    for group, count in sorted(metrics["group_phase_user"].items(), key=lambda x: x[0]):
        if not group:
            continue
            
        story.append(PageBreak())
        story.append(Paragraph(f"Group {group} Details", heading_style))
        story.append(Paragraph(f"Total changes: {metrics['groups'].get(group, 0)}", normal_style))
        
        # Stacked bar chart for this group
        phase_user_data = metrics["group_phase_user"].get(group, {})
        if phase_user_data:
            stacked_chart = make_stacked_bar_chart_by_user(
                group, 
                phase_user_data, 
                f"User Activity by Phase for Group {group}"
            )
            story.append(stacked_chart)
        else:
            story.append(Paragraph("No detailed data available for this group", normal_style))
    
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
        
    # Try to load all changes if no data for this period
    all_changes = changes if has_data else load_changes()
    
    # Collect metrics
    metrics = collect_metrics(changes if has_data else all_changes)
    
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
        if all_changes:
            story.append(Paragraph(f"<i>No data for this period. Showing sample with data from all available history.</i>", normal_style))
        else:
            story.append(Paragraph(f"<i>Sample report - no data available yet</i>", normal_style))
    else:
        story.append(Paragraph(f"Period: {start_date.strftime('%B %Y')}", normal_style))
    
    story.append(Spacer(1, 10*mm))
    
    # Summary
    story.append(Paragraph("Monthly Summary", heading_style))
    summary_data = [
        ["Total Changes", str(metrics["total_changes"])],
        ["Groups with Activity", str(len(metrics["groups"]))],
        ["Users Active", str(len(metrics["users"]))],
    ]
    summary_table = Table(summary_data, colWidths=[100*mm, 50*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (1,-1), 'RIGHT')
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10*mm))
    
    # Main page charts
    
    # Chart 1: Group activity bar chart
    story.append(Paragraph("Activity by Product Group", heading_style))
    group_chart = make_group_bar_chart(metrics["groups"], "Changes by Group")
    story.append(group_chart)
    story.append(Spacer(1, 10*mm))
    
    # Chart 2: Phase activity bar chart
    story.append(Paragraph("Activity by Phase", heading_style))
    phase_chart = make_phase_bar_chart(metrics["phases"], "Changes by Phase")
    story.append(phase_chart)
    
    # Group detail pages with stacked bar charts
    for group, count in sorted(metrics["group_phase_user"].items(), key=lambda x: x[0]):
        if not group:
            continue
            
        story.append(PageBreak())
        story.append(Paragraph(f"Group {group} Details", heading_style))
        story.append(Paragraph(f"Total changes: {metrics['groups'].get(group, 0)}", normal_style))
        
        # Stacked bar chart for this group
        phase_user_data = metrics["group_phase_user"].get(group, {})
        if phase_user_data:
            stacked_chart = make_stacked_bar_chart_by_user(
                group, 
                phase_user_data, 
                f"User Activity by Phase for Group {group}"
            )
            story.append(stacked_chart)
        else:
            story.append(Paragraph("No detailed data available for this group", normal_style))
    
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

def get_current_week():
    """Get the start and end dates for the current week (Monday to today)."""
    today = date.today()
    start = today - timedelta(days=today.weekday())  # Monday
    end = today
    return start, end

def get_previous_month():
    """Get the year and month for the previous month."""
    today = date.today()
    if today.month == 1:
        return today.year - 1, 12
    else:
        return today.year, today.month - 1

def get_current_month():
    """Get the year and month for the current month."""
    today = date.today()
    return today.year, today.month

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
    parser.add_argument("--current", action="store_true", help="Generate report for current week/month to date")
    parser.add_argument("--force", action="store_true", help="Force report generation even with no data")
    
    args = parser.parse_args()
    
    try:
        if args.weekly:
            if args.previous:
                start_date, end_date = get_previous_week()
                filename = create_weekly_report(start_date, end_date, force=args.force)
            elif args.current:
                start_date, end_date = get_current_week()
                filename = create_weekly_report(start_date, end_date, force=args.force)
            elif args.year and args.week:
                # Calculate date from ISO week
                start_date = datetime.fromisocalendar(args.year, args.week, 1).date()
                end_date = start_date + timedelta(days=6)
                filename = create_weekly_report(start_date, end_date, force=args.force)
            else:
                logger.error("For weekly reports, specify --previous OR --current OR (--year and --week)")
                exit(1)
                
        elif args.monthly:
            if args.previous:
                year, month = get_previous_month()
                filename = create_monthly_report(year, month, force=args.force)
            elif args.current:
                year, month = get_current_month()
                filename = create_monthly_report(year, month, force=args.force)
            elif args.year and args.month:
                filename = create_monthly_report(args.year, args.month, force=args.force)
            else:
                logger.error("For monthly reports, specify --previous OR --current OR (--year and --month)")
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

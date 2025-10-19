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
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.linecharts import HorizontalLineChart
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

GROUP_COLORS = {
    "NA": colors.HexColor("#E63946"),
    "NF": colors.HexColor("#457B9D"),
    "NH": colors.HexColor("#2A9D8F"),
    "NM": colors.HexColor("#E9C46A"),
    "NP": colors.HexColor("#F4A261"),
    "NT": colors.HexColor("#9D4EDD"),
    "NV": colors.HexColor("#00B4D8"),
}

PHASE_COLORS = {
    "1": colors.HexColor("#1F77B4"),
    "2": colors.HexColor("#FF7F0E"),
    "3": colors.HexColor("#2CA02C"),
    "4": colors.HexColor("#9467BD"),
    "5": colors.HexColor("#D62728"),
}

PHASE_NAMES = {
    "1": "Kontrolle",
    "2": "BE",
    "3": "K2",
    "4": "C",
    "5": "Reopen C2",
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
        "group_phase": defaultdict(lambda: defaultdict(int)),
        "user_groups": defaultdict(lambda: defaultdict(int)),
        "user_phases": defaultdict(lambda: defaultdict(int)),
        "marketplace_groups": defaultdict(lambda: defaultdict(int)),
    }
    
    # Add sample data if no changes
    if not changes:
        # Sample data for empty report
        metrics["groups"] = {"NA": 5, "NF": 3, "NH": 2}
        metrics["phases"] = {"1": 3, "2": 4, "3": 2, "4": 1}
        metrics["users"] = {"Sample User 1": 5, "Sample User 2": 3, "Sample User 3": 2}
        metrics["marketplaces"] = {"DE": 4, "FR": 3, "UK": 2, "ES": 1}
        
        # Sample daily data for timeline chart
        today = date.today()
        for i in range(7, 0, -1):
            day = today - timedelta(days=i)
            metrics["daily"][day.isoformat()] = i % 3 + 1
            
        return metrics
    
    for change in changes:
        group = change.get('Group', '')
        phase = change.get('Phase', '')
        user = change.get('User', '')
        marketplace = change.get('Marketplace', '')
        
        metrics["groups"][group] += 1
        metrics["phases"][phase] += 1
        metrics["users"][user] += 1
        metrics["marketplaces"][marketplace] += 1
        
        # User metrics
        if user:
            metrics["user_groups"][user][group] += 1
            metrics["user_phases"][user][phase] += 1
            
        # Marketplace metrics
        if marketplace:
            metrics["marketplace_groups"][marketplace][group] += 1
            
        # Daily timeline data
        if change.get('ParsedDate'):
            date_key = change['ParsedDate'].isoformat()
            metrics["daily"][date_key] += 1
            
        # Group-phase breakdown
        if group and phase:
            metrics["group_phase"][group][phase] += 1
            
    return metrics

def create_changes_table(changes, max_rows=20):
    """Create a table showing detailed change information."""
    if not changes:
        return None
        
    # Prepare data
    table_data = [["Date", "Group", "Phase", "User", "Marketplace"]]
    
    # Sort changes by date (newest first)
    sorted_changes = sorted(changes, 
                           key=lambda x: x.get('Timestamp', ''), 
                           reverse=True)
    
    # Add rows (limit to max_rows)
    for change in sorted_changes[:max_rows]:
        date_str = change.get('Date', '')
        if not date_str and change.get('ParsedDate'):
            date_str = change['ParsedDate'].strftime('%Y-%m-%d')
            
        row = [
            date_str,
            change.get('Group', ''),
            PHASE_NAMES.get(change.get('Phase', ''), change.get('Phase', '')),
            change.get('User', ''),
            change.get('Marketplace', '')
        ]
        table_data.append(row)
    
    # Create table
    colWidths = [30*mm, 20*mm, 25*mm, 45*mm, 30*mm]
    table = Table(table_data, colWidths=colWidths)
    
    # Style
    style = TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
    ])
    table.setStyle(style)
    
    return table

def make_pie_chart(data_dict, title, width=400, height=200, colors_dict=None):
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
        
        # Set color if available
        if colors_dict and labels[i] in colors_dict:
            pie.slices[i].fillColor = colors_dict[labels[i]]
    
    drawing.add(pie)
    return drawing

def make_bar_chart(data_dict, title, width=500, height=250, colors_dict=None):
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
    
    # Set custom colors if available
    if colors_dict:
        for i, key in enumerate(data_dict.keys()):
            if key in colors_dict:
                chart.bars[0].fillColor = colors_dict[key]
    
    # Adjust labels
    chart.categoryAxis.labels.angle = 30 if len(data_dict) > 5 else 0
    chart.categoryAxis.labels.fontSize = 8
    chart.categoryAxis.labels.dx = -10
    chart.categoryAxis.labels.dy = -5
    
    drawing.add(chart)
    return drawing

def make_line_chart(data_dict, title, width=500, height=250):
    """Create a line chart from the given data."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-15, title,
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # If data is empty, add sample data
    if not data_dict:
        # Sample data with dates
        from datetime import date, timedelta
        today = date.today()
        data_dict = {
            (today - timedelta(days=i)).isoformat(): i % 3 + 1 
            for i in range(7, 0, -1)
        }
    
    # Sort by date
    sorted_dates = sorted(data_dict.keys())
    
    # Create the line chart
    chart = HorizontalLineChart()
    chart.x = 50
    chart.y = 50
    chart.height = 150
    chart.width = 400
    
    # Format x-axis dates for better display
    display_dates = []
    for d in sorted_dates:
        try:
            dt = datetime.fromisoformat(d).date()
            display_dates.append(dt.strftime('%d/%m'))
        except:
            display_dates.append(str(d))
    
    chart.categoryAxis.categoryNames = display_dates
    chart.categoryAxis.labels.angle = 30
    chart.categoryAxis.labels.fontSize = 8
    
    # Set data
    chart.data = [[data_dict[d] for d in sorted_dates]]
    
    # Style
    chart.lines[0].strokeColor = colors.blue
    chart.lines[0].symbol = makeMarker('FilledCircle')
    chart.lines[0].strokeWidth = 2
    
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
    
    # Try to load all changes if no data for this period
    all_changes = changes if has_data else load_changes()
    
    # Collect metrics (will use sample data if no changes)
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
    small_style = ParagraphStyle(
        'Small',
        parent=styles['Normal'],
        fontSize=8,
        leading=10
    )
    
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
    if not has_data and not all_changes:
        story.append(create_sample_image("Product Group Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["groups"], "Changes by Group", colors_dict=GROUP_COLORS))
    story.append(Spacer(1, 5*mm))
    
    # Phase activity
    story.append(Paragraph("Activity by Phase", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("Phase Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["phases"], "Changes by Phase", colors_dict=PHASE_COLORS))
    story.append(Spacer(1, 5*mm))
    
    # Daily activity timeline
    story.append(PageBreak())
    story.append(Paragraph("Daily Activity Timeline", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("Daily Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_line_chart(metrics["daily"], "Changes by Day"))
    story.append(Spacer(1, 10*mm))
    
    # Most active users
    story.append(Paragraph("Most Active Users", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("User Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_users = dict(sorted(metrics["users"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_users, "Top 10 Users by Activity"))
    story.append(Spacer(1, 5*mm))
    
    # Most active marketplaces
    story.append(PageBreak())
    story.append(Paragraph("Most Active Marketplaces", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("Marketplace Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_markets = dict(sorted(metrics["marketplaces"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_markets, "Top 10 Marketplaces by Activity"))
    story.append(Spacer(1, 10*mm))
    
    # Detailed table of changes
    if has_data or (all_changes and len(all_changes) > 0):
        story.append(PageBreak())
        story.append(Paragraph("Recent Changes Detail", heading_style))
        
        if not has_data and all_changes:
            story.append(Paragraph("<i>Showing data from entire history as no changes were found for the selected period</i>", normal_style))
            
        changes_table = create_changes_table(changes if has_data else all_changes)
        if changes_table:
            story.append(changes_table)
            
            if len(changes if has_data else all_changes) > 20:
                story.append(Paragraph(f"<i>Showing 20 most recent of {len(changes if has_data else all_changes)} total changes</i>", small_style))
        else:
            story.append(Paragraph("No detailed change data available", normal_style))
    
    # User details
    if has_data or (all_changes and len(all_changes) > 0):
        story.append(PageBreak())
        story.append(Paragraph("User Activity Details", heading_style))
        
        if not has_data and all_changes:
            story.append(Paragraph("<i>Showing data from entire history as no changes were found for the selected period</i>", normal_style))
        
        # Create user activity breakdowns
        user_data = {}
        for user, count in sorted(metrics["users"].items(), key=lambda x: x[1], reverse=True)[:5]:
            if not user:
                continue
                
            user_data[user] = {
                "total": count,
                "groups": metrics["user_groups"].get(user, {}),
                "phases": metrics["user_phases"].get(user, {})
            }
        
        # Show top user details
        for user, data in user_data.items():
            story.append(Paragraph(f"User: {user}", subheading_style))
            
            # User summary
            user_summary = [
                ["Total Changes", str(data["total"])],
            ]
            
            # Add group breakdown
            for group, count in sorted(data["groups"].items(), key=lambda x: x[1], reverse=True):
                if group:
                    user_summary.append([f"Group {group}", str(count)])
            
            # Create table
            user_table = Table(user_summary, colWidths=[80*mm, 30*mm])
            user_table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ]))
            story.append(user_table)
            
            # Phase breakdown chart
            if data["phases"]:
                story.append(Spacer(1, 5*mm))
                phase_chart = make_pie_chart(data["phases"], f"Phase Distribution for {user}", colors_dict=PHASE_COLORS)
                story.append(phase_chart)
            
            story.append(Spacer(1, 10*mm))
    
    # Group detail pages
    story.append(PageBreak())
    story.append(Paragraph("Product Group Details", heading_style))
    
    if not has_data and not all_changes:
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
            story.append(Paragraph(f"Total changes: {count}", normal_style))
            
            # Phase distribution for this group
            phase_data = metrics["group_phase"].get(group, {})
            if phase_data:
                story.append(make_pie_chart(phase_data, f"Phase Distribution for Group {group}", colors_dict=PHASE_COLORS))
                
                # Phase breakdown table
                phase_table_data = [["Phase", "Count", "Percentage"]]
                for phase, phase_count in sorted(phase_data.items(), key=lambda x: x[1], reverse=True):
                    percentage = (phase_count / count) * 100 if count > 0 else 0
                    phase_table_data.append([
                        PHASE_NAMES.get(phase, phase),
                        str(phase_count),
                        f"{percentage:.1f}%"
                    ])
                    
                phase_table = Table(phase_table_data, colWidths=[50*mm, 30*mm, 30*mm])
                phase_table.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                    ('ALIGN', (1,1), (2,-1), 'RIGHT')
                ]))
                story.append(Spacer(1, 5*mm))
                story.append(phase_table)
            
            story.append(Spacer(1, 10*mm))
    
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
    
    # Collect metrics (will use sample data if no changes)
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
    small_style = ParagraphStyle(
        'Small',
        parent=styles['Normal'],
        fontSize=8,
        leading=10
    )
    
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
    if not has_data and not all_changes:
        story.append(create_sample_image("Product Group Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["groups"], "Changes by Group", colors_dict=GROUP_COLORS))
        
        # Add group table
        group_table_data = [["Group", "Changes", "Percentage"]]
        for group, count in sorted(metrics["groups"].items(), key=lambda x: x[1], reverse=True):
            if not group:
                continue
            percentage = (count / metrics["total_changes"]) * 100 if metrics["total_changes"] > 0 else 0
            group_table_data.append([group, str(count), f"{percentage:.1f}%"])
            
        group_table = Table(group_table_data, colWidths=[30*mm, 30*mm, 40*mm])
        group_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (1,1), (2,-1), 'RIGHT')
        ]))
        story.append(Spacer(1, 5*mm))
        story.append(group_table)
            
    story.append(Spacer(1, 5*mm))
    
    # Phase activity
    story.append(PageBreak())
    story.append(Paragraph("Activity by Phase", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("Phase Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_pie_chart(metrics["phases"], "Changes by Phase", colors_dict=PHASE_COLORS))
        
        # Add phase table with names
        phase_table_data = [["Phase", "Name", "Changes", "Percentage"]]
        for phase, count in sorted(metrics["phases"].items(), key=lambda x: x[1], reverse=True):
            if not phase:
                continue
            percentage = (count / metrics["total_changes"]) * 100 if metrics["total_changes"] > 0 else 0
            phase_table_data.append([
                phase, 
                PHASE_NAMES.get(phase, "Unknown"),
                str(count), 
                f"{percentage:.1f}%"
            ])
            
        phase_table = Table(phase_table_data, colWidths=[20*mm, 40*mm, 30*mm, 30*mm])
        phase_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (2,1), (3,-1), 'RIGHT')
        ]))
        story.append(Spacer(1, 5*mm))
        story.append(phase_table)
            
    story.append(Spacer(1, 5*mm))
    
    # Daily activity timeline
    story.append(PageBreak())
    story.append(Paragraph("Daily Activity Timeline", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("Daily Activity", "Sample data shown - no actual changes in this period"))
    else:
        story.append(make_line_chart(metrics["daily"], "Changes by Day"))
    story.append(Spacer(1, 10*mm))
    
    # Most active users
    story.append(Paragraph("Most Active Users", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("User Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_users = dict(sorted(metrics["users"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_users, "Top 10 Users by Activity"))
    story.append(Spacer(1, 5*mm))
    
    # Most active marketplaces
    story.append(PageBreak())
    story.append(Paragraph("Most Active Marketplaces", heading_style))
    if not has_data and not all_changes:
        story.append(create_sample_image("Marketplace Activity", "Sample data shown - no actual changes in this period"))
    else:
        top_markets = dict(sorted(metrics["marketplaces"].items(), key=lambda x: x[1], reverse=True)[:10])
        story.append(make_bar_chart(top_markets, "Top 10 Marketplaces by Activity"))
    story.append(Spacer(1, 10*mm))
    
    # Detailed table of changes
    if has_data or (all_changes and len(all_changes) > 0):
        story.append(PageBreak())
        story.append(Paragraph("Recent Changes Detail", heading_style))
        
        if not has_data and all_changes:
            story.append(Paragraph("<i>Showing data from entire history as no changes were found for the selected period</i>", normal_style))
            
        changes_table = create_changes_table(changes if has_data else all_changes, max_rows=30)
        if changes_table:
            story.append(changes_table)
            
            if len(changes if has_data else all_changes) > 30:
                story.append(Paragraph(f"<i>Showing 30 most recent of {len(changes if has_data else all_changes)} total changes</i>", small_style))
        else:
            story.append(Paragraph("No detailed change data available", normal_style))
    
    # User details
    if has_data or (all_changes and len(all_changes) > 0):
        story.append(PageBreak())
        story.append(Paragraph("User Activity Details", heading_style))
        
        if not has_data and all_changes:
            story.append(Paragraph("<i>Showing data from entire history as no changes were found for the selected period</i>", normal_style))
        
        # Create user activity breakdowns
        user_data = {}
        for user, count in sorted(metrics["users"].items(), key=lambda x: x[1], reverse=True)[:8]:
            if not user:
                continue
                
            user_data[user] = {
                "total": count,
                "groups": metrics["user_groups"].get(user, {}),
                "phases": metrics["user_phases"].get(user, {})
            }
        
        # Show top user details
        for user, data in user_data.items():
            story.append(Paragraph(f"User: {user}", subheading_style))
            
            # User summary
            user_summary = [
                ["Total Changes", str(data["total"])],
            ]
            
            # Add group breakdown
            for group, count in sorted(data["groups"].items(), key=lambda x: x[1], reverse=True):
                if group:
                    user_summary.append([f"Group {group}", str(count)])
            
            # Create table
            user_table = Table(user_summary, colWidths=[80*mm, 30*mm])
            user_table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ]))
            story.append(user_table)
            
            # Phase breakdown chart
            if data["phases"]:
                story.append(Spacer(1, 5*mm))
                phase_chart = make_pie_chart(data["phases"], f"Phase Distribution for {user}", colors_dict=PHASE_COLORS)
                story.append(phase_chart)
            
            story.append(Spacer(1, 10*mm))
    
    # Group detail pages
    story.append(PageBreak())
    story.append(Paragraph("Product Group Details", heading_style))
    
    if not has_data and not all_changes:
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
            story.append(Paragraph(f"Total changes: {count}", normal_style))
            
            # Phase distribution for this group
            phase_data = metrics["group_phase"].get(group, {})
            if phase_data:
                story.append(make_pie_chart(phase_data, f"Phase Distribution for Group {group}", colors_dict=PHASE_COLORS))
                
                # Phase breakdown table
                phase_table_data = [["Phase", "Name", "Count", "Percentage"]]
                for phase, phase_count in sorted(phase_data.items(), key=lambda x: x[1], reverse=True):
                    percentage = (phase_count / count) * 100 if count > 0 else 0
                    phase_table_data.append([
                        phase,
                        PHASE_NAMES.get(phase, "Unknown"),
                        str(phase_count),
                        f"{percentage:.1f}%"
                    ])
                    
                phase_table = Table(phase_table_data, colWidths=[20*mm, 40*mm, 25*mm, 30*mm])
                phase_table.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                    ('ALIGN', (2,1), (3,-1), 'RIGHT')
                ]))
                story.append(Spacer(1, 5*mm))
                story.append(phase_table)
            
            story.append(Spacer(1, 10*mm))
    
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

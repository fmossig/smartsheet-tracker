import os
import csv
import json
from datetime import datetime, timedelta, date
from collections import defaultdict, Counter
import logging
import math

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
from reportlab.graphics.shapes import Line, Rect
from reportlab.graphics.shapes import Path, Circle
from reportlab.graphics.shapes import Wedge

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
    "SPECIAL": 5261724614610820,  # Adding special activities sheet ID
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

# Phase names for reference (now using Phase 1, Phase 2, etc. for display)
PHASE_NAMES = {
    "1": "Phase 1",
    "2": "Phase 2",
    "3": "Phase 3",
    "4": "Phase 4",
    "5": "Phase 5",
}

# Phase colors
PHASE_COLORS = {
    "1": colors.HexColor("#1f77b4"),  # Blue
    "2": colors.HexColor("#ff7f0e"),  # Orange
    "3": colors.HexColor("#2ca02c"),  # Green
    "4": colors.HexColor("#9467bd"),  # Purple
    "5": colors.HexColor("#d62728"),  # Red
}

# Fixed total product counts for each group
TOTAL_PRODUCTS = {
    "NA": 1779,
    "NF": 1716,
    "NM": 391,
    "NH": 893,
    "NP": 394,
    "NT": 119,
    "NV": 0  # Adding NV with 0 since it wasn't provided
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
        
    # Clean up common typos
    cleaned = str(date_str).strip()
    if cleaned and not cleaned[-1].isdigit():
        cleaned = cleaned.rstrip('abcdefghijklmnopqrstuvwxyz')
        
    # Try various formats
    for fmt in ('%Y-%m-%d', '%d.%m.%Y'):
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
            
    try:
        # Try ISO format (catches many variations)
        return datetime.fromisoformat(cleaned).date()
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
        
        # Sample user data for each group - MODIFIED: Use actual user initials with counts
        sample_users = {"DM": 8, "EK": 7, "HI": 6, "SM": 5, "JHU": 9, "LK": 4}
        metrics["users"] = sample_users  # Add sample users to metrics
        
        # Create random distribution of users across phases
        for group in metrics["groups"]:
            for phase in metrics["phases"]:
                for user, count in sample_users.items():
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

def get_special_activities(days=30):
    """
    Fetch and analyze special activities from the dedicated Smartsheet.
    
    Args:
        days: Number of days to look back (default 30)
    
    Returns:
        Tuple of (sorted_category_hours, total_hours) containing activity data
    """
    client = smartsheet.Smartsheet(token)
    
    try:
        # Get the special activities sheet
        sheet_id = SHEET_IDS.get("SPECIAL")
        sheet = client.Sheets.get_sheet(sheet_id)
        logger.info(f"Retrieved special activities sheet with {len(sheet.rows)} rows")
        
        # Current date for calculating the date range
        current_date = datetime.now().date()
        cutoff_date = current_date - timedelta(days=days)
        
        # Map column titles to IDs
        col_map = {col.title: col.id for col in sheet.columns}
        
        # Check if required columns exist
        required_columns = ["Mitarbeiter", "Datum", "Kategorie", "Arbeitszeit in Std"]
        missing_columns = [col for col in required_columns if col not in col_map]
        if missing_columns:
            logger.warning(f"Missing columns in special activities sheet: {missing_columns}")
            return generate_sample_special_activities()
        
        # Extract data from rows
        activities = []
        
        for row in sheet.rows:
            activity = {}
            
            # Extract cell values
            for cell in row.cells:
                for col_name, col_id in col_map.items():
                    if cell.column_id == col_id and cell.value is not None:
                        activity[col_name] = cell.value
            
            # Skip rows without category or hours
            if not activity.get("Kategorie") or not activity.get("Arbeitszeit in Std"):
                continue
                
            # Parse date
            if "Datum" in activity:
                try:
                    activity_date = parse_date(activity["Datum"])
                    if activity_date and activity_date >= cutoff_date:
                        # Convert hours from string to float (handling comma as decimal separator)
                        if "Arbeitszeit in Std" in activity:
                            hours_str = str(activity["Arbeitszeit in Std"]).replace(',', '.')
                            try:
                                activity["Hours"] = float(hours_str)
                            except ValueError:
                                activity["Hours"] = 0
                        else:
                            activity["Hours"] = 0
                            
                        activities.append(activity)
                except Exception as e:
                    logger.debug(f"Error parsing date in special activities: {e}")
        
        # Group by category and sum hours
        category_hours = defaultdict(float)
        for activity in activities:
            category = activity.get("Kategorie", "Unknown")
            hours = activity.get("Hours", 0)
            category_hours[category] += hours
        
        # Sort by hours descending
        sorted_categories = sorted(category_hours.items(), key=lambda x: x[1], reverse=True)
        
        # Calculate total hours
        total_hours = sum(category_hours.values())
        
        logger.info(f"Processed {len(activities)} special activities in the last {days} days")
        logger.info(f"Found {len(category_hours)} categories with {total_hours:.1f} total hours")
        
        # If no activities found, return sample data
        if not sorted_categories:
            logger.info("No special activities found, using sample data")
            return generate_sample_special_activities()
            
        return sorted_categories, total_hours
        
    except Exception as e:
        logger.error(f"Error fetching special activities: {e}", exc_info=True)
        return generate_sample_special_activities()

def generate_sample_special_activities():
    """Generate sample special activities data for testing and fallback."""
    # Sample data similar to the example image
    sample_data = [
        ("Compliance", 17.25),
        ("Meetings", 23.25),
        ("Meeting Vor- & Nachbereitung", 9.75),
        ("Organisatorische Aufgaben", 20.25),
        ("Primary Case/Gerd (Technical Account Manager)", 17.25),
        ("Produkte anlegen", 10.50),
        ("Research", 1.50),
        ("Search Suppressed", 1.25),
        ("Feed File Upload", 1.00),
        ("Anderes", 9.00),
        ("A+", 3.00),
    ]
    
    # Calculate total hours
    total_hours = sum(hours for _, hours in sample_data)
    
    return sample_data, total_hours

def create_activities_pie_chart(category_hours, total_hours, width=500, height=400):
    """Create a pie chart showing hours by activity category."""
    from reportlab.graphics.charts.piecharts import Pie
    
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-20, 
                      f"Summe Stunden Sonderaktivitäten letzte 30T",
                      fontName='Helvetica-Bold', fontSize=14, textAnchor='middle'))
    
    # Create the pie chart
    pie = Pie()
    pie.x = width * 0.1
    pie.y = height / 2.5
    pie.width = min(width, height) * 0.45
    pie.height = min(width, height) * 0.45
    
    # Limit to top categories if there are too many
    max_slices = 12
    if len(category_hours) > max_slices:
        top_categories = category_hours[:max_slices-1]
        other_hours = sum(hours for _, hours in category_hours[max_slices-1:])
        if other_hours > 0:
            chart_data = top_categories + [("Other", other_hours)]
        else:
            chart_data = top_categories
    else:
        chart_data = category_hours
    
    # Set data and labels - don't set any labels on the pie directly
    pie.data = [hours for _, hours in chart_data]
    
    # Custom colorful palette for slices
    colorful_palette = [
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
        colors.HexColor("#e6ab02"),  # Gold
        colors.HexColor("#a6761d"),  # Brown
    ]
    
    # Assign colors to slices - only set fillColor which is safe
    for i in range(len(chart_data)):
        if i < len(pie.slices):
            pie.slices[i].fillColor = colorful_palette[i % len(colorful_palette)]
    
    # Add the pie to the drawing
    drawing.add(pie)
    
    # Add legend manually - positioned to the right
    legend_x = width * 0.55
    legend_y = height * 0.75
    legend_font_size = 8
    line_height = legend_font_size * 1.5
    
    for i, (category, hours) in enumerate(chart_data):
        # Skip if no hours
        if hours <= 0:
            continue
            
        y_pos = legend_y - (i * line_height)
        color_idx = i % len(colorful_palette)
        
        # Add color box
        drawing.add(Rect(
            legend_x, 
            y_pos, 
            8, 
            8, 
            fillColor=colorful_palette[color_idx],
            strokeColor=colors.black,
            strokeWidth=0.5
        ))
        
        # Add category name with percentage
        percentage = (hours / total_hours * 100) if total_hours > 0 else 0
        drawing.add(String(
            legend_x + 12,
            y_pos + 4,
            f"{category} ({percentage:.1f}%)",
            fontName='Helvetica',
            fontSize=legend_font_size
        ))
    
    # Add total hours at the bottom
    drawing.add(String(
        width/2,
        30,
        f"Gesamtstunden: {total_hours:.1f}",
        fontName='Helvetica-Bold',
        fontSize=12,
        textAnchor='middle'
    ))
    
    return drawing

def create_special_activities_breakdown(category_hours, total_hours):
    """Create a detailed table showing the breakdown of special activities."""
    # Create table data with header
    table_data = [["Category", "Hours", "% of Total"]]
    
    # Add data rows sorted by hours descending
    for category, hours in category_hours:
        percentage = (hours / total_hours * 100) if total_hours > 0 else 0
        table_data.append([category, f"{hours:.1f}", f"{percentage:.1f}%"])
    
    # Add total row
    table_data.append(["Total", f"{total_hours:.1f}", "100.0%"])
    
    # Create the table
    table = Table(table_data, colWidths=[100*mm, 30*mm, 30*mm])
    
    # Style the table
    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ('ALIGN', (1,0), (2,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ]))
    
    return table

def get_marketplace_activity(group):
    """
    For a specific group, calculate average days since last activity for each marketplace/country.
    Returns two lists: most active and most inactive marketplaces.
    """
    client = smartsheet.Smartsheet(token)
    
    try:
        # Get sheet by group ID
        sheet_id = SHEET_IDS.get(group)
        if not sheet_id:
            logger.warning(f"Sheet ID not found for group {group}")
            return [], []
            
        sheet = client.Sheets.get_sheet(sheet_id)
        
        # Find the marketplace column (now specifically "Amazon") and date columns
        marketplace_col_id = None
        date_cols = {}
        
        # Log all available column titles for debugging
        column_titles = [col.title for col in sheet.columns]
        logger.info(f"Available columns in sheet {group}: {column_titles}")
        
        # Look specifically for the "Amazon" column
        for col in sheet.columns:
            if col.title == "Amazon":  # Look for exact match with "Amazon"
                marketplace_col_id = col.id
                logger.info(f"Found marketplace column: {col.title}")
                break
                
        # Search for date columns
        for col in sheet.columns:
            if col.title in ["Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"] or "am" in col.title or "date" in col.title.lower():
                date_cols[col.title] = col.id
                
        if not marketplace_col_id:
            logger.warning(f"Marketplace column 'Amazon' not found for group {group}")
            # Generate sample data since no marketplace column was found
            return generate_sample_marketplace_data()
            
        if not date_cols:
            logger.warning(f"No date columns found for group {group}")
            return generate_sample_marketplace_data()
            
        logger.info(f"Found {len(date_cols)} date columns for group {group}")
        
        # Current date for calculating days since
        current_date = datetime.now().date()
        
        # Dictionary to collect marketplace activity data
        marketplace_data = defaultdict(list)
        
        # Process each row
        row_count = 0
        valid_rows = 0
        
        for row in sheet.rows:
            row_count += 1
            marketplace = None
            most_recent_date = None
            
            # Get the marketplace value
            for cell in row.cells:
                if cell.column_id == marketplace_col_id and cell.value:
                    marketplace = str(cell.value).strip()
                    break
            
            if not marketplace:
                continue
                
            # Find the most recent date across all date columns
            for col_name, col_id in date_cols.items():
                for cell in row.cells:
                    if cell.column_id == col_id and cell.value:
                        try:
                            date_val = parse_date(cell.value)
                            if date_val and (most_recent_date is None or date_val > most_recent_date):
                                most_recent_date = date_val
                        except Exception as e:
                            logger.debug(f"Error parsing date: {e}")
            
            # Calculate days since last activity
            if most_recent_date:
                days_since = (current_date - most_recent_date).days
                marketplace_data[marketplace].append(days_since)
                valid_rows += 1
        
        logger.info(f"Processed {row_count} rows in group {group}, found {valid_rows} valid rows with dates")
        logger.info(f"Found {len(marketplace_data)} unique marketplaces")
        
        # Calculate average days for each marketplace
        marketplace_averages = []
        for marketplace, days_list in marketplace_data.items():
            if days_list:  # Make sure we have data
                avg_days = sum(days_list) / len(days_list)
                marketplace_averages.append((marketplace, avg_days, len(days_list)))
        
        # Sort by average days (ascending for most active, descending for most inactive)
        most_active = sorted(marketplace_averages, key=lambda x: x[1])[:5]  # Lowest average days
        most_inactive = sorted(marketplace_averages, key=lambda x: x[1], reverse=True)[:5]  # Highest average days
        
        # If we don't have enough data, generate sample data
        if len(most_active) < 2 or len(most_inactive) < 2:
            logger.info(f"Not enough marketplace data found for group {group}, using sample data")
            return generate_sample_marketplace_data()
            
        return most_active, most_inactive
        
    except Exception as e:
        logger.error(f"Error getting marketplace activity for group {group}: {e}", exc_info=True)
        return generate_sample_marketplace_data()

def generate_sample_marketplace_data():
    """Generate sample marketplace activity data."""
    # Sample marketplaces with different activity levels (marketplace, days, count)
    sample_marketplaces = [
        ("Amazon DE", 12, 45),
        ("Amazon FR", 8, 32),
        ("Amazon UK", 5, 28),
        ("Amazon IT", 18, 25),
        ("Amazon ES", 22, 20),
        ("Amazon US", 3, 50),
        ("Amazon CA", 7, 15),
        ("Amazon JP", 30, 10),
        ("Amazon AU", 15, 8),
        ("Amazon BR", 25, 12)
    ]
    
    # Create sample active and inactive lists
    active_samples = [(m, d, c) for m, d, c in sample_marketplaces if d < 20][:5]
    inactive_samples = [(m, d, c) for m, d, c in sample_marketplaces if d >= 20][:5]
    
    return active_samples, inactive_samples

def create_activity_table(activity_data, title):
    """Create a table showing marketplace activity data."""
    if not activity_data:
        # Instead of using a custom paragraph style with Helvetica-Italic,
        # create a simple table with a "No data" message
        table_data = [["No marketplace data available"]]
        table = Table(table_data, colWidths=[70*mm])  # Reduced width
        table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('BACKGROUND', (0,0), (-1,-1), colors.lightgrey),
        ]))
        return table
    
    # Create table data with header - SHORTENED HEADERS
    table_data = [["Country", "Avg Days", "Products"]]  # Shorter header text
    
    # Add data rows
    for country, avg_days, count in activity_data:
        table_data.append([country, f"{avg_days:.1f}", str(count)])
    
    # Create the table with REDUCED COLUMN WIDTHS
    table = Table(table_data, colWidths=[40*mm, 20*mm, 15*mm])  # Narrower columns
    
    # Style the table
    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (2,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),  # Reduced font size
    ]))
    
    return table

# Update the generate_user_colors function to use your custom colors
def generate_user_colors(users):
    """Generate consistent colors for users with custom colors for specific users."""
    # Custom colors for specific users
    custom_colors = {
        "DM": colors.HexColor("#223459"),
        "EK": colors.HexColor("#6A5AAA"),
        "HI": colors.HexColor("#B45082"),
        "SM": colors.HexColor("#F9767F"),
        "JHU": colors.HexColor("#FFB142"),
        "LK": colors.HexColor("#FFDE70")
    }
    
    # Base colors for other users
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
    
    # First assign the custom colors
    for user in users.keys():
        if user in custom_colors:
            USER_COLORS[user] = custom_colors[user]
    
    # Then assign colors to any remaining users
    color_index = 0
    for user in sorted(users.keys()):
        if user and user not in USER_COLORS:
            USER_COLORS[user] = base_colors[color_index % len(base_colors)]
            color_index += 1
    
    return USER_COLORS

def make_group_bar_chart(data_dict, title, width=250, height=200):
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
    chart.x = 30
    chart.y = 30
    chart.height = 130
    chart.width = width - 60
    
    # Sort groups alphabetically
    sorted_keys = sorted(data_dict.keys())
    chart.categoryAxis.categoryNames = sorted_keys
    chart.data = [list(data_dict[k] for k in sorted_keys)]
    
    # Add colors for each group
    for i, key in enumerate(sorted_keys):
        if key in GROUP_COLORS:
            chart.bars[0].fillColor = GROUP_COLORS[key]
            
    # Adjust labels
    chart.categoryAxis.labels.fontSize = 8
    chart.categoryAxis.labels.boxAnchor = 'n'
    chart.categoryAxis.labels.angle = 0
    chart.categoryAxis.labels.dy = -10
    
    # Value axis
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max(data_dict.values()) * 1.1 if data_dict else 10
    chart.valueAxis.valueStep = max(1, int(max(data_dict.values()) / 5)) if data_dict else 2
    chart.valueAxis.labels.fontSize = 8
    
    # Add group colors
    for i, group in enumerate(sorted_keys):
        if group in GROUP_COLORS:
            chart.bars[(0, i)].fillColor = GROUP_COLORS.get(group, colors.steelblue)
    
    drawing.add(chart)
    return drawing

def make_phase_bar_chart(data_dict, title, width=250, height=200):
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
    chart.x = 30
    chart.y = 30
    chart.height = 130
    chart.width = width - 60
    
    # Sort phases numerically
    sorted_keys = sorted(data_dict.keys(), key=lambda x: int(x) if x.isdigit() else 999)
    
    # Use phase names for display
    chart.categoryAxis.categoryNames = [f"{PHASE_NAMES.get(k, '')}" for k in sorted_keys]
    chart.data = [list(data_dict[k] for k in sorted_keys)]
    
    # Adjust labels
    chart.categoryAxis.labels.fontSize = 8
    chart.categoryAxis.labels.boxAnchor = 'n'
    chart.categoryAxis.labels.angle = 0
    chart.categoryAxis.labels.dy = -10
    
    # Value axis
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max(data_dict.values()) * 1.1 if data_dict else 10
    chart.valueAxis.valueStep = max(1, int(max(data_dict.values()) / 5)) if data_dict else 2
    chart.valueAxis.labels.fontSize = 8
    
    # Add phase colors
    for i, phase in enumerate(sorted_keys):
        chart.bars[(0, i)].fillColor = PHASE_COLORS.get(phase, colors.steelblue)
    
    drawing.add(chart)
    return drawing

def make_group_detail_chart(group, phase_user_data, title, width=500, height=200): # Further reduced height
    """Create a horizontal stacked bar chart showing user contributions per phase."""
    from reportlab.graphics.shapes import Rect
    
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-10, title, # Reduced from height-15
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # Sort phases numerically
    phases = sorted(phase_user_data.keys(), key=lambda x: int(x) if x.isdigit() else 999)
    
    # Get all users across all phases
    all_users = set()
    for phase_data in phase_user_data.values():
        all_users.update(phase_data.keys())
    all_users = sorted(all_users)
    
    # Generate consistent colors for users
    user_colors = generate_user_colors({user: 1 for user in all_users})
    
    # Chart dimensions - FURTHER REDUCED
    chart_x = 120
    chart_y = 30  # Reduced from 40
    chart_width = 320
    chart_height = 130  # Reduced from 160
    bar_height = 14     # Reduced from 16
    spacing = 6         # Reduced from 8
    
    # Calculate maximum total for scale
    max_total = 1  # Minimum value to avoid division by zero
    for phase in phases:
        phase_total = sum(phase_user_data.get(phase, {}).values())
        if phase_total > max_total:
            max_total = phase_total
    
    # Draw each phase as a stacked bar
    for i, phase in enumerate(phases):
        y_position = chart_y + (bar_height + spacing) * i
        
        # Add phase label
        drawing.add(String(
            chart_x - 10, 
            y_position + bar_height/2, 
            PHASE_NAMES.get(phase, f"Phase {phase}"),
            fontName='Helvetica', 
            fontSize=8,  # Reduced from 9
            textAnchor='end'
        ))
        
        # Get user data for this phase
        phase_data = phase_user_data.get(phase, {})
        
        # Calculate total for this phase
        phase_total = sum(phase_data.values())
        
        # Starting position for first segment
        x_start = chart_x
        
        # Draw each user's contribution as a colored segment
        for user in all_users:
            value = phase_data.get(user, 0)
            if value > 0:
                # Calculate width of this segment proportional to its value
                segment_width = (value / max_total) * chart_width
                
                # Draw segment
                rect = Rect(
                    x_start, 
                    y_position, 
                    segment_width, 
                    bar_height, 
                    fillColor=user_colors.get(user, colors.steelblue),
                    strokeColor=colors.black,
                    strokeWidth=0.5
                )
                drawing.add(rect)
                
                # Add value label if segment is wide enough
                if segment_width > 20:
                    drawing.add(String(
                        x_start + segment_width/2,
                        y_position + bar_height/2,
                        str(value),
                        fontName='Helvetica',
                        fontSize=6,  # Reduced from 7
                        textAnchor='middle'
                    ))
                
                # Move x position for next segment
                x_start += segment_width
    
    # Draw axis line
    drawing.add(Line(
        chart_x, chart_y - 6,  # Reduced from chart_y - 8
        chart_x + chart_width, chart_y - 6,
        strokeWidth=1,
        strokeColor=colors.black
    ))
    
    # Add scale markers
    scale_steps = 5
    for i in range(scale_steps + 1):
        x_pos = chart_x + (i / scale_steps) * chart_width
        value = int((i / scale_steps) * max_total)
        
        # Tick mark
        drawing.add(Line(
            x_pos, chart_y - 6,
            x_pos, chart_y - 10,
            strokeWidth=1,
            strokeColor=colors.black
        ))
        
        # Value label
        drawing.add(String(
            x_pos, chart_y - 18,  # Reduced from chart_y - 20
            str(value),
            fontName='Helvetica',
            fontSize=6,  # Reduced from 7
            textAnchor='middle'
        ))
    
    # Return the chart and legend data
    return drawing, [(user_colors.get(user, colors.steelblue), user) for user in all_users]
    
def create_horizontal_legend(color_name_pairs, width=500, height=30):
    """Create a horizontal legend with the given color-name pairs with adjusted spacing."""
    drawing = Drawing(width, height)
    
    # Calculate spacing based on number of items to ensure they all fit on one row
    num_items = len(color_name_pairs)
    if num_items <= 0:
        return drawing
    
    # Calculate spacing to fit all items in one row
    # Subtract some padding from the width to ensure it doesn't go off-edge
    usable_width = width - 20  # Padding on both sides
    item_width = usable_width / num_items
    
    # Adjust font size based on number of items to ensure readability
    font_size = 10 if num_items <= 6 else (8 if num_items <= 8 else 6)
    
    # Use smaller color boxes if many items
    box_size = 8 if num_items <= 8 else 6
    
    for i, (color, name) in enumerate(color_name_pairs):
        # Calculate x position for this item
        x_pos = 10 + (i * item_width)
        
        # Draw color box (square)
        drawing.add(Rect(
            x_pos,
            height/2 - box_size/2,
            box_size,
            box_size,
            fillColor=color,
            strokeColor=colors.black,
            strokeWidth=0.5
        ))
        
        # Draw name with compact spacing
        drawing.add(String(
            x_pos + box_size + 2,  # Reduced space after color box
            height/2,
            name,
            fontName='Helvetica', 
            fontSize=font_size
        ))
    
    return drawing

def collect_user_group_data(metrics, target_user):
    """Collect data about a user's activity across different product groups."""
    user_data = defaultdict(lambda: defaultdict(int))
    
    # Go through the group_phase_user data and flip it for user-centric view
    for group, phase_data in metrics["group_phase_user"].items():
        for phase, user_data_dict in phase_data.items():
            if target_user in user_data_dict:
                user_data[group][phase] = user_data_dict[target_user]
    
    return user_data

def make_user_detail_chart(user, group_phase_data, width=500, height=200):
    """Create a horizontal stacked bar chart showing user's work across phases with groups as segments."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-10,
                      f"Activity by Phase for {user}",
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # Get all phases across all groups
    all_phases = set()
    for phase_data in group_phase_data.values():
        all_phases.update(phase_data.keys())
    all_phases = sorted(all_phases, key=lambda x: int(x) if x.isdigit() else 999)
    
    # Sort groups alphabetically
    groups = sorted(group_phase_data.keys())
    
    # Chart dimensions
    chart_x = 120
    chart_y = 30
    chart_width = 320
    chart_height = 130
    bar_height = 14
    spacing = 6
    
    # Calculate maximum total for scale
    max_total = 1  # Minimum value to avoid division by zero
    for phase in all_phases:
        phase_total = 0
        for group in groups:
            phase_total += group_phase_data.get(group, {}).get(phase, 0)
        if phase_total > max_total:
            max_total = phase_total
    
    # Draw each phase as a stacked bar
    for i, phase in enumerate(all_phases):
        y_position = chart_y + (bar_height + spacing) * i
        
        # Add phase label
        drawing.add(String(
            chart_x - 10, 
            y_position + bar_height/2, 
            PHASE_NAMES.get(phase, f"Phase {phase}"),
            fontName='Helvetica-Bold', 
            fontSize=8,
            textAnchor='end'
        ))
        
        # Starting position for first segment
        x_start = chart_x
        
        # Draw each group contribution as a colored segment
        for group in groups:
            value = group_phase_data.get(group, {}).get(phase, 0)
            if value > 0:
                # Calculate width of this segment proportional to its value
                segment_width = (value / max_total) * chart_width
                
                # Draw segment
                rect = Rect(
                    x_start, 
                    y_position, 
                    segment_width, 
                    bar_height, 
                    fillColor=GROUP_COLORS.get(group, colors.steelblue),
                    strokeColor=colors.black,
                    strokeWidth=0.5
                )
                drawing.add(rect)
                
                # Add value label if segment is wide enough
                if segment_width > 20:
                    drawing.add(String(
                        x_start + segment_width/2,
                        y_position + bar_height/2,
                        str(value),
                        fontName='Helvetica',
                        fontSize=6,
                        textAnchor='middle'
                    ))
                
                # Move x position for next segment
                x_start += segment_width
    
    # Draw axis line
    drawing.add(Line(
        chart_x, chart_y - 6,
        chart_x + chart_width, chart_y - 6,
        strokeWidth=1,
        strokeColor=colors.black
    ))
    
    # Add scale markers
    scale_steps = 5
    for i in range(scale_steps + 1):
        x_pos = chart_x + (i / scale_steps) * chart_width
        value = int((i / scale_steps) * max_total)
        
        # Tick mark
        drawing.add(Line(
            x_pos, chart_y - 6,
            x_pos, chart_y - 10,
            strokeWidth=1,
            strokeColor=colors.black
        ))
        
        # Value label
        drawing.add(String(
            x_pos, chart_y - 18,
            str(value),
            fontName='Helvetica',
            fontSize=6,
            textAnchor='middle'
        ))
    
    # Create legend data for groups
    legend_data = [(GROUP_COLORS.get(group, colors.steelblue), f"Group {group}") for group in groups]
    
    return drawing, legend_data

def create_user_group_distribution_chart(group_phase_data, user, width=500, height=250):
    """Create a pie chart showing distribution of user's work across product groups."""
    drawing = Drawing(width, height)
    
    # Add title
    drawing.add(String(width/2, height-20,
                      f"Group Distribution for {user} (Last 30 Days)",
                      fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    
    # Calculate totals for each group
    group_totals = {}
    for group, phase_data in group_phase_data.items():
        group_totals[group] = sum(phase_data.values())
    
    # Skip if no data
    if not group_totals or sum(group_totals.values()) == 0:
        drawing.add(String(width/2, height/2, "No data available",
                          fontName='Helvetica-Italic', fontSize=12, textAnchor='middle'))
        return drawing
    
    # Create half-circle gauges for each group
    total_changes = sum(group_totals.values())
    
    # Create pie chart data
    from reportlab.graphics.charts.piecharts import Pie
    
    pie = Pie()
    pie.x = width * 0.3  # Position left of center
    pie.y = height / 2
    pie.width = min(width, height) * 0.4
    pie.height = min(width, height) * 0.4
    
    # Sort groups by count
    sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)
    
    # Set data
    pie.data = [count for _, count in sorted_groups]
    
    # Set colors
    for i, (group, _) in enumerate(sorted_groups):
        if i < len(pie.slices):
            pie.slices[i].fillColor = GROUP_COLORS.get(group, colors.steelblue)
    
    # Add the pie to the drawing
    drawing.add(pie)
    
    # Add legend manually
    legend_x = width * 0.65  # Position to the right of pie
    legend_y = height - 50
    legend_font_size = 8
    line_height = legend_font_size * 1.5
    
    for i, (group, count) in enumerate(sorted_groups):
        y_pos = legend_y - (i * line_height)
        
        # Add color box
        drawing.add(Rect(
            legend_x, 
            y_pos, 
            8, 
            8, 
            fillColor=GROUP_COLORS.get(group, colors.steelblue),
            strokeColor=colors.black,
            strokeWidth=0.5
        ))
        
        # Add group name with percentage
        percentage = (count / total_changes * 100) if total_changes > 0 else 0
        drawing.add(String(
            legend_x + 12,
            y_pos + 4,
            f"Group {group}: {count} ({percentage:.1f}%)",
            fontName='Helvetica',
            fontSize=legend_font_size
        ))
    
    # Add total at the bottom
    drawing.add(String(
        width/2,
        30,
        f"Total changes: {total_changes}",
        fontName='Helvetica-Bold',
        fontSize=10,
        textAnchor='middle'
    ))
    
    return drawing
    
def query_smartsheet_data(group=None):
    """Query raw Smartsheet data to get activity metrics, optionally filtered by group."""
    client = smartsheet.Smartsheet(token)
    
    # Track counts
    total_items = 0
    recent_activity_items = 0
    thirty_days_ago = datetime.now() - timedelta(days=30)
    
    # Process sheets
    if group and group in SHEET_IDS:
        # Process only the specified group
        sheet_ids_to_process = {group: SHEET_IDS[group]}
    else:
        # Process all groups
        sheet_ids_to_process = SHEET_IDS
    
    # Process each sheet
    for sheet_group, sheet_id in sheet_ids_to_process.items():
        if sheet_group == "SPECIAL":  # Skip special activities sheet for this query
            continue
            
        try:
            # Get the sheet with all rows and columns
            sheet = client.Sheets.get_sheet(sheet_id)
            logger.info(f"Processing sheet {sheet_group} for activity metrics")
            
            # Map column titles to IDs for the phase columns
            phase_cols = {}
            for col in sheet.columns:
                if col.title in ["Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"]:
                    phase_cols[col.title] = col.id
            
            # Process each row
            for row in sheet.rows:
                total_items += 1
                
                # Find the most recent date across all phase columns
                most_recent_date = None
                
                for col_title, col_id in phase_cols.items():
                    for cell in row.cells:
                        if cell.column_id == col_id and cell.value:
                            try:
                                date_val = parse_date(cell.value)
                                if date_val and (most_recent_date is None or date_val > most_recent_date):
                                    most_recent_date = date_val
                            except:
                                pass
                
                # Check if the most recent date is within the last 30 days
                if most_recent_date and most_recent_date >= thirty_days_ago.date():
                    recent_activity_items += 1
                    
        except Exception as e:
            logger.error(f"Error processing sheet {sheet_group} for metrics: {e}")
    
    # Calculate percentage
    recent_percentage = (recent_activity_items / total_items * 100) if total_items > 0 else 0
    
    return {
        "total_items": total_items,
        "recent_activity_items": recent_activity_items,
        "recent_percentage": recent_percentage
    }

def draw_half_circle_gauge(value_percentage, total_value, label, width=250, height=150, 
                          color=colors.steelblue, empty_color=colors.lightgrey):
    """Draw a half-circle gauge chart showing a percentage."""
    import math
    from reportlab.graphics.shapes import Wedge, String
    
    drawing = Drawing(width, height)
    
    # Center of the half-circle
    cx = width / 2
    cy = height * 0.2  # Position near the bottom to leave room for labels
    
    # Radius of the half-circle
    radius = min(width, height) * 0.7 / 2
    
    # Create the background (empty) half-circle using a Wedge
    background = Wedge(cx, cy, radius, 0, 180, fillColor=empty_color, strokeColor=colors.grey, strokeWidth=1)
    drawing.add(background)
    
    # Calculate angle for the filled portion (0 to 180 degrees)
    filled_angle = min(180, value_percentage * 1.8)  # 100% = 180 degrees
    
    # Create the filled portion using a Wedge
    if filled_angle > 0:
        filled = Wedge(cx, cy, radius, 0, filled_angle, 
                     fillColor=color, strokeColor=colors.black, strokeWidth=0.5)
        drawing.add(filled)
    
    # Add labels
    drawing.add(String(cx, cy + radius * 1.2, label,
                     fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))
    drawing.add(String(cx, cy - radius * 0.3, f"{value_percentage:.1f}%",
                     fontName='Helvetica-Bold', fontSize=14, textAnchor='middle'))
    drawing.add(String(cx, cy - radius * 0.6, f"({total_value} items)",
                     fontName='Helvetica', fontSize=10, textAnchor='middle'))
    
    return drawing

def draw_full_gauge(total_value, label, width=250, height=150, color=colors.steelblue):
    """Draw a decorative half-circle gauge showing the total count."""
    # This is essentially draw_half_circle_gauge with 100% value
    return draw_half_circle_gauge(100, total_value, label, width, height, color)

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

def add_special_activities_section(story):
    """Add special activities section to the report."""
    styles = getSampleStyleSheet()
    heading_style = styles['Heading1']
    subheading_style = styles['Heading2']
    normal_style = styles['Normal']
    
    # Add page break before special activities section
    story.append(PageBreak())
    
    # Add section header
    story.append(Paragraph("Special Activities", heading_style))
    story.append(Spacer(1, 5*mm))
    
    # Get special activities data
    category_hours, total_hours = get_special_activities(days=30)
    
    # Add summary text
    story.append(Paragraph(
        f"Overview of special activities in the last 30 days. Total hours: {total_hours:.1f}",
        normal_style
    ))
    story.append(Spacer(1, 5*mm))
    
    # Create and add pie chart - THIS MUST REMAIN
    pie_chart = create_activities_pie_chart(category_hours, total_hours)
    story.append(pie_chart)
    
    # Add table with details
    story.append(Spacer(1, 10*mm))
    story.append(Paragraph("Detailed Breakdown", subheading_style))
    
    # Create and add breakdown table
    breakdown_table = create_special_activities_breakdown(category_hours, total_hours)
    story.append(breakdown_table)

def add_user_details_section(story, metrics):
    """Add a section showing details for each active user."""
    styles = getSampleStyleSheet()
    heading_style = styles['Heading1']
    subheading_style = styles['Heading2']
    normal_style = styles['Normal']
    
    # Add page break before user details section
    story.append(PageBreak())
    
    # Add section header
    story.append(Paragraph("User Activity Analysis", heading_style))
    story.append(Spacer(1, 5*mm))
    
    # Add explanation
    story.append(Paragraph(
        "Detailed breakdown of activity by user across product groups in the last 30 days.",
        normal_style
    ))
    story.append(Spacer(1, 10*mm))
    
    # Get active users (excluding empty user names)
    active_users = {user: count for user, count in metrics["users"].items() if user and user.strip()}
    
    if not active_users:
        story.append(Paragraph("No active users found in this period.", normal_style))
        return
    
    # Sort users by activity count
    sorted_users = sorted(active_users.items(), key=lambda x: x[1], reverse=True)
    
    # Process each user
    for i, (user, count) in enumerate(sorted_users):
        # Add page break between users (but not before the first one)
        if i > 0:
            story.append(PageBreak())
        
        # Create colored header for user
        user_color = USER_COLORS.get(user, colors.steelblue)
        user_header_data = [[f"User: {user}"]]
        user_header = Table(user_header_data, colWidths=[150*mm], rowHeights=[10*mm])
        user_header.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), user_color),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 14),
        ]))
        story.append(user_header)
        story.append(Spacer(1, 5*mm))
        
        # Add user activity summary
        story.append(Paragraph(f"Total changes: {count}", normal_style))
        story.append(Spacer(1, 5*mm))
        
        # Create user's activity by group stacked bar chart
        user_group_data = collect_user_group_data(metrics, user)
        if user_group_data:
            chart, legend_data = make_user_detail_chart(user, user_group_data)
            story.append(chart)
            
            # Add legend for groups
            if legend_data:
                legend = create_horizontal_legend(legend_data, width=400)
                story.append(legend)
                
            # No group distribution pie chart as requested
        else:
            story.append(Paragraph("No detailed data available for this user.", normal_style))
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
    
    # Create a colored heading style for group headers
    colored_group_style = ParagraphStyle(
        'ColoredGroupHeading',
        parent=heading_style,
        textColor=colors.white,  # White text
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
    ]
    summary_table = Table(summary_data, colWidths=[100*mm, 50*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (1,-1), 'RIGHT')
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10*mm))
    
    # Main page charts - side by side
    story.append(Paragraph("Activity Overview", heading_style))
    
    # Create both charts
    group_chart = make_group_bar_chart(metrics["groups"], "Changes by Group")
    phase_chart = make_phase_bar_chart(metrics["phases"], "Changes by Phase")
    
    # Put them in a table side by side
    chart_table_data = [[group_chart, phase_chart]]
    chart_table = Table(chart_table_data)
    story.append(chart_table)
    story.append(Spacer(1, 15*mm))
    
    # Group detail pages with grouped bar charts
    for group, count in sorted(metrics["group_phase_user"].items(), key=lambda x: x[0]):
        if not group:
            continue
            
        story.append(PageBreak())
        
        # Create colored header for group
        group_color = GROUP_COLORS.get(group, colors.steelblue)
        group_header_data = [[f"Group {group} Details"]]
        group_header = Table(group_header_data, colWidths=[150*mm], rowHeights=[10*mm])
        group_header.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), group_color),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 14),
        ]))
        story.append(group_header)
        story.append(Spacer(1, 5*mm))
        
        story.append(Paragraph(f"Total changes: {metrics['groups'].get(group, 0)}", normal_style))
        
        # Grouped bar chart for this group
        phase_user_data = metrics["group_phase_user"].get(group, {})
        if phase_user_data:
            chart, legend_data = make_group_detail_chart(
                group, 
                phase_user_data, 
                f"User Activity by Phase for Group {group}"
            )
            story.append(chart)
            
            # Add horizontal legend below
            if legend_data:
                # Split legend into chunks of 5 if there are many users
                chunk_size = 5
                legend_chunks = [legend_data[i:i+chunk_size] for i in range(0, len(legend_data), chunk_size)]
                
                for chunk in legend_chunks:
                    legend = create_horizontal_legend(chunk, width=400)
                    story.append(legend)
                    
            # Add the gauge charts for this group - side by side
            story.append(Spacer(1, 8*mm))  # REDUCED from 15mm to 8mm
            story.append(Paragraph("Activity Metrics", subheading_style))
                
            # Get smartsheet data for gauges filtered by group
            try:
                metrics_data = query_smartsheet_data(group)
                
                # Get color for this group
                group_color = GROUP_COLORS.get(group, colors.HexColor("#2ecc71"))
                
                # Get the fixed total products for this group
                total_products = TOTAL_PRODUCTS.get(group, 0)
                
                # Calculate correct percentage based on fixed product count
                if total_products > 0:
                    correct_percentage = (metrics_data["recent_activity_items"] / total_products) * 100
                else:
                    correct_percentage = 0
                
                # Create both gauge charts with same color
                recent_gauge = draw_half_circle_gauge(
                    correct_percentage,  # Use our recalculated percentage
                    metrics_data["recent_activity_items"],
                    "30-Day Activity",
                    color=group_color
                )
                
                # Use fixed total products value
                total_gauge = draw_full_gauge(
                    total_products,
                    "Total Products",
                    color=group_color
                )
                
                # Put them in a table side by side
                gauge_table_data = [[recent_gauge, total_gauge]]
                gauge_table = Table(gauge_table_data)
                story.append(gauge_table)
                
                # Add marketplace activity metrics after the gauge charts
                story.append(Spacer(1, 8*mm))  # REDUCED from 15mm to 8mm
                story.append(Paragraph("Marketplace Activity", subheading_style))
                
                # Get marketplace activity data
                most_active, most_inactive = get_marketplace_activity(group)
                
                # Create tables for most active and inactive marketplaces - SIDE BY SIDE
                active_table = create_activity_table(most_active, "Most Active")
                inactive_table = create_activity_table(most_inactive, "Most Inactive")
                
                # Place tables side by side with FIXED WIDTHS to prevent overflow
                marketplace_table_data = [
                    [Paragraph("Most Active", subheading_style), 
                     Paragraph("Most Inactive", subheading_style)],
                    [active_table, inactive_table]
                ]
                marketplace_table = Table(marketplace_table_data, colWidths=[75*mm, 75*mm])  # Fixed column widths
                marketplace_table.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ]))
                story.append(marketplace_table)
                    
            except Exception as e:
                logger.error(f"Error creating gauge charts for group {group}: {e}")
                # Add a placeholder if there's an error
                story.append(Paragraph(f"Could not generate gauge charts: {str(e)}", normal_style))
        else:
            story.append(Paragraph("No detailed data available for this group", normal_style))
    
    # Add user details section
    add_user_details_section(story, metrics)
    # Add special activities section
    add_special_activities_section(story)
    
    # Build the PDF
    doc.build(story)
    logger.info(f"Weekly report created: {filename}")
    # Log the absolute path for debugging
    logger.info(f"Report file created at: {os.path.abspath(filename)}")
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
    
    # Create a colored heading style for group headers
    colored_group_style = ParagraphStyle(
        'ColoredGroupHeading',
        parent=heading_style,
        textColor=colors.white,  # White text
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
    ]
    summary_table = Table(summary_data, colWidths=[100*mm, 50*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,0), (1,-1), 'RIGHT')
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10*mm))
    
    # Main page charts - side by side
    story.append(Paragraph("Activity Overview", heading_style))
    
    # Create both charts
    group_chart = make_group_bar_chart(metrics["groups"], "Changes by Group")
    phase_chart = make_phase_bar_chart(metrics["phases"], "Changes by Phase")
    
    # Put them in a table side by side
    chart_table_data = [[group_chart, phase_chart]]
    chart_table = Table(chart_table_data)
    story.append(chart_table)
    story.append(Spacer(1, 15*mm))
    
    # Group detail pages with grouped bar charts
    for group, count in sorted(metrics["group_phase_user"].items(), key=lambda x: x[0]):
        if not group:
            continue
            
        story.append(PageBreak())
        
        # Create colored header for group
        group_color = GROUP_COLORS.get(group, colors.steelblue)
        group_header_data = [[f"Group {group} Details"]]
        group_header = Table(group_header_data, colWidths=[150*mm], rowHeights=[10*mm])
        group_header.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), group_color),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 14),
        ]))
        story.append(group_header)
        story.append(Spacer(1, 5*mm))
        
        story.append(Paragraph(f"Total changes: {metrics['groups'].get(group, 0)}", normal_style))
        
        # Grouped bar chart for this group
        phase_user_data = metrics["group_phase_user"].get(group, {})
        if phase_user_data:
            chart, legend_data = make_group_detail_chart(
                group, 
                phase_user_data, 
                f"User Activity by Phase for Group {group}"
            )
            story.append(chart)
            
        # Add horizontal legend below - ALL USERS IN ONE ROW
        if legend_data:
            legend = create_horizontal_legend(legend_data, width=400)
            story.append(legend)
                
            # Add the gauge charts for this group - side by side
            story.append(Spacer(1, 8*mm))  # REDUCED from 15mm to 8mm
            story.append(Paragraph("Activity Metrics", subheading_style))
                
            # Get smartsheet data for gauges filtered by group
            try:
                metrics_data = query_smartsheet_data(group)
                
                # Get color for this group
                group_color = GROUP_COLORS.get(group, colors.HexColor("#2ecc71"))
                
                # Get the fixed total products for this group
                total_products = TOTAL_PRODUCTS.get(group, 0)
                
                # Calculate correct percentage based on fixed product count
                if total_products > 0:
                    correct_percentage = (metrics_data["recent_activity_items"] / total_products) * 100
                else:
                    correct_percentage = 0
                
                # Create both gauge charts with same color
                recent_gauge = draw_half_circle_gauge(
                    correct_percentage,  # Use our recalculated percentage
                    metrics_data["recent_activity_items"],
                    "30-Day Activity",
                    color=group_color
                )
                
                # Use fixed total products value
                total_gauge = draw_full_gauge(
                    total_products,
                    "Total Products",
                    color=group_color
                )
                
                # Put them in a table side by side
                gauge_table_data = [[recent_gauge, total_gauge]]
                gauge_table = Table(gauge_table_data)
                story.append(gauge_table)
                
                # Add marketplace activity metrics after the gauge charts
                story.append(Spacer(1, 8*mm))  # REDUCED from 15mm to 8mm
                story.append(Paragraph("Marketplace Activity", subheading_style))
                
                # Get marketplace activity data
                most_active, most_inactive = get_marketplace_activity(group)
                
                # Create tables for most active and inactive marketplaces - SIDE BY SIDE
                active_table = create_activity_table(most_active, "Most Active")
                inactive_table = create_activity_table(most_inactive, "Most Inactive")
                
                # Place tables side by side with FIXED WIDTHS to prevent overflow
                marketplace_table_data = [
                    [Paragraph("Most Active", subheading_style), 
                     Paragraph("Most Inactive", subheading_style)],
                    [active_table, inactive_table]
                ]
                marketplace_table = Table(marketplace_table_data, colWidths=[75*mm, 75*mm])  # Fixed column widths
                marketplace_table.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ]))
                story.append(marketplace_table)
                    
            except Exception as e:
                logger.error(f"Error creating gauge charts for group {group}: {e}")
                # Add a placeholder if there's an error
                story.append(Paragraph(f"Could not generate metrics: {str(e)}", normal_style))
        else:
            story.append(Paragraph("No detailed data available for this group", normal_style))
    
    # Add user details section
    add_user_details_section(story, metrics)
    # Add special activities section
    add_special_activities_section(story)
    
    # Build the PDF
    doc.build(story)
    logger.info(f"Monthly report created: {filename}")
    # Log the absolute path for debugging
    logger.info(f"Report file created at: {os.path.abspath(filename)}")
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

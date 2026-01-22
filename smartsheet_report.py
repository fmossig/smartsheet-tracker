"""
Smartsheet Report Generator v2.0
================================
A professionally designed PDF report generator for Smartsheet change tracking.

Design Philosophy:
- Clean, modern visual hierarchy
- Consistent spacing system (8pt grid)
- Professional color palette
- Clear data visualization
- Reliable, tested output

Author: Redesigned for Noctua Returns Department
"""

import os
import csv
import json
from datetime import datetime, timedelta, date
from collections import defaultdict, Counter
import logging
import math

# Optional imports for Smartsheet API
try:
    import smartsheet
    SMARTSHEET_AVAILABLE = True
except ImportError:
    SMARTSHEET_AVAILABLE = False
    smartsheet = None

from dotenv import load_dotenv
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
pt = 1  # 1 point = 1 unit in reportlab
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, KeepTogether, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, String, Line, Rect, Circle, Wedge
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.pdfbase.pdfmetrics import stringWidth

# =============================================================================
# LOGGING SETUP
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("smartsheet_report.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# =============================================================================
# ENVIRONMENT & CONFIGURATION
# =============================================================================
load_dotenv()
token = os.getenv("SMARTSHEET_TOKEN")
if not token and SMARTSHEET_AVAILABLE:
    logger.warning("SMARTSHEET_TOKEN not found - API features disabled")

# =============================================================================
# CONSTANTS
# =============================================================================

# Sheet IDs
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
    "SPECIAL": 5261724614610820,
}

# Report metadata sheet
REPORT_METADATA_SHEET_ID = 7888169555939204
MONTHLY_REPORT_ATTACHMENT_ROW_ID = 5089581251235716
WEEKLY_REPORT_ATTACHMENT_ROW_ID = 1192484760260484
MONTHLY_METADATA_ROW_ID = 5089581251235716
WEEKLY_METADATA_ROW_ID = 1192484760260484

# Fixed total product counts
TOTAL_PRODUCTS = {
    "NA": 1779,
    "NF": 1716,
    "NM": 391,
    "NH": 893,
    "NP": 394,
    "NT": 119,
    "NV": 0
}

# Phase configuration
PHASE_NAMES = {
    "1": "Phase 1 - Kontrolle",
    "2": "Phase 2 - BE",
    "3": "Phase 3 - K2",
    "4": "Phase 4 - C",
    "5": "Phase 5 - Reopen",
}

PHASE_SHORT = {
    "1": "P1",
    "2": "P2", 
    "3": "P3",
    "4": "P4",
    "5": "P5",
}

# Directories
DATA_DIR = "tracking_data"
REPORTS_DIR = "reports"
CHANGES_FILE = os.path.join(DATA_DIR, "change_history.csv")
os.makedirs(REPORTS_DIR, exist_ok=True)

# =============================================================================
# DESIGN SYSTEM - Colors, Typography, Spacing
# =============================================================================

class DesignSystem:
    """Centralized design system for consistent styling."""
    
    # Primary color palette - Professional blue-based scheme
    PRIMARY = colors.HexColor("#1E3A5F")      # Deep navy blue
    PRIMARY_LIGHT = colors.HexColor("#2E5077") # Lighter navy
    SECONDARY = colors.HexColor("#3D7EAA")    # Medium blue
    ACCENT = colors.HexColor("#65B891")       # Teal accent
    
    # Neutral palette
    WHITE = colors.HexColor("#FFFFFF")
    GRAY_50 = colors.HexColor("#F8FAFC")      # Lightest gray
    GRAY_100 = colors.HexColor("#F1F5F9")     # Very light gray
    GRAY_200 = colors.HexColor("#E2E8F0")     # Light gray
    GRAY_300 = colors.HexColor("#CBD5E1")     # Medium light gray
    GRAY_400 = colors.HexColor("#94A3B8")     # Medium gray
    GRAY_500 = colors.HexColor("#64748B")     # Gray
    GRAY_600 = colors.HexColor("#475569")     # Dark gray
    GRAY_700 = colors.HexColor("#334155")     # Darker gray
    GRAY_800 = colors.HexColor("#1E293B")     # Very dark gray
    BLACK = colors.HexColor("#0F172A")        # Near black
    
    # Semantic colors
    SUCCESS = colors.HexColor("#10B981")      # Green
    WARNING = colors.HexColor("#F59E0B")      # Amber
    ERROR = colors.HexColor("#EF4444")        # Red
    INFO = colors.HexColor("#3B82F6")         # Blue
    
    # Group colors - Original Noctua color scheme
    GROUP_COLORS = {
        "NA": colors.HexColor("#E63946"),     # Red
        "NF": colors.HexColor("#457B9D"),     # Blue
        "NH": colors.HexColor("#2A9D8F"),     # Teal
        "NM": colors.HexColor("#E9C46A"),     # Gold
        "NP": colors.HexColor("#F4A261"),     # Orange
        "NT": colors.HexColor("#9D4EDD"),     # Purple
        "NV": colors.HexColor("#00B4D8"),     # Cyan
    }
    
    # Phase colors - Sequential blue gradient
    PHASE_COLORS = {
        "1": colors.HexColor("#1E40AF"),      # Dark blue
        "2": colors.HexColor("#2563EB"),      # Blue
        "3": colors.HexColor("#3B82F6"),      # Light blue
        "4": colors.HexColor("#60A5FA"),      # Lighter blue
        "5": colors.HexColor("#93C5FD"),      # Lightest blue
    }
    
    # User colors - Original Noctua team colors
    USER_COLORS = {
        "DM": colors.HexColor("#223459"),     # Dark navy
        "EK": colors.HexColor("#6A5AAA"),     # Purple
        "HI": colors.HexColor("#B45082"),     # Magenta
        "SM": colors.HexColor("#F9767F"),     # Coral
        "JHU": colors.HexColor("#FFB142"),    # Orange
        "LK": colors.HexColor("#FFDE70"),     # Yellow
    }
    
    # Base fallback colors for additional users
    FALLBACK_USER_COLORS = [
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
    
    # Status colors for overdue
    STATUS_COLORS = {
        "Aktuell": colors.HexColor("#10B981"),   # Green - current
        "<30": colors.HexColor("#F59E0B"),       # Amber - warning
        "31 - 60": colors.HexColor("#F97316"),   # Orange - attention
        ">60": colors.HexColor("#DC2626"),       # Red - critical
    }
    
    # Spacing system (8pt grid)
    SPACE_XS = 4 * pt
    SPACE_SM = 8 * pt
    SPACE_MD = 16 * pt
    SPACE_LG = 24 * pt
    SPACE_XL = 32 * pt
    SPACE_2XL = 48 * pt
    
    # Page margins
    MARGIN_LEFT = 20 * mm
    MARGIN_RIGHT = 20 * mm
    MARGIN_TOP = 25 * mm
    MARGIN_BOTTOM = 20 * mm
    
    # Typography
    FONT_FAMILY = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"
    
    FONT_SIZE_XS = 8
    FONT_SIZE_SM = 9
    FONT_SIZE_BASE = 10
    FONT_SIZE_MD = 11
    FONT_SIZE_LG = 14
    FONT_SIZE_XL = 18
    FONT_SIZE_2XL = 24
    FONT_SIZE_3XL = 32

    @classmethod
    def get_user_color(cls, user):
        """Get color for user, with fallback generation."""
        if user in cls.USER_COLORS:
            return cls.USER_COLORS[user]
        # Generate consistent color based on user name hash
        return cls.FALLBACK_USER_COLORS[hash(user) % len(cls.FALLBACK_USER_COLORS)]


# =============================================================================
# CUSTOM STYLES
# =============================================================================

def create_styles():
    """Create custom paragraph styles for the report."""
    styles = getSampleStyleSheet()
    
    # Report title
    styles.add(ParagraphStyle(
        name='ReportTitle',
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_2XL,
        textColor=DesignSystem.PRIMARY,
        spaceAfter=DesignSystem.SPACE_SM,
        alignment=TA_LEFT,
    ))
    
    # Report subtitle
    styles.add(ParagraphStyle(
        name='ReportSubtitle',
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_MD,
        textColor=DesignSystem.GRAY_500,
        spaceAfter=DesignSystem.SPACE_LG,
        alignment=TA_LEFT,
    ))
    
    # Section header
    styles.add(ParagraphStyle(
        name='SectionHeader',
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_LG,
        textColor=DesignSystem.PRIMARY,
        spaceBefore=DesignSystem.SPACE_LG,
        spaceAfter=DesignSystem.SPACE_MD,
        borderPadding=0,
    ))
    
    # Subsection header
    styles.add(ParagraphStyle(
        name='SubsectionHeader',
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_MD,
        textColor=DesignSystem.GRAY_700,
        spaceBefore=DesignSystem.SPACE_MD,
        spaceAfter=DesignSystem.SPACE_SM,
    ))
    
    # Body text
    styles.add(ParagraphStyle(
        name='ReportBody',
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_BASE,
        textColor=DesignSystem.GRAY_700,
        spaceAfter=DesignSystem.SPACE_SM,
        leading=14,
    ))
    
    # Small text
    styles.add(ParagraphStyle(
        name='SmallText',
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_SM,
        textColor=DesignSystem.GRAY_500,
        spaceAfter=DesignSystem.SPACE_XS,
    ))
    
    # Caption
    styles.add(ParagraphStyle(
        name='Caption',
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_XS,
        textColor=DesignSystem.GRAY_400,
        alignment=TA_CENTER,
        spaceBefore=DesignSystem.SPACE_XS,
    ))
    
    # KPI value
    styles.add(ParagraphStyle(
        name='KPIValue',
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_XL,
        textColor=DesignSystem.PRIMARY,
        alignment=TA_CENTER,
    ))
    
    # KPI label
    styles.add(ParagraphStyle(
        name='KPILabel',
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_SM,
        textColor=DesignSystem.GRAY_500,
        alignment=TA_CENTER,
    ))
    
    return styles


# =============================================================================
# CUSTOM FLOWABLES
# =============================================================================

class SectionDivider(Flowable):
    """A horizontal line divider between sections."""
    
    def __init__(self, width, color=None, thickness=1):
        Flowable.__init__(self)
        self.width = width
        self.color = color or DesignSystem.GRAY_200
        self.thickness = thickness
        self.height = DesignSystem.SPACE_MD
    
    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        y = self.height / 2
        self.canv.line(0, y, self.width, y)


class KPICard(Flowable):
    """A card displaying a key performance indicator."""
    
    def __init__(self, value, label, width=80*mm, height=50*mm, 
                 color=None, show_border=True):
        Flowable.__init__(self)
        self.value = str(value)
        self.label = label
        self.card_width = width
        self.card_height = height
        self.color = color or DesignSystem.PRIMARY
        self.show_border = show_border
        self.width = width
        self.height = height
    
    def draw(self):
        # Background
        self.canv.setFillColor(DesignSystem.WHITE)
        self.canv.setStrokeColor(DesignSystem.GRAY_200 if self.show_border else DesignSystem.WHITE)
        self.canv.setLineWidth(1)
        self.canv.roundRect(0, 0, self.card_width, self.card_height, 4*mm, fill=1, stroke=1)
        
        # Top accent line
        self.canv.setFillColor(self.color)
        self.canv.rect(0, self.card_height - 3*mm, self.card_width, 3*mm, fill=1, stroke=0)
        
        # Value
        self.canv.setFillColor(self.color)
        self.canv.setFont(DesignSystem.FONT_BOLD, DesignSystem.FONT_SIZE_2XL)
        value_width = stringWidth(self.value, DesignSystem.FONT_BOLD, DesignSystem.FONT_SIZE_2XL)
        self.canv.drawString(
            (self.card_width - value_width) / 2,
            self.card_height / 2 + 2*mm,
            self.value
        )
        
        # Label
        self.canv.setFillColor(DesignSystem.GRAY_500)
        self.canv.setFont(DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_SM)
        label_width = stringWidth(self.label, DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_SM)
        self.canv.drawString(
            (self.card_width - label_width) / 2,
            self.card_height / 2 - 8*mm,
            self.label
        )


class GroupHeader(Flowable):
    """A styled header for group sections."""
    
    def __init__(self, group_name, total_changes, width=170*mm):
        Flowable.__init__(self)
        self.group_name = group_name
        self.total_changes = total_changes
        self.box_width = width
        self.box_height = 18 * mm
        self.width = width
        self.height = self.box_height
        self.color = DesignSystem.GROUP_COLORS.get(group_name, DesignSystem.PRIMARY)
    
    def draw(self):
        # Background
        self.canv.setFillColor(self.color)
        self.canv.roundRect(0, 0, self.box_width, self.box_height, 3*mm, fill=1, stroke=0)
        
        # Group name
        self.canv.setFillColor(DesignSystem.WHITE)
        self.canv.setFont(DesignSystem.FONT_BOLD, DesignSystem.FONT_SIZE_LG)
        self.canv.drawString(5*mm, self.box_height/2 - 2*mm, f"Group {self.group_name}")
        
        # Total changes badge
        badge_text = f"{self.total_changes} changes"
        badge_width = stringWidth(badge_text, DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_SM) + 8*mm
        badge_x = self.box_width - badge_width - 5*mm
        
        # Badge background (semi-transparent white)
        self.canv.setFillColor(colors.Color(1, 1, 1, alpha=0.2))
        self.canv.roundRect(badge_x, self.box_height/2 - 4*mm, badge_width, 8*mm, 2*mm, fill=1, stroke=0)
        
        # Badge text
        self.canv.setFillColor(DesignSystem.WHITE)
        self.canv.setFont(DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_SM)
        self.canv.drawString(badge_x + 4*mm, self.box_height/2 - 1.5*mm, badge_text)


class UserHeader(Flowable):
    """A styled header for user sections."""
    
    def __init__(self, user_name, total_changes, width=170*mm):
        Flowable.__init__(self)
        self.user_name = user_name
        self.total_changes = total_changes
        self.box_width = width
        self.box_height = 14 * mm
        self.width = width
        self.height = self.box_height
        self.color = DesignSystem.get_user_color(user_name)
    
    def draw(self):
        # Background
        self.canv.setFillColor(self.color)
        self.canv.roundRect(0, 0, self.box_width, self.box_height, 3*mm, fill=1, stroke=0)
        
        # User name
        self.canv.setFillColor(DesignSystem.WHITE)
        self.canv.setFont(DesignSystem.FONT_BOLD, DesignSystem.FONT_SIZE_MD)
        self.canv.drawString(5*mm, self.box_height/2 - 2*mm, f"User: {self.user_name}")
        
        # Total changes
        changes_text = f"{self.total_changes} total changes"
        self.canv.setFont(DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_SM)
        text_width = stringWidth(changes_text, DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_SM)
        self.canv.drawString(self.box_width - text_width - 5*mm, self.box_height/2 - 1.5*mm, changes_text)


# =============================================================================
# CHART COMPONENTS
# =============================================================================

def create_bar_chart(data_dict, title, width=220, height=180, 
                     color_map=None, show_values=True):
    """Create a clean vertical bar chart."""
    drawing = Drawing(width, height)
    
    if not data_dict:
        data_dict = {"No Data": 0}
    
    # Sort keys
    sorted_keys = sorted(data_dict.keys())
    values = [data_dict[k] for k in sorted_keys]
    max_value = max(values) if values and max(values) > 0 else 1
    
    # Chart dimensions
    chart_x = 35
    chart_y = 30
    chart_width = width - 50
    chart_height = height - 60
    
    # Title
    drawing.add(String(
        width / 2, height - 12,
        title,
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_BASE,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_700
    ))
    
    # Draw bars
    bar_count = len(sorted_keys)
    if bar_count == 0:
        return drawing
    
    bar_spacing = chart_width / bar_count
    bar_width = bar_spacing * 0.7
    
    for i, key in enumerate(sorted_keys):
        value = data_dict[key]
        bar_height = (value / max_value) * chart_height if max_value > 0 else 0
        
        x = chart_x + i * bar_spacing + (bar_spacing - bar_width) / 2
        y = chart_y
        
        # Get color
        if color_map and key in color_map:
            bar_color = color_map[key]
        else:
            bar_color = DesignSystem.PRIMARY
        
        # Draw bar with rounded top
        if bar_height > 0:
            drawing.add(Rect(
                x, y, bar_width, bar_height,
                fillColor=bar_color,
                strokeColor=None,
                strokeWidth=0
            ))
        
        # Value label
        if show_values and value > 0:
            drawing.add(String(
                x + bar_width / 2,
                y + bar_height + 3,
                str(value),
                fontName=DesignSystem.FONT_FAMILY,
                fontSize=DesignSystem.FONT_SIZE_XS,
                textAnchor='middle',
                fillColor=DesignSystem.GRAY_600
            ))
        
        # Category label
        drawing.add(String(
            x + bar_width / 2,
            chart_y - 12,
            str(key),
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_XS,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_500
        ))
    
    # Y-axis line
    drawing.add(Line(
        chart_x - 2, chart_y,
        chart_x - 2, chart_y + chart_height,
        strokeColor=DesignSystem.GRAY_300,
        strokeWidth=1
    ))
    
    # X-axis line
    drawing.add(Line(
        chart_x - 2, chart_y,
        chart_x + chart_width, chart_y,
        strokeColor=DesignSystem.GRAY_300,
        strokeWidth=1
    ))
    
    return drawing


def create_horizontal_stacked_bar(phase_user_data, title, width=480, height=160):
    """Create a horizontal stacked bar chart showing user contributions per phase."""
    drawing = Drawing(width, height)
    
    if not phase_user_data:
        drawing.add(String(
            width / 2, height / 2,
            "No data available",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_400
        ))
        return drawing, []
    
    # Title
    drawing.add(String(
        width / 2, height - 10,
        title,
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_BASE,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_700
    ))
    
    # Get all users and phases
    phases = sorted(phase_user_data.keys(), key=lambda x: int(x) if x.isdigit() else 999)
    all_users = set()
    for phase_data in phase_user_data.values():
        all_users.update(phase_data.keys())
    all_users = sorted(all_users)
    
    # Calculate max total
    max_total = 1
    for phase in phases:
        phase_total = sum(phase_user_data.get(phase, {}).values())
        max_total = max(max_total, phase_total)
    
    # Chart dimensions
    chart_x = 80
    chart_y = 25
    chart_width = width - 100
    bar_height = min(18, (height - 60) / len(phases) - 4) if phases else 18
    spacing = 4
    
    # Draw bars
    for i, phase in enumerate(phases):
        y_pos = chart_y + (bar_height + spacing) * i
        
        # Phase label
        phase_label = PHASE_SHORT.get(phase, f"P{phase}")
        drawing.add(String(
            chart_x - 8,
            y_pos + bar_height / 2 - 3,
            phase_label,
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='end',
            fillColor=DesignSystem.GRAY_600
        ))
        
        # Draw stacked segments
        x_start = chart_x
        phase_data = phase_user_data.get(phase, {})
        
        for user in all_users:
            value = phase_data.get(user, 0)
            if value > 0:
                segment_width = (value / max_total) * chart_width
                
                drawing.add(Rect(
                    x_start, y_pos,
                    segment_width, bar_height,
                    fillColor=DesignSystem.get_user_color(user),
                    strokeColor=DesignSystem.WHITE,
                    strokeWidth=0.5
                ))
                
                # Value label if wide enough
                if segment_width > 18:
                    drawing.add(String(
                        x_start + segment_width / 2,
                        y_pos + bar_height / 2 - 3,
                        str(value),
                        fontName=DesignSystem.FONT_FAMILY,
                        fontSize=DesignSystem.FONT_SIZE_XS,
                        textAnchor='middle',
                        fillColor=DesignSystem.WHITE
                    ))
                
                x_start += segment_width
    
    # Build legend data
    legend_data = [(DesignSystem.get_user_color(user), user) for user in all_users]
    
    return drawing, legend_data


def create_legend_row(color_name_pairs, width=480, height=20):
    """Create a horizontal legend row."""
    drawing = Drawing(width, height)
    
    if not color_name_pairs:
        return drawing
    
    num_items = len(color_name_pairs)
    available_width = width - 20
    item_width = available_width / num_items
    
    box_size = 8
    font_size = DesignSystem.FONT_SIZE_XS
    
    for i, (color, name) in enumerate(color_name_pairs):
        x = 10 + i * item_width
        y = height / 2 - box_size / 2
        
        # Color box
        drawing.add(Rect(
            x, y, box_size, box_size,
            fillColor=color,
            strokeColor=DesignSystem.GRAY_300,
            strokeWidth=0.5
        ))
        
        # Label
        label = name if len(name) <= 8 else name[:7] + "…"
        drawing.add(String(
            x + box_size + 4,
            y + 1,
            label,
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=font_size,
            fillColor=DesignSystem.GRAY_600
        ))
    
    return drawing


def create_donut_chart(data_dict, title, width=200, height=180, color_map=None):
    """Create a donut chart with center label."""
    drawing = Drawing(width, height)
    
    if not data_dict or sum(data_dict.values()) == 0:
        drawing.add(String(
            width / 2, height / 2,
            "No data",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_400
        ))
        return drawing
    
    # Title
    drawing.add(String(
        width / 2, height - 10,
        title,
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_SM,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_700
    ))
    
    # Chart center and radius
    cx, cy = width / 2, height / 2 - 5
    outer_radius = min(width, height) * 0.32
    inner_radius = outer_radius * 0.55
    
    # Calculate total and draw segments
    total = sum(data_dict.values())
    start_angle = 90  # Start from top
    
    sorted_items = sorted(data_dict.items(), key=lambda x: x[1], reverse=True)
    
    for key, value in sorted_items:
        if value <= 0:
            continue
        
        angle_extent = (value / total) * 360
        
        # Get color
        if color_map and key in color_map:
            segment_color = color_map[key]
        else:
            segment_color = DesignSystem.PRIMARY
        
        # Draw wedge
        drawing.add(Wedge(
            cx, cy, outer_radius,
            start_angle - angle_extent, start_angle,
            fillColor=segment_color,
            strokeColor=DesignSystem.WHITE,
            strokeWidth=2
        ))
        
        start_angle -= angle_extent
    
    # Inner circle (creates donut effect)
    drawing.add(Circle(
        cx, cy, inner_radius,
        fillColor=DesignSystem.WHITE,
        strokeColor=None
    ))
    
    # Center text - total
    drawing.add(String(
        cx, cy + 4,
        str(total),
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_LG,
        textAnchor='middle',
        fillColor=DesignSystem.PRIMARY
    ))
    
    drawing.add(String(
        cx, cy - 8,
        "Total",
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_XS,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_500
    ))
    
    return drawing


def create_gauge_chart(percentage, label, width=160, height=100, color=None):
    """Create a half-circle gauge chart."""
    drawing = Drawing(width, height)
    
    color = color or DesignSystem.PRIMARY
    
    # Center and radius
    cx = width / 2
    cy = 25
    radius = min(width * 0.4, height * 0.65)
    
    # Background arc
    drawing.add(Wedge(
        cx, cy, radius,
        0, 180,
        fillColor=DesignSystem.GRAY_200,
        strokeColor=None
    ))
    
    # Value arc
    filled_angle = min(180, percentage * 1.8)
    if filled_angle > 0:
        drawing.add(Wedge(
            cx, cy, radius,
            180 - filled_angle, 180,
            fillColor=color,
            strokeColor=None
        ))
    
    # Inner circle (hollow center)
    inner_radius = radius * 0.6
    drawing.add(Circle(
        cx, cy, inner_radius,
        fillColor=DesignSystem.WHITE,
        strokeColor=None
    ))
    
    # Percentage text
    drawing.add(String(
        cx, cy - 5,
        f"{percentage:.0f}%",
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_MD,
        textAnchor='middle',
        fillColor=color
    ))
    
    # Label
    drawing.add(String(
        cx, height - 12,
        label,
        fontName=DesignSystem.FONT_FAMILY,
        fontSize=DesignSystem.FONT_SIZE_XS,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_600
    ))
    
    return drawing


def create_status_bar(status_values, width=450, height=60):
    """Create a horizontal stacked bar for status breakdown."""
    drawing = Drawing(width, height)
    
    total = sum(status_values.values())
    if total == 0:
        drawing.add(String(
            width / 2, height / 2,
            "No status data available",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_400
        ))
        return drawing
    
    # Title
    drawing.add(String(
        width / 2, height - 8,
        "Product Status Overview",
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_SM,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_700
    ))
    
    # Bar dimensions
    bar_x = 20
    bar_y = 22
    bar_width = width - 40
    bar_height = 16
    
    # Draw segments
    x_start = bar_x
    status_order = ["Aktuell", "<30", "31 - 60", ">60"]
    
    for status in status_order:
        value = status_values.get(status, 0)
        if value > 0:
            segment_width = (value / total) * bar_width
            color = DesignSystem.STATUS_COLORS.get(status, DesignSystem.GRAY_400)
            
            drawing.add(Rect(
                x_start, bar_y,
                segment_width, bar_height,
                fillColor=color,
                strokeColor=DesignSystem.WHITE,
                strokeWidth=1
            ))
            
            # Value label if wide enough
            if segment_width > 25:
                drawing.add(String(
                    x_start + segment_width / 2,
                    bar_y + bar_height / 2 - 3,
                    str(value),
                    fontName=DesignSystem.FONT_FAMILY,
                    fontSize=DesignSystem.FONT_SIZE_XS,
                    textAnchor='middle',
                    fillColor=DesignSystem.WHITE
                ))
            
            x_start += segment_width
    
    # Legend
    legend_y = 5
    legend_x = bar_x
    for status in status_order:
        color = DesignSystem.STATUS_COLORS.get(status, DesignSystem.GRAY_400)
        value = status_values.get(status, 0)
        pct = (value / total * 100) if total > 0 else 0
        
        drawing.add(Rect(legend_x, legend_y, 6, 6, fillColor=color, strokeColor=None))
        label = f"{status}: {pct:.0f}%"
        drawing.add(String(
            legend_x + 9, legend_y,
            label,
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=7,
            fillColor=DesignSystem.GRAY_600
        ))
        legend_x += stringWidth(label, DesignSystem.FONT_FAMILY, 7) + 20
    
    return drawing


# =============================================================================
# DATA FUNCTIONS
# =============================================================================

def parse_date(date_str):
    """Parse date from string, supporting multiple formats."""
    if not date_str:
        return None
    
    if isinstance(date_str, date):
        return date_str
    if isinstance(date_str, datetime):
        return date_str.date()
    
    cleaned = str(date_str).strip()
    if cleaned and not cleaned[-1].isdigit():
        cleaned = cleaned.rstrip('abcdefghijklmnopqrstuvwxyz')
    
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%Y-%m-%dT%H:%M:%S'):
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    
    try:
        return datetime.fromisoformat(cleaned).date()
    except:
        return None


def load_changes(start_date=None, end_date=None):
    """Load changes from CSV file within date range."""
    if not os.path.exists(CHANGES_FILE):
        logger.error(f"Changes file not found: {CHANGES_FILE}")
        return []
    
    changes = []
    try:
        with open(CHANGES_FILE, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    ts = datetime.strptime(row['Timestamp'], "%Y-%m-%d %H:%M:%S").date()
                    
                    if start_date and end_date:
                        if start_date <= ts <= end_date:
                            row['ParsedDate'] = parse_date(row.get('Date'))
                            changes.append(row)
                    else:
                        row['ParsedDate'] = parse_date(row.get('Date'))
                        changes.append(row)
                except (ValueError, KeyError) as e:
                    logger.warning(f"Error parsing row: {e}")
                    continue
    except Exception as e:
        logger.error(f"Error reading changes file: {e}")
    
    logger.info(f"Loaded {len(changes)} changes")
    return changes


def collect_metrics(changes):
    """Collect metrics from changes data."""
    metrics = {
        "total_changes": len(changes),
        "groups": defaultdict(int),
        "phases": defaultdict(int),
        "users": defaultdict(int),
        "group_phase_user": defaultdict(lambda: defaultdict(lambda: defaultdict(int))),
        "marketplaces": defaultdict(int),
    }
    
    if not changes:
        return metrics
    
    for change in changes:
        group = change.get('Group', '')
        phase = change.get('Phase', '')
        user = change.get('User', '')
        marketplace = change.get('Marketplace', '')
        
        if group:
            metrics["groups"][group] += 1
        if phase:
            metrics["phases"][phase] += 1
        if user:
            metrics["users"][user] += 1
        if marketplace:
            metrics["marketplaces"][marketplace] += 1
        
        if group and phase and user:
            metrics["group_phase_user"][group][phase][user] += 1
    
    return metrics


def get_sheet_summary_data(sheet_id):
    """Fetch sheet summary fields."""
    if not SMARTSHEET_AVAILABLE or not token:
        return None
    try:
        client = smartsheet.Smartsheet(token)
        summary = client.Sheets.get_sheet_summary(sheet_id)
        return {field.title: field.display_value for field in summary.fields}
    except Exception as e:
        logger.error(f"Error fetching sheet summary: {e}")
        return None


def get_column_map(sheet_id):
    """Get column name to ID mapping."""
    if not SMARTSHEET_AVAILABLE or not token:
        return None
    try:
        client = smartsheet.Smartsheet(token)
        sheet = client.Sheets.get_sheet(sheet_id, include=['columns'])
        return {col.title: col.id for col in sheet.columns}
    except Exception as e:
        logger.error(f"Error getting column map: {e}")
        return None


def query_smartsheet_data(group=None):
    """Query Smartsheet for activity metrics."""
    if not SMARTSHEET_AVAILABLE or not token:
        return {"total_items": 0, "recent_activity_items": 0, "recent_percentage": 0}
    
    client = smartsheet.Smartsheet(token)
    
    total_items = 0
    recent_activity_items = 0
    thirty_days_ago = datetime.now() - timedelta(days=30)
    
    sheet_ids = {group: SHEET_IDS[group]} if group and group in SHEET_IDS else SHEET_IDS
    
    for sheet_group, sheet_id in sheet_ids.items():
        if sheet_group == "SPECIAL":
            continue
        
        try:
            sheet = client.Sheets.get_sheet(sheet_id)
            
            phase_cols = {}
            for col in sheet.columns:
                if col.title in ["Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"]:
                    phase_cols[col.title] = col.id
            
            for row in sheet.rows:
                total_items += 1
                most_recent = None
                
                for col_title, col_id in phase_cols.items():
                    for cell in row.cells:
                        if cell.column_id == col_id and cell.value:
                            try:
                                date_val = parse_date(cell.value)
                                if date_val and (most_recent is None or date_val > most_recent):
                                    most_recent = date_val
                            except:
                                pass
                
                if most_recent and most_recent >= thirty_days_ago.date():
                    recent_activity_items += 1
        except Exception as e:
            logger.error(f"Error processing sheet {sheet_group}: {e}")
    
    return {
        "total_items": total_items,
        "recent_activity_items": recent_activity_items,
        "recent_percentage": (recent_activity_items / total_items * 100) if total_items > 0 else 0
    }


def get_special_activities(start_date, end_date):
    """Fetch special activities from designated sheet."""
    sheet_id = SHEET_IDS.get("SPECIAL")
    if not sheet_id:
        return {}, 0, 0
    
    if not SMARTSHEET_AVAILABLE or not token:
        return {}, 0, 0
    
    try:
        client = smartsheet.Smartsheet(token)
        sheet = client.Sheets.get_sheet(sheet_id)
        
        col_map = {col.title: col.id for col in sheet.columns}
        user_col_id = col_map.get("Mitarbeiter")
        date_col_id = col_map.get("Datum")
        category_col_id = col_map.get("Kategorie")
        duration_col_id = col_map.get("Arbeitszeit in Std")
        
        if not all([user_col_id, date_col_id, category_col_id, duration_col_id]):
            return {}, 0, 0
        
        user_activity = {}
        total_activities = 0
        total_hours = 0
        
        for row in sheet.rows:
            date_cell = row.get_column(date_col_id)
            if date_cell and date_cell.value:
                try:
                    # Handle multiple date formats from Smartsheet
                    date_str = str(date_cell.value)
                    if 'T' in date_str:
                        # ISO format: 2025-02-05T00:00:00Z or 2025-02-05T00:00:00
                        activity_date = datetime.fromisoformat(date_str.replace('Z', '+00:00')).date()
                    else:
                        activity_date = datetime.strptime(date_str[:10], '%Y-%m-%d').date()
                    if start_date <= activity_date <= end_date:
                        user = row.get_column(user_col_id)
                        user = user.value if user else "Unassigned"
                        
                        category = row.get_column(category_col_id)
                        category = category.value if category else "Uncategorized"
                        
                        duration_cell = row.get_column(duration_col_id)
                        duration = 0
                        if duration_cell and duration_cell.value:
                            try:
                                duration = float(str(duration_cell.value).replace(',', '.'))
                            except:
                                pass
                        
                        if user not in user_activity:
                            user_activity[user] = {"count": 0, "hours": 0, "categories": {}}
                        
                        user_activity[user]["count"] += 1
                        user_activity[user]["hours"] += duration
                        user_activity[user]["categories"][category] = \
                            user_activity[user]["categories"].get(category, 0) + duration
                        
                        total_activities += 1
                        total_hours += duration
                except:
                    continue
        
        return user_activity, total_activities, total_hours
    except Exception as e:
        logger.error(f"Error fetching special activities: {e}")
        return {}, 0, 0


def get_marketplace_activity(group_name, sheet_id, start_date, end_date):
    """Get marketplace activity metrics."""
    if not SMARTSHEET_AVAILABLE or not token:
        return [], []
    
    try:
        client = smartsheet.Smartsheet(token)
        sheet = client.Sheets.get_sheet(sheet_id)
        
        col_map = {col.title: col.id for col in sheet.columns}
        marketplace_col_id = col_map.get("Amazon")
        date_cols = {t: i for t, i in col_map.items() if " am" in t or "Kontrolle" in t}
        
        if not marketplace_col_id or not date_cols:
            return [], []
        
        product_last_activity = {}
        for row in sheet.rows:
            last_date = None
            for cell in row.cells:
                if cell.column_id in date_cols.values():
                    try:
                        cell_date = parse_date(cell.value)
                        if cell_date and (last_date is None or cell_date > last_date):
                            last_date = cell_date
                    except:
                        continue
            if last_date:
                product_last_activity[row.id] = last_date
        
        marketplace_data = defaultdict(lambda: {"count": 0, "days": []})
        today = datetime.now().date()
        
        for row in sheet.rows:
            if row.id in product_last_activity:
                mp_cell = row.get_column(marketplace_col_id)
                if mp_cell and mp_cell.value:
                    mp = mp_cell.value.strip().upper()
                    marketplace_data[mp]["count"] += 1
                    marketplace_data[mp]["days"].append((today - product_last_activity[row.id]).days)
        
        # Calculate averages and format
        combined = []
        for mp, data in marketplace_data.items():
            avg_days = sum(data["days"]) / len(data["days"]) if data["days"] else 0
            combined.append((mp, avg_days, data["count"]))
        
        most_active = sorted(combined, key=lambda x: x[1])[:5]
        most_inactive = sorted(combined, key=lambda x: x[1], reverse=True)[:5]
        
        return most_active, most_inactive
    except Exception as e:
        logger.error(f"Error getting marketplace activity: {e}")
        return [], []


# =============================================================================
# TABLE CREATION
# =============================================================================

def create_summary_table(metrics, styles):
    """Create the summary statistics table."""
    data = [
        ["Metric", "Value"],
        ["Total Changes", str(metrics["total_changes"])],
        ["Active Groups", str(len(metrics["groups"]))],
        ["Active Users", str(len(metrics["users"]))],
        ["Active Phases", str(len(metrics["phases"]))],
    ]
    
    table = Table(data, colWidths=[80*mm, 40*mm])
    table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.WHITE),
        ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), DesignSystem.FONT_SIZE_SM),
        ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
        ('PADDING', (0, 0), (-1, 0), 8),
        
        # Data rows
        ('FONTNAME', (0, 1), (-1, -1), DesignSystem.FONT_FAMILY),
        ('FONTSIZE', (0, 1), (-1, -1), DesignSystem.FONT_SIZE_SM),
        ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_700),
        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
        ('FONTNAME', (1, 1), (1, -1), DesignSystem.FONT_BOLD),
        
        # Alternating rows
        ('BACKGROUND', (0, 1), (-1, 1), DesignSystem.WHITE),
        ('BACKGROUND', (0, 2), (-1, 2), DesignSystem.GRAY_50),
        ('BACKGROUND', (0, 3), (-1, 3), DesignSystem.WHITE),
        ('BACKGROUND', (0, 4), (-1, 4), DesignSystem.GRAY_50),
        
        # Grid
        ('LINEBELOW', (0, 0), (-1, 0), 1, DesignSystem.PRIMARY),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, DesignSystem.GRAY_200),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
    ]))
    
    return table


def create_activity_table(activity_data, title):
    """Create a marketplace activity table."""
    if not activity_data:
        data = [[f"No {title.lower()} data"]]
        table = Table(data, colWidths=[70*mm])
        table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), DesignSystem.FONT_FAMILY),
            ('FONTSIZE', (0, 0), (-1, -1), DesignSystem.FONT_SIZE_SM),
            ('TEXTCOLOR', (0, 0), (-1, -1), DesignSystem.GRAY_400),
            ('BACKGROUND', (0, 0), (-1, -1), DesignSystem.GRAY_50),
            ('PADDING', (0, 0), (-1, -1), 10),
        ]))
        return table
    
    data = [["Market", "Avg Days", "Count"]]
    for market, avg_days, count in activity_data:
        data.append([market, f"{avg_days:.1f}", str(count)])
    
    table = Table(data, colWidths=[30*mm, 22*mm, 18*mm])
    table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.GRAY_100),
        ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.GRAY_700),
        ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), DesignSystem.FONT_SIZE_XS),
        ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
        
        # Data
        ('FONTNAME', (0, 1), (-1, -1), DesignSystem.FONT_FAMILY),
        ('FONTSIZE', (0, 1), (-1, -1), DesignSystem.FONT_SIZE_XS),
        ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_600),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
        
        # Grid
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, DesignSystem.GRAY_200),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    
    return table


# =============================================================================
# PAGE TEMPLATES
# =============================================================================

def add_header_footer(canvas, doc, report_title, period_str):
    """Add header and footer to each page."""
    canvas.saveState()
    
    page_width, page_height = A4
    
    # Header line
    canvas.setStrokeColor(DesignSystem.PRIMARY)
    canvas.setLineWidth(2)
    canvas.line(
        DesignSystem.MARGIN_LEFT, page_height - 15*mm,
        page_width - DesignSystem.MARGIN_RIGHT, page_height - 15*mm
    )
    
    # Header text
    canvas.setFillColor(DesignSystem.PRIMARY)
    canvas.setFont(DesignSystem.FONT_BOLD, DesignSystem.FONT_SIZE_SM)
    canvas.drawString(DesignSystem.MARGIN_LEFT, page_height - 12*mm, "Amazon Content Management")
    
    canvas.setFillColor(DesignSystem.GRAY_500)
    canvas.setFont(DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_XS)
    canvas.drawRightString(
        page_width - DesignSystem.MARGIN_RIGHT, 
        page_height - 12*mm, 
        period_str
    )
    
    # Footer line
    canvas.setStrokeColor(DesignSystem.GRAY_200)
    canvas.setLineWidth(0.5)
    canvas.line(
        DesignSystem.MARGIN_LEFT, 12*mm,
        page_width - DesignSystem.MARGIN_RIGHT, 12*mm
    )
    
    # Footer text
    canvas.setFillColor(DesignSystem.GRAY_400)
    canvas.setFont(DesignSystem.FONT_FAMILY, DesignSystem.FONT_SIZE_XS)
    canvas.drawString(DesignSystem.MARGIN_LEFT, 8*mm, report_title)
    canvas.drawRightString(
        page_width - DesignSystem.MARGIN_RIGHT, 
        8*mm, 
        f"Page {doc.page}"
    )
    
    canvas.restoreState()


# =============================================================================
# REPORT BUILDERS
# =============================================================================

def build_title_section(story, styles, report_type, start_date, end_date):
    """Build the title section of the report."""
    if report_type == "weekly":
        week_num = start_date.isocalendar()[1]
        title = f"Weekly Activity Report"
        subtitle = f"Week {week_num} · {start_date.strftime('%B %d')} – {end_date.strftime('%B %d, %Y')}"
    else:
        title = f"Monthly Activity Report"
        subtitle = f"{start_date.strftime('%B %Y')}"
    
    story.append(Paragraph(title, styles['ReportTitle']))
    story.append(Paragraph(subtitle, styles['ReportSubtitle']))
    
    return title, subtitle


def build_kpi_section(story, metrics, content_width):
    """Build the KPI cards section."""
    # Calculate card dimensions - compact for page 1
    card_width = (content_width - 10*mm) / 3
    
    cards = [
        KPICard(
            metrics["total_changes"], 
            "Total Changes",
            width=card_width,
            height=38*mm,  # Reduced from 45mm
            color=DesignSystem.PRIMARY
        ),
        KPICard(
            len(metrics["groups"]), 
            "Active Groups",
            width=card_width,
            height=38*mm,
            color=DesignSystem.SECONDARY
        ),
        KPICard(
            len(metrics["users"]), 
            "Active Users",
            width=card_width,
            height=38*mm,
            color=DesignSystem.ACCENT
        ),
    ]
    
    # Create table with cards
    card_table = Table([cards], colWidths=[card_width] * 3)
    card_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2*mm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2*mm),
    ]))
    
    story.append(card_table)
    story.append(Spacer(1, DesignSystem.SPACE_MD))  # Reduced from SPACE_LG


def build_overview_charts(story, styles, metrics, content_width):
    """Build the overview charts section with user summary - all on first page."""
    story.append(Paragraph("Activity Overview", styles['SectionHeader']))
    
    # Create smaller charts to fit on first page
    chart_width = (content_width - 10*mm) / 2
    chart_height = 130  # Reduced from 160
    
    group_chart = create_bar_chart(
        dict(metrics["groups"]),
        "Changes by Group",
        width=chart_width,
        height=chart_height,
        color_map=DesignSystem.GROUP_COLORS
    )
    
    phase_chart = create_bar_chart(
        dict(metrics["phases"]),
        "Changes by Phase",
        width=chart_width,
        height=chart_height,
        color_map=DesignSystem.PHASE_COLORS
    )
    
    chart_table = Table([[group_chart, phase_chart]], colWidths=[chart_width, chart_width])
    chart_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    
    story.append(chart_table)
    story.append(Spacer(1, DesignSystem.SPACE_MD))
    
    # === USER ACTIVITY SUMMARY SECTION (on first page) ===
    story.append(Paragraph("User Activity Summary", styles['SectionHeader']))
    
    # Get users sorted by activity
    active_users = {u: c for u, c in metrics["users"].items() if u and u.strip()}
    if not active_users:
        story.append(Paragraph("No user activity recorded in this period.", styles['ReportBody']))
        return
    
    sorted_users = sorted(active_users.items(), key=lambda x: x[1], reverse=True)
    total = sum(active_users.values())
    
    # Create compact user table
    table_data = [["User", "Changes", "Share"]]
    for user, count in sorted_users:
        share = f"{count/total*100:.1f}%" if total > 0 else "0%"
        table_data.append([user, str(count), share])
    
    # Compact table with smaller column widths
    user_table = Table(table_data, colWidths=[35*mm, 25*mm, 25*mm])
    user_table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.WHITE),
        ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), DesignSystem.FONT_SIZE_XS),
        ('ALIGN', (1, 0), (-1, 0), 'CENTER'),
        
        # Data
        ('FONTNAME', (0, 1), (-1, -1), DesignSystem.FONT_FAMILY),
        ('FONTSIZE', (0, 1), (-1, -1), DesignSystem.FONT_SIZE_XS),
        ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_700),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        
        # Alternating rows
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [DesignSystem.WHITE, DesignSystem.GRAY_50]),
        
        # Grid
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, DesignSystem.GRAY_200),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    
    # Create compact donut chart
    user_donut = None
    if len(active_users) > 1:
        user_donut = create_donut_chart(
            active_users,
            "Distribution",
            width=160,
            height=140,
            color_map={u: DesignSystem.get_user_color(u) for u in active_users}
        )
    
    # Layout table and chart side by side
    if user_donut:
        summary_layout = Table(
            [[user_table, user_donut]],
            colWidths=[90*mm, content_width - 95*mm]
        )
        summary_layout.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ]))
        story.append(summary_layout)
    else:
        story.append(user_table)


def build_group_detail_page(story, styles, group, group_data, metrics, content_width, start_date, end_date):
    """Build a detailed page for a specific group."""
    story.append(PageBreak())
    
    # Group header
    total_changes = metrics["groups"].get(group, 0)
    story.append(GroupHeader(group, total_changes, width=content_width))
    story.append(Spacer(1, DesignSystem.SPACE_MD))
    
    # Phase-user breakdown chart
    phase_user_data = group_data
    if phase_user_data:
        story.append(Paragraph("User Activity by Phase", styles['SubsectionHeader']))
        
        chart, legend_data = create_horizontal_stacked_bar(
            phase_user_data,
            "",
            width=content_width,
            height=140
        )
        story.append(chart)
        
        if legend_data:
            legend = create_legend_row(legend_data, width=content_width, height=18)
            story.append(legend)
        
        story.append(Spacer(1, DesignSystem.SPACE_MD))
    
    # Try to get sheet summary data for status breakdown
    try:
        sheet_id = SHEET_IDS.get(group)
        if sheet_id:
            summary_data = get_sheet_summary_data(sheet_id)
            if summary_data:
                story.append(Paragraph("Product Status", styles['SubsectionHeader']))
                
                # Extract status values
                status_values = {}
                for cat in ["Aktuell", "<30", "31 - 60", ">60"]:
                    try:
                        val = summary_data.get(cat, '0') or '0'
                        status_values[cat] = int(str(val).replace('.', ''))
                    except:
                        status_values[cat] = 0
                
                if sum(status_values.values()) > 0:
                    status_bar = create_status_bar(status_values, width=content_width, height=55)
                    story.append(status_bar)
                    story.append(Spacer(1, DesignSystem.SPACE_MD))
    except Exception as e:
        logger.warning(f"Could not get status data for group {group}: {e}")
    
    # Marketplace activity
    try:
        sheet_id = SHEET_IDS.get(group)
        if sheet_id:
            most_active, most_inactive = get_marketplace_activity(group, sheet_id, start_date, end_date)
            
            if most_active or most_inactive:
                story.append(Paragraph("Marketplace Activity", styles['SubsectionHeader']))
                
                active_table = create_activity_table(most_active, "Most Active")
                inactive_table = create_activity_table(most_inactive, "Least Active")
                
                # Headers and tables
                header_style = ParagraphStyle(
                    'TableHeader',
                    fontName=DesignSystem.FONT_BOLD,
                    fontSize=DesignSystem.FONT_SIZE_SM,
                    textColor=DesignSystem.GRAY_600,
                    alignment=TA_CENTER,
                    spaceAfter=DesignSystem.SPACE_XS,
                )
                
                mp_table = Table([
                    [Paragraph("Most Active", header_style), Paragraph("Least Active", header_style)],
                    [active_table, inactive_table]
                ], colWidths=[content_width/2 - 5*mm, content_width/2 - 5*mm])
                
                mp_table.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 2*mm),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 2*mm),
                ]))
                
                story.append(mp_table)
    except Exception as e:
        logger.warning(f"Could not get marketplace data for group {group}: {e}")


def build_user_summary_page(story, styles, metrics, content_width):
    """Build a summary page for user activity - now integrated into first page."""
    # User summary is now included in build_overview_charts on page 1
    # This function is kept for compatibility but doesn't add a separate page
    pass


def create_user_activity_by_group_chart(user_group_data, user_name, width=480, height=140):
    """Create a horizontal bar chart showing user's activity by group."""
    drawing = Drawing(width, height)
    
    if not user_group_data:
        drawing.add(String(
            width / 2, height / 2,
            "No group activity data",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_400
        ))
        return drawing
    
    # Title
    drawing.add(String(
        width / 2, height - 10,
        f"Product Changes by Group",
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_BASE,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_700
    ))
    
    # Sort groups by activity
    sorted_groups = sorted(user_group_data.items(), key=lambda x: x[1], reverse=True)
    max_value = max(v for _, v in sorted_groups) if sorted_groups else 1
    
    # Chart dimensions
    chart_x = 50
    chart_y = 20
    chart_width = width - 80
    bar_height = min(16, (height - 50) / len(sorted_groups) - 4) if sorted_groups else 16
    spacing = 4
    
    for i, (group, count) in enumerate(sorted_groups):
        y_pos = chart_y + (bar_height + spacing) * (len(sorted_groups) - 1 - i)
        
        # Group label
        drawing.add(String(
            chart_x - 8,
            y_pos + bar_height / 2 - 3,
            group,
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='end',
            fillColor=DesignSystem.GRAY_600
        ))
        
        # Bar
        bar_width = (count / max_value) * chart_width if max_value > 0 else 0
        color = DesignSystem.GROUP_COLORS.get(group, DesignSystem.PRIMARY)
        
        if bar_width > 0:
            drawing.add(Rect(
                chart_x, y_pos,
                bar_width, bar_height,
                fillColor=color,
                strokeColor=None
            ))
        
        # Value label
        drawing.add(String(
            chart_x + bar_width + 5,
            y_pos + bar_height / 2 - 3,
            str(count),
            fontName=DesignSystem.FONT_BOLD,
            fontSize=DesignSystem.FONT_SIZE_XS,
            textAnchor='start',
            fillColor=DesignSystem.GRAY_600
        ))
    
    return drawing


def create_user_phase_breakdown_chart(user_phase_data, width=220, height=160):
    """Create a donut chart showing user's activity by phase."""
    if not user_phase_data or sum(user_phase_data.values()) == 0:
        drawing = Drawing(width, height)
        drawing.add(String(
            width / 2, height / 2,
            "No phase data",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_400
        ))
        return drawing
    
    return create_donut_chart(
        user_phase_data,
        "Changes by Phase",
        width=width,
        height=height,
        color_map=DesignSystem.PHASE_COLORS
    )


def create_special_activities_mini_chart(category_hours, total_hours, width=250, height=140):
    """Create a compact horizontal bar chart for special activities categories."""
    drawing = Drawing(width, height)
    
    if not category_hours or total_hours == 0:
        drawing.add(String(
            width / 2, height / 2,
            "No special activities",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=DesignSystem.FONT_SIZE_SM,
            textAnchor='middle',
            fillColor=DesignSystem.GRAY_400
        ))
        return drawing
    
    # Title
    drawing.add(String(
        width / 2, height - 8,
        "Special Activities (Hours)",
        fontName=DesignSystem.FONT_BOLD,
        fontSize=DesignSystem.FONT_SIZE_SM,
        textAnchor='middle',
        fillColor=DesignSystem.GRAY_700
    ))
    
    # Sort and limit categories
    sorted_cats = sorted(category_hours.items(), key=lambda x: x[1], reverse=True)[:6]
    max_hours = max(h for _, h in sorted_cats) if sorted_cats else 1
    
    # Chart dimensions - adjusted for smaller width
    label_width = 65  # Space for category labels
    chart_x = label_width
    chart_y = 12
    chart_width = width - label_width - 35  # Leave space for value labels
    bar_height = min(12, (height - 35) / len(sorted_cats) - 3) if sorted_cats else 12
    spacing = 3
    
    # Colors for categories
    cat_colors = [
        DesignSystem.SECONDARY,
        DesignSystem.ACCENT,
        colors.HexColor("#8B5CF6"),
        colors.HexColor("#EC4899"),
        colors.HexColor("#F59E0B"),
        colors.HexColor("#10B981"),
    ]
    
    for i, (cat, hours) in enumerate(sorted_cats):
        y_pos = chart_y + (bar_height + spacing) * (len(sorted_cats) - 1 - i)
        
        # Category label (truncated)
        label = cat if len(cat) <= 10 else cat[:8] + "…"
        drawing.add(String(
            chart_x - 5,
            y_pos + bar_height / 2 - 3,
            label,
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=7,
            textAnchor='end',
            fillColor=DesignSystem.GRAY_600
        ))
        
        # Bar
        bar_width = (hours / max_hours) * chart_width if max_hours > 0 else 0
        color = cat_colors[i % len(cat_colors)]
        
        if bar_width > 0:
            drawing.add(Rect(
                chart_x, y_pos,
                bar_width, bar_height,
                fillColor=color,
                strokeColor=None
            ))
        
        # Hours label
        drawing.add(String(
            chart_x + bar_width + 4,
            y_pos + bar_height / 2 - 3,
            f"{hours:.1f}h",
            fontName=DesignSystem.FONT_FAMILY,
            fontSize=7,
            textAnchor='start',
            fillColor=DesignSystem.GRAY_600
        ))
    
    return drawing


def collect_user_activity_data(metrics, user):
    """Collect comprehensive activity data for a specific user."""
    user_data = {
        "total_changes": metrics["users"].get(user, 0),
        "groups": defaultdict(int),
        "phases": defaultdict(int),
        "group_phase": defaultdict(lambda: defaultdict(int)),
    }
    
    # Collect from group_phase_user structure
    for group, phase_data in metrics["group_phase_user"].items():
        for phase, user_counts in phase_data.items():
            if user in user_counts:
                count = user_counts[user]
                user_data["groups"][group] += count
                user_data["phases"][phase] += count
                user_data["group_phase"][group][phase] = count
    
    return user_data


def get_user_special_activities_data(user_name, start_date, end_date):
    """Get special activities data for a specific user within date range."""
    sheet_id = SHEET_IDS.get("SPECIAL")
    if not sheet_id or not SMARTSHEET_AVAILABLE or not token:
        return {}, 0, 0
    
    try:
        client = smartsheet.Smartsheet(token)
        sheet = client.Sheets.get_sheet(sheet_id)
        
        col_map = {col.title: col.id for col in sheet.columns}
        user_col_id = col_map.get("Mitarbeiter")
        date_col_id = col_map.get("Datum")
        category_col_id = col_map.get("Kategorie")
        duration_col_id = col_map.get("Arbeitszeit in Std")
        
        if not all([user_col_id, date_col_id, category_col_id, duration_col_id]):
            return {}, 0, 0
        
        category_hours = defaultdict(float)
        total_count = 0
        total_hours = 0
        
        for row in sheet.rows:
            # Check user
            user_cell = row.get_column(user_col_id)
            if not user_cell or user_cell.value != user_name:
                continue
            
            # Check date
            date_cell = row.get_column(date_col_id)
            if date_cell and date_cell.value:
                try:
                    # Handle multiple date formats from Smartsheet
                    date_str = str(date_cell.value)
                    if 'T' in date_str:
                        activity_date = datetime.fromisoformat(date_str.replace('Z', '+00:00')).date()
                    else:
                        activity_date = datetime.strptime(date_str[:10], '%Y-%m-%d').date()
                    if not (start_date <= activity_date <= end_date):
                        continue
                except:
                    continue
            else:
                continue
            
            # Get category and duration
            category_cell = row.get_column(category_col_id)
            category = category_cell.value if category_cell and category_cell.value else "Other"
            
            duration_cell = row.get_column(duration_col_id)
            duration = 0
            if duration_cell and duration_cell.value:
                try:
                    duration = float(str(duration_cell.value).replace(',', '.'))
                except:
                    pass
            
            category_hours[category] += duration
            total_count += 1
            total_hours += duration
        
        return dict(category_hours), total_count, total_hours
    
    except Exception as e:
        logger.error(f"Error fetching special activities for {user_name}: {e}")
        return {}, 0, 0


def build_employee_detail_pages(story, styles, metrics, content_width, start_date, end_date, 
                                 special_activities_data=None):
    """Build detailed pages for each employee showing their work and special activities."""
    
    # Get users sorted by activity
    active_users = {u: c for u, c in metrics["users"].items() if u and u.strip()}
    if not active_users:
        return
    
    sorted_users = sorted(active_users.items(), key=lambda x: x[1], reverse=True)
    
    # Section header
    story.append(PageBreak())
    story.append(Paragraph("Employee Activity Details", styles['SectionHeader']))
    story.append(Paragraph(
        f"Individual breakdown of product changes and special activities for each team member.",
        styles['ReportBody']
    ))
    story.append(Spacer(1, DesignSystem.SPACE_MD))
    
    for idx, (user, total_changes) in enumerate(sorted_users):
        # Page break between users (not before first)
        if idx > 0:
            story.append(PageBreak())
        
        # User header
        story.append(UserHeader(user, total_changes, width=content_width))
        story.append(Spacer(1, DesignSystem.SPACE_MD))
        
        # Collect user's activity data
        user_activity = collect_user_activity_data(metrics, user)
        
        # Get special activities for this user
        if special_activities_data and user in special_activities_data:
            user_special = special_activities_data[user]
            special_categories = user_special.get("categories", {})
            special_count = user_special.get("count", 0)
            special_hours = user_special.get("hours", 0)
        else:
            # Try to fetch if not provided
            special_categories, special_count, special_hours = get_user_special_activities_data(
                user, start_date, end_date
            )
        
        # === ROW 1: KPI Cards ===
        card_width = (content_width - 10*mm) / 4
        
        total_share = (total_changes / sum(active_users.values()) * 100) if sum(active_users.values()) > 0 else 0
        
        cards = [
            KPICard(
                total_changes,
                "Product Changes",
                width=card_width,
                height=35*mm,
                color=DesignSystem.get_user_color(user)
            ),
            KPICard(
                f"{total_share:.1f}%",
                "Team Share",
                width=card_width,
                height=35*mm,
                color=DesignSystem.SECONDARY
            ),
            KPICard(
                special_count if special_count > 0 else "—",
                "Special Activities",
                width=card_width,
                height=35*mm,
                color=DesignSystem.ACCENT
            ),
            KPICard(
                f"{special_hours:.1f}h" if special_hours > 0 else "—",
                "Special Hours",
                width=card_width,
                height=35*mm,
                color=colors.HexColor("#8B5CF6")
            ),
        ]
        
        card_table = Table([cards], colWidths=[card_width] * 4)
        card_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 1*mm),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1*mm),
        ]))
        story.append(card_table)
        story.append(Spacer(1, DesignSystem.SPACE_MD))
        
        # === ROW 2: Product Activity Section ===
        story.append(Paragraph("Product Activity Breakdown", styles['SubsectionHeader']))
        
        # Activity by group chart
        group_chart = create_user_activity_by_group_chart(
            dict(user_activity["groups"]),
            user,
            width=content_width * 0.6,
            height=120
        )
        
        # Phase breakdown chart
        phase_chart = create_user_phase_breakdown_chart(
            dict(user_activity["phases"]),
            width=content_width * 0.38,
            height=120
        )
        
        charts_table = Table(
            [[group_chart, phase_chart]],
            colWidths=[content_width * 0.6, content_width * 0.4]
        )
        charts_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(charts_table)
        story.append(Spacer(1, DesignSystem.SPACE_MD))
        
        # === ROW 3: Detailed Tables ===
        # Group-Phase breakdown table
        if user_activity["group_phase"]:
            story.append(Paragraph("Activity by Group and Phase", styles['SubsectionHeader']))
            
            # Build table data
            phases = sorted(set(
                p for gp in user_activity["group_phase"].values() for p in gp.keys()
            ), key=lambda x: int(x) if x.isdigit() else 999)
            
            groups = sorted(user_activity["group_phase"].keys())
            
            # Header row
            header = ["Group"] + [PHASE_SHORT.get(p, f"P{p}") for p in phases] + ["Total"]
            table_data = [header]
            
            # Data rows
            for group in groups:
                row = [group]
                group_total = 0
                for phase in phases:
                    val = user_activity["group_phase"][group].get(phase, 0)
                    row.append(str(val) if val > 0 else "—")
                    group_total += val
                row.append(str(group_total))
                table_data.append(row)
            
            # Total row
            total_row = ["Total"]
            grand_total = 0
            for phase in phases:
                phase_total = sum(user_activity["group_phase"][g].get(phase, 0) for g in groups)
                total_row.append(str(phase_total))
                grand_total += phase_total
            total_row.append(str(grand_total))
            table_data.append(total_row)
            
            # Calculate column widths
            num_cols = len(header)
            first_col = 25*mm
            last_col = 20*mm
            middle_cols = (content_width - first_col - last_col) / (num_cols - 2)
            col_widths = [first_col] + [middle_cols] * (num_cols - 2) + [last_col]
            
            detail_table = Table(table_data, colWidths=col_widths)
            detail_table.setStyle(TableStyle([
                # Header
                ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.PRIMARY),
                ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.WHITE),
                ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
                ('FONTSIZE', (0, 0), (-1, 0), DesignSystem.FONT_SIZE_XS),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                
                # Total row
                ('BACKGROUND', (0, -1), (-1, -1), DesignSystem.GRAY_100),
                ('FONTNAME', (0, -1), (-1, -1), DesignSystem.FONT_BOLD),
                
                # Total column
                ('BACKGROUND', (-1, 1), (-1, -2), DesignSystem.GRAY_50),
                ('FONTNAME', (-1, 1), (-1, -1), DesignSystem.FONT_BOLD),
                
                # Data
                ('FONTNAME', (0, 1), (-2, -2), DesignSystem.FONT_FAMILY),
                ('FONTSIZE', (0, 1), (-1, -1), DesignSystem.FONT_SIZE_XS),
                ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_700),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                
                # Grid
                ('LINEBELOW', (0, 0), (-1, -1), 0.5, DesignSystem.GRAY_200),
                ('LINEAFTER', (-1, 0), (-1, -1), 1, DesignSystem.GRAY_300),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ]))
            
            story.append(detail_table)
            story.append(Spacer(1, DesignSystem.SPACE_MD))
        
        # === ROW 4: Special Activities Section ===
        if special_hours > 0 and special_categories:
            story.append(Paragraph("Special Activities Breakdown", styles['SubsectionHeader']))
            
            # Calculate proper widths to avoid overlap
            # content_width is approximately 170mm on A4
            chart_width_mm = 85*mm  # Fixed width for chart
            table_width_mm = 80*mm  # Fixed width for table
            gap = 5*mm  # Gap between chart and table
            
            # Special activities chart - reduced width
            special_chart = create_special_activities_mini_chart(
                special_categories,
                special_hours,
                width=chart_width_mm,
                height=100
            )
            
            # Special activities table - compact columns
            sorted_cats = sorted(special_categories.items(), key=lambda x: x[1], reverse=True)[:6]
            
            sa_data = [["Category", "Hours", "%"]]
            for cat, hours in sorted_cats:
                pct = (hours / special_hours * 100) if special_hours > 0 else 0
                display_cat = cat if len(cat) <= 18 else cat[:15] + "..."
                sa_data.append([display_cat, f"{hours:.1f}", f"{pct:.0f}%"])
            
            # Add total row
            sa_data.append(["Total", f"{special_hours:.1f}", "100%"])
            
            sa_table = Table(sa_data, colWidths=[45*mm, 16*mm, 14*mm])
            sa_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.GRAY_100),
                ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.GRAY_700),
                ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
                ('FONTSIZE', (0, 0), (-1, -1), DesignSystem.FONT_SIZE_XS),
                ('FONTNAME', (0, 1), (-1, -2), DesignSystem.FONT_FAMILY),
                ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_600),
                ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('BACKGROUND', (0, -1), (-1, -1), DesignSystem.GRAY_100),
                ('FONTNAME', (0, -1), (-1, -1), DesignSystem.FONT_BOLD),
                ('LINEBELOW', (0, 0), (-1, -1), 0.5, DesignSystem.GRAY_200),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ]))
            
            # Layout with explicit widths and gap
            sa_layout = Table(
                [[special_chart, sa_table]],
                colWidths=[chart_width_mm, table_width_mm]
            )
            sa_layout.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'LEFT'),
                ('LEFTPADDING', (0, 0), (0, 0), 0),
                ('RIGHTPADDING', (0, 0), (0, 0), gap),
                ('LEFTPADDING', (1, 0), (1, 0), 0),
            ]))
            
            story.append(sa_layout)
        elif special_count == 0:
            story.append(Paragraph("Special Activities", styles['SubsectionHeader']))
            story.append(Paragraph(
                "No special activities recorded for this period.",
                styles['SmallText']
            ))


def build_special_activities_page(story, styles, start_date, end_date, content_width):
    """Build the special activities page."""
    user_activity, total_activities, total_hours = get_special_activities(start_date, end_date)
    
    if not user_activity:
        return
    
    story.append(PageBreak())
    story.append(Paragraph("Special Activities", styles['SectionHeader']))
    story.append(Paragraph(
        f"Summary of special activities from {start_date.strftime('%B %d')} to {end_date.strftime('%B %d')}",
        styles['ReportBody']
    ))
    
    # KPI cards for special activities
    card_width = (content_width - 5*mm) / 2
    
    cards = [
        KPICard(
            total_activities,
            "Total Activities",
            width=card_width,
            height=40*mm,
            color=DesignSystem.SECONDARY
        ),
        KPICard(
            f"{total_hours:.1f}h",
            "Total Hours",
            width=card_width,
            height=40*mm,
            color=DesignSystem.ACCENT
        ),
    ]
    
    card_table = Table([cards], colWidths=[card_width, card_width])
    card_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    
    story.append(Spacer(1, DesignSystem.SPACE_SM))
    story.append(card_table)
    story.append(Spacer(1, DesignSystem.SPACE_LG))

    # Category breakdown
    category_hours = defaultdict(float)
    for user_data in user_activity.values():
        for cat, hours in user_data.get("categories", {}).items():
            category_hours[cat] += hours

    if category_hours:
        # Define colors for categories
        category_colors = [
            colors.HexColor("#223459"),  # Dark Navy
            colors.HexColor("#6A5AAA"),  # Purple
            colors.HexColor("#B45082"),  # Magenta
            colors.HexColor("#F9767F"),  # Coral
            colors.HexColor("#FFB142"),  # Orange
            colors.HexColor("#10B981"),  # Emerald
            colors.HexColor("#3B82F6"),  # Blue
            colors.HexColor("#8B5CF6"),  # Violet
            colors.HexColor("#EC4899"),  # Pink
            colors.HexColor("#F59E0B"),  # Amber
        ]

        # Sort categories by hours
        sorted_cats = sorted(category_hours.items(), key=lambda x: x[1], reverse=True)[:10]

        # Create color map for donut chart
        color_map = {}
        for i, (cat, _) in enumerate(sorted_cats):
            color_map[cat] = category_colors[i % len(category_colors)]

        # Create donut chart for category distribution
        story.append(Paragraph("Category Distribution", styles['SubsectionHeader']))
        story.append(Spacer(1, DesignSystem.SPACE_SM))

        # Donut chart with legend side by side
        donut_data = {cat: hours for cat, hours in sorted_cats}
        donut_chart = create_donut_chart(
            donut_data,
            f"Total: {total_hours:.1f}h",
            width=180,
            height=160,
            color_map=color_map
        )

        # Create legend
        legend_data = []
        for i, (cat, hours) in enumerate(sorted_cats):
            color_cell = Table(
                [[""]],
                colWidths=[8*mm],
                rowHeights=[4*mm]
            )
            color_cell.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, 0), category_colors[i % len(category_colors)]),
                ('LINEBELOW', (0, 0), (0, 0), 0, DesignSystem.WHITE),
            ]))
            display_cat = cat if len(cat) <= 25 else cat[:22] + "..."
            share = f"{hours/total_hours*100:.1f}%" if total_hours > 0 else "0%"
            legend_data.append([color_cell, display_cat, f"{hours:.1f}h", share])

        legend_table = Table(legend_data, colWidths=[10*mm, 55*mm, 18*mm, 15*mm])
        legend_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), DesignSystem.FONT_FAMILY),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('TEXTCOLOR', (1, 0), (-1, -1), DesignSystem.GRAY_600),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))

        # Combine chart and legend in a table
        chart_legend_table = Table(
            [[donut_chart, legend_table]],
            colWidths=[65*mm, content_width - 70*mm]
        )
        chart_legend_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ]))

        story.append(chart_legend_table)
        story.append(Spacer(1, DesignSystem.SPACE_LG))

        # Detailed table
        story.append(Paragraph("Hours by Category", styles['SubsectionHeader']))
        
        # Sort and limit categories
        sorted_cats = sorted(category_hours.items(), key=lambda x: x[1], reverse=True)[:10]
        
        data = [["Category", "Hours", "Share"]]
        for cat, hours in sorted_cats:
            share = f"{hours/total_hours*100:.1f}%" if total_hours > 0 else "0%"
            # Truncate long category names
            display_cat = cat if len(cat) <= 40 else cat[:37] + "..."
            data.append([display_cat, f"{hours:.1f}", share])
        
        table = Table(data, colWidths=[90*mm, 30*mm, 30*mm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.GRAY_100),
            ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.GRAY_700),
            ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
            ('FONTSIZE', (0, 0), (-1, -1), DesignSystem.FONT_SIZE_SM),
            ('FONTNAME', (0, 1), (-1, -1), DesignSystem.FONT_FAMILY),
            ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_600),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [DesignSystem.WHITE, DesignSystem.GRAY_50]),
            ('LINEBELOW', (0, 0), (-1, -1), 0.5, DesignSystem.GRAY_200),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ]))

        story.append(table)

    # Individual breakdown per employee
    if user_activity:
        story.append(Spacer(1, DesignSystem.SPACE_XL))
        story.append(Paragraph("Individual Special Activities", styles['SubsectionHeader']))
        story.append(Paragraph(
            "Detailed breakdown of special activities per team member",
            styles['ReportBody']
        ))
        story.append(Spacer(1, DesignSystem.SPACE_MD))

        # Sort users by total hours (descending)
        sorted_users = sorted(
            user_activity.items(),
            key=lambda x: x[1].get("hours", 0),
            reverse=True
        )

        # Create a table for each user
        for user_name, user_data in sorted_users:
            user_hours = user_data.get("hours", 0)
            user_count = user_data.get("count", 0)
            user_categories = user_data.get("categories", {})

            if not user_categories:
                continue

            # Get user color from DesignSystem
            user_color = DesignSystem.get_user_color(user_name)

            # User header with colored bar
            user_header = Table(
                [[f"  {user_name}", f"{user_hours:.1f}h total  |  {user_count} activities"]],
                colWidths=[content_width * 0.5, content_width * 0.5]
            )
            user_header.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), user_color),
                ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.WHITE),
                ('FONTNAME', (0, 0), (0, 0), DesignSystem.FONT_BOLD),
                ('FONTNAME', (1, 0), (1, 0), DesignSystem.FONT_FAMILY),
                ('FONTSIZE', (0, 0), (-1, -1), DesignSystem.FONT_SIZE_SM),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))

            story.append(user_header)

            # User's category breakdown table
            sorted_user_cats = sorted(user_categories.items(), key=lambda x: x[1], reverse=True)
            user_table_data = [["Category", "Hours", "Share"]]

            for cat, hours in sorted_user_cats:
                share = f"{hours/user_hours*100:.1f}%" if user_hours > 0 else "0%"
                display_cat = cat if len(cat) <= 40 else cat[:37] + "..."
                user_table_data.append([display_cat, f"{hours:.1f}", share])

            user_table = Table(user_table_data, colWidths=[90*mm, 30*mm, 30*mm])
            user_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), DesignSystem.GRAY_100),
                ('TEXTCOLOR', (0, 0), (-1, 0), DesignSystem.GRAY_700),
                ('FONTNAME', (0, 0), (-1, 0), DesignSystem.FONT_BOLD),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('FONTNAME', (0, 1), (-1, -1), DesignSystem.FONT_FAMILY),
                ('TEXTCOLOR', (0, 1), (-1, -1), DesignSystem.GRAY_600),
                ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [DesignSystem.WHITE, DesignSystem.GRAY_50]),
                ('LINEBELOW', (0, 0), (-1, -1), 0.5, DesignSystem.GRAY_200),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ]))

            story.append(user_table)
            story.append(Spacer(1, DesignSystem.SPACE_MD))


# =============================================================================
# MAIN REPORT FUNCTIONS
# =============================================================================

def create_weekly_report(start_date, end_date, force=False):
    """Create a weekly PDF report."""
    week_str = f"{start_date.isocalendar()[0]}-W{start_date.isocalendar()[1]:02d}"
    out_dir = os.path.join(REPORTS_DIR, "weekly")
    os.makedirs(out_dir, exist_ok=True)
    filename = os.path.join(out_dir, f"weekly_report_{week_str}.pdf")
    
    logger.info(f"Creating weekly report for {week_str}")
    
    # Load data
    changes = load_changes(start_date, end_date)
    
    if not changes and not force:
        logger.warning(f"No changes found for week {week_str}")
        return None
    
    metrics = collect_metrics(changes)
    styles = create_styles()
    
    # Calculate content width
    page_width, page_height = A4
    content_width = page_width - DesignSystem.MARGIN_LEFT - DesignSystem.MARGIN_RIGHT
    
    # Build document
    period_str = f"Week {start_date.isocalendar()[1]}, {start_date.year}"
    
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        leftMargin=DesignSystem.MARGIN_LEFT,
        rightMargin=DesignSystem.MARGIN_RIGHT,
        topMargin=DesignSystem.MARGIN_TOP,
        bottomMargin=DesignSystem.MARGIN_BOTTOM
    )
    
    story = []
    
    # Title section
    title, subtitle = build_title_section(story, styles, "weekly", start_date, end_date)
    
    # KPI section
    build_kpi_section(story, metrics, content_width)
    
    # Overview charts
    build_overview_charts(story, styles, metrics, content_width)
    
    # Group detail pages
    for group in sorted(metrics["group_phase_user"].keys()):
        if group:
            build_group_detail_page(
                story, styles, group, 
                metrics["group_phase_user"][group],
                metrics, content_width,
                start_date, end_date
            )
    
    # User summary page
    build_user_summary_page(story, styles, metrics, content_width)
    
    # Employee detail pages (individual breakdown for each user)
    special_activities_data, _, _ = get_special_activities(start_date, end_date)
    build_employee_detail_pages(
        story, styles, metrics, content_width,
        start_date, end_date, special_activities_data
    )
    
    # Special activities summary page (team overview)
    build_special_activities_page(story, styles, start_date, end_date, content_width)
    
    # Build PDF with header/footer
    def add_page_elements(canvas, doc):
        add_header_footer(canvas, doc, "Weekly Activity Report", period_str)
    
    doc.build(story, onFirstPage=add_page_elements, onLaterPages=add_page_elements)
    
    logger.info(f"Weekly report created: {filename}")
    
    # Upload to Smartsheet if configured
    if REPORT_METADATA_SHEET_ID:
        upload_pdf_to_smartsheet(filename, WEEKLY_REPORT_ATTACHMENT_ROW_ID)
        column_map = get_column_map(REPORT_METADATA_SHEET_ID)
        if column_map:
            date_range_str = f"{start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}"
            update_smartsheet_cells(
                REPORT_METADATA_SHEET_ID, 
                WEEKLY_METADATA_ROW_ID, 
                column_map, 
                os.path.basename(filename), 
                date_range_str
            )
    
    return filename


def create_monthly_report(year, month, force=False):
    """Create a monthly PDF report."""
    from calendar import monthrange
    
    start_date = date(year, month, 1)
    end_date = date(year, month, monthrange(year, month)[1])
    
    month_str = f"{year}-{month:02d}"
    out_dir = os.path.join(REPORTS_DIR, "monthly")
    os.makedirs(out_dir, exist_ok=True)
    filename = os.path.join(out_dir, f"monthly_report_{month_str}.pdf")
    
    logger.info(f"Creating monthly report for {month_str}")
    
    # Load data
    changes = load_changes(start_date, end_date)
    
    if not changes and not force:
        logger.warning(f"No changes found for {month_str}")
        return None
    
    metrics = collect_metrics(changes)
    styles = create_styles()
    
    # Calculate content width
    page_width, page_height = A4
    content_width = page_width - DesignSystem.MARGIN_LEFT - DesignSystem.MARGIN_RIGHT
    
    # Build document
    period_str = start_date.strftime("%B %Y")
    
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        leftMargin=DesignSystem.MARGIN_LEFT,
        rightMargin=DesignSystem.MARGIN_RIGHT,
        topMargin=DesignSystem.MARGIN_TOP,
        bottomMargin=DesignSystem.MARGIN_BOTTOM
    )
    
    story = []
    
    # Title section
    title, subtitle = build_title_section(story, styles, "monthly", start_date, end_date)
    
    # KPI section
    build_kpi_section(story, metrics, content_width)
    
    # Overview charts
    build_overview_charts(story, styles, metrics, content_width)
    
    # Group detail pages
    for group in sorted(metrics["group_phase_user"].keys()):
        if group:
            build_group_detail_page(
                story, styles, group,
                metrics["group_phase_user"][group],
                metrics, content_width,
                start_date, end_date
            )
    
    # User summary page
    build_user_summary_page(story, styles, metrics, content_width)
    
    # Employee detail pages (individual breakdown for each user)
    special_activities_data, _, _ = get_special_activities(start_date, end_date)
    build_employee_detail_pages(
        story, styles, metrics, content_width,
        start_date, end_date, special_activities_data
    )
    
    # Special activities summary page (team overview)
    build_special_activities_page(story, styles, start_date, end_date, content_width)
    
    # Build PDF with header/footer
    def add_page_elements(canvas, doc):
        add_header_footer(canvas, doc, "Monthly Activity Report", period_str)
    
    doc.build(story, onFirstPage=add_page_elements, onLaterPages=add_page_elements)
    
    logger.info(f"Monthly report created: {filename}")
    
    # Upload to Smartsheet if configured
    if REPORT_METADATA_SHEET_ID:
        upload_pdf_to_smartsheet(filename, MONTHLY_REPORT_ATTACHMENT_ROW_ID)
        column_map = get_column_map(REPORT_METADATA_SHEET_ID)
        if column_map:
            update_smartsheet_cells(
                REPORT_METADATA_SHEET_ID,
                MONTHLY_METADATA_ROW_ID,
                column_map,
                os.path.basename(filename),
                period_str
            )
    
    return filename


# =============================================================================
# SMARTSHEET UPLOAD FUNCTIONS
# =============================================================================

def upload_pdf_to_smartsheet(file_path, row_id):
    """Upload PDF to Smartsheet row."""
    if not REPORT_METADATA_SHEET_ID or not row_id:
        logger.warning("Smartsheet upload not configured")
        return
    
    if not SMARTSHEET_AVAILABLE or not token:
        logger.warning("Smartsheet SDK not available - skipping upload")
        return
    
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        return
    
    try:
        client = smartsheet.Smartsheet(token)
        client.Attachments.attach_file_to_row(
            REPORT_METADATA_SHEET_ID,
            row_id,
            (os.path.basename(file_path), open(file_path, 'rb'), 'application/pdf')
        )
        logger.info(f"Uploaded PDF to row {row_id}")
    except Exception as e:
        logger.error(f"Failed to upload PDF: {e}")


def update_smartsheet_cells(sheet_id, row_id, column_map, filename, date_range_str):
    """Update Smartsheet cells with report metadata."""
    if not all([sheet_id, row_id, column_map]):
        return
    
    if not SMARTSHEET_AVAILABLE or not token:
        return
    
    try:
        client = smartsheet.Smartsheet(token)
        
        primary_col = column_map.get("Primäre Spalte")
        secondary_col = column_map.get("Spalte2")
        
        if not primary_col or not secondary_col:
            logger.error("Required columns not found")
            return
        
        cells = [
            smartsheet.models.Cell({'column_id': primary_col, 'value': filename}),
            smartsheet.models.Cell({'column_id': secondary_col, 'value': date_range_str})
        ]
        
        row = smartsheet.models.Row({'id': row_id, 'cells': cells})
        client.Sheets.update_rows(sheet_id, [row])
        
        logger.info("Updated Smartsheet cells")
    except Exception as e:
        logger.error(f"Failed to update cells: {e}")


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def get_previous_week():
    """Get previous week date range."""
    today = date.today()
    end = today - timedelta(days=today.weekday() + 1)
    start = end - timedelta(days=6)
    return start, end


def get_current_week():
    """Get current week date range."""
    today = date.today()
    start = today - timedelta(days=today.weekday())
    return start, today


def get_previous_month():
    """Get previous month year and month."""
    today = date.today()
    if today.month == 1:
        return today.year - 1, 12
    return today.year, today.month - 1


def get_current_month():
    """Get current month year and month."""
    today = date.today()
    return today.year, today.month


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Generate Smartsheet change reports (v2)")
    report_type = parser.add_mutually_exclusive_group(required=True)
    report_type.add_argument("--weekly", action="store_true", help="Generate weekly report")
    report_type.add_argument("--monthly", action="store_true", help="Generate monthly report")
    
    parser.add_argument("--year", type=int, help="Year for report")
    parser.add_argument("--month", type=int, help="Month number for monthly report")
    parser.add_argument("--week", type=int, help="ISO week number for weekly report")
    parser.add_argument("--previous", action="store_true", help="Generate report for previous period")
    parser.add_argument("--current", action="store_true", help="Generate report for current period")
    parser.add_argument("--force", action="store_true", help="Force report generation")
    
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
                start_date = datetime.fromisocalendar(args.year, args.week, 1).date()
                end_date = start_date + timedelta(days=6)
                filename = create_weekly_report(start_date, end_date, force=args.force)
            else:
                logger.error("Specify --previous, --current, or --year and --week")
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
                logger.error("Specify --previous, --current, or --year and --month")
                exit(1)
        
        if filename:
            logger.info(f"Report generated: {filename}")
            exit(0)
        else:
            logger.warning("No report generated")
            exit(1)
    
    except Exception as e:
        logger.error(f"Error generating report: {e}", exc_info=True)
        exit(1)

"""
Centralized Configuration for Smartsheet Tracker

This module contains all shared configuration used across the tracking system.
All constants, IDs, and settings should be defined here to ensure consistency.

IMPORTANT: The groups (NA, NF, NH, NM, NP, NT, NV) are PRODUCT GROUPS, not regional groups.
"""

from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional
from enum import Enum
import os

# =============================================================================
# PRODUCT GROUPS
# =============================================================================

class ProductGroup(Enum):
    """Product group identifiers."""
    NA = "NA"
    NF = "NF"
    NH = "NH"
    NM = "NM"
    NP = "NP"
    NT = "NT"
    NV = "NV"


# Sheet IDs for each product group
SHEET_IDS: Dict[str, int] = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NM": 4275419734822788,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
}

# Total product counts per group (for percentage calculations)
TOTAL_PRODUCTS: Dict[str, int] = {
    "NA": 1779,
    "NF": 1716,
    "NH": 893,
    "NM": 391,
    "NP": 394,
    "NT": 119,
    "NV": 314,
}

# =============================================================================
# PHASE DEFINITIONS
# =============================================================================

class Phase(Enum):
    """Processing phase identifiers."""
    KONTROLLE = 1
    BE = 2
    K2 = 3
    C = 4
    REOPEN = 5


@dataclass(frozen=True)
class PhaseField:
    """Definition of a phase field to track."""
    date_column: str
    user_column: str
    phase_number: int
    display_name: str


# Fields to track - (date_column, user_column, phase_number)
PHASE_FIELDS: List[Tuple[str, str, int]] = [
    ("Kontrolle", "K von", 1),
    ("BE am", "BE von", 2),
    ("K am", "K2 von", 3),
    ("C am", "C von", 4),
    ("Reopen C2 am", "Reopen C2 von", 5),
]

# Phase display names
PHASE_NAMES: Dict[str, str] = {
    "1": "Phase 1 (Kontrolle)",
    "2": "Phase 2 (BE)",
    "3": "Phase 3 (K2)",
    "4": "Phase 4 (C)",
    "5": "Phase 5 (Reopen)",
}

# =============================================================================
# USERS
# =============================================================================

# Known users in the system
USERS: List[str] = ["DM", "EK", "HI", "JHU", "LK", "SM"]

# =============================================================================
# STATUS & STATS SHEET IDs
# =============================================================================

STATUS_SHEET_ID: int = 1175618067582852
WEEKLY_STATS_SHEET_ID: int = 5679217694953348
DAILY_STATS_SHEET_ID: int = 8081126615633796

# Report metadata sheet
REPORT_METADATA_SHEET_ID: int = 7888169555939204
MONTHLY_REPORT_ATTACHMENT_ROW_ID: int = 5089581251235716
WEEKLY_REPORT_ATTACHMENT_ROW_ID: int = 1192484760260484

# =============================================================================
# FILE PATHS
# =============================================================================

DATA_DIR: str = "tracking_data"
REPORTS_DIR: str = "reports"
WEEKLY_REPORTS_DIR: str = os.path.join(REPORTS_DIR, "weekly")
MONTHLY_REPORTS_DIR: str = os.path.join(REPORTS_DIR, "monthly")

STATE_FILE: str = os.path.join(DATA_DIR, "tracker_state.json")
CHANGES_FILE: str = os.path.join(DATA_DIR, "change_history.csv")

# =============================================================================
# COLORS (Hex values - to be converted by reportlab when needed)
# =============================================================================

# Product group colors
GROUP_COLORS_HEX: Dict[str, str] = {
    "NA": "#E63946",  # Red
    "NF": "#457B9D",  # Blue
    "NH": "#2A9D8F",  # Teal
    "NM": "#E9C46A",  # Gold
    "NP": "#F4A261",  # Orange
    "NT": "#9D4EDD",  # Purple
    "NV": "#00B4D8",  # Cyan
}

# Phase colors - Sequential blue gradient
PHASE_COLORS_HEX: Dict[str, str] = {
    "1": "#1E40AF",  # Dark blue
    "2": "#2563EB",  # Blue
    "3": "#3B82F6",  # Light blue
    "4": "#60A5FA",  # Lighter blue
    "5": "#93C5FD",  # Lightest blue
}

# User colors - Noctua team colors
USER_COLORS_HEX: Dict[str, str] = {
    "DM": "#223459",   # Dark navy
    "EK": "#6A5AAA",   # Purple
    "HI": "#B45082",   # Magenta
    "SM": "#F9767F",   # Coral
    "JHU": "#FFB142",  # Orange
    "LK": "#FFDE70",   # Yellow
}

# Fallback colors for additional users
FALLBACK_COLORS_HEX: List[str] = [
    "#1f77b4",  # Blue
    "#ff7f0e",  # Orange
    "#2ca02c",  # Green
    "#d62728",  # Red
    "#9467bd",  # Purple
    "#8c564b",  # Brown
    "#e377c2",  # Pink
    "#7f7f7f",  # Gray
    "#bcbd22",  # Yellow-green
    "#17becf",  # Cyan
]

# =============================================================================
# CSV COLUMN HEADERS
# =============================================================================

CHANGE_HISTORY_COLUMNS: List[str] = [
    "Timestamp",
    "Group",
    "RowID",
    "Phase",
    "DateField",
    "Date",
    "User",
    "Marketplace",
]

# =============================================================================
# DATE FORMATS
# =============================================================================

# Supported date formats for parsing
DATE_FORMATS: List[str] = [
    '%Y-%m-%dT%H:%M:%S',  # ISO with time
    '%Y-%m-%d',           # ISO date
    '%d.%m.%Y',           # German format
    '%m/%d/%Y',           # US format
    '%Y/%m/%d',           # Alternative ISO
]

TIMESTAMP_FORMAT: str = "%Y-%m-%d %H:%M:%S"
DATE_FORMAT: str = "%Y-%m-%d"

# =============================================================================
# API CONFIGURATION
# =============================================================================

# Retry configuration for API calls
API_MAX_RETRIES: int = 3
API_RETRY_DELAY: float = 2.0  # Base delay in seconds
API_RETRY_MULTIPLIER: float = 2.0  # Exponential backoff multiplier
API_RETRY_MAX_DELAY: float = 10.0  # Maximum delay between retries

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_product_groups() -> List[str]:
    """Get list of all product group codes."""
    return list(SHEET_IDS.keys())


def get_sheet_id(group: str) -> Optional[int]:
    """Get sheet ID for a product group."""
    return SHEET_IDS.get(group)


def get_group_color(group: str) -> str:
    """Get color hex code for a product group."""
    return GROUP_COLORS_HEX.get(group, "#808080")


def get_user_color(user: str) -> str:
    """Get color hex code for a user."""
    if user in USER_COLORS_HEX:
        return USER_COLORS_HEX[user]
    # Generate consistent color based on user name hash
    return FALLBACK_COLORS_HEX[hash(user) % len(FALLBACK_COLORS_HEX)]


def get_phase_color(phase: str) -> str:
    """Get color hex code for a phase."""
    return PHASE_COLORS_HEX.get(str(phase), "#808080")


def ensure_directories() -> None:
    """Create required directories if they don't exist."""
    for directory in [DATA_DIR, REPORTS_DIR, WEEKLY_REPORTS_DIR, MONTHLY_REPORTS_DIR]:
        os.makedirs(directory, exist_ok=True)

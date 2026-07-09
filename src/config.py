"""
Configuration settings for Hotel Sales Analytics application.

This module contains all configurable parameters including file paths,
logging settings, and application constants.

Author: Hotel Sales Analytics Team
Version: 1.0.0
"""

from pathlib import Path
from datetime import datetime

# ============================================================================
# PROJECT STRUCTURE
# ============================================================================

PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "data"
HISTORICAL_DATA_DIR = DATA_DIR / "historical"
OUTPUT_DIR = PROJECT_ROOT / "output"
LOGS_DIR = PROJECT_ROOT / "logs"

# ============================================================================
# FILE SETTINGS
# ============================================================================

ANALISIS_FILE = DATA_DIR / "Analisis.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "Company_Analysis.xlsx"
LOG_FILE = LOGS_DIR / "app.log"

# ============================================================================
# EXCEL SHEET NAMES
# ============================================================================

DETAIL_SHEET_NAME = "Detail Analysis"
SUMMARY_SHEET_NAME = "Summary"

# ============================================================================
# COLUMN DEFINITIONS
# ============================================================================

HISTORICAL_COLUMNS = {
    "company": "Company",
    "room_nights": "Exclude Compliment",
    "revenue": "Room Revenue"
}

ANALYSIS_COLUMNS = {
    "company": "Company",
    "status": "Status",
    "last_visit_year": "Last Visit Year",
    "room_nights": "Last Room Night",
    "revenue": "Last Revenue"
}

# ============================================================================
# ACCOUNT CLASSIFICATION STATUS
# ============================================================================

STATUS_REPEATER = "Repeater"
STATUS_SLEEPING = "Sleeping Account"
STATUS_NEWCOMER = "Newcomer"

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================

LOG_FORMAT = "%(asctime)s - %(levelname)s - %(name)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
LOG_LEVEL = "INFO"

# ============================================================================
# DATA PROCESSING SETTINGS
# ============================================================================

# Minimum year files to consider valid
MIN_HISTORICAL_YEARS = 1

# Number formatting
CURRENCY_FORMAT = "#,##0.00"
INTEGER_FORMAT = "0"

# ============================================================================
# SUMMARY STATISTICS LABELS
# ============================================================================

SUMMARY_LABELS = {
    "total_analyzed": "Total Companies Analyzed",
    "total_repeater": "Total Repeater",
    "total_sleeping": "Total Sleeping Accounts",
    "total_newcomer": "Total Newcomers",
    "sleeping_room_nights": "Total Sleeping Account Room Nights",
    "sleeping_revenue": "Total Sleeping Account Revenue"
}

# ============================================================================
# NORMALIZATION SETTINGS
# ============================================================================

# Characters to remove during normalization
CHARS_TO_REMOVE = ['\r', '\n', '\t', '\x00', '\xa0']

# ============================================================================
# PERFORMANCE SETTINGS
# ============================================================================

# DataFrame chunk size for large files
CHUNK_SIZE = 10000

# ============================================================================
# DATA VALIDATION SETTINGS
# ============================================================================

REQUIRED_COLUMNS = list(HISTORICAL_COLUMNS.values())
YEAR_FILE_PATTERN = r"^\d{4}\.xlsx$"

"""
Utility functions for Hotel Sales Analytics application.

This module provides helper functions for company name normalization,
logging setup, and general utilities.

Author: Hotel Sales Analytics Team
Version: 1.0.0
"""

import logging
import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import Optional
from config import (
    LOG_FILE, LOG_FORMAT, LOG_DATE_FORMAT, LOG_LEVEL,
    CHARS_TO_REMOVE, YEAR_FILE_PATTERN
)

# ============================================================================
# LOGGING SETUP
# ============================================================================

def setup_logging() -> logging.Logger:
    """
    Configure logging for the application.

    Creates logs directory if it doesn't exist and sets up both
    file and console handlers with appropriate formatting.

    Returns:
        logging.Logger: Configured logger instance

    Example:
        >>> logger = setup_logging()
        >>> logger.info("Application started")
    """
    # Create logs directory if it doesn't exist
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

    # Create logger
    logger = logging.getLogger(__name__)
    logger.setLevel(getattr(logging, LOG_LEVEL))

    # Check if handlers already exist to avoid duplicates
    if logger.hasHandlers():
        return logger

    # Create formatter
    formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)

    # File handler
    file_handler = logging.FileHandler(LOG_FILE)
    file_handler.setLevel(getattr(logging, LOG_LEVEL))
    file_handler.setFormatter(formatter)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(getattr(logging, LOG_LEVEL))
    console_handler.setFormatter(formatter)

    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


# ============================================================================
# COMPANY NAME NORMALIZATION
# ============================================================================

def normalize_company_name(name: str) -> str:
    """
    Normalize company name for consistent matching.

    Applies the following transformations:
    1. Convert to uppercase
    2. Trim leading/trailing spaces
    3. Replace multiple spaces with single space
    4. Remove invisible characters
    5. Remove line breaks and tabs
    6. Remove special control characters

    Args:
        name (str): Original company name

    Returns:
        str: Normalized company name

    Example:
        >>> normalize_company_name("  pt   abc  ")
        'PT ABC'
        >>> normalize_company_name("pt abc\r\n")
        'PT ABC'
    """
    if not isinstance(name, str):
        return ""

    # Convert to uppercase
    normalized = name.upper()

    # Trim leading and trailing spaces
    normalized = normalized.strip()

    # Remove invisible characters and control characters
    for char in CHARS_TO_REMOVE:
        normalized = normalized.replace(char, "")

    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)

    return normalized


def normalize_company_names_series(series: pd.Series) -> pd.Series:
    """
    Vectorized normalization of company names in a pandas Series.

    This is more efficient than applying normalize_company_name
    row by row when dealing with large datasets.

    Args:
        series (pd.Series): Series containing company names

    Returns:
        pd.Series: Series with normalized company names

    Example:
        >>> companies = pd.Series(["PT ABC", "  pt abc  "])
        >>> normalize_company_names_series(companies)
        0    PT ABC
        1    PT ABC
        dtype: object
    """
    # Convert to uppercase
    result = series.str.upper()

    # Trim spaces
    result = result.str.strip()

    # Remove invisible characters
    for char in CHARS_TO_REMOVE:
        result = result.str.replace(char, "", regex=False)

    # Replace multiple spaces with single space
    result = result.str.replace(r'\s+', ' ', regex=True)

    return result


# ============================================================================
# FILE DETECTION AND VALIDATION
# ============================================================================

def detect_year_files(directory: Path) -> dict:
    """
    Detect all year-based Excel files in a directory.

    Scans directory for files matching pattern YYYY.xlsx and returns
    them sorted by year.

    Args:
        directory (Path): Directory to scan

    Returns:
        dict: Dictionary with year as key and file path as value
              Example: {2013: Path(...), 2014: Path(...), ...}

    Example:
        >>> files = detect_year_files(Path("data/historical"))
        >>> print(files)
        {2013: PosixPath('data/historical/2013.xlsx'), ...}
    """
    year_files = {}

    if not directory.exists():
        return year_files

    for file_path in sorted(directory.glob("*.xlsx")):
        # Extract year from filename
        filename = file_path.name
        match = re.match(r'^(\d{4})\.xlsx$', filename)

        if match:
            year = int(match.group(1))
            year_files[year] = file_path

    return year_files


def get_current_and_previous_year(year_files: dict) -> tuple:
    """
    Get current year and previous year from detected year files.

    Args:
        year_files (dict): Dictionary of year files from detect_year_files()

    Returns:
        tuple: (current_year, previous_year)
               Returns (None, None) if less than 1 year file exists

    Example:
        >>> year_files = {2025: Path(...), 2026: Path(...)}
        >>> current, previous = get_current_and_previous_year(year_files)
        >>> print(current, previous)
        2026 2025
    """
    if not year_files:
        return None, None

    years = sorted(year_files.keys())
    current_year = years[-1]

    previous_year = years[-2] if len(years) > 1 else None

    return current_year, previous_year


# ============================================================================
# DATA VALIDATION
# ============================================================================

def validate_required_columns(df: pd.DataFrame, required_columns: list,
                             file_name: str = "file") -> bool:
    """
    Validate that required columns exist in DataFrame.

    Args:
        df (pd.DataFrame): DataFrame to validate
        required_columns (list): List of required column names
        file_name (str): Name of file for error messages

    Returns:
        bool: True if all columns exist, False otherwise

    Example:
        >>> df = pd.DataFrame({"Company": [...], "Room Revenue": [...]})
        >>> validate_required_columns(df, ["Company", "Room Revenue"])
        True
    """
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        logger = logging.getLogger(__name__)
        logger.error(
            f"Missing columns in {file_name}: {missing_columns}. "
            f"Found columns: {list(df.columns)}"
        )
        return False

    return True


def safe_read_excel(file_path: Path, sheet_name: int = 0) -> Optional[pd.DataFrame]:
    """
    Safely read Excel file with error handling.

    Args:
        file_path (Path): Path to Excel file
        sheet_name (int): Sheet index to read (default: 0)

    Returns:
        Optional[pd.DataFrame]: DataFrame if successful, None otherwise

    Example:
        >>> df = safe_read_excel(Path("data/2013.xlsx"))
    """
    logger = logging.getLogger(__name__)

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        logger.info(f"Successfully loaded {file_path.name}: {len(df)} rows")
        return df

    except FileNotFoundError:
        logger.error(f"File not found: {file_path}")
        return None

    except Exception as e:
        logger.error(f"Error reading {file_path}: {type(e).__name__}: {str(e)}")
        return None


# ============================================================================
# DATA TYPE CONVERSION
# ============================================================================

def safe_convert_numeric(series: pd.Series, data_type: type = float,
                        fill_value: float = 0.0) -> pd.Series:
    """
    Safely convert Series to numeric type with error handling.

    Args:
        series (pd.Series): Series to convert
        data_type (type): Target data type (float or int)
        fill_value (float): Value to use for invalid entries

    Returns:
        pd.Series: Converted Series

    Example:
        >>> revenue = pd.Series(["100.50", "200", None])
        >>> safe_convert_numeric(revenue, float)
        0    100.50
        1    200.00
        2      0.00
        dtype: float64
    """
    return pd.to_numeric(series, errors='coerce').fillna(fill_value)


# ============================================================================
# SUMMARY STATISTICS
# ============================================================================

def format_currency(value: float) -> str:
    """
    Format value as currency string.

    Args:
        value (float): Value to format

    Returns:
        str: Formatted currency string

    Example:
        >>> format_currency(1234567.89)
        '1,234,567.89'
    """
    return f"{value:,.2f}"


def format_integer(value: float) -> str:
    """
    Format value as integer string.

    Args:
        value (float): Value to format

    Returns:
        str: Formatted integer string

    Example:
        >>> format_integer(1234567.89)
        '1,234,568'
    """
    return f"{int(round(value)):,}"

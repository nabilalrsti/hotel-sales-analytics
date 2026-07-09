"""
Data loading and validation module for Hotel Sales Analytics.

Handles loading historical data files, validating file structure,
and preparing data for analysis.

Author: Hotel Sales Analytics Team
Version: 1.0.0
"""

import logging
import pandas as pd
from pathlib import Path
from typing import Optional, Dict, Tuple
from config import (
    HISTORICAL_DATA_DIR, ANALISIS_FILE, HISTORICAL_COLUMNS,
    REQUIRED_COLUMNS, MIN_HISTORICAL_YEARS
)
from utils import (
    detect_year_files, get_current_and_previous_year,
    validate_required_columns, safe_read_excel,
    normalize_company_names_series, safe_convert_numeric
)

logger = logging.getLogger(__name__)


# ============================================================================
# FILE DETECTION AND LOADING
# ============================================================================

class DataLoader:
    """
    Handles loading and validation of Excel files.

    This class manages:
    - Detection of historical year files
    - Loading and validation of data files
    - Data preparation and normalization
    - Extraction of current and previous year information
    """

    def __init__(self, historical_dir: Path = HISTORICAL_DATA_DIR,
                 analisis_file: Path = ANALISIS_FILE):
        """
        Initialize DataLoader.

        Args:
            historical_dir (Path): Path to historical data directory
            analisis_file (Path): Path to Analisis.xlsx file
        """
        self.historical_dir = historical_dir
        self.analisis_file = analisis_file
        self.year_files: Dict[int, Path] = {}
        self.current_year: Optional[int] = None
        self.previous_year: Optional[int] = None
        self.logger = logging.getLogger(self.__class__.__name__)

    def validate_input_files(self) -> bool:
        """
        Validate that all required input files and directories exist.

        Returns:
            bool: True if all validations pass, False otherwise
        """
        # Check historical directory exists
        if not self.historical_dir.exists():
            self.logger.error(
                f"Historical data directory not found: {self.historical_dir}"
            )
            return False

        # Check Analisis file exists
        if not self.analisis_file.exists():
            self.logger.error(
                f"Analisis file not found: {self.analisis_file}"
            )
            return False

        # Detect year files
        self.year_files = detect_year_files(self.historical_dir)

        if not self.year_files:
            self.logger.error(
                f"No year files (YYYY.xlsx) found in: {self.historical_dir}"
            )
            return False

        # Check minimum number of year files
        if len(self.year_files) < MIN_HISTORICAL_YEARS:
            self.logger.error(
                f"Insufficient year files. Expected at least {MIN_HISTORICAL_YEARS}, "
                f"found {len(self.year_files)}"
            )
            return False

        # Get current and previous year
        self.current_year, self.previous_year = get_current_and_previous_year(
            self.year_files
        )

        self.logger.info(f"Detected {len(self.year_files)} year files")
        self.logger.info(f"Current year: {self.current_year}")
        self.logger.info(f"Previous year: {self.previous_year}")

        return True

    def load_historical_files(self) -> Optional[pd.DataFrame]:
        """
        Load and combine all historical year files.

        Returns:
            Optional[pd.DataFrame]: Combined historical data or None if error

        The returned DataFrame has columns:
        - Company: Original company name
        - Normalized_Company: Normalized for matching
        - Year: Year from filename
        - Room_Nights: From 'Exclude Compliment' column
        - Revenue: From 'Room Revenue' column
        """
        combined_data = []

        for year in sorted(self.year_files.keys()):
            file_path = self.year_files[year]
            self.logger.info(f"Loading {year} data from {file_path.name}")

            df = safe_read_excel(file_path)
            if df is None:
                return None

            # Validate columns
            if not validate_required_columns(df, REQUIRED_COLUMNS, file_path.name):
                return None

            # Process data
            df_processed = self._process_yearly_data(df, year)
            combined_data.append(df_processed)

        # Combine all data
        if not combined_data:
            self.logger.error("No data loaded from historical files")
            return None

        historical_df = pd.concat(combined_data, ignore_index=True)
        self.logger.info(
            f"Combined {len(self.year_files)} files into {len(historical_df)} rows"
        )

        return historical_df

    def _process_yearly_data(self, df: pd.DataFrame, year: int) -> pd.DataFrame:
        """
        Process and standardize yearly data.

        Args:
            df (pd.DataFrame): Raw data from Excel file
            year (int): Year of the data

        Returns:
            pd.DataFrame: Processed data with standardized columns
        """
        # Create copy to avoid modifying original
        df = df.copy()

        # Rename columns to standardized names
        df = df.rename(columns=HISTORICAL_COLUMNS)

        # Normalize company names
        df['Normalized_Company'] = normalize_company_names_series(df['Company'])

        # Convert numeric columns
        df['Room_Nights'] = safe_convert_numeric(df['Exclude Compliment'], int, 0)
        df['Revenue'] = safe_convert_numeric(df['Room Revenue'], float, 0.0)

        # Add year
        df['Year'] = year

        # Select only needed columns
        df = df[['Company', 'Normalized_Company', 'Year', 'Room_Nights', 'Revenue']]

        # Remove rows with empty company names
        df = df[df['Normalized_Company'].str.len() > 0]

        return df

    def load_analisis_file(self) -> Optional[pd.DataFrame]:
        """
        Load companies to analyze from Analisis.xlsx.

        Returns:
            Optional[pd.DataFrame]: DataFrame with companies to analyze

        The returned DataFrame has columns:
        - Company: Original company name
        - Normalized_Company: Normalized for matching
        - Room_Nights: Room nights to analyze
        - Revenue: Revenue to analyze
        """
        self.logger.info(f"Loading Analisis file: {self.analisis_file.name}")

        df = safe_read_excel(self.analisis_file)
        if df is None:
            return None

        # Validate columns
        if not validate_required_columns(df, REQUIRED_COLUMNS, "Analisis.xlsx"):
            return None

        # Rename columns
        df = df.rename(columns=HISTORICAL_COLUMNS)

        # Normalize company names
        df['Normalized_Company'] = normalize_company_names_series(df['Company'])

        # Convert numeric columns
        df['Room_Nights'] = safe_convert_numeric(df['Exclude Compliment'], int, 0)
        df['Revenue'] = safe_convert_numeric(df['Room Revenue'], float, 0.0)

        # Select only needed columns
        df = df[['Company', 'Normalized_Company', 'Room_Nights', 'Revenue']]

        # Remove rows with empty company names
        df = df[df['Normalized_Company'].str.len() > 0]

        self.logger.info(f"Loaded {len(df)} companies from Analisis.xlsx")

        return df

    def get_year_range(self) -> Tuple[Optional[int], Optional[int]]:
        """
        Get the current and previous year.

        Returns:
            Tuple[int, int]: (current_year, previous_year)
        """
        return self.current_year, self.previous_year

    def get_all_years(self) -> list:
        """
        Get sorted list of all available years.

        Returns:
            list: Sorted list of years
        """
        return sorted(self.year_files.keys())

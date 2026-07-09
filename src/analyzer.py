"""
Account classification and analysis module for Hotel Sales Analytics.

Classifies companies into Repeater, Sleeping Account, or Newcomer categories
based on historical data.

Author: Hotel Sales Analytics Team
Version: 1.0.0
"""

import logging
import pandas as pd
import numpy as np
from typing import Optional, Dict, Tuple
from config import (
    STATUS_REPEATER, STATUS_SLEEPING, STATUS_NEWCOMER,
    ANALYSIS_COLUMNS
)

logger = logging.getLogger(__name__)


# ============================================================================
# ACCOUNT CLASSIFICATION
# ============================================================================

class AccountAnalyzer:
    """
    Analyzes and classifies company accounts based on historical data.

    Classification categories:
    1. REPEATER: Appears in current or previous year
    2. SLEEPING ACCOUNT: Appeared in older years but not recent
    3. NEWCOMER: Never appeared in historical data
    """

    def __init__(self, historical_df: pd.DataFrame, current_year: int,
                 previous_year: Optional[int] = None):
        """
        Initialize AccountAnalyzer.

        Args:
            historical_df (pd.DataFrame): Combined historical data
            current_year (int): Current year
            previous_year (Optional[int]): Previous year (if available)
        """
        self.historical_df = historical_df.copy()
        self.current_year = current_year
        self.previous_year = previous_year
        self.logger = logging.getLogger(self.__class__.__name__)

        # Create lookup dictionaries for performance
        self._build_lookups()

    def _build_lookups(self):
        """Build lookup dictionaries for fast company searches."""
        # Companies in current year
        self.current_year_companies = set(
            self.historical_df[self.historical_df['Year'] == self.current_year]
            ['Normalized_Company'].unique()
        )

        # Companies in previous year (if available)
        if self.previous_year:
            self.previous_year_companies = set(
                self.historical_df[self.historical_df['Year'] == self.previous_year]
                ['Normalized_Company'].unique()
            )
        else:
            self.previous_year_companies = set()

        # All companies in history
        self.all_companies = set(
            self.historical_df['Normalized_Company'].unique()
        )

        self.logger.info(f"Current year companies: {len(self.current_year_companies)}")
        self.logger.info(f"Previous year companies: {len(self.previous_year_companies)}")
        self.logger.info(f"Total unique companies in history: {len(self.all_companies)}")

    def classify_company(self, normalized_company: str) -> Dict:
        """
        Classify a single company into one of three categories.

        Args:
            normalized_company (str): Normalized company name

        Returns:
            dict: Classification result with keys:
                - status: Classification status
                - last_visit_year: Last year company appeared
                - room_nights: Room nights in last visit
                - revenue: Revenue in last visit
        """
        # Check if Repeater (current or previous year)
        if normalized_company in self.current_year_companies:
            return self._classify_repeater(normalized_company, self.current_year)

        if normalized_company in self.previous_year_companies:
            return self._classify_repeater(normalized_company, self.previous_year)

        # Check if Sleeping Account (in history but not recent)
        if normalized_company in self.all_companies:
            return self._classify_sleeping_account(normalized_company)

        # Newcomer (never in history)
        return self._classify_newcomer()

    def _classify_repeater(self, normalized_company: str, year: int) -> Dict:
        """
        Classify company as Repeater.

        Args:
            normalized_company (str): Normalized company name
            year (int): Year of last visit

        Returns:
            dict: Classification result
        """
        # Get data for this company in this year
        company_data = self.historical_df[
            (self.historical_df['Normalized_Company'] == normalized_company) &
            (self.historical_df['Year'] == year)
        ]

        # Sum up multiple entries if they exist
        total_room_nights = company_data['Room_Nights'].sum()
        total_revenue = company_data['Revenue'].sum()

        return {
            'status': STATUS_REPEATER,
            'last_visit_year': year,
            'room_nights': total_room_nights,
            'revenue': total_revenue
        }

    def _classify_sleeping_account(self, normalized_company: str) -> Dict:
        """
        Classify company as Sleeping Account.

        Args:
            normalized_company (str): Normalized company name

        Returns:
            dict: Classification result
        """
        # Get all records for this company
        company_data = self.historical_df[
            self.historical_df['Normalized_Company'] == normalized_company
        ]

        # Find the most recent year
        last_year = company_data['Year'].max()

        # Get data from last visit
        last_visit_data = company_data[company_data['Year'] == last_year]

        total_room_nights = last_visit_data['Room_Nights'].sum()
        total_revenue = last_visit_data['Revenue'].sum()

        return {
            'status': STATUS_SLEEPING,
            'last_visit_year': last_year,
            'room_nights': total_room_nights,
            'revenue': total_revenue
        }

    def _classify_newcomer(self) -> Dict:
        """
        Classify company as Newcomer.

        Returns:
            dict: Classification result
        """
        return {
            'status': STATUS_NEWCOMER,
            'last_visit_year': None,
            'room_nights': 0,
            'revenue': 0.0
        }

    def analyze_companies(self, analisis_df: pd.DataFrame) -> pd.DataFrame:
        """
        Classify all companies from Analisis file.

        Args:
            analisis_df (pd.DataFrame): DataFrame with companies to analyze

        Returns:
            pd.DataFrame: DataFrame with classification results
        """
        self.logger.info(f"Analyzing {len(analisis_df)} companies...")

        results = []

        for idx, row in analisis_df.iterrows():
            normalized_company = row['Normalized_Company']
            classification = self.classify_company(normalized_company)

            results.append({
                'Company': row['Company'],
                'Status': classification['status'],
                'Last Visit Year': classification['last_visit_year'],
                'Last Room Night': classification['room_nights'],
                'Last Revenue': classification['revenue']
            })

            # Log progress every 100 companies
            if (idx + 1) % 100 == 0:
                self.logger.debug(f"Classified {idx + 1} companies...")

        analysis_df = pd.DataFrame(results)

        # Log classification summary
        status_counts = analysis_df['Status'].value_counts()
        self.logger.info("Classification Summary:")
        for status, count in status_counts.items():
            self.logger.info(f"  {status}: {count}")

        return analysis_df


# ============================================================================
# SUMMARY STATISTICS
# ============================================================================

class SummaryGenerator:
    """
    Generates summary statistics from analysis results.
    """

    def __init__(self, analysis_df: pd.DataFrame):
        """
        Initialize SummaryGenerator.

        Args:
            analysis_df (pd.DataFrame): Analysis results DataFrame
        """
        self.analysis_df = analysis_df
        self.logger = logging.getLogger(self.__class__.__name__)

    def generate_summary(self) -> Dict:
        """
        Generate comprehensive summary statistics.

        Returns:
            dict: Summary statistics
        """
        summary = {
            'total_analyzed': len(self.analysis_df),
            'total_repeater': len(self.analysis_df[
                self.analysis_df['Status'] == STATUS_REPEATER
            ]),
            'total_sleeping': len(self.analysis_df[
                self.analysis_df['Status'] == STATUS_SLEEPING
            ]),
            'total_newcomer': len(self.analysis_df[
                self.analysis_df['Status'] == STATUS_NEWCOMER
            ]),
        }

        # Calculate revenue opportunity from sleeping accounts
        sleeping_df = self.analysis_df[
            self.analysis_df['Status'] == STATUS_SLEEPING
        ]

        summary['sleeping_room_nights'] = sleeping_df['Last Room Night'].sum()
        summary['sleeping_revenue'] = sleeping_df['Last Revenue'].sum()

        self.logger.info("Summary Statistics:")
        self.logger.info(f"  Total Analyzed: {summary['total_analyzed']}")
        self.logger.info(f"  Repeater: {summary['total_repeater']}")
        self.logger.info(f"  Sleeping Accounts: {summary['total_sleeping']}")
        self.logger.info(f"  Newcomers: {summary['total_newcomer']}")
        self.logger.info(
            f"  Sleeping Account Room Nights: {summary['sleeping_room_nights']}"
        )
        self.logger.info(
            f"  Sleeping Account Revenue: ${summary['sleeping_revenue']:,.2f}"
        )

        return summary

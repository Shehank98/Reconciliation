"""
TC PDF to Excel Converter
Converts Telecast Certificate PDFs to structured DataFrame
Adapted from func/main.py PDFToExcelConverter
"""

import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os


class TCPDFConverter:
    """Converts TC PDF files to structured data"""

    def __init__(self):
        """Initialize converter"""
        self.extracted_data = []

    def convert_pdf_to_dataframe(self, pdf_path):
        """
        Convert TC PDF to pandas DataFrame

        Args:
            pdf_path: Path to TC PDF file

        Returns:
            pandas DataFrame with extracted TC data
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        self.extracted_data = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text()

                    if text:
                        self._extract_from_page(text, page_num)

            # Convert to DataFrame
            if not self.extracted_data:
                raise Exception("No data extracted from PDF. Please check PDF format.")

            df = pd.DataFrame(self.extracted_data)

            # Remove duplicates
            df = df.drop_duplicates()

            # Sort by date and time
            if 'Date' in df.columns and 'Advt_time' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                df = df.sort_values(['Date', 'Advt_time'])

            # Rename columns to match system standards
            column_mapping = {
                'Advt_theme': 'Theme',
                'Advt_time': 'Time',
                'Duration': 'Duration'
            }

            df = df.rename(columns=column_mapping)

            return df

        except Exception as e:
            raise Exception(f"Error converting PDF: {str(e)}")

    def _extract_from_page(self, text, page_num):
        """
        Extract data from a single PDF page

        Args:
            text: Extracted text from page
            page_num: Page number
        """
        lines = text.split('\n')

        for line in lines:
            # Main pattern: Date | Program | Time | Theme | Duration
            # Example: 01/01/2024 News 10:30:15:00 Coca Cola 30

            pattern1 = r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d{2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)\s*$'
            match = re.search(pattern1, line)

            if match:
                date_str = match.group(1)
                program = match.group(2).strip()
                time_str = match.group(3)
                theme = match.group(4).strip()
                duration = match.group(5)

                # Convert time format HH:MM:SS:FF to HH:MM:SS
                time_parts = time_str.split(':')
                if len(time_parts) == 4:
                    time_formatted = f"{time_parts[0]}:{time_parts[1]}:{time_parts[2]}"
                else:
                    time_formatted = time_str

                self.extracted_data.append({
                    'Date': date_str,
                    'Program': program,
                    'Advt_time': time_formatted,
                    'Advt_theme': theme,
                    'Duration': int(duration)
                })
                continue

            # Alternative pattern without program name
            pattern2 = r'(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)\s*$'
            match = re.search(pattern2, line)

            if match:
                date_str = match.group(1)
                time_str = match.group(2)
                theme = match.group(3).strip()
                duration = match.group(4)

                # Convert time format
                time_parts = time_str.split(':')
                if len(time_parts) == 4:
                    time_formatted = f"{time_parts[0]}:{time_parts[1]}:{time_parts[2]}"
                else:
                    time_formatted = time_str

                self.extracted_data.append({
                    'Date': date_str,
                    'Program': 'Unknown',
                    'Advt_time': time_formatted,
                    'Advt_theme': theme,
                    'Duration': int(duration)
                })
                continue

            # Pattern 3: More flexible for different formats
            # Try to find date first
            date_pattern = r'(\d{2}/\d{2}/\d{4})'
            time_pattern = r'(\d{2}:\d{2}:\d{2}(?::\d{2})?)'
            duration_pattern = r'(\d+)\s*(?:sec|seconds|s)?\s*$'

            date_match = re.search(date_pattern, line)
            time_match = re.search(time_pattern, line)
            duration_match = re.search(duration_pattern, line)

            if date_match and time_match and duration_match:
                date_str = date_match.group(1)
                time_str = time_match.group(1)
                duration = duration_match.group(1)

                # Extract theme (text between time and duration)
                time_end = time_match.end()
                duration_start = duration_match.start()
                theme = line[time_end:duration_start].strip()

                # Extract program (text between date and time)
                date_end = date_match.end()
                time_start = time_match.start()
                program = line[date_end:time_start].strip()

                # Convert time format
                time_parts = time_str.split(':')
                if len(time_parts) == 4:
                    time_formatted = f"{time_parts[0]}:{time_parts[1]}:{time_parts[2]}"
                else:
                    time_formatted = time_str

                if theme:  # Only add if theme is not empty
                    self.extracted_data.append({
                        'Date': date_str,
                        'Program': program if program else 'Unknown',
                        'Advt_time': time_formatted,
                        'Advt_theme': theme,
                        'Duration': int(duration)
                    })

    def get_extraction_summary(self):
        """
        Get summary of extracted data

        Returns:
            Dictionary with extraction statistics
        """
        if not self.extracted_data:
            return {
                'total_records': 0,
                'date_range': None,
                'themes': [],
                'programs': []
            }

        df = pd.DataFrame(self.extracted_data)

        return {
            'total_records': len(df),
            'date_range': {
                'min': df['Date'].min() if 'Date' in df.columns else None,
                'max': df['Date'].max() if 'Date' in df.columns else None
            },
            'themes': sorted(df['Advt_theme'].unique().tolist()) if 'Advt_theme' in df.columns else [],
            'programs': sorted(df['Program'].unique().tolist()) if 'Program' in df.columns else [],
            'avg_duration': df['Duration'].mean() if 'Duration' in df.columns else 0
        }

    def validate_extracted_data(self):
        """
        Validate extracted data for completeness

        Returns:
            Dictionary with validation results
        """
        if not self.extracted_data:
            return {
                'is_valid': False,
                'errors': ['No data extracted'],
                'warnings': []
            }

        df = pd.DataFrame(self.extracted_data)
        errors = []
        warnings = []

        # Check required columns
        required_columns = ['Date', 'Advt_time', 'Advt_theme', 'Duration']
        for col in required_columns:
            if col not in df.columns:
                errors.append(f"Missing required column: {col}")

        # Check for empty values
        if 'Advt_theme' in df.columns:
            empty_themes = df['Advt_theme'].isna().sum()
            if empty_themes > 0:
                warnings.append(f"{empty_themes} records have empty themes")

        # Check date format
        if 'Date' in df.columns:
            try:
                pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='raise')
            except:
                errors.append("Invalid date format. Expected DD/MM/YYYY")

        # Check duration values
        if 'Duration' in df.columns:
            if (df['Duration'] <= 0).any():
                warnings.append("Some records have invalid duration (<=0)")

        return {
            'is_valid': len(errors) == 0,
            'errors': errors,
            'warnings': warnings,
            'total_records': len(df)
        }

"""
Google Sheets Integration for LMRB Data
Fetches and filters data from Google Sheets with date range support
"""

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
import os


class GoogleSheetsService:
    """Service class for Google Sheets operations"""

    def __init__(self, credentials_path=None):
        """
        Initialize Google Sheets service

        Args:
            credentials_path: Path to service account JSON file
                            If None, looks for GOOGLE_CREDENTIALS_PATH env variable
        """
        self.credentials_path = credentials_path or os.getenv(
            'GOOGLE_CREDENTIALS_PATH',
            'credentials/service-account.json'
        )
        self.service = None
        self._initialize_service()

    def _initialize_service(self):
        """Initialize Google Sheets API service"""
        try:
            # Define the scopes
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

            # Create credentials
            if os.path.exists(self.credentials_path):
                credentials = service_account.Credentials.from_service_account_file(
                    self.credentials_path,
                    scopes=SCOPES
                )
            else:
                # For development/testing without credentials
                print(f"Warning: Credentials file not found at {self.credentials_path}")
                print("Using mock data for testing. Set GOOGLE_CREDENTIALS_PATH for production.")
                self.service = None
                return

            # Build the service
            self.service = build('sheets', 'v4', credentials=credentials)

        except Exception as e:
            print(f"Error initializing Google Sheets service: {e}")
            self.service = None

    def get_sheet_metadata(self, sheet_id):
        """
        Get metadata about the Google Sheet

        Args:
            sheet_id: Google Sheet ID from URL

        Returns:
            Dictionary with sheet metadata
        """
        if not self.service:
            return self._get_mock_metadata()

        try:
            sheet_metadata = self.service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()

            sheets = sheet_metadata.get('sheets', [])

            return {
                'title': sheet_metadata.get('properties', {}).get('title', 'Unknown'),
                'sheets': [
                    {
                        'name': sheet.get('properties', {}).get('title'),
                        'row_count': sheet.get('properties', {}).get('gridProperties', {}).get('rowCount'),
                        'column_count': sheet.get('properties', {}).get('gridProperties', {}).get('columnCount')
                    }
                    for sheet in sheets
                ],
                'url': f"https://docs.google.com/spreadsheets/d/{sheet_id}"
            }

        except HttpError as e:
            raise Exception(f"Error accessing Google Sheet: {e}")

    def get_lmrb_data(self, sheet_id, sheet_name='Sheet1', start_date=None, end_date=None):
        """
        Fetch LMRB data from Google Sheet with optional date filtering

        Args:
            sheet_id: Google Sheet ID
            sheet_name: Name of the sheet tab (default: 'Sheet1')
            start_date: Start date for filtering (YYYY-MM-DD format)
            end_date: End date for filtering (YYYY-MM-DD format)

        Returns:
            pandas DataFrame with LMRB data
        """
        if not self.service:
            return self._get_mock_lmrb_data(start_date, end_date)

        try:
            # Fetch all data from the sheet
            range_name = f'{sheet_name}!A:Z'  # Adjust range as needed

            result = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])

            if not values:
                raise Exception("No data found in sheet")

            # Convert to DataFrame
            df = pd.DataFrame(values[1:], columns=values[0])

            # Convert date column to datetime
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

            # Filter by date range if provided
            if start_date and 'Date' in df.columns:
                start_dt = pd.to_datetime(start_date)
                df = df[df['Date'] >= start_dt]

            if end_date and 'Date' in df.columns:
                end_dt = pd.to_datetime(end_date)
                df = df[df['Date'] <= end_dt]

            # Remove completely empty rows
            df = df.dropna(how='all')

            return df

        except HttpError as e:
            raise Exception(f"Error fetching data from Google Sheet: {e}")

    def _get_mock_metadata(self):
        """Mock metadata for testing without credentials"""
        return {
            'title': 'LMRB Data (Mock)',
            'sheets': [
                {
                    'name': 'Sheet1',
                    'row_count': 1000,
                    'column_count': 20
                }
            ],
            'url': 'https://docs.google.com/spreadsheets/d/mock-sheet-id'
        }

    def _get_mock_lmrb_data(self, start_date=None, end_date=None):
        """
        Generate mock LMRB data for testing without Google Sheets access

        Returns:
            pandas DataFrame with sample LMRB data
        """
        import random
        from datetime import timedelta

        # Sample data
        channels = ['Sirasa TV', 'Hiru TV', 'Derana TV', 'TV1']
        products = ['Coca Cola', 'Pepsi', 'Sprite', 'Fanta', 'Red Bull']
        themes = ['Beverages', 'Soft Drinks', 'Energy Drinks']
        programs = ['News', 'Drama', 'Reality Show', 'Movies', 'Sports']
        durations = [15, 30, 45, 60]

        # Generate date range
        if start_date:
            start_dt = pd.to_datetime(start_date)
        else:
            start_dt = pd.to_datetime('2024-01-01')

        if end_date:
            end_dt = pd.to_datetime(end_date)
        else:
            end_dt = pd.to_datetime('2024-12-31')

        # Generate random records
        records = []
        current_date = start_dt

        while current_date <= end_dt:
            num_records = random.randint(5, 15)

            for _ in range(num_records):
                hour = random.randint(6, 23)
                minute = random.randint(0, 59)
                second = random.randint(0, 59)

                records.append({
                    'Date': current_date.strftime('%Y-%m-%d'),
                    'Channel': random.choice(channels),
                    'Product': random.choice(products),
                    'Theme': random.choice(themes),
                    'Program': random.choice(programs),
                    'Time': f"{hour:02d}:{minute:02d}:{second:02d}",
                    'Duration': random.choice(durations),
                    'Advertiser': f"Advertiser {random.randint(1, 10)}"
                })

            current_date += timedelta(days=1)

        df = pd.DataFrame(records)
        df['Date'] = pd.to_datetime(df['Date'])

        return df

    def get_last_updated_date(self, sheet_id):
        """
        Get the last updated date from LMRB sheet

        Args:
            sheet_id: Google Sheet ID

        Returns:
            Date string (YYYY-MM-DD) of the most recent record
        """
        if not self.service:
            return datetime.now().strftime('%Y-%m-%d')

        try:
            # Fetch data and find max date
            df = self.get_lmrb_data(sheet_id)

            if 'Date' in df.columns:
                max_date = df['Date'].max()
                return max_date.strftime('%Y-%m-%d')

            return None

        except Exception as e:
            raise Exception(f"Error getting last updated date: {e}")

    def create_copy_sheet(self, sheet_id, start_date, end_date, new_sheet_name=None):
        """
        Create a copy of the sheet with filtered data for reconciliation

        Args:
            sheet_id: Source Google Sheet ID
            start_date: Start date for filtering
            end_date: End date for filtering
            new_sheet_name: Name for the new sheet (auto-generated if None)

        Returns:
            New sheet ID and URL
        """
        if not self.service:
            raise Exception("Google Sheets service not initialized. Cannot create copy.")

        try:
            # Generate new sheet name if not provided
            if not new_sheet_name:
                new_sheet_name = f"Recon_{start_date}_to_{end_date}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

            # Duplicate the sheet
            request_body = {
                'requests': [{
                    'duplicateSheet': {
                        'sourceSheetId': 0,  # First sheet
                        'newSheetName': new_sheet_name
                    }
                }]
            }

            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body=request_body
            ).execute()

            new_sheet_id = response['replies'][0]['duplicateSheet']['properties']['sheetId']

            return {
                'sheet_id': sheet_id,
                'new_sheet_id': new_sheet_id,
                'new_sheet_name': new_sheet_name,
                'url': f"https://docs.google.com/spreadsheets/d/{sheet_id}#gid={new_sheet_id}"
            }

        except HttpError as e:
            raise Exception(f"Error creating copy sheet: {e}")

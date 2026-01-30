import gspread
from google.oauth2.service_account import Credentials
import json
import os

class GoogleSheetsReader:
    def __init__(self, config_file="config.json"):
        """
        Initialize Google Sheets reader with credentials from config file
        """
        self.config = self._load_config(config_file)
        self.client = self._authenticate()
        self.spreadsheet = None
        self.company_name = ""
        self.app_name = ""
        self._open_spreadsheet()
    
    def _load_config(self, config_file):
        """Load configuration from JSON file"""
        if not os.path.exists(config_file):
            raise FileNotFoundError(
                f"Config file '{config_file}' not found. "
                "Please create it with your Google Sheets credentials."
            )
        
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        required_keys = ['spreadsheet_id', 'credentials_file']
        for key in required_keys:
            if key not in config:
                raise ValueError(f"Missing required config key: {key}")
        
        return config
    
    def _authenticate(self):
        """Authenticate with Google Sheets API"""
        credentials_file = self.config['credentials_file']
        
        if not os.path.exists(credentials_file):
            raise FileNotFoundError(
                f"Credentials file '{credentials_file}' not found. "
                "Please download your service account credentials from Google Cloud Console."
            )
        
        # Define the scope
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # Authenticate
        creds = Credentials.from_service_account_file(
            credentials_file,
            scopes=scope
        )
        
        return gspread.authorize(creds)
    
    def _open_spreadsheet(self):
        """Open the Google Spreadsheet"""
        spreadsheet_id = self.config['spreadsheet_id']
        try:
            self.spreadsheet = self.client.open_by_key(spreadsheet_id)
        except Exception as e:
            raise Exception(f"Failed to open spreadsheet: {str(e)}")
    
    def get_month_data(self, month_name):
        """
        Get all records from the specified month's sheet

        Args:
            month_name: Name of the month (e.g., "January", "February")

        Returns:
            List of dictionaries, each representing a record
        """
        worksheet = None

        # Try different case variations
        variations = [month_name, month_name.lower(), month_name.upper(), month_name.capitalize()]

        for variation in variations:
            try:
                worksheet = self.spreadsheet.worksheet(variation)
                break
            except gspread.exceptions.WorksheetNotFound:
                continue

        if worksheet is None:
            # List available worksheets
            available_sheets = self.get_all_sheets()
            raise Exception(
                f"Worksheet '{month_name}' not found in the spreadsheet.\n"
                f"Available sheets: {', '.join(available_sheets)}"
            )

        # Get all values from the sheet
        try:
            all_values = worksheet.get_all_values()
        except Exception as e:
            raise Exception(f"Error reading data from worksheet: {str(e)}")

        # First two rows are reserved: row 1 = company name, row 2 = app name
        if len(all_values) >= 1 and all_values[0]:
            self.company_name = all_values[0][0] if all_values[0][0] else ""
        if len(all_values) >= 2 and all_values[1]:
            self.app_name = all_values[1][0] if all_values[1][0] else ""

        # Row 3 is the header row, data starts from row 4
        if len(all_values) < 4:
            return []

        headers = all_values[2]  # Row 3 (index 2) is header
        data_rows = all_values[3:]  # Data starts from row 4 (index 3)

        # Convert to list of dictionaries
        records = []
        for row in data_rows:
            # Pad row with empty strings if it's shorter than headers
            while len(row) < len(headers):
                row.append('')

            record = {headers[i]: row[i] for i in range(len(headers)) if headers[i]}

            # Only include non-empty rows
            if any(str(v).strip() for v in record.values() if v):
                records.append(record)

        return records

    def get_company_info(self):
        """Get company name and app name from the sheet"""
        return {
            'company_name': self.company_name,
            'app_name': self.app_name
        }
    
    def get_all_sheets(self):
        """Get list of all sheet names in the spreadsheet"""
        worksheets = self.spreadsheet.worksheets()
        return [ws.title for ws in worksheets]


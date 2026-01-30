"""Quick test script to verify Google Sheets connection and email configuration"""

import os
import sys

def test_config():
    """Test configuration file"""
    print("=" * 50)
    print("Testing Configuration...")
    print("=" * 50)

    try:
        import json
        with open('config.json', 'r') as f:
            config = json.load(f)

        print(f"[OK] Config file loaded successfully")
        print(f"     Spreadsheet ID: {config.get('spreadsheet_id', 'NOT SET')[:20]}...")
        print(f"     Credentials file: {config.get('credentials_file', 'NOT SET')}")
        print(f"     Sender email: {config.get('sender_email', 'NOT SET')}")
        print(f"     SMTP server: {config.get('smtp_server', 'NOT SET')}")

        # Check if credentials file exists
        cred_file = config.get('credentials_file', '')
        if os.path.exists(cred_file):
            print(f"[OK] Credentials file exists at: {cred_file}")
        else:
            print(f"[ERROR] Credentials file not found at: {cred_file}")
            return False

        return True
    except Exception as e:
        print(f"[ERROR] Config test failed: {e}")
        return False

def test_google_sheets():
    """Test Google Sheets connection"""
    print("\n" + "=" * 50)
    print("Testing Google Sheets Connection...")
    print("=" * 50)

    try:
        from google_sheets_reader import GoogleSheetsReader

        reader = GoogleSheetsReader()
        print("[OK] Google Sheets reader initialized")

        # Get all sheets
        sheets = reader.get_all_sheets()
        print(f"[OK] Connected to spreadsheet successfully!")
        print(f"     Available sheets: {', '.join(sheets)}")

        # Try to get data from the first sheet
        if sheets:
            test_sheet = sheets[0]
            print(f"\nTrying to read from sheet: '{test_sheet}'...")
            records = reader.get_month_data(test_sheet)

            # Get company info
            company_info = reader.get_company_info()
            if company_info.get('company_name'):
                print(f"[OK] Company Name: {company_info['company_name']}")
            if company_info.get('app_name'):
                print(f"[OK] App Name: {company_info['app_name']}")

            print(f"[OK] Found {len(records)} employee records")

            if records:
                print(f"\nSample record fields: {list(records[0].keys())}")

        return True
    except Exception as e:
        print(f"[ERROR] Google Sheets test failed: {e}")
        return False

def test_email_config():
    """Test email configuration (without sending)"""
    print("\n" + "=" * 50)
    print("Testing Email Configuration...")
    print("=" * 50)

    try:
        from email_sender import EmailSender

        sender = EmailSender()
        sender._ensure_configured()
        print("[OK] Email sender configured successfully")
        print(f"     SMTP Server: {sender.smtp_server}:{sender.smtp_port}")
        print(f"     Sender Email: {sender.sender_email}")

        return True
    except Exception as e:
        print(f"[ERROR] Email config test failed: {e}")
        return False

def test_pdf_generator():
    """Test PDF generator"""
    print("\n" + "=" * 50)
    print("Testing PDF Generator...")
    print("=" * 50)

    try:
        from pdf_generator import PDFGenerator

        generator = PDFGenerator()
        print("[OK] PDF generator initialized")

        # Create a test record
        test_record = {
            'Name': 'Test Employee',
            'email address': 'test@example.com',
            'CNIC': '12345-1234567-1',
            'Designation': 'Software Engineer',
            'Basic Salary': 50000,
            'Food Allowance': 5000,
            'Travel Allowance': 3000,
            'Net Salary': 58000
        }

        # Set company info
        generator.set_company_info("Test Company", "Test App")

        # Create test directory
        os.makedirs("pdfs", exist_ok=True)

        # Generate test PDF
        pdf_path = generator.create_pdf(test_record, "Test Month 2025")
        print(f"[OK] Test PDF created: {pdf_path}")

        # Clean up test PDF
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
            print("[OK] Test PDF cleaned up")

        return True
    except Exception as e:
        print(f"[ERROR] PDF generator test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("   SALARY AUTOMATION - CONNECTION TEST")
    print("=" * 60)

    results = {
        'Config': test_config(),
        'Google Sheets': test_google_sheets(),
        'Email Config': test_email_config(),
        'PDF Generator': test_pdf_generator()
    }

    print("\n" + "=" * 60)
    print("   TEST SUMMARY")
    print("=" * 60)

    all_passed = True
    for test_name, passed in results.items():
        status = "[PASS]" if passed else "[FAIL]"
        print(f"  {status} {test_name}")
        if not passed:
            all_passed = False

    print("=" * 60)

    if all_passed:
        print("\nAll tests passed! The automation is ready to use.")
        print("Run 'python main.py' to start the application.")
    else:
        print("\nSome tests failed. Please check the errors above.")

    sys.exit(0 if all_passed else 1)

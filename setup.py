"""
Setup helper script to create necessary directories and check configuration
"""
import os
import json
import sys

def create_directories():
    """Create necessary directories"""
    directories = ['pdfs']
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        print(f"✓ Created directory: {directory}/")

def check_config():
    """Check if config.json exists"""
    if not os.path.exists('config.json'):
        print("⚠ config.json not found!")
        print("Please copy config.json.example to config.json and fill in your details.")
        return False
    else:
        print("✓ config.json found")
        return True

def check_credentials():
    """Check if credentials.json exists"""
    if not os.path.exists('credentials.json'):
        print("⚠ credentials.json not found!")
        print("Please download your Google Service Account credentials and save as credentials.json")
        return False
    else:
        print("✓ credentials.json found")
        return True

def validate_config():
    """Validate config.json structure"""
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
        
        required_keys = ['spreadsheet_id', 'credentials_file', 'sender_email', 'sender_password']
        missing_keys = [key for key in required_keys if key not in config]
        
        if missing_keys:
            print(f"⚠ Missing keys in config.json: {', '.join(missing_keys)}")
            return False
        
        # Check for placeholder values
        if 'YOUR_GOOGLE_SHEET_ID_HERE' in str(config.get('spreadsheet_id', '')):
            print("⚠ Please update spreadsheet_id in config.json")
            return False
        
        if 'your-email@gmail.com' in str(config.get('sender_email', '')):
            print("⚠ Please update sender_email in config.json")
            return False
        
        print("✓ config.json is properly configured")
        return True
        
    except json.JSONDecodeError:
        print("⚠ config.json is not valid JSON")
        return False
    except Exception as e:
        print(f"⚠ Error reading config.json: {str(e)}")
        return False

def main():
    print("=" * 50)
    print("Salary Automation System - Setup Check")
    print("=" * 50)
    print()

    # Create directories
    print("Creating directories...")
    create_directories()
    print()

    # Check files
    print("Checking configuration files...")
    config_exists = check_config()
    credentials_exists = check_credentials()
    print()

    config_valid = False
    if config_exists:
        print("Validating configuration...")
        config_valid = validate_config()
        print()

    print("=" * 50)
    if config_exists and credentials_exists and config_valid:
        print("Setup complete! You can now run: python main.py")
    else:
        print("Please complete the setup before running the application")
        print()
        print("Next steps:")
        if not config_exists:
            print("  1. Copy config.json.example to config.json")
            print("  2. Fill in your Google Sheet ID and email credentials")
        if not credentials_exists:
            print("  3. Download Google Service Account credentials as credentials.json")
        if config_exists and not config_valid:
            print("  4. Fix the configuration issues mentioned above")
        print("=" * 50)
        sys.exit(1)
    print("=" * 50)

if __name__ == "__main__":
    main()



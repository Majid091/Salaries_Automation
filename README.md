# Salary Automation System

An automated Python application that reads salary data from Google Sheets, generates PDF statements for each employee, and sends them via email.

## Features

- 📊 **Google Sheets Integration**: Automatically reads data from Google Sheets organized by month
- 📄 **PDF Generation**: Creates professional PDF salary statements for each employee
- 📧 **Email Automation**: Automatically sends PDFs to employee email addresses
- 🖥️ **User-Friendly UI**: Simple graphical interface for month selection and automation control
- ⚡ **Progress Tracking**: Real-time progress bar and status updates

## Prerequisites

- Python 3.7 or higher
- Google Cloud Project with Sheets API enabled
- Gmail account (or other SMTP email service)

## Setup Instructions

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Google Sheets Setup

1. **Create a Google Cloud Project**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one

2. **Enable Google Sheets API**:
   - Navigate to "APIs & Services" > "Library"
   - Search for "Google Sheets API" and enable it
   - Search for "Google Drive API" and enable it

3. **Create Service Account**:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "Service Account"
   - Create a service account and download the JSON key file
   - Rename it to `credentials.json` and place it in the project root

4. **Share Google Sheet**:
   - Open your Google Sheet
   - Click "Share" and add the service account email (found in credentials.json)
   - Give it "Editor" access
   - Copy the Spreadsheet ID from the URL:
     ```
     https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
     ```

### 3. Google Sheet Structure

Your Google Sheet should have:
- **One worksheet per month** (e.g., "January", "February", etc.)
- **First row as headers** with column names
- **Required columns** (adjust in code if needed):
  - `name` - Employee name
  - `email` - Employee email address
  - `employee_id` - Employee ID
  - `basic_salary`, `gross_salary`, `net_salary` - Salary fields
  - Any other fields you want to include in the PDF

### 4. Email Configuration

For Gmail:
1. Enable 2-Step Verification on your Google account
2. Generate an App Password:
   - Go to Google Account settings
   - Security > 2-Step Verification > App passwords
   - Generate a password for "Mail"
   - Use this password in `config.json`

For other email providers, update SMTP settings in `config.json`.

### 5. Configuration File

1. Copy `config.json.example` to `config.json`:
   ```bash
   copy config.json.example config.json
   ```

2. Edit `config.json` with your details:
   ```json
   {
     "spreadsheet_id": "your-spreadsheet-id-here",
     "credentials_file": "credentials.json",
     "sender_email": "your-email@gmail.com",
     "sender_password": "your-app-password",
     "smtp_server": "smtp.gmail.com",
     "smtp_port": 587
   }
   ```

## Usage

1. **Run the application**:
   ```bash
   python main.py
   ```

2. **Select a month** from the dropdown menu

3. **Generate PDFs**:
   - Click "Generate PDFs" button
   - The system will fetch data from the selected month's sheet
   - PDFs will be created in the `pdfs` folder

4. **Send Emails**:
   - Click "Send Emails" button
   - The system will automatically:
     - Read email addresses from the sheet
     - Attach corresponding PDFs
     - Send emails to each employee

## Project Structure

```
Salaries_Automation/
│
├── main.py                      # Main application with UI
├── google_sheets_reader.py      # Google Sheets integration
├── pdf_generator.py             # PDF generation module
├── email_sender.py              # Email sending module
├── config.json                  # Configuration file (create from example)
├── credentials.json             # Google service account credentials
├── requirements.txt             # Python dependencies
├── README.md                    # This file
└── pdfs/                        # Generated PDFs folder (auto-created)
```

## Customization

### Adding More Fields to PDF

Edit `pdf_generator.py`:
- Modify `_get_employee_info()` to add employee fields
- Modify `_get_salary_details()` to add salary fields

### Customizing Email Template

Edit `main.py`, method `_get_email_body()` to customize the email content.

### Changing PDF Layout

Edit `pdf_generator.py` to modify colors, fonts, and layout.

## Troubleshooting

### "Worksheet not found" error
- Ensure your sheet tabs are named exactly as months (e.g., "January", "February")
- Check case sensitivity

### "Authentication failed" error
- Verify `credentials.json` is correct
- Ensure service account has access to the sheet

### "Email sending failed" error
- Check SMTP credentials in `config.json`
- For Gmail, ensure you're using an App Password, not your regular password
- Check firewall/network settings

### PDFs not generating
- Ensure the `pdfs` folder exists or can be created
- Check that all required fields are present in the Google Sheet

## Security Notes

- **Never commit** `config.json` or `credentials.json` to version control
- Add them to `.gitignore`
- Keep your service account credentials secure
- Use App Passwords for email, not your main account password

## License

This project is for internal use. Modify as needed for your organization.

## Support

For issues or questions, please contact your IT department or the project maintainer.



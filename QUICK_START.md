# Quick Start Guide - How to Run the Salary Automation System

## Prerequisites Checklist

Before running, make sure you have:
- ✅ Python 3.7+ installed (you have Python 3.14.2 ✓)
- ✅ Google Cloud Project with Sheets API enabled
- ✅ Gmail account (or email service with SMTP)

---

## Step-by-Step Setup

### Step 1: Install Dependencies

Open PowerShell/Command Prompt in this folder and run:

```bash
python -m pip install -r requirements.txt
```

---

### Step 2: Set Up Google Sheets Access

#### 2.1 Create Google Cloud Project & Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or select existing)
3. Enable APIs:
   - Go to **APIs & Services** > **Library**
   - Enable **Google Sheets API**
   - Enable **Google Drive API**
4. Create Service Account:
   - Go to **APIs & Services** > **Credentials**
   - Click **Create Credentials** > **Service Account**
   - Give it a name (e.g., "salary-automation")
   - Click **Create and Continue**
   - Skip role assignment, click **Done**
5. Download Credentials:
   - Click on the created service account
   - Go to **Keys** tab
   - Click **Add Key** > **Create new key**
   - Choose **JSON** format
   - Download the file
   - **Rename it to `credentials.json`** and place it in this project folder

#### 2.2 Share Your Google Sheet

1. Open your Google Sheet
2. Click the **Share** button (top right)
3. Find the service account email in `credentials.json` (look for `"client_email"` field)
4. Paste that email address in the share dialog
5. Give it **Editor** access
6. Click **Send**

#### 2.3 Get Your Spreadsheet ID

From your Google Sheet URL:
```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
```
Copy the `SPREADSHEET_ID_HERE` part.

---

### Step 3: Configure Email Settings

#### 3.1 For Gmail Users:

1. Enable 2-Step Verification on your Google account
2. Generate App Password:
   - Go to [Google Account Settings](https://myaccount.google.com/)
   - Security > **2-Step Verification** > **App passwords**
   - Select **Mail** and **Other (Custom name)**
   - Enter "Salary Automation"
   - Click **Generate**
   - **Copy the 16-character password** (you'll need this)

#### 3.2 Create Config File

1. Copy the example config:
   ```bash
   copy config.json.example config.json
   ```

2. Open `config.json` in a text editor and fill in:

```json
{
  "spreadsheet_id": "YOUR_SPREADSHEET_ID_FROM_STEP_2.3",
  "credentials_file": "credentials.json",
  "sender_email": "your-email@gmail.com",
  "sender_password": "your-16-character-app-password",
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Important:** 
- Use the **App Password** (16 characters), NOT your regular Gmail password
- Keep `credentials_file` as `"credentials.json"`

---

### Step 4: Prepare Your Google Sheet

Your Google Sheet should have:
- **One worksheet per month** (named "January", "February", etc.)
- **First row as headers** with column names
- **Recommended columns:**
  - `name` or `Name` - Employee name
  - `email` or `Email` - Employee email
  - `employee_id` or `Employee ID` - Employee ID
  - `basic_salary` or `Basic Salary` - Basic salary
  - `gross_salary` or `Gross Salary` - Gross salary
  - `net_salary` or `Net Salary` - Net salary
  - Any other fields you want in the PDF

**Note:** The system is case-insensitive and handles various field name formats.

---

### Step 5: Verify Setup

Run the setup check script:

```bash
python setup.py
```

This will verify:
- ✅ All required files exist
- ✅ Configuration is correct
- ✅ Directories are created

---

### Step 6: Run the Application

```bash
python main.py
```

A window will open with:
- Month selection dropdown
- **Generate PDFs** button
- **Send Emails** button
- Progress bar

---

## How to Use

### Generate PDFs:

1. Select a month from the dropdown
2. Click **Generate PDFs**
3. Wait for processing (progress bar will show status)
4. PDFs will be saved in the `pdfs/` folder

### Send Emails:

1. **First generate PDFs** (see above)
2. Select the same month
3. Click **Send Emails**
4. Confirm when prompted
5. The system will automatically:
   - Read email addresses from the sheet
   - Attach corresponding PDFs
   - Send emails to each employee

---

## Troubleshooting

### "Config file not found"
- Make sure you copied `config.json.example` to `config.json`
- Check that `config.json` is in the same folder as `main.py`

### "Credentials file not found"
- Make sure `credentials.json` is in the project folder
- Check the filename is exactly `credentials.json` (not `credentials.json.json`)

### "Worksheet not found"
- Check that your sheet tabs are named exactly as months (e.g., "January", "February")
- The system tries different cases, but exact match is best

### "Email sending failed"
- For Gmail: Make sure you're using an **App Password**, not your regular password
- Check that 2-Step Verification is enabled
- Verify SMTP settings in `config.json`

### "Authentication failed"
- Make sure you shared the Google Sheet with the service account email
- Verify the service account email in `credentials.json`
- Check that Google Sheets API is enabled in Google Cloud Console

---

## File Structure

```
Salaries_Automation/
├── main.py                 # Main application (run this!)
├── google_sheets_reader.py # Google Sheets integration
├── pdf_generator.py        # PDF creation
├── email_sender.py         # Email sending
├── setup.py                # Setup verification
├── config.json             # Your configuration (create this)
├── credentials.json         # Google credentials (download this)
├── requirements.txt        # Python dependencies
├── pdfs/                   # Generated PDFs (auto-created)
└── README.md              # Full documentation
```

---

## Need Help?

1. Run `python setup.py` to check your configuration
2. Check the error messages in the application window
3. Verify all files are in the correct location
4. Make sure your Google Sheet is properly formatted

---

## Quick Command Reference

```bash
# Install dependencies
python -m pip install -r requirements.txt

# Check setup
python setup.py

# Run application
python main.py
```

Good luck! 🚀



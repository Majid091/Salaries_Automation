import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import os

class EmailSender:
    def __init__(self, config_file="config.json"):
        """
        Initialize email sender with configuration from config file
        """
        self.config_file = config_file
        self.config = None
        self.smtp_server = None
        self.smtp_port = None
        self.sender_email = None
        self.sender_password = None
        self.is_configured = False

    def _ensure_configured(self):
        """Load config if not already loaded, raise error if not configured"""
        if self.is_configured:
            return True

        self.config = self._load_config(self.config_file)
        self.smtp_server = self.config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = self.config.get('smtp_port', 587)
        self.sender_email = self.config.get('sender_email', '').strip()
        # Get app password (try both with and without spaces)
        self.sender_password = self.config.get('sender_password', '').strip()

        if not self.sender_email or not self.sender_password:
            raise ValueError("Email credentials not configured in config.json")

        self.is_configured = True
        return True
    
    def _load_config(self, config_file):
        """Load configuration from JSON file"""
        if not os.path.exists(config_file):
            raise FileNotFoundError(
                f"Config file '{config_file}' not found. "
                "Please create it with your email credentials."
            )
        
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        return config
    
    def send_email(self, to_email, subject, body, pdf_path=None):
        """
        Send an email with optional PDF attachment

        Args:
            to_email: Recipient email address
            subject: Email subject
            body: Email body text
            pdf_path: Path to PDF file to attach (optional)

        Returns:
            True if successful
        """
        # Ensure config is loaded
        self._ensure_configured()

        # Create message
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Add body to email
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach PDF if provided
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            
            filename = os.path.basename(pdf_path)
            part.add_header(
                'Content-Disposition',
                'attachment',
                filename=filename
            )
            
            msg.attach(part)
        
        # Create SMTP session
        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()  # Enable security
            server.login(self.sender_email, self.sender_password)
            
            # Send email
            text = msg.as_string()
            server.sendmail(self.sender_email, to_email, text)
            server.quit()
            
            return True
            
        except Exception as e:
            raise Exception(f"Failed to send email: {str(e)}")



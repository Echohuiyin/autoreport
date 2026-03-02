#!/usr/bin/env python3
"""
Configuration manager for the weekly report system.
Handles loading configuration from environment variables and config file.
"""

import os
import sys
import logging
from src.exceptions import ReportConfigurationError
from src.constants import DEFAULT_EXCEL_FILE, DEFAULT_EMAIL_SUBJECT

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))

try:
    from dotenv import load_dotenv
    # Load environment variables from .env file if it exists
    load_dotenv()
except ImportError:
    logging.warning("python-dotenv not installed. Using default configuration.")
    pass

logger = logging.getLogger(__name__)

class ConfigManager:
    """Manages configuration for the weekly report system."""
    
    def __init__(self):
        self._load_config()
    
    def _load_config(self):
        """Load configuration from environment variables."""
        # Email configuration
        self.email_config = {
            'sender_email': os.getenv('SENDER_EMAIL'),
            'sender_password': os.getenv('SENDER_PASSWORD'),
            'smtp_server': os.getenv('SMTP_SERVER', 'smtp.163.com'),
            'smtp_port': int(os.getenv('SMTP_PORT', '465'))
        }
        
        # Recipients configuration
        to_emails = os.getenv('TO_EMAILS')
        cc_emails = os.getenv('CC_EMAILS')
        
        self.recipients_config = {
            'to_emails': [email.strip() for email in to_emails.split(',')] if to_emails else [],
            'cc_emails': [email.strip() for email in cc_emails.split(',')] if cc_emails else []
        }
        
        # File configuration
        self.file_config = {
            'excel_file_path': os.getenv('EXCEL_FILE_PATH', DEFAULT_EXCEL_FILE),
            'subject': os.getenv('EMAIL_SUBJECT', DEFAULT_EMAIL_SUBJECT),
            'body_template': os.getenv('EMAIL_BODY_TEMPLATE', '''
Dear Team,

Please find attached the weekly report.

Best regards,
Automated Report System
''')
        }
    
    def get_email_config(self):
        """Get email configuration."""
        return self.email_config
    
    def get_recipients_config(self):
        """Get recipients configuration."""
        return self.recipients_config
    
    def get_file_config(self):
        """Get file configuration."""
        return self.file_config
    
    def validate_config(self):
        """Validate configuration values."""
        # Validate email configuration
        if not self.email_config['sender_email']:
            raise ReportConfigurationError("Please configure sender_email in environment variables or config file")
        
        if not self.email_config['sender_password']:
            raise ReportConfigurationError("Please configure sender_password in environment variables")
        
        # Validate recipients configuration
        if not self.recipients_config['to_emails']:
            raise ReportConfigurationError("Please configure to_emails in environment variables or config file")
        
        # Validate file configuration
        if not os.path.exists(self.file_config['excel_file_path']):
            # Check for alternative filename without typo
            alt_filename = self.file_config['excel_file_path'].replace('weekyly', 'weekly')
            if os.path.exists(alt_filename):
                self.file_config['excel_file_path'] = alt_filename
                logger.info(f"Using alternative filename: {alt_filename}")
            else:
                raise FileNotFoundError(f"Excel file not found: {self.file_config['excel_file_path']}")
        
        logger.info("Configuration validated successfully")

# Create a singleton instance
config_manager = ConfigManager()
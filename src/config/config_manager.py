#!/usr/bin/env python3
"""
Configuration manager for the weekly report system.
Handles loading configuration from environment variables and config file.
"""

import os
import logging

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
        """Load configuration from environment variables and config file."""
        # Email configuration
        self.email_config = {
            'sender_email': os.getenv('SENDER_EMAIL', 'mingruiliu99@163.com'),
            'sender_password': os.getenv('SENDER_PASSWORD', 'QQheCPvibk9AEFDf'),
            'smtp_server': os.getenv('SMTP_SERVER', 'smtp.163.com'),
            'smtp_port': int(os.getenv('SMTP_PORT', '465'))
        }
        
        # Recipients configuration
        to_emails = os.getenv('TO_EMAILS', 'lmr09232007@163.com')
        cc_emails = os.getenv('CC_EMAILS', 'mingruiliu99@163.com')
        
        self.recipients_config = {
            'to_emails': [email.strip() for email in to_emails.split(',')],
            'cc_emails': [email.strip() for email in cc_emails.split(',')]
        }
        
        # File configuration
        self.file_config = {
            'excel_file_path': os.getenv('EXCEL_FILE_PATH', 'weekly_report.xlsx'),
            'subject': os.getenv('EMAIL_SUBJECT', 'Weekly Report'),
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
        if not self.email_config['sender_email'] or self.email_config['sender_email'] == 'your_email@example.com':
            raise ValueError("Please configure sender_email in environment variables or config file")
        
        if not self.email_config['sender_password'] or self.email_config['sender_password'] == 'your_app_password':
            raise ValueError("Please configure sender_password in environment variables")
        
        # Validate recipients configuration
        if not self.recipients_config['to_emails'] or self.recipients_config['to_emails'] == ['recipient1@example.com', 'recipient2@example.com']:
            raise ValueError("Please configure to_emails in environment variables or config file")
        
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
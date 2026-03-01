#!/usr/bin/env python3
"""
Test script to verify Excel processing functionality without sending email.
"""

import os
import sys
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Import the test config
import config_test as config

# Add debug logging
import logging
logging.basicConfig(level=logging.DEBUG)

# Import the WeeklyReportSender class
from weekly_report_sender import WeeklyReportSender

class TestWeeklyReportSender(WeeklyReportSender):
    """Test version that skips email sending."""
    def __init__(self):
        # Override to use test config
        self.sender_email = config.EMAIL_CONFIG['sender_email']
        self.sender_password = config.EMAIL_CONFIG['sender_password']
        self.smtp_server = config.EMAIL_CONFIG['smtp_server']
        self.smtp_port = config.EMAIL_CONFIG['smtp_port']
        self.to_emails = config.RECIPIENTS_CONFIG['to_emails']
        self.cc_emails = config.RECIPIENTS_CONFIG['cc_emails']
        self.excel_file_path = config.FILE_CONFIG['excel_file_path']
        self.subject = config.FILE_CONFIG['subject']
        self.body_template = config.FILE_CONFIG['body_template']
    
    def send_email(self, msg):
        """Skip sending email for testing."""
        logger.info("Skipping email sending for testing")
        # Print the email content instead
        logger.info("Email body preview:")
        # Extract and print the HTML part
        for part in msg.get_payload():
            if part.get_content_type() == 'text/html':
                html_content = part.get_payload(decode=True).decode('utf-8')
                # Print first 500 characters to see the structure
                logger.info(html_content[:500] + "...")
        return

def main():
    """Test the Excel processing functionality."""
    try:
        logger.info("Starting Excel processing test...")
        
        # Create test sender
        sender = TestWeeklyReportSender()
        
        # Validate configuration
        sender.validate_config()
        logger.info("Configuration validated successfully")
        
        # Read and parse Excel content
        logger.info("Reading Excel file...")
        excel_content = sender.read_excel_content()
        logger.info("Excel content processed successfully")
        
        # Create email message
        logger.info("Creating email message...")
        msg = sender.create_email_message(excel_content)
        logger.info("Email message created successfully")
        
        # Skip sending email, just print preview
        sender.send_email(msg)
        
        logger.info("Excel processing test completed successfully!")
        
    except Exception as e:
        logger.error(f"Test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
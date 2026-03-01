#!/usr/bin/env python3
"""
Test script to generate HTML output and save it to a file for inspection.
"""

import os
import sys
import config_test as config
from weekly_report_sender import WeeklyReportSender

class TestWeeklyReportSender(WeeklyReportSender):
    """Test version that generates HTML output."""
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

def main():
    """Generate HTML output and save to file."""
    try:
        print("Generating HTML output...")
        
        # Create test sender
        sender = TestWeeklyReportSender()
        
        # Validate configuration
        sender.validate_config()
        print("Configuration validated successfully")
        
        # Read and parse Excel content
        print("Reading Excel file...")
        excel_content = sender.read_excel_content()
        print("Excel content processed successfully")
        
        # Save HTML to file
        with open('email_body.html', 'w', encoding='utf-8') as f:
            f.write(excel_content)
        print("HTML output saved to email_body.html")
        
        # Print some debug info
        print("\nDebug information:")
        print(f"HTML length: {len(excel_content)} characters")
        print("First 500 characters of HTML:")
        print(excel_content[:500] + "...")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
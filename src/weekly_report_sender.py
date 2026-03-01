#!/usr/bin/env python3
"""
Main module for the weekly report system.
Coordinates the different components to send weekly reports.
"""

import os
import sys

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

import logging
from src.config.config_manager import config_manager
from src.excel.excel_reader import ExcelReader
from src.html.html_generator import HtmlGenerator
from src.email.email_sender import EmailSender

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('weekly_report.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class WeeklyReportSender:
    """Handles reading Excel files and sending weekly reports via email."""
    
    def __init__(self, excel_file_path=None):
        # Get configuration
        self.config = config_manager
        self.email_config = self.config.get_email_config()
        self.recipients_config = self.config.get_recipients_config()
        self.file_config = self.config.get_file_config()
        
        # Override excel file path if provided
        if excel_file_path:
            self.file_config['excel_file_path'] = excel_file_path
        
        # Initialize components
        self.excel_reader = ExcelReader(self.file_config['excel_file_path'])
        self.html_generator = HtmlGenerator()
        self.email_sender = EmailSender(self.email_config, self.recipients_config)
    
    def validate_config(self):
        """Validate that all required configuration values are set."""
        self.config.validate_config()
    
    def read_excel_content(self):
        """Read Excel content and convert to HTML."""
        return self.excel_reader.read_excel_content()
    
    def create_email_message(self, excel_content):
        """Create email message with Excel content."""
        html_body = self.html_generator.generate_html_from_excel(excel_content)
        return self.email_sender.create_email_message(
            self.file_config['subject'],
            html_body
        )
    
    def send_email(self, msg):
        """Send email."""
        return self.email_sender.send_email(msg)
    
    def run(self):
        """Main execution method."""
        try:
            logger.info("Starting weekly report automation...")
            
            # Validate configuration
            self.validate_config()
            
            # Read and parse Excel content with merged cell support
            excel_content = self.read_excel_content()
            
            # Create email message
            msg = self.create_email_message(excel_content)
            
            # Send email
            self.send_email(msg)
            
            logger.info("Weekly report automation completed successfully!")
            
        except Exception as e:
            logger.error(f"Weekly report automation failed: {str(e)}")
            raise

def main():
    """Main function to run the weekly report sender."""
    try:
        sender = WeeklyReportSender()
        sender.run()
    except KeyboardInterrupt:
        logger.info("Process interrupted by user")
        exit(0)
    except Exception as e:
        logger.error(f"Application failed: {str(e)}")
        exit(1)

if __name__ == "__main__":
    main()
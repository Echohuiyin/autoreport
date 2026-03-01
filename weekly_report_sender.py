#!/usr/bin/env python3
"""
Automated Weekly Report Sender
Reads Excel file content and sends it via email to configured recipients.
"""

import os
import sys
import logging
import smtplib
import pandas as pd
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('weekly_report.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

try:
    import pandas as pd
    import openpyxl
except ImportError:
    logger.error("Required libraries missing. Please install: pip install pandas openpyxl")
    sys.exit(1)

try:
    from config import EMAIL_CONFIG, RECIPIENTS_CONFIG, FILE_CONFIG
except ImportError:
    logger.error("config.py file not found. Please create the configuration file first.")
    sys.exit(1)


class WeeklyReportSender:
    """Handles reading Excel files and sending weekly reports via email."""
    
    def __init__(self):
        self.sender_email = EMAIL_CONFIG['sender_email']
        self.sender_password = EMAIL_CONFIG['sender_password']
        self.smtp_server = EMAIL_CONFIG['smtp_server']
        self.smtp_port = EMAIL_CONFIG['smtp_port']
        self.to_emails = RECIPIENTS_CONFIG['to_emails']
        self.cc_emails = RECIPIENTS_CONFIG['cc_emails']
        self.excel_file_path = FILE_CONFIG['excel_file_path']
        self.subject = FILE_CONFIG['subject']
        self.body_template = FILE_CONFIG['body_template']
    
    def validate_config(self):
        """Validate that all required configuration values are set."""
        if not self.sender_email or self.sender_email == 'your_email@example.com':
            raise ValueError("Please configure sender_email in config.py")
        if not self.sender_password or self.sender_password == 'your_app_password':
            raise ValueError("Please configure sender_password in config.py")
        if not self.to_emails or self.to_emails == ['recipient1@example.com', 'recipient2@example.com']:
            raise ValueError("Please configure to_emails in config.py")
        if not os.path.exists(self.excel_file_path):
            # Check for alternative filename without typo
            alt_filename = self.excel_file_path.replace('weekyly', 'weekly')
            if os.path.exists(alt_filename):
                self.excel_file_path = alt_filename
                logger.info(f"Using alternative filename: {alt_filename}")
            else:
                raise FileNotFoundError(f"Excel file not found: {self.excel_file_path}")
    
    def read_excel_with_merged_cells(self, file_path):
        """
        Read Excel file and handle merged cells properly.
        This function fills in the missing values that result from merged cells.
        """
        # Load the workbook
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        
        # Get the data starting from row 2 (skip title row)
        data_rows = []
        headers = None
        
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if row_idx == 1:
                # Skip the title row "周报"
                continue
            elif row_idx == 2:
                # This is the header row
                headers = [cell if cell is not None else f'Column_{i}' for i, cell in enumerate(row)]
            else:
                # Data rows
                data_rows.append(row)
        
        wb.close()
        
        if headers is None:
            raise ValueError("Could not find header row in Excel file")
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Handle merged cells by forward-filling the '项目' column
        # This will propagate category names down through empty cells
        if '项目' in df.columns:
            df['项目'] = df['项目'].ffill()  # Use ffill() instead of fillna(method='ffill')
        
        return df
    
    def parse_excel_structure(self, df):
        """
        Parse the hierarchical Excel structure to extract meaningful content.
        After handling merged cells, the '项目' column should contain category names
        for all relevant rows.
        """
        parsed_content = []
        current_category = None
        
        # Group by category in the '项目' column
        for idx, row in df.iterrows():
            project_val = row['项目'] if pd.notna(row['项目']) else None
            name_val = row['名称'] if pd.notna(row['名称']) else None
            
            # Skip rows with no name (likely empty or category-only rows)
            if name_val is None:
                continue
            
            # Check if we have a new category
            if project_val is not None and project_val != current_category:
                current_category = project_val
                parsed_content.append(f"\n## {current_category}")
            
            # Extract other column values
            progress_val = row['进展'] if pd.notna(row['进展']) else ''
            handler_val = row['处理人'] if pd.notna(row['处理人']) else ''
            status_val = row['状态'] if pd.notna(row['状态']) else ''
            
            # Format the data row
            data_row = f"- **{name_val}**"
            if progress_val:
                data_row += f" | Progress: {progress_val}"
            if handler_val:
                data_row += f" | Handler: {handler_val}"
            if status_val:
                data_row += f" | Status: {status_val}"
            
            parsed_content.append(data_row)
        
        return "\n".join(parsed_content) if parsed_content else "No data found in Excel file."
    
    def read_excel_content(self):
        """Read and return the content of the Excel file as a formatted string."""
        try:
            # Read Excel file with merged cell handling
            df = self.read_excel_with_merged_cells(self.excel_file_path)
            
            # Parse the hierarchical structure
            parsed_content = self.parse_excel_structure(df)
            
            # Log summary information
            logger.info(f"Successfully processed Excel file: {len(df)} rows processed")
            if len(df) > 0:
                categories = df['项目'].dropna().unique() if '项目' in df.columns else []
                logger.info(f"Categories found: {list(categories)}")
            
            return parsed_content
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            raise
    
    def create_email_message(self, excel_content):
        """Create the email message with formatted Excel content in the body."""
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = ', '.join(self.to_emails)
        msg['Cc'] = ', '.join(self.cc_emails) if self.cc_emails else ''
        msg['Subject'] = self.subject
        
        # Email body with parsed content
        body = self.body_template + f"\n\n---\n## Weekly Report Content:\n\n{excel_content}"
        msg.attach(MIMEText(body, 'plain'))
        
        # Remove attachment functionality as per requirement
        # Original attachment code commented out
        # try:
        #     with open(self.excel_file_path, "rb") as attachment:
        #         part = MIMEBase('application', 'octet-stream')
        #         part.set_payload(attachment.read())
        #     
        #     encoders.encode_base64(part)
        #     part.add_header(
        #         'Content-Disposition',
        #         f'attachment; filename= {os.path.basename(self.excel_file_path)}'
        #     )
        #     msg.attach(part)
        #     
        # except Exception as e:
        #     logger.error(f"Error attaching Excel file: {str(e)}")
        #     raise
        
        return msg
    
    def send_email(self, msg):
        """Send the email using SMTP."""
        try:
            # Create SMTP session
            if self.smtp_port == 465:
                # Use SSL for port 465
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            else:
                # Use TLS for other ports (25, 587)
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.starttls()  # Enable TLS encryption
            
            server.login(self.sender_email, self.sender_password)
            
            # Get all recipients (to + cc)
            all_recipients = self.to_emails + (self.cc_emails if self.cc_emails else [])
            
            # Send email
            text = msg.as_string()
            server.sendmail(self.sender_email, all_recipients, text)
            server.quit()
            
            logger.info(f"Email sent successfully to {len(all_recipients)} recipients")
            logger.info(f"To: {', '.join(self.to_emails)}")
            if self.cc_emails:
                logger.info(f"CC: {', '.join(self.cc_emails)}")
                
        except smtplib.SMTPAuthenticationError:
            logger.error("SMTP authentication failed. Please check your email and password.")
            raise
        except smtplib.SMTPException as e:
            logger.error(f"SMTP error occurred: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error while sending email: {str(e)}")
            raise
    
    def run(self):
        """Main execution method."""
        try:
            logger.info("Starting weekly report automation...")
            
            # Validate configuration
            self.validate_config()
            logger.info("Configuration validated successfully")
            
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
        sys.exit(0)
    except Exception as e:
        logger.error(f"Application failed: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
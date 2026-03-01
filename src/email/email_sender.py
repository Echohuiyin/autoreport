#!/usr/bin/env python3
"""
Email sender module for the weekly report system.
Handles sending emails with Excel content.
"""

import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

logger = logging.getLogger(__name__)

class EmailSender:
    """Handles sending emails with Excel content."""
    
    def __init__(self, email_config, recipients_config):
        self.sender_email = email_config['sender_email']
        self.sender_password = email_config['sender_password']
        self.smtp_server = email_config['smtp_server']
        self.smtp_port = email_config['smtp_port']
        self.to_emails = recipients_config['to_emails']
        self.cc_emails = recipients_config['cc_emails']
    
    def create_email_message(self, subject, html_content):
        """
        Create the email message with formatted Excel content in the body.
        
        Args:
            subject: Email subject
            html_content: HTML content for the email body
            
        Returns:
            MIMEMultipart: Email message object
        """
        try:
            msg = MIMEMultipart('alternative')
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(self.to_emails)
            msg['Cc'] = ', '.join(self.cc_emails) if self.cc_emails else ''
            msg['Subject'] = subject
            
            # Attach HTML part
            msg.attach(MIMEText(html_content, 'html'))
            
            return msg
            
        except Exception as e:
            logger.error(f"Error creating email message: {str(e)}")
            raise
    
    def send_email(self, msg):
        """
        Send the email using SMTP.
        
        Args:
            msg: Email message object
            
        Returns:
            bool: True if email sent successfully
        """
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
                
            return True
            
        except smtplib.SMTPAuthenticationError:
            logger.error("SMTP authentication failed. Please check your email and password.")
            raise
        except smtplib.SMTPException as e:
            logger.error(f"SMTP error occurred: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error while sending email: {str(e)}")
            raise
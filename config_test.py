"""
Test configuration file for the automated weekly report system.
"""

# Email configuration
EMAIL_CONFIG = {
    'sender_email': 'mingruiliu99@163.com',  # Your 163 email address
    'sender_password': 'test_password',   # Test password
    'smtp_server': 'smtp.163.com',           # 163 SMTP server
    'smtp_port': 465                         # 163 SMTP port for SSL
}

# Recipients configuration
RECIPIENTS_CONFIG = {
    'to_emails': ['lmr09232007@163.com'],  # Primary recipients
    'cc_emails': ['mingruiliu99@163.com']                 # CC recipients
}

# File configuration
FILE_CONFIG = {
    'excel_file_path': 'weekly_report.xlsx',  # Path to the Excel file
    'subject': 'Weekly Report',                # Email subject
    'body_template': '''
Dear Team,

Please find attached the weekly report.

Best regards,
Automated Report System
'''
}
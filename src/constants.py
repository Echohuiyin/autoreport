#!/usr/bin/env python3
"""
Constants for the weekly report system.
"""

# Default Excel file name
DEFAULT_EXCEL_FILE = 'weekly_report.xlsx'

# Default email subject
DEFAULT_EMAIL_SUBJECT = 'Weekly Report'

# HTML template constants
HTML_TABLE_CLASS = 'excel-table'

# Configuration keys
CONFIG_KEYS = {
    'SENDER_EMAIL': 'SENDER_EMAIL',
    'SENDER_PASSWORD': 'SENDER_PASSWORD',
    'SMTP_SERVER': 'SMTP_SERVER',
    'SMTP_PORT': 'SMTP_PORT',
    'TO_EMAILS': 'TO_EMAILS',
    'CC_EMAILS': 'CC_EMAILS',
    'EXCEL_FILE_PATH': 'EXCEL_FILE_PATH',
    'EMAIL_SUBJECT': 'EMAIL_SUBJECT'
}
# Automated Weekly Report System

This system automatically reads an Excel file (`weekyly report.xlsx`) and sends it via email to configured recipients.

## Features

- Reads Excel files using pandas
- Sends emails with attachments via SMTP
- Configurable sender, recipients, and email content
- Comprehensive logging and error handling
- Supports both primary recipients and CC recipients

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

Required packages:
- `pandas` - for reading Excel files
- `openpyxl` - Excel file support for pandas

### 2. Configure Email Settings

Edit the `config.py` file to set up your email configuration:

```python
# Email configuration
EMAIL_CONFIG = {
    'sender_email': 'your_email@gmail.com',     # Your Gmail address
    'sender_password': 'your_app_password',     # Gmail App Password (not regular password)
    'smtp_server': 'smtp.gmail.com',           # Gmail SMTP server
    'smtp_port': 587                           # Gmail TLS port
}

# Recipients configuration
RECIPIENTS_CONFIG = {
    'to_emails': ['recipient1@example.com', 'recipient2@example.com'],
    'cc_emails': ['cc1@example.com', 'cc2@example.com']
}

# File configuration
FILE_CONFIG = {
    'excel_file_path': 'weekyly report.xlsx',
    'subject': 'Weekly Report',
    'body_template': '''
Dear Team,

Please find attached the weekly report.

Best regards,
Automated Report System
'''
}
```

**Important Notes:**
- For Gmail, you need to use an [App Password](https://support.google.com/accounts/answer/185833) instead of your regular password
- For other email providers, update the SMTP server and port accordingly:
  - Outlook/Hotmail: `smtp-mail.outlook.com`, port 587
  - Yahoo: `smtp.mail.yahoo.com`, port 587

### 3. Prepare Your Excel File

Ensure your Excel file is named `weekyly report.xlsx` and placed in the same directory as the script.

## Usage

Run the automated weekly report sender:

```bash
python weekly_report_sender.py
```

The system will:
1. Validate your configuration
2. Read the Excel file content
3. Create an email with the file attached
4. Send the email to all configured recipients
5. Log all activities to both console and `weekly_report.log`

## Logging

All activities are logged to:
- Console output
- `weekly_report.log` file in the same directory

## Error Handling

The system includes comprehensive error handling for:
- Missing configuration values
- Invalid email credentials
- Missing Excel files
- Network/SMTP issues
- File reading errors

## Security Considerations

- Never commit your email password to version control
- Use environment variables or secure credential storage in production
- Consider using OAuth2 for production applications instead of app passwords

## Customization

You can easily customize:
- Email subject and body template
- Excel file path
- SMTP settings for different email providers
- Logging configuration
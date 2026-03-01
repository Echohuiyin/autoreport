# Automated Weekly Report System

This system automatically reads an Excel file (`weekly_report.xlsx`) and sends it via email to configured recipients, with the Excel content embedded in the email body while preserving formatting.

## Features

- Reads Excel files using pandas and openpyxl
- Embeds Excel content in email body with preserved formatting
- Preserves background fill colors, font colors, and font sizes
- Handles merged cells correctly
- Sends emails via SMTP
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
    'excel_file_path': 'weekly_report.xlsx',
    'subject': 'Weekly Report',
    'body_template': '''
Dear Team,

Please find the weekly report below.

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

Ensure your Excel file is named `weekly_report.xlsx` and placed in the same directory as the script.

The system will automatically detect and preserve the following formatting from your Excel file:
- Background fill colors
- Font colors
- Font sizes
- Merged cells
- Text alignment

## Usage

Run the automated weekly report sender:

```bash
python weekly_report_sender.py
```

The system will:
1. Validate your configuration
2. Read the Excel file content with formatting
3. Convert the Excel content to HTML while preserving formatting
4. Create an email with the Excel content embedded in the body
5. Send the email to all configured recipients
6. Log all activities to both console and `weekly_report.log`

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
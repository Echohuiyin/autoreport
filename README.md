# Automated Weekly Report System

This system automatically reads an Excel file (`weekly_report.xlsx`) and sends it via email to configured recipients, with the Excel content embedded in the email body while preserving formatting.

## Features

- Reads Excel files using openpyxl
- Embeds Excel content in email body with preserved formatting
- Preserves background fill colors, font colors, and font sizes
- Handles merged cells correctly
- Sends emails via SMTP
- Configurable sender, recipients, and email content
- Comprehensive logging and error handling
- Supports both primary recipients and CC recipients
- Modular architecture for easy maintenance and extension
- Environment variable support for secure configuration

## Project Structure

```
├── src/
│   ├── config/          # Configuration management
│   ├── excel/           # Excel file reading and processing
│   ├── html/            # HTML generation
│   ├── email/           # Email sending
│   └── weekly_report_sender.py  # Main coordinator
├── README.md            # Project documentation
├── requirements.txt     # Dependencies
├── config.py            # Configuration file
├── refactoring_analysis.md  # Refactoring analysis report
└── weekly_report.xlsx   # Sample Excel report
```

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

Required packages:
- `openpyxl` - for reading Excel files and preserving formatting
- `python-dotenv` - for environment variable support

### 2. Configure Email Settings

#### Using Environment Variables (Recommended)

The system now exclusively uses environment variables for configuration. Create a `.env` file in the project root:

```
# Email configuration
SENDER_EMAIL=your_email@163.com
SENDER_PASSWORD=your_email_password
SMTP_SERVER=smtp.163.com
SMTP_PORT=465

# Recipients configuration
TO_EMAILS=recipient1@example.com,recipient2@example.com
CC_EMAILS=cc1@example.com,cc2@example.com

# File configuration
EXCEL_FILE_PATH=weekly_report.xlsx
EMAIL_SUBJECT=Weekly Report
```

**Environment Variable Details:**

| Variable Name | Required | Description | Default Value |
|---------------|----------|-------------|---------------|
| `SENDER_EMAIL` | Yes | Your email address | None |
| `SENDER_PASSWORD` | Yes | Your email password or app password | None |
| `SMTP_SERVER` | No | SMTP server address | smtp.163.com |
| `SMTP_PORT` | No | SMTP server port | 465 |
| `TO_EMAILS` | Yes | Comma-separated list of primary recipients | None |
| `CC_EMAILS` | No | Comma-separated list of CC recipients | Empty list |
| `EXCEL_FILE_PATH` | No | Path to the Excel file | weekly_report.xlsx |
| `EMAIL_SUBJECT` | No | Email subject line | Weekly Report |
| `EMAIL_BODY_TEMPLATE` | No | Custom email body template | Default template |

**Setting Environment Variables Manually:**

If you prefer not to use a `.env` file, you can set environment variables directly:

- **Windows Command Prompt:**
  ```cmd
  set SENDER_EMAIL=your_email@163.com
  set SENDER_PASSWORD=your_password
  set TO_EMAILS=recipient@example.com
  python src/weekly_report_sender.py
  ```

- **Windows PowerShell:**
  ```powershell
  $env:SENDER_EMAIL = "your_email@163.com"
  $env:SENDER_PASSWORD = "your_password"
  $env:TO_EMAILS = "recipient@example.com"
  python src/weekly_report_sender.py
  ```

- **Linux/Mac Terminal:**
  ```bash
  export SENDER_EMAIL=your_email@163.com
  export SENDER_PASSWORD=your_password
  export TO_EMAILS=recipient@example.com
  python src/weekly_report_sender.py
  ```

**Email Provider Configuration:**

| Provider | SMTP Server | Port | Notes |
|----------|-------------|------|-------|
| 163.com | smtp.163.com | 465 | Use SSL |
| Gmail | smtp.gmail.com | 587 | Use TLS, requires App Password |
| Outlook | smtp-mail.outlook.com | 587 | Use TLS |
| Yahoo | smtp.mail.yahoo.com | 587 | Use TLS, requires App Password |

**Important Notes:**
- For Gmail and Yahoo, you need to use an [App Password](https://support.google.com/accounts/answer/185833) instead of your regular password
- For 163.com, use port 465 with SSL
- Never commit your `.env` file to version control
- The `.env` file takes precedence over manually set environment variables

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
python src/weekly_report_sender.py
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

## Module Structure

### 1. Configuration Manager (`src/config/config_manager.py`)
- Loads configuration from environment variables and config file
- Validates configuration values
- Provides configuration to other modules

### 2. Excel Reader (`src/excel/excel_reader.py`)
- Reads Excel files using openpyxl
- Handles merged cells
- Preserves cell formatting
- Converts Excel content to HTML

### 3. HTML Generator (`src/html/html_generator.py`)
- Generates HTML content from Excel data
- Adds CSS styles for better presentation

### 4. Email Sender (`src/email/email_sender.py`)
- Creates email messages
- Sends emails via SMTP
- Handles email authentication and delivery

### 5. Main Coordinator (`src/weekly_report_sender.py`)
- Coordinates the different components
- Executes the main workflow
- Handles error logging and reporting

## Testing

The project includes comprehensive test scripts to verify functionality:
- `test/test_system.py` - Tests the complete system functionality
- `test/test_config.py` - Tests configuration management
- `test/test_excel.py` - Tests Excel reading with random row/column additions
- `test/test_html.py` - Tests HTML generation
- `test/test_email.py` - Tests email sending
- `test/test_main.py` - Tests main coordinator

To run the tests:

```bash
# Run all tests
python test/test_system.py

# Run specific tests
python test/test_config.py
python test/test_excel.py
python test/test_html.py
python test/test_email.py
python test/test_main.py
```
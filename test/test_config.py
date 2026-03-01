#!/usr/bin/env python3
"""
Test script for configuration management functionality.
"""

import os
import sys
import tempfile

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

from src.config.config_manager import config_manager

print("Testing configuration management...")
print("=" * 60)

# Test 1: Configuration loading
try:
    print("\nTest 1: Configuration loading")
    print("-" * 40)
    
    # Get email configuration
    email_config = config_manager.get_email_config()
    print("✓ Email configuration loaded successfully")
    print(f"  Sender email: {email_config['sender_email']}")
    print(f"  SMTP server: {email_config['smtp_server']}")
    print(f"  SMTP port: {email_config['smtp_port']}")
    
    # Get recipients configuration
    recipients_config = config_manager.get_recipients_config()
    print("✓ Recipients configuration loaded successfully")
    print(f"  To emails: {recipients_config['to_emails']}")
    print(f"  CC emails: {recipients_config['cc_emails']}")
    
    # Get file configuration
    file_config = config_manager.get_file_config()
    print("✓ File configuration loaded successfully")
    print(f"  Excel file path: {file_config['excel_file_path']}")
    print(f"  Subject: {file_config['subject']}")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

# Test 2: Environment variable support
try:
    print("\nTest 2: Environment variable support")
    print("-" * 40)
    
    # Test with temporary environment variables
    with tempfile.NamedTemporaryFile(mode='w', suffix='.env', delete=False) as f:
        f.write("SENDER_EMAIL=test@example.com\n")
        f.write("SMTP_SERVER=smtp.example.com\n")
        f.write("SMTP_PORT=587\n")
        env_file = f.name
    
    # Set environment variable for dotenv
    os.environ['DOTENV_FILE'] = env_file
    
    # Reload configuration
    from src.config.config_manager import ConfigManager
    temp_config = ConfigManager()
    email_config = temp_config.get_email_config()
    print("✓ Environment variables loaded successfully")
    print(f"  Sender email: {email_config['sender_email']}")
    print(f"  SMTP server: {email_config['smtp_server']}")
    print(f"  SMTP port: {email_config['smtp_port']}")
    
    # Clean up
    os.unlink(env_file)
    del os.environ['DOTENV_FILE']
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Configuration tests completed!")
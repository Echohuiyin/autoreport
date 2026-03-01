#!/usr/bin/env python3
"""
Test script for email sending functionality.
"""

import os
import sys

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

from src.email.email_sender import EmailSender
from src.config.config_manager import config_manager

print("Testing email sending functionality...")
print("=" * 60)

# Test 1: Email sender initialization
try:
    print("\nTest 1: Email sender initialization")
    print("-" * 40)
    
    # Get configurations
    email_config = config_manager.get_email_config()
    recipients_config = config_manager.get_recipients_config()
    
    # Create email sender instance
    sender = EmailSender(email_config, recipients_config)
    print("✓ EmailSender instance created successfully")
    print(f"  Sender email: {sender.sender_email}")
    print(f"  SMTP server: {sender.smtp_server}")
    print(f"  SMTP port: {sender.smtp_port}")
    print(f"  To emails: {sender.to_emails}")
    print(f"  CC emails: {sender.cc_emails}")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

# Test 2: Email message creation
try:
    print("\nTest 2: Email message creation")
    print("-" * 40)
    
    # Get configurations
    email_config = config_manager.get_email_config()
    recipients_config = config_manager.get_recipients_config()
    
    # Create email sender instance
    sender = EmailSender(email_config, recipients_config)
    
    # Create test HTML content
    html_content = "<h1>Test Email</h1><p>This is a test email.</p>"
    
    # Create email message
    msg = sender.create_email_message("Test Subject", html_content)
    print("✓ Email message created successfully")
    print(f"  From: {msg['From']}")
    print(f"  To: {msg['To']}")
    if msg['Cc']:
        print(f"  CC: {msg['Cc']}")
    print(f"  Subject: {msg['Subject']}")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Email tests completed!")
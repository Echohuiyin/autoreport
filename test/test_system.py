#!/usr/bin/env python3
"""
Test script to verify the functionality of the refactored weekly report system.
"""

import os
import sys

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath('.'))

from src.weekly_report_sender import WeeklyReportSender

print("Testing refactored weekly report system...")
print("=" * 60)

# Test 1: Excel file reading
try:
    print("\nTest 1: Excel file reading")
    print("-" * 40)
    
    # Create a sender instance
    sender = WeeklyReportSender()
    print("✓ WeeklyReportSender instance created successfully")
    
    # Read Excel content
    html_content = sender.read_excel_content()
    print("✓ Excel content read successfully")
    print(f"  HTML length: {len(html_content)} characters")
    
    # Check if key columns are present
    if '项目' in html_content and '名称' in html_content and '进展' in html_content:
        print("✓ Key columns found in HTML output")
    else:
        print("✗ Key columns not found in HTML output")
        
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

# Test 2: Email message creation
try:
    print("\nTest 2: Email message creation")
    print("-" * 40)
    
    # Create a sender instance
    sender = WeeklyReportSender()
    
    # Read Excel content
    html_content = sender.read_excel_content()
    
    # Create email message
    msg = sender.create_email_message(html_content)
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

# Test 3: Configuration validation
try:
    print("\nTest 3: Configuration validation")
    print("-" * 40)
    
    # Create a sender instance
    sender = WeeklyReportSender()
    
    # Validate configuration
    sender.validate_config()
    print("✓ Configuration validated successfully")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Test completed!")
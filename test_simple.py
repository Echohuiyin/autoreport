#!/usr/bin/env python3
"""
Simple test script to check if the WeeklyReportSender is working.
"""

from weekly_report_sender import WeeklyReportSender

print("Starting test...")

try:
    print("Creating WeeklyReportSender instance...")
    sender = WeeklyReportSender()
    print("Instance created successfully")
    
    print("Reading Excel content...")
    html_content = sender.read_excel_content()
    print("Excel content read successfully")
    
    print(f"HTML length: {len(html_content)}")
    print("First 500 characters:")
    print(html_content[:500])
    
    if '备注' in html_content:
        print("✓ '备注' column found")
    else:
        print("✗ '备注' column not found")
        
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()

print("Test completed!")
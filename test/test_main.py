#!/usr/bin/env python3
"""
Test script for main coordinator functionality.
"""

import os
import sys

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

from src.weekly_report_sender import WeeklyReportSender

print("Testing main coordinator functionality...")
print("=" * 60)

# Test 1: Main coordinator initialization
try:
    print("\nTest 1: Main coordinator initialization")
    print("-" * 40)
    
    # Create main coordinator instance
    sender = WeeklyReportSender()
    print("✓ WeeklyReportSender instance created successfully")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

# Test 2: Full workflow
try:
    print("\nTest 2: Full workflow")
    print("-" * 40)
    
    # Create main coordinator instance
    sender = WeeklyReportSender()
    
    # Run the full workflow
    print("  Running full workflow...")
    sender.run()
    print("✓ Full workflow completed successfully")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Main coordinator tests completed!")
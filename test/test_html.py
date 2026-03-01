#!/usr/bin/env python3
"""
Test script for HTML generation functionality.
"""

import os
import sys

# Add the project root directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

from src.html.html_generator import HtmlGenerator
from src.excel.excel_reader import ExcelReader

print("Testing HTML generation functionality...")
print("=" * 60)

# Test 1: Basic HTML generation
try:
    print("\nTest 1: Basic HTML generation")
    print("-" * 40)
    
    # Create Excel reader instance
    reader = ExcelReader("weekly_report.xlsx")
    print("✓ ExcelReader instance created successfully")
    
    # Read Excel content and generate HTML
    html_content = reader.read_excel_content()
    print("✓ HTML generated successfully")
    print(f"  HTML length: {len(html_content)} characters")
    
    # Check if HTML structure is correct
    if '<table' in html_content and '</table>' in html_content:
        print("✓ HTML table structure found")
    else:
        print("✗ HTML table structure not found")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

# Test 2: HTML with merged cells
try:
    print("\nTest 2: HTML with merged cells")
    print("-" * 40)
    
    # Read Excel content with merged cells
    reader = ExcelReader("weekly_report.xlsx")
    html_content = reader.read_excel_content()
    print("✓ HTML with merged cells generated successfully")
    
    # Check if merged cells are present
    if 'colspan' in html_content or 'rowspan' in html_content:
        print("✓ Merged cells found in HTML")
    else:
        print("✗ Merged cells not found in HTML")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("HTML tests completed!")
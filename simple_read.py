#!/usr/bin/env python3
"""
Simple script to read Excel file and print content.
"""

import os
import pandas as pd

# Check if the Excel file exists
excel_file = 'weekly_report.xlsx'
if os.path.exists(excel_file):
    print(f"Excel file found: {excel_file}")
    
    # Try to read with pandas
    try:
        print("Reading Excel file with pandas...")
        df = pd.read_excel(excel_file, header=1, engine='openpyxl')
        print(f"Successfully read Excel file")
        print(f"DataFrame shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print("\nData:")
        print(df)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"Excel file not found: {excel_file}")
#!/usr/bin/env python3
"""
Script to read Excel file and write content to a file.
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
        
        # Write to file
        with open('excel_output.txt', 'w', encoding='utf-8') as f:
            f.write(f"DataFrame shape: {df.shape}\n")
            f.write(f"Columns: {list(df.columns)}\n\n")
            f.write("Data:\n")
            f.write(df.to_string())
        print("Excel content written to excel_output.txt")
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        import traceback
        traceback.print_exc()
        # Write error to file
        with open('error_log.txt', 'w', encoding='utf-8') as f:
            f.write(f"Error reading Excel file: {e}\n")
            f.write(traceback.format_exc())
else:
    print(f"Excel file not found: {excel_file}")
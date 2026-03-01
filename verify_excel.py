#!/usr/bin/env python3
"""
Simple script to verify Excel file can be read and write content to a text file.
"""

import os
import pandas as pd
import openpyxl

print("Current directory:", os.getcwd())
print("Files in directory:", os.listdir('.'))

# Check if the Excel file exists
excel_file = 'weekly_report.xlsx'
if os.path.exists(excel_file):
    print(f"Excel file found: {excel_file}")
    
    # Try to read with pandas
    try:
        df = pd.read_excel(excel_file, header=1, engine='openpyxl')
        print(f"Successfully read Excel file with pandas")
        print(f"DataFrame shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        # Write content to a text file
        with open('excel_content.txt', 'w', encoding='utf-8') as f:
            f.write(f"DataFrame shape: {df.shape}\n")
            f.write(f"Columns: {list(df.columns)}\n\n")
            f.write("Data:\n")
            f.write(df.to_string())
        
        print("Excel content written to excel_content.txt")
        
    except Exception as e:
        print(f"Error reading Excel file with pandas: {e}")
else:
    print(f"Excel file not found: {excel_file}")
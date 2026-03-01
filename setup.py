#!/usr/bin/env python3
"""
Setup script to install required dependencies for the weekly report system.
"""

import subprocess
import sys

def install_requirements():
    """Install required packages from requirements.txt."""
    try:
        print("Installing required dependencies...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("Dependencies installed successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error installing dependencies: {e}")
        print("Please install manually using: pip install -r requirements.txt")
        return False
    except FileNotFoundError:
        print("pip not found. Please ensure Python and pip are installed.")
        return False
    return True

def test_imports():
    """Test if required modules can be imported."""
    try:
        import pandas as pd
        import openpyxl
        print("All required modules imported successfully!")
        return True
    except ImportError as e:
        print(f"Import error: {e}")
        return False

if __name__ == "__main__":
    if install_requirements():
        test_imports()
    else:
        print("Setup failed. Please check your Python environment.")
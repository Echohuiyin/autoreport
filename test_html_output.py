#!/usr/bin/env python3
"""
Test script to generate HTML output and display it.
"""

from weekly_report_sender import WeeklyReportSender
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_html_generation():
    """
    Test HTML generation with the current Excel file.
    """
    try:
        # Create report sender instance
        sender = WeeklyReportSender()
        
        # Read Excel content and generate HTML
        html_content = sender.read_excel_content()
        
        # Save HTML to file
        with open('email_body.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        logger.info("Excel content processed successfully")
        logger.info(f"HTML output saved to email_body.html")
        
        # Print debug information
        logger.info("\nDebug information:")
        logger.info(f"HTML length: {len(html_content)} characters")
        logger.info(f"First 500 characters of HTML:\n{html_content[:500]}...")
        
        # Check for specific columns
        if '备注' in html_content:
            logger.info("✓ '备注' column found in HTML")
        else:
            logger.warning("✗ '备注' column not found in HTML")
            
    except Exception as e:
        logger.error(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    print("Generating HTML output...")
    test_html_generation()
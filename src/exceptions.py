#!/usr/bin/env python3
"""
Custom exception classes for the weekly report system.
"""


class ReportConfigurationError(Exception):
    """Raised when there's an issue with the configuration."""
    pass


class ExcelParsingError(Exception):
    """Raised when there's an error parsing the Excel file."""
    pass


class EmailDeliveryError(Exception):
    """Raised when there's an error delivering the email."""
    pass
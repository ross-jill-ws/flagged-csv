"""
Excel format string parser and formatter.
"""

import re
from decimal import Decimal
from datetime import datetime, date
from typing import Any, Union
import pandas as pd


class ExcelFormatter:
    """Format values according to Excel number format strings."""
    
    @staticmethod
    def format_value(value: Any, format_string: str) -> str:
        """
        Format a value according to an Excel format string.
        
        Args:
            value: The value to format
            format_string: Excel format string
            
        Returns:
            Formatted string representation
        """
        if value is None:
            return ""
        
        # Handle common formats
        if format_string == "General" or not format_string:
            return str(value)
        
        # Currency formats
        if "$" in format_string or "¥" in format_string or "€" in format_string:
            return ExcelFormatter._format_currency(value, format_string)
        
        # Percentage formats
        if "%" in format_string:
            return ExcelFormatter._format_percentage(value, format_string)
        
        # Date/time formats
        if any(x in format_string.upper() for x in ["Y", "M", "D", "H", "S"]):
            return ExcelFormatter._format_datetime(value, format_string)
        
        # Number formats with thousand separators
        if "#,##" in format_string:
            return ExcelFormatter._format_number_with_separator(value, format_string)
        
        # Fraction formats
        if "/" in format_string and "?" in format_string:
            return ExcelFormatter._format_fraction(value, format_string)
        
        # Default to string representation
        return str(value)
    
    @staticmethod
    def _format_currency(value: Any, format_string: str) -> str:
        """Format as currency."""
        try:
            # Extract currency symbol
            currency_match = re.search(r'([$¥€£])', format_string)
            currency = currency_match.group(1) if currency_match else "$"
            
            # Extract decimal places
            decimal_match = re.search(r'\.(\d+)', format_string)
            decimals = len(decimal_match.group(1)) if decimal_match else 2
            
            # Format the number
            num_value = float(value)
            
            # Check for negative numbers in parentheses
            if num_value < 0 and "(" in format_string:
                return f"({currency}{abs(num_value):,.{decimals}f})"
            
            # Check if currency comes after number
            if format_string.index(currency) > format_string.index("#") if "#" in format_string else 0:
                return f"{num_value:,.{decimals}f}{currency}"
            
            return f"{currency}{num_value:,.{decimals}f}"
            
        except (ValueError, TypeError):
            return str(value)
    
    @staticmethod
    def _format_percentage(value: Any, format_string: str) -> str:
        """Format as percentage."""
        try:
            # Extract decimal places
            decimal_match = re.search(r'\.(\d+)', format_string)
            decimals = len(decimal_match.group(1)) if decimal_match else 0
            
            num_value = float(value) * 100
            return f"{num_value:.{decimals}f}%"
            
        except (ValueError, TypeError):
            return str(value)
    
    @staticmethod
    def _format_datetime(value: Any, format_string: str) -> str:
        """Format as date/time."""
        try:
            # Convert Excel date number to datetime if needed
            if isinstance(value, (int, float)):
                # Excel epoch starts from 1899-12-30
                excel_epoch = datetime(1899, 12, 30)
                value = excel_epoch + pd.Timedelta(days=value)
            
            if not isinstance(value, (datetime, date)):
                return str(value)
            
            # Convert Excel format to Python format
            format_map = {
                'yyyy': '%Y', 'yy': '%y',
                'mmmm': '%B', 'mmm': '%b', 'mm': '%m', 'm': '%-m',
                'dddd': '%A', 'ddd': '%a', 'dd': '%d', 'd': '%-d',
                'hh': '%H', 'h': '%-H',
                'mm': '%M', 'm': '%-M',
                'ss': '%S', 's': '%-S'
            }
            
            py_format = format_string
            for excel, python in sorted(format_map.items(), key=lambda x: -len(x[0])):
                py_format = py_format.replace(excel, python)
            
            return value.strftime(py_format)
            
        except Exception:
            return str(value)
    
    @staticmethod
    def _format_number_with_separator(value: Any, format_string: str) -> str:
        """Format number with thousand separators."""
        try:
            # Extract decimal places
            decimal_match = re.search(r'\.(\d+)', format_string)
            decimals = len(decimal_match.group(1)) if decimal_match else 0
            
            num_value = float(value)
            
            # Check for negative numbers in parentheses
            if num_value < 0 and "(" in format_string:
                return f"({abs(num_value):,.{decimals}f})"
            
            return f"{num_value:,.{decimals}f}"
            
        except (ValueError, TypeError):
            return str(value)
    
    @staticmethod
    def _format_fraction(value: Any, format_string: str) -> str:
        """Format as fraction."""
        try:
            from fractions import Fraction
            
            num_value = float(value)
            frac = Fraction(num_value).limit_denominator(100)
            
            if frac.denominator == 1:
                return str(frac.numerator)
            
            whole = frac.numerator // frac.denominator
            remainder = frac.numerator % frac.denominator
            
            if whole == 0:
                return f"{remainder}/{frac.denominator}"
            else:
                return f"{whole} {remainder}/{frac.denominator}"
                
        except (ValueError, TypeError):
            return str(value)
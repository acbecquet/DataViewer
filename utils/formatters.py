# utils/formatters.py
"""
utils/formatters.py
Data formatting utilities for the DataViewer application.
Provides consistent formatting for numbers, dates, and other data types.
"""

import pandas as pd
import numpy as np
from datetime import datetime, date
from typing import Any, Optional, Union

try:
    from .debug import debug_print, error_print
except ImportError:
    def debug_print(msg): print(f"DEBUG: {msg}")
    def error_print(msg, exc=None): print(f"ERROR: {msg}")

def format_number(value: Any, decimals: int = 2, thousands_sep: str = ',') -> str:
    """
    Format a number with specified decimal places and thousands separator.
    
    Args:
        value: Number to format
        decimals (int): Number of decimal places
        thousands_sep (str): Thousands separator character
        
    Returns:
        str: Formatted number string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        # Convert to float
        num_value = float(value)
        
        # Handle infinity and very large numbers
        if np.isinf(num_value):
            return "∞" if num_value > 0 else "-∞"
        
        # Format with specified decimals
        if decimals == 0:
            formatted = f"{int(num_value):,}".replace(',', thousands_sep)
        else:
            formatted = f"{num_value:,.{decimals}f}".replace(',', thousands_sep)
        
        return formatted
        
    except (ValueError, TypeError, OverflowError) as e:
        debug_print(f"Error formatting number {value}: {e}")
        return str(value) if value is not None else ""

def format_percentage(value: Any, decimals: int = 1) -> str:
    """
    Format a value as a percentage.
    
    Args:
        value: Value to format (0.15 -> 15.0%)
        decimals (int): Number of decimal places
        
    Returns:
        str: Formatted percentage string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        num_value = float(value)
        
        # Handle infinity
        if np.isinf(num_value):
            return "∞%" if num_value > 0 else "-∞%"
        
        # Convert to percentage and format
        percentage = num_value * 100
        return f"{percentage:.{decimals}f}%"
        
    except (ValueError, TypeError, OverflowError) as e:
        debug_print(f"Error formatting percentage {value}: {e}")
        return str(value) if value is not None else ""

def format_scientific(value: Any, precision: int = 2) -> str:
    """
    Format a number in scientific notation.
    
    Args:
        value: Number to format
        precision (int): Number of significant digits
        
    Returns:
        str: Formatted scientific notation string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        num_value = float(value)
        
        # Handle infinity and zero
        if np.isinf(num_value):
            return "∞" if num_value > 0 else "-∞"
        
        if num_value == 0:
            return "0.00e+00"
        
        # Format in scientific notation
        return f"{num_value:.{precision}e}"
        
    except (ValueError, TypeError, OverflowError) as e:
        debug_print(f"Error formatting scientific {value}: {e}")
        return str(value) if value is not None else ""

def format_datetime(value: Any, format_string: str = "%Y-%m-%d %H:%M:%S") -> str:
    """
    Format a datetime value.
    
    Args:
        value: Datetime value to format
        format_string (str): Format string for datetime
        
    Returns:
        str: Formatted datetime string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        # Handle different types of datetime inputs
        if isinstance(value, (datetime, date)):
            dt_value = value
        elif isinstance(value, pd.Timestamp):
            dt_value = value.to_pydatetime()
        elif isinstance(value, str):
            # Try to parse string as datetime
            dt_value = pd.to_datetime(value)
        else:
            # Try to convert to datetime
            dt_value = pd.to_datetime(value)
        
        return dt_value.strftime(format_string)
        
    except (ValueError, TypeError) as e:
        debug_print(f"Error formatting datetime {value}: {e}")
        return str(value) if value is not None else ""

def format_currency(value: Any, currency_symbol: str = "$", decimals: int = 2) -> str:
    """
    Format a value as currency.
    
    Args:
        value: Value to format
        currency_symbol (str): Currency symbol to use
        decimals (int): Number of decimal places
        
    Returns:
        str: Formatted currency string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        num_value = float(value)
        
        # Handle negative values
        if num_value < 0:
            return f"-{currency_symbol}{format_number(abs(num_value), decimals)}"
        else:
            return f"{currency_symbol}{format_number(num_value, decimals)}"
        
    except (ValueError, TypeError) as e:
        debug_print(f"Error formatting currency {value}: {e}")
        return str(value) if value is not None else ""

def format_file_size(size_bytes: Union[int, float]) -> str:
    """
    Format file size in human-readable format.
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        str: Formatted file size string
    """
    try:
        if size_bytes == 0:
            return "0 B"
        
        size_bytes = float(size_bytes)
        
        # Define size units
        units = ['B', 'KB', 'MB', 'GB', 'TB']
        unit_index = 0
        
        while size_bytes >= 1024 and unit_index < len(units) - 1:
            size_bytes /= 1024
            unit_index += 1
        
        # Format with appropriate decimal places
        if unit_index == 0:  # Bytes
            return f"{int(size_bytes)} {units[unit_index]}"
        else:
            return f"{size_bytes:.1f} {units[unit_index]}"
        
    except (ValueError, TypeError) as e:
        debug_print(f"Error formatting file size {size_bytes}: {e}")
        return str(size_bytes) if size_bytes is not None else ""

def format_duration(seconds: Union[int, float]) -> str:
    """
    Format duration in human-readable format.
    
    Args:
        seconds: Duration in seconds
        
    Returns:
        str: Formatted duration string
    """
    try:
        if pd.isna(seconds) or seconds is None:
            return ""
        
        total_seconds = int(float(seconds))
        
        if total_seconds < 60:
            return f"{total_seconds}s"
        elif total_seconds < 3600:  # Less than 1 hour
            minutes = total_seconds // 60
            remaining_seconds = total_seconds % 60
            return f"{minutes}m {remaining_seconds}s"
        else:  # 1 hour or more
            hours = total_seconds // 3600
            remaining_minutes = (total_seconds % 3600) // 60
            remaining_seconds = total_seconds % 60
            return f"{hours}h {remaining_minutes}m {remaining_seconds}s"
        
    except (ValueError, TypeError) as e:
        debug_print(f"Error formatting duration {seconds}: {e}")
        return str(seconds) if seconds is not None else ""

def format_table_cell(value: Any, max_width: int = 20) -> str:
    """
    Format a cell value for table display with width truncation.
    
    Args:
        value: Cell value to format
        max_width (int): Maximum width for cell content
        
    Returns:
        str: Formatted cell string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        # Convert to string
        str_value = str(value)
        
        # Truncate if too long
        if len(str_value) > max_width:
            str_value = str_value[:max_width-3] + "..."
        
        return str_value
        
    except Exception as e:
        debug_print(f"Error formatting table cell {value}: {e}")
        return str(value) if value is not None else ""

def format_dataframe_for_display(df: pd.DataFrame, 
                                max_rows: int = 100, 
                                max_cols: int = 20,
                                cell_width: int = 15) -> pd.DataFrame:
    """
    Format a DataFrame for display with size and content limits.
    
    Args:
        df (pd.DataFrame): DataFrame to format
        max_rows (int): Maximum number of rows to display
        max_cols (int): Maximum number of columns to display
        cell_width (int): Maximum width for cell content
        
    Returns:
        pd.DataFrame: Formatted DataFrame
    """
    try:
        if df.empty:
            return df
        
        # Limit rows and columns
        display_df = df.iloc[:max_rows, :max_cols].copy()
        
        # Format cell contents
        for col in display_df.columns:
            if display_df[col].dtype == 'object':
                display_df[col] = display_df[col].apply(
                    lambda x: format_table_cell(x, cell_width)
                )
            elif np.issubdtype(display_df[col].dtype, np.number):
                display_df[col] = display_df[col].apply(
                    lambda x: format_number(x, 2) if not pd.isna(x) else ""
                )
        
        debug_print(f"Formatted DataFrame for display: {display_df.shape}")
        return display_df
        
    except Exception as e:
        error_print(f"Error formatting DataFrame for display: {e}")
        return df

def auto_format_value(value: Any, column_name: str = "") -> str:
    """
    Automatically format a value based on its type and column name.
    
    Args:
        value: Value to format
        column_name (str): Column name for context-aware formatting
        
    Returns:
        str: Formatted value string
    """
    try:
        if pd.isna(value) or value is None:
            return ""
        
        column_lower = column_name.lower()
        
        # Check for percentage columns
        if any(keyword in column_lower for keyword in ['percent', '%', 'efficiency', 'rate']):
            if isinstance(value, (int, float)) and 0 <= value <= 1:
                return format_percentage(value)
            elif isinstance(value, (int, float)) and value > 1:
                return format_percentage(value / 100)
        
        # Check for currency columns
        if any(keyword in column_lower for keyword in ['price', 'cost', 'amount', '$']):
            return format_currency(value)
        
        # Check for date columns
        if any(keyword in column_lower for keyword in ['date', 'time', 'created', 'updated']):
            return format_datetime(value)
        
        # Check for size columns
        if any(keyword in column_lower for keyword in ['size', 'bytes', 'length']):
            if isinstance(value, (int, float)) and value > 1000:
                return format_file_size(value)
        
        # Default numeric formatting
        if isinstance(value, (int, float)):
            if abs(value) >= 1000000:
                return format_scientific(value)
            else:
                return format_number(value)
        
        # Default string formatting
        return format_table_cell(value)
        
    except Exception as e:
        debug_print(f"Error auto-formatting value {value}: {e}")
        return str(value) if value is not None else ""

print("DEBUG: formatters.py - Formatting functions loaded successfully")
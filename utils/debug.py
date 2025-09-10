# utils/formatters.py
"""
utils/formatters.py
Data formatting utilities for the DataViewer application.
"""

from typing import Any, Optional, Union
import pandas as pd
import numpy as np
from datetime import datetime, date
from decimal import Decimal


def format_number(value: Any, decimals: int = 2, thousands_sep: bool = True) -> str:
    """Format a number with specified decimal places and optional thousands separator."""
    try:
        if pd.isna(value) or value is None:
            return "N/A"
        
        # Convert to float
        num_val = float(value)
        
        # Handle infinity and NaN
        if np.isinf(num_val):
            return "∞" if num_val > 0 else "-∞"
        
        if np.isnan(num_val):
            return "NaN"
        
        # Format with specified decimals
        if thousands_sep:
            formatted = f"{num_val:,.{decimals}f}"
        else:
            formatted = f"{num_val:.{decimals}f}"
        
        print(f"DEBUG: format_number({value}) -> {formatted}")
        return formatted
        
    except (ValueError, TypeError, OverflowError):
        print(f"WARNING: format_number failed for value: {value}")
        return str(value)


def format_percentage(value: Any, decimals: int = 1) -> str:
    """Format a value as a percentage."""
    try:
        if pd.isna(value) or value is None:
            return "N/A"
        
        # Convert to float and multiply by 100
        num_val = float(value) * 100
        
        # Handle infinity and NaN
        if np.isinf(num_val):
            return "∞%" if num_val > 0 else "-∞%"
        
        if np.isnan(num_val):
            return "NaN%"
        
        formatted = f"{num_val:.{decimals}f}%"
        
        print(f"DEBUG: format_percentage({value}) -> {formatted}")
        return formatted
        
    except (ValueError, TypeError, OverflowError):
        print(f"WARNING: format_percentage failed for value: {value}")
        return f"{value}%"


def format_scientific(value: Any, precision: int = 2) -> str:
    """Format a number in scientific notation."""
    try:
        if pd.isna(value) or value is None:
            return "N/A"
        
        # Convert to float
        num_val = float(value)
        
        # Handle infinity and NaN
        if np.isinf(num_val):
            return "∞" if num_val > 0 else "-∞"
        
        if np.isnan(num_val):
            return "NaN"
        
        # Handle zero
        if num_val == 0:
            return "0.00E+00"
        
        formatted = f"{num_val:.{precision}E}"
        
        print(f"DEBUG: format_scientific({value}) -> {formatted}")
        return formatted
        
    except (ValueError, TypeError, OverflowError):
        print(f"WARNING: format_scientific failed for value: {value}")
        return str(value)


def format_datetime(dt: Any, format_string: str = "%Y-%m-%d %H:%M:%S") -> str:
    """Format a datetime object to string."""
    try:
        if pd.isna(dt) or dt is None:
            return "N/A"
        
        # Handle different datetime types
        if isinstance(dt, str):
            # Try to parse string to datetime
            try:
                dt = pd.to_datetime(dt)
            except:
                return dt  # Return original string if parsing fails
        
        if isinstance(dt, (datetime, date, pd.Timestamp)):
            formatted = dt.strftime(format_string)
            print(f"DEBUG: format_datetime({dt}) -> {formatted}")
            return formatted
        
        # Try pandas datetime conversion
        try:
            pd_dt = pd.to_datetime(dt)
            formatted = pd_dt.strftime(format_string)
            print(f"DEBUG: format_datetime({dt}) -> {formatted} (via pandas)")
            return formatted
        except:
            pass
        
        # Fallback to string representation
        return str(dt)
        
    except Exception as e:
        print(f"WARNING: format_datetime failed for value {dt}: {e}")
        return str(dt)


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    try:
        if size_bytes < 0:
            return "Invalid size"
        
        if size_bytes == 0:
            return "0 B"
        
        # Define size units
        units = ['B', 'KB', 'MB', 'GB', 'TB']
        
        # Calculate appropriate unit
        for i, unit in enumerate(units):
            if size_bytes < 1024 or i == len(units) - 1:
                if i == 0:
                    return f"{size_bytes} {unit}"
                else:
                    return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        
    except (ValueError, TypeError):
        print(f"WARNING: format_file_size failed for value: {size_bytes}")
        return "Unknown size"


def format_duration(seconds: float) -> str:
    """Format duration in seconds to human-readable format."""
    try:
        if pd.isna(seconds) or seconds is None:
            return "N/A"
        
        total_seconds = int(float(seconds))
        
        if total_seconds < 0:
            return "Invalid duration"
        
        if total_seconds < 60:
            return f"{total_seconds}s"
        
        minutes = total_seconds // 60
        remaining_seconds = total_seconds % 60
        
        if minutes < 60:
            return f"{minutes}m {remaining_seconds}s"
        
        hours = minutes // 60
        remaining_minutes = minutes % 60
        
        if hours < 24:
            return f"{hours}h {remaining_minutes}m {remaining_seconds}s"
        
        days = hours // 24
        remaining_hours = hours % 24
        
        return f"{days}d {remaining_hours}h {remaining_minutes}m"
        
    except (ValueError, TypeError):
        print(f"WARNING: format_duration failed for value: {seconds}")
        return str(seconds)


def format_temperature(value: Any, unit: str = "C", decimals: int = 1) -> str:
    """Format temperature with unit."""
    try:
        if pd.isna(value) or value is None:
            return "N/A"
        
        temp_val = float(value)
        
        # Handle infinity and NaN
        if np.isinf(temp_val):
            return "∞°" + unit if temp_val > 0 else "-∞°" + unit
        
        if np.isnan(temp_val):
            return "NaN°" + unit
        
        formatted = f"{temp_val:.{decimals}f}°{unit}"
        
        print(f"DEBUG: format_temperature({value}) -> {formatted}")
        return formatted
        
    except (ValueError, TypeError):
        print(f"WARNING: format_temperature failed for value: {value}")
        return f"{value}°{unit}"


def format_viscosity(value: Any, unit: str = "Pa·s", decimals: int = 3) -> str:
    """Format viscosity value with unit."""
    try:
        if pd.isna(value) or value is None:
            return "N/A"
        
        visc_val = float(value)
        
        # Handle infinity and NaN
        if np.isinf(visc_val):
            return "∞ " + unit if visc_val > 0 else "-∞ " + unit
        
        if np.isnan(visc_val):
            return "NaN " + unit
        
        # Use scientific notation for very small or large values
        if visc_val < 0.001 or visc_val > 1000:
            formatted = f"{visc_val:.{decimals}E} {unit}"
        else:
            formatted = f"{visc_val:.{decimals}f} {unit}"
        
        print(f"DEBUG: format_viscosity({value}) -> {formatted}")
        return formatted
        
    except (ValueError, TypeError):
        print(f"WARNING: format_viscosity failed for value: {value}")
        return f"{value} {unit}"


def format_data_frame_for_display(df: pd.DataFrame, max_rows: int = 100, max_cols: int = 20) -> pd.DataFrame:
    """Format DataFrame for display in UI components."""
    try:
        if df.empty:
            return df
        
        # Limit size for display
        display_df = df.copy()
        
        if len(display_df) > max_rows:
            display_df = display_df.head(max_rows)
            print(f"DEBUG: Truncated DataFrame to {max_rows} rows for display")
        
        if len(display_df.columns) > max_cols:
            display_df = display_df.iloc[:, :max_cols]
            print(f"DEBUG: Truncated DataFrame to {max_cols} columns for display")
        
        # Format numeric columns
        for col in display_df.columns:
            if pd.api.types.is_numeric_dtype(display_df[col]):
                display_df[col] = display_df[col].apply(lambda x: format_number(x, decimals=3))
            elif pd.api.types.is_datetime64_any_dtype(display_df[col]):
                display_df[col] = display_df[col].apply(lambda x: format_datetime(x))
        
        print(f"DEBUG: Formatted DataFrame for display: {display_df.shape}")
        return display_df
        
    except Exception as e:
        print(f"ERROR: format_data_frame_for_display failed: {e}")
        return df


def format_model_summary(model_info: dict) -> str:
    """Format model information for display."""
    try:
        if not model_info:
            return "No model information available"
        
        summary_lines = []
        
        # Model name and type
        name = model_info.get('name', 'Unknown')
        model_type = model_info.get('type', 'Unknown')
        summary_lines.append(f"Model: {name} ({model_type})")
        
        # Training information
        if 'training_size' in model_info:
            summary_lines.append(f"Training data: {model_info['training_size']} points")
        
        # Accuracy
        if 'r2_score' in model_info:
            r2 = format_number(model_info['r2_score'], decimals=4)
            summary_lines.append(f"R² score: {r2}")
        
        # Model parameters
        if model_type == 'polynomial' and 'degree' in model_info:
            summary_lines.append(f"Degree: {model_info['degree']}")
        
        if model_type == 'arrhenius':
            if 'A' in model_info:
                A = format_scientific(model_info['A'])
                summary_lines.append(f"Pre-exponential factor: {A}")
            if 'E_over_R' in model_info:
                E_R = format_number(model_info['E_over_R'], decimals=1)
                summary_lines.append(f"E/R: {E_R} K")
        
        # Training time
        if 'trained_at' in model_info:
            trained_at = format_datetime(model_info['trained_at'])
            summary_lines.append(f"Trained: {trained_at}")
        
        formatted_summary = '\n'.join(summary_lines)
        print(f"DEBUG: Formatted model summary for {name}")
        return formatted_summary
        
    except Exception as e:
        print(f"ERROR: format_model_summary failed: {e}")
        return "Error formatting model summary"
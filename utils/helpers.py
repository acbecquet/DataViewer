# utils/helpers.py
"""
utils/helpers.py
General helper functions.
This will contain utility functions from the existing utils.py.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from typing import Optional, Any, Union
import pandas as pd


def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        print(f"DEBUG: Using PyInstaller resource path: {base_path}")
    except AttributeError:
        # Running in normal Python environment
        base_path = os.path.abspath(".")
        print(f"DEBUG: Using development resource path: {base_path}")
    
    full_path = os.path.join(base_path, relative_path)
    print(f"DEBUG: Resource path resolved: {relative_path} -> {full_path}")
    return full_path


def clean_columns(data: pd.DataFrame) -> pd.DataFrame:
    """Clean column names in a DataFrame."""
    if data.empty:
        print("DEBUG: clean_columns - DataFrame is empty")
        return data
    
    original_columns = list(data.columns)
    
    # Clean column names
    cleaned_data = data.copy()
    cleaned_data.columns = [str(col).strip().replace('\n', ' ').replace('\r', ' ') for col in cleaned_data.columns]
    
    # Remove unnamed columns
    cleaned_data = cleaned_data.loc[:, ~cleaned_data.columns.str.contains('^Unnamed')]
    
    print(f"DEBUG: clean_columns - {len(original_columns)} -> {len(cleaned_data.columns)} columns")
    return cleaned_data


def get_save_path(file_path: str, suffix: str = "_processed") -> str:
    """Generate a save path with suffix."""
    path = Path(file_path)
    stem = path.stem
    extension = path.suffix
    parent = path.parent
    
    save_path = parent / f"{stem}{suffix}{extension}"
    
    print(f"DEBUG: Generated save path: {file_path} -> {save_path}")
    return str(save_path)


def is_standard_file(file_path: str) -> bool:
    """Check if a file follows the standard testing format."""
    try:
        # Simple heuristic - check if file has expected sheets/structure
        if file_path.lower().endswith(('.xlsx', '.xls')):
            # Could check for specific sheet names or patterns
            # For now, assume Excel files are standard
            result = True
        else:
            result = False
        
        print(f"DEBUG: is_standard_file({file_path}): {result}")
        return result
        
    except Exception as e:
        print(f"ERROR: is_standard_file failed for {file_path}: {e}")
        return False


def plotting_sheet_test(sheet_name: str) -> bool:
    """Test if a sheet should have plotting capabilities."""
    if not sheet_name:
        return False
    
    # Keywords that indicate a plotting sheet
    plotting_keywords = [
        'test', 'data', 'measurement', 'results', 'analysis',
        'tpm', 'resistance', 'pressure', 'efficiency'
    ]
    
    sheet_lower = sheet_name.lower()
    is_plotting = any(keyword in sheet_lower for keyword in plotting_keywords)
    
    print(f"DEBUG: plotting_sheet_test({sheet_name}): {is_plotting}")
    return is_plotting


def wrap_text(text: str, width: int = 50) -> str:
    """Wrap text to specified width."""
    if not text or width <= 0:
        return text
    
    words = text.split()
    lines = []
    current_line = []
    current_length = 0
    
    for word in words:
        if current_length + len(word) + len(current_line) <= width:
            current_line.append(word)
            current_length += len(word)
        else:
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]
            current_length = len(word)
    
    if current_line:
        lines.append(' '.join(current_line))
    
    wrapped = '\n'.join(lines)
    print(f"DEBUG: wrap_text - {len(text)} chars -> {len(lines)} lines")
    return wrapped


def center_window(window: tk.Toplevel, width: int = None, height: int = None):
    """Center a window on screen."""
    try:
        window.update_idletasks()
        
        # Get window dimensions
        if width and height:
            w, h = width, height
        else:
            w = window.winfo_width()
            h = window.winfo_height()
        
        # Get screen dimensions
        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()
        
        # Calculate center position
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        
        # Set window geometry
        window.geometry(f"{w}x{h}+{x}+{y}")
        
        print(f"DEBUG: center_window - {w}x{h} at ({x},{y}) on {screen_w}x{screen_h} screen")
        
    except Exception as e:
        print(f"ERROR: center_window failed: {e}")


def show_success_message(title: str, message: str, parent: tk.Tk = None):
    """Show a success message dialog."""
    try:
        messagebox.showinfo(title, message, parent=parent)
        print(f"DEBUG: show_success_message - {title}: {message}")
    except Exception as e:
        print(f"ERROR: show_success_message failed: {e}")


def validate_dataframe(df: pd.DataFrame, min_rows: int = 1, min_cols: int = 1) -> tuple[bool, str]:
    """Validate a DataFrame meets minimum requirements."""
    if df is None:
        return False, "DataFrame is None"
    
    if df.empty:
        return False, "DataFrame is empty"
    
    if len(df) < min_rows:
        return False, f"DataFrame has {len(df)} rows, minimum {min_rows} required"
    
    if len(df.columns) < min_cols:
        return False, f"DataFrame has {len(df.columns)} columns, minimum {min_cols} required"
    
    print(f"DEBUG: validate_dataframe - {len(df)} rows, {len(df.columns)} cols - VALID")
    return True, "DataFrame is valid"


def safe_numeric_conversion(value: Any, default: float = 0.0) -> float:
    """Safely convert a value to numeric, with fallback."""
    try:
        if pd.isna(value):
            return default
        
        # Try direct float conversion
        return float(value)
        
    except (ValueError, TypeError):
        try:
            # Try pandas numeric conversion
            result = pd.to_numeric(value, errors='coerce')
            return float(result) if not pd.isna(result) else default
        except:
            return default


def ensure_directory_exists(directory_path: str) -> bool:
    """Ensure a directory exists, create if necessary."""
    try:
        Path(directory_path).mkdir(parents=True, exist_ok=True)
        print(f"DEBUG: ensure_directory_exists - {directory_path} OK")
        return True
    except Exception as e:
        print(f"ERROR: ensure_directory_exists failed for {directory_path}: {e}")
        return False


def get_file_extension(file_path: str) -> str:
    """Get file extension in lowercase."""
    ext = Path(file_path).suffix.lower()
    print(f"DEBUG: get_file_extension({file_path}): {ext}")
    return ext


def is_file_accessible(file_path: str) -> bool:
    """Check if a file is accessible for reading."""
    try:
        path = Path(file_path)
        if not path.exists():
            print(f"DEBUG: is_file_accessible({file_path}): File does not exist")
            return False
        
        if not path.is_file():
            print(f"DEBUG: is_file_accessible({file_path}): Not a file")
            return False
        
        # Try to open file
        with open(file_path, 'rb') as f:
            f.read(1)  # Read just one byte
        
        print(f"DEBUG: is_file_accessible({file_path}): ACCESSIBLE")
        return True
        
    except Exception as e:
        print(f"DEBUG: is_file_accessible({file_path}): ERROR - {e}")
        return False
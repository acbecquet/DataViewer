# utils/validators.py
"""
utils/validators.py
Data validation functions for the DataViewer application.
"""

import os
import re
from pathlib import Path
from typing import Tuple, List, Any, Optional
import pandas as pd
import numpy as np
from .constants import (
    EXCEL_EXTENSIONS, VAP3_EXTENSIONS, CSV_EXTENSIONS,
    MIN_DATA_ROWS, MIN_DATA_COLUMNS, MAX_FILE_SIZE_MB,
    MAX_COLUMN_NAME_LENGTH, MAX_SHEET_NAME_LENGTH
)


def validate_file_path(file_path: str) -> Tuple[bool, str]:
    """Validate if a file path is valid and accessible."""
    if not file_path:
        return False, "File path is empty"
    
    try:
        path = Path(file_path)
        
        # Check if file exists
        if not path.exists():
            return False, f"File does not exist: {file_path}"
        
        # Check if it's actually a file
        if not path.is_file():
            return False, f"Path is not a file: {file_path}"
        
        # Check file size
        file_size_mb = path.stat().st_size / (1024 * 1024)
        if file_size_mb > MAX_FILE_SIZE_MB:
            return False, f"File too large: {file_size_mb:.1f}MB (max {MAX_FILE_SIZE_MB}MB)"
        
        # Check if file is readable
        try:
            with open(file_path, 'rb') as f:
                f.read(1)  # Try to read one byte
        except PermissionError:
            return False, f"Permission denied: {file_path}"
        except Exception as e:
            return False, f"Cannot read file: {e}"
        
        print(f"DEBUG: validate_file_path({file_path}): VALID ({file_size_mb:.1f}MB)")
        return True, "File path is valid"
        
    except Exception as e:
        error_msg = f"File path validation error: {e}"
        print(f"ERROR: validate_file_path({file_path}): {error_msg}")
        return False, error_msg


def validate_excel_file(file_path: str) -> Tuple[bool, str]:
    """Validate if a file is a valid Excel file."""
    # First validate basic file path
    is_valid, message = validate_file_path(file_path)
    if not is_valid:
        return False, message
    
    try:
        path = Path(file_path)
        
        # Check file extension
        if path.suffix.lower() not in EXCEL_EXTENSIONS:
            return False, f"Not an Excel file: {path.suffix}"
        
        # Try to read the Excel file
        try:
            # Just try to get sheet names without loading data
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            excel_file.close()
            
            if not sheet_names:
                return False, "Excel file contains no sheets"
            
            print(f"DEBUG: validate_excel_file({file_path}): VALID ({len(sheet_names)} sheets)")
            return True, f"Valid Excel file with {len(sheet_names)} sheets"
            
        except Exception as e:
            return False, f"Cannot read Excel file: {e}"
        
    except Exception as e:
        error_msg = f"Excel file validation error: {e}"
        print(f"ERROR: validate_excel_file({file_path}): {error_msg}")
        return False, error_msg


def validate_data_frame(df: pd.DataFrame, min_rows: int = None, min_cols: int = None) -> Tuple[bool, List[str]]:
    """Validate a DataFrame for basic data quality requirements."""
    issues = []
    
    # Use default minimums if not specified
    min_rows = min_rows or MIN_DATA_ROWS
    min_cols = min_cols or MIN_DATA_COLUMNS
    
    try:
        # Check if DataFrame exists
        if df is None:
            issues.append("DataFrame is None")
            return False, issues
        
        # Check if DataFrame is empty
        if df.empty:
            issues.append("DataFrame is empty")
            return False, issues
        
        # Check minimum dimensions
        if len(df) < min_rows:
            issues.append(f"Insufficient rows: {len(df)} < {min_rows}")
        
        if len(df.columns) < min_cols:
            issues.append(f"Insufficient columns: {len(df.columns)} < {min_cols}")
        
        # Check for all-null columns
        null_columns = df.columns[df.isnull().all()].tolist()
        if null_columns:
            issues.append(f"Columns with all null values: {', '.join(null_columns)}")
        
        # Check for duplicate column names
        duplicate_columns = df.columns[df.columns.duplicated()].tolist()
        if duplicate_columns:
            issues.append(f"Duplicate column names: {', '.join(duplicate_columns)}")
        
        # Check for excessively long column names
        long_columns = [col for col in df.columns if len(str(col)) > MAX_COLUMN_NAME_LENGTH]
        if long_columns:
            issues.append(f"Column names too long: {', '.join(long_columns[:3])}{'...' if len(long_columns) > 3 else ''}")
        
        # Check for unusual data types
        object_columns = df.select_dtypes(include=['object']).columns.tolist()
        if len(object_columns) == len(df.columns):
            issues.append("All columns are object type (possible parsing issue)")
        
        # Check for excessive missing data
        missing_percentage = (df.isnull().sum().sum() / (len(df) * len(df.columns))) * 100
        if missing_percentage > 50:
            issues.append(f"High percentage of missing data: {missing_percentage:.1f}%")
        
        is_valid = len(issues) == 0
        print(f"DEBUG: validate_data_frame - {len(df)} rows, {len(df.columns)} cols - {'VALID' if is_valid else 'INVALID'} ({len(issues)} issues)")
        
        return is_valid, issues
        
    except Exception as e:
        error_msg = f"DataFrame validation error: {e}"
        print(f"ERROR: validate_data_frame: {error_msg}")
        return False, [error_msg]


def validate_sheet_data(sheet_data: dict, sheet_name: str) -> Tuple[bool, List[str]]:
    """Validate sheet data dictionary structure."""
    issues = []
    
    try:
        # Check if sheet data exists
        if not sheet_data:
            issues.append(f"Sheet data is empty for '{sheet_name}'")
            return False, issues
        
        # Check for required keys
        required_keys = ['data']
        for key in required_keys:
            if key not in sheet_data:
                issues.append(f"Missing required key '{key}' in sheet '{sheet_name}'")
        
        # Validate DataFrame if present
        if 'data' in sheet_data and isinstance(sheet_data['data'], pd.DataFrame):
            df_valid, df_issues = validate_data_frame(sheet_data['data'])
            if not df_valid:
                issues.extend([f"Sheet '{sheet_name}': {issue}" for issue in df_issues])
        
        # Validate sheet name
        if len(sheet_name) > MAX_SHEET_NAME_LENGTH:
            issues.append(f"Sheet name too long: '{sheet_name}' ({len(sheet_name)} > {MAX_SHEET_NAME_LENGTH})")
        
        # Check for invalid characters in sheet name
        invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
        if any(char in sheet_name for char in invalid_chars):
            issues.append(f"Sheet name contains invalid characters: '{sheet_name}'")
        
        is_valid = len(issues) == 0
        print(f"DEBUG: validate_sheet_data('{sheet_name}') - {'VALID' if is_valid else 'INVALID'} ({len(issues)} issues)")
        
        return is_valid, issues
        
    except Exception as e:
        error_msg = f"Sheet data validation error for '{sheet_name}': {e}"
        print(f"ERROR: validate_sheet_data: {error_msg}")
        return False, [error_msg]


def validate_numeric_data(data: Any, allow_nan: bool = True) -> Tuple[bool, str]:
    """Validate if data can be converted to numeric."""
    try:
        if data is None:
            return False, "Data is None"
        
        if pd.isna(data):
            return allow_nan, "Data is NaN" if not allow_nan else "Data is NaN (allowed)"
        
        # Try to convert to float
        try:
            float_val = float(data)
            
            # Check for infinity
            if np.isinf(float_val):
                return False, "Data is infinite"
            
            print(f"DEBUG: validate_numeric_data({data}): VALID (converted to {float_val})")
            return True, "Valid numeric data"
            
        except (ValueError, TypeError):
            # Try pandas numeric conversion
            numeric_val = pd.to_numeric(data, errors='coerce')
            if pd.isna(numeric_val):
                return False, f"Cannot convert to numeric: {data}"
            
            return True, "Valid numeric data (converted)"
        
    except Exception as e:
        error_msg = f"Numeric validation error: {e}"
        print(f"ERROR: validate_numeric_data({data}): {error_msg}")
        return False, error_msg


def validate_temperature_range(temperature: float) -> Tuple[bool, str]:
    """Validate temperature is within reasonable range."""
    try:
        # Convert to float if needed
        temp = float(temperature)
        
        # Check for reasonable temperature range (celsius)
        min_temp = -273.15  # Absolute zero
        max_temp = 1000.0   # Reasonable upper limit for most applications
        
        if temp < min_temp:
            return False, f"Temperature below absolute zero: {temp}蚓"
        
        if temp > max_temp:
            return False, f"Temperature unreasonably high: {temp}蚓"
        
        print(f"DEBUG: validate_temperature_range({temp}蚓): VALID")
        return True, "Valid temperature"
        
    except (ValueError, TypeError):
        return False, f"Invalid temperature value: {temperature}"


def validate_viscosity_value(viscosity: float) -> Tuple[bool, str]:
    """Validate viscosity value is positive and reasonable."""
    try:
        # Convert to float if needed
        visc = float(viscosity)
        
        # Viscosity must be positive
        if visc <= 0:
            return False, f"Viscosity must be positive: {visc}"
        
        # Check for reasonable range (Pa新)
        min_visc = 1e-6   # Very thin liquids
        max_visc = 1e6    # Very thick materials
        
        if visc < min_visc:
            return False, f"Viscosity unreasonably low: {visc} Pa新"
        
        if visc > max_visc:
            return False, f"Viscosity unreasonably high: {visc} Pa新"
        
        print(f"DEBUG: validate_viscosity_value({visc} Pa新): VALID")
        return True, "Valid viscosity"
        
    except (ValueError, TypeError):
        return False, f"Invalid viscosity value: {viscosity}"


def validate_file_extension(file_path: str, allowed_extensions: List[str]) -> Tuple[bool, str]:
    """Validate file has an allowed extension."""
    try:
        path = Path(file_path)
        extension = path.suffix.lower()
        
        if not extension:
            return False, "File has no extension"
        
        allowed_lower = [ext.lower() for ext in allowed_extensions]
        
        if extension not in allowed_lower:
            return False, f"Extension '{extension}' not in allowed list: {', '.join(allowed_extensions)}"
        
        print(f"DEBUG: validate_file_extension({file_path}): VALID ({extension})")
        return True, f"Valid extension: {extension}"
        
    except Exception as e:
        error_msg = f"Extension validation error: {e}"
        print(f"ERROR: validate_file_extension({file_path}): {error_msg}")
        return False, error_msg


def validate_model_name(model_name: str) -> Tuple[bool, str]:
    """Validate model name follows naming conventions."""
    try:
        if not model_name:
            return False, "Model name is empty"
        
        if len(model_name) < 3:
            return False, "Model name too short (minimum 3 characters)"
        
        if len(model_name) > 50:
            return False, "Model name too long (maximum 50 characters)"
        
        # Check for valid characters (alphanumeric, underscore, hyphen)
        if not re.match(r'^[a-zA-Z0-9_-]+$', model_name):
            return False, "Model name contains invalid characters (use only letters, numbers, underscore, hyphen)"
        
        # Must start with letter
        if not model_name[0].isalpha():
            return False, "Model name must start with a letter"
        
        print(f"DEBUG: validate_model_name('{model_name}'): VALID")
        return True, "Valid model name"
        
    except Exception as e:
        error_msg = f"Model name validation error: {e}"
        print(f"ERROR: validate_model_name('{model_name}'): {error_msg}")
        return False, error_msg
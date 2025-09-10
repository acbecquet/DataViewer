# utils/validators.py
"""
utils/validators.py
Validation functions for the DataViewer application.
Consolidates all validation logic from refactor_files.
"""

import os
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Optional, List, Tuple, Any, Union

try:
    from .debug import debug_print, error_print
except ImportError:
    def debug_print(msg): print(f"DEBUG: {msg}")
    def error_print(msg, exc=None): print(f"ERROR: {msg}")

def validate_file_path(file_path: str) -> Tuple[bool, str]:
    """
    Validate if a file path exists and is accessible.
    
    Args:
        file_path (str): Path to validate
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        if not file_path:
            return False, "File path is empty"
        
        path = Path(file_path)
        
        if not path.exists():
            return False, f"File does not exist: {file_path}"
        
        if not path.is_file():
            return False, f"Path is not a file: {file_path}"
        
        if not os.access(file_path, os.R_OK):
            return False, f"File is not readable: {file_path}"
        
        debug_print(f"File path validation passed: {file_path}")
        return True, "Valid file path"
        
    except Exception as e:
        error_msg = f"Error validating file path {file_path}: {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_excel_file(file_path: str) -> Tuple[bool, str]:
    """
    Validate if a file is a valid Excel file.
    
    Args:
        file_path (str): Path to Excel file
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        # First validate the file path
        path_valid, path_error = validate_file_path(file_path)
        if not path_valid:
            return False, path_error
        
        # Check file extension
        ext = Path(file_path).suffix.lower()
        if ext not in ['.xlsx', '.xls']:
            return False, f"Not an Excel file: {file_path} (extension: {ext})"
        
        # Check for temporary Excel files
        filename = Path(file_path).name
        if filename.startswith('~$'):
            return False, f"Temporary Excel file: {file_path}"
        
        # Try to load the file to verify it's a valid Excel file
        try:
            sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            if not sheets:
                return False, f"Excel file contains no sheets: {file_path}"
            
            debug_print(f"Excel file validation passed: {file_path} ({len(sheets)} sheets)")
            return True, "Valid Excel file"
            
        except Exception as excel_error:
            return False, f"Cannot read Excel file {file_path}: {str(excel_error)}"
        
    except Exception as e:
        error_msg = f"Error validating Excel file {file_path}: {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_data_frame(df: pd.DataFrame, name: str = "DataFrame") -> Tuple[bool, str]:
    """
    Validate a pandas DataFrame for common issues.
    
    Args:
        df (pd.DataFrame): DataFrame to validate
        name (str): Name for error messages
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        if df is None:
            return False, f"{name} is None"
        
        if not isinstance(df, pd.DataFrame):
            return False, f"{name} is not a pandas DataFrame (type: {type(df)})"
        
        if df.empty:
            return False, f"{name} is empty"
        
        # Check for basic structure
        if len(df.columns) == 0:
            return False, f"{name} has no columns"
        
        if len(df) == 0:
            return False, f"{name} has no rows"
        
        # Check for all-NaN columns
        all_nan_cols = df.columns[df.isnull().all()].tolist()
        if all_nan_cols:
            debug_print(f"{name} has all-NaN columns: {all_nan_cols}")
        
        # Check for duplicate column names
        duplicate_cols = df.columns[df.columns.duplicated()].tolist()
        if duplicate_cols:
            debug_print(f"{name} has duplicate column names: {duplicate_cols}")
        
        debug_print(f"{name} validation passed: {df.shape}")
        return True, f"Valid {name}"
        
    except Exception as e:
        error_msg = f"Error validating {name}: {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_sheet_data(data: pd.DataFrame, 
                       required_columns: Optional[List[str]] = None, 
                       required_rows: Optional[int] = None,
                       sheet_name: str = "Sheet") -> Tuple[bool, str]:
    """
    Validate sheet data for specific requirements.
    
    Args:
        data (pd.DataFrame): Sheet data to validate
        required_columns (List[str], optional): Required column names
        required_rows (int, optional): Minimum number of rows
        sheet_name (str): Sheet name for error messages
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        # First validate the basic DataFrame
        df_valid, df_error = validate_data_frame(data, f"Sheet '{sheet_name}'")
        if not df_valid:
            return False, df_error
        
        # Check required columns
        if required_columns:
            missing_columns = []
            for col in required_columns:
                if col not in data.columns:
                    missing_columns.append(col)
            
            if missing_columns:
                return False, f"Sheet '{sheet_name}' missing required columns: {missing_columns}"
        
        # Check minimum rows
        if required_rows and len(data) < required_rows:
            return False, f"Sheet '{sheet_name}' has {len(data)} rows, minimum {required_rows} required"
        
        # Check for data quality issues
        total_cells = data.size
        null_cells = data.isnull().sum().sum()
        null_percentage = (null_cells / total_cells) * 100 if total_cells > 0 else 0
        
        if null_percentage > 90:
            debug_print(f"Sheet '{sheet_name}' has {null_percentage:.1f}% null values")
        
        debug_print(f"Sheet '{sheet_name}' validation passed: {data.shape}, {null_percentage:.1f}% null")
        return True, f"Valid sheet '{sheet_name}'"
        
    except Exception as e:
        error_msg = f"Error validating sheet '{sheet_name}': {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_numeric_column(data: pd.Series, 
                          column_name: str,
                          allow_negative: bool = True,
                          min_value: Optional[float] = None,
                          max_value: Optional[float] = None) -> Tuple[bool, str]:
    """
    Validate a numeric column for specific constraints.
    
    Args:
        data (pd.Series): Column data to validate
        column_name (str): Column name for error messages
        allow_negative (bool): Whether negative values are allowed
        min_value (float, optional): Minimum allowed value
        max_value (float, optional): Maximum allowed value
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        if data.empty:
            return False, f"Column '{column_name}' is empty"
        
        # Convert to numeric, coercing errors to NaN
        numeric_data = pd.to_numeric(data, errors='coerce')
        
        # Check for non-numeric values
        non_numeric_count = numeric_data.isnull().sum() - data.isnull().sum()
        if non_numeric_count > 0:
            debug_print(f"Column '{column_name}' has {non_numeric_count} non-numeric values")
        
        # Get valid numeric values for validation
        valid_values = numeric_data.dropna()
        
        if len(valid_values) == 0:
            return False, f"Column '{column_name}' has no valid numeric values"
        
        # Check negative values
        if not allow_negative and (valid_values < 0).any():
            negative_count = (valid_values < 0).sum()
            return False, f"Column '{column_name}' has {negative_count} negative values (not allowed)"
        
        # Check minimum value
        if min_value is not None and valid_values.min() < min_value:
            return False, f"Column '{column_name}' minimum value {valid_values.min()} is below {min_value}"
        
        # Check maximum value
        if max_value is not None and valid_values.max() > max_value:
            return False, f"Column '{column_name}' maximum value {valid_values.max()} is above {max_value}"
        
        debug_print(f"Numeric column '{column_name}' validation passed: {len(valid_values)} valid values")
        return True, f"Valid numeric column '{column_name}'"
        
    except Exception as e:
        error_msg = f"Error validating numeric column '{column_name}': {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_plotting_data(data: pd.DataFrame, plot_type: str) -> Tuple[bool, str]:
    """
    Validate data for plotting capabilities.
    
    Args:
        data (pd.DataFrame): Data to validate for plotting
        plot_type (str): Type of plot being created
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        # Basic DataFrame validation
        df_valid, df_error = validate_data_frame(data, f"Plotting data for {plot_type}")
        if not df_valid:
            return False, df_error
        
        # Check minimum requirements for plotting
        if len(data) < 2:
            return False, f"Insufficient data for {plot_type} plot: need at least 2 rows, got {len(data)}"
        
        if len(data.columns) < 2:
            return False, f"Insufficient columns for {plot_type} plot: need at least 2 columns, got {len(data.columns)}"
        
        # Check for numeric data
        numeric_columns = data.select_dtypes(include=[np.number]).columns
        if len(numeric_columns) == 0:
            return False, f"No numeric columns found for {plot_type} plot"
        
        # Check for sufficient non-null values
        max_valid_values = 0
        for col in numeric_columns:
            valid_count = data[col].dropna().count()
            max_valid_values = max(max_valid_values, valid_count)
        
        if max_valid_values < 2:
            return False, f"Insufficient valid data points for {plot_type} plot: need at least 2, got {max_valid_values}"
        
        debug_print(f"Plotting data validation passed for {plot_type}: {data.shape}, {len(numeric_columns)} numeric columns")
        return True, f"Valid plotting data for {plot_type}"
        
    except Exception as e:
        error_msg = f"Error validating plotting data for {plot_type}: {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_database_connection(connection_string: str) -> Tuple[bool, str]:
    """
    Validate database connection parameters.
    
    Args:
        connection_string (str): Database connection string
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        if not connection_string:
            return False, "Database connection string is empty"
        
        # Basic validation of connection string format
        required_params = ['host', 'database']
        for param in required_params:
            if param not in connection_string.lower():
                return False, f"Database connection string missing '{param}' parameter"
        
        debug_print("Database connection string validation passed")
        return True, "Valid database connection string"
        
    except Exception as e:
        error_msg = f"Error validating database connection: {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

def validate_image_file(file_path: str) -> Tuple[bool, str]:
    """
    Validate if a file is a valid image file.
    
    Args:
        file_path (str): Path to image file
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        # First validate the file path
        path_valid, path_error = validate_file_path(file_path)
        if not path_valid:
            return False, path_error
        
        # Check file extension
        ext = Path(file_path).suffix.lower()
        valid_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif']
        
        if ext not in valid_extensions:
            return False, f"Not a valid image file: {file_path} (extension: {ext})"
        
        # Check file size (basic validation)
        file_size = Path(file_path).stat().st_size
        if file_size == 0:
            return False, f"Image file is empty: {file_path}"
        
        if file_size > 50 * 1024 * 1024:  # 50MB limit
            return False, f"Image file too large: {file_path} ({file_size / 1024 / 1024:.1f}MB)"
        
        debug_print(f"Image file validation passed: {file_path}")
        return True, "Valid image file"
        
    except Exception as e:
        error_msg = f"Error validating image file {file_path}: {str(e)}"
        error_print(error_msg, e)
        return False, error_msg

print("DEBUG: validators.py - Validation functions loaded successfully")
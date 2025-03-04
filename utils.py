"""
Utility module for the DataViewer Application. Developed by Charlie Becquet.
Provides a Tkinter-based interface for interacting with Excel data, generating reports,
and plotting graphs.
"""
import numpy as np
import os
import re
import sys
import tempfile
import pandas as pd
from tkinter import filedialog, messagebox, Toplevel, Label, Button
from typing import Optional
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

FONT = ('Arial', 12)
# Global constants
APP_BACKGROUND_COLOR = '#0504AA'
BUTTON_COLOR = '#4169E1'
PLOT_CHECKBOX_TITLE = "Click Checkbox to \nAdd/Remove Item \nFrom Plot"

def round_values(value, decimals=2):
    """Round the values to a specified number of decimal places."""
    try:
        return round(float(value), decimals)
    except (ValueError, TypeError):
        return value

def get_save_path(self, default_extension: str = ".xlsx") -> Optional[str]:
    """
    Prompt the user to select a file save location and return the path.

    Args:
        default_extension (str): Default file extension for the save dialog.

    Returns:
        Optional[str]: The selected file path or None if canceled.
    """
    return filedialog.asksaveasfilename(
        defaultextension=default_extension,
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for PyInstaller and Python execution.

    Args:
        relative_path (str): Relative path of the resource.

    Returns:
        str: Absolute path to the resource.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def generate_temp_image(self, figure):
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
        figure.savefig(tmpfile.name)
        return tmpfile.name

def remove_empty_columns(data: pd.DataFrame) -> pd.DataFrame:
    """
    Remove empty columns from a DataFrame. An empty column is defined as one where all values are NaN or 0.
    This process continues until a non-empty column is encountered.

    Args:
        data (pd.DataFrame): The input DataFrame.

    Returns:
        pd.DataFrame: The DataFrame with empty columns removed.
    """
    # Iterate over the columns and drop any empty columns (filled with NaN or 0)
    while True:
        # Check if all values in the first column are NaN or 0
        if data.iloc[:, 0].replace(0, pd.NA).isna().all():
            # If the column is empty, drop it
            data = data.iloc[:, 1:]  # Drop the first column
        else:
            # If we find a non-empty column, break the loop
            break
        # Stop if we have dropped all columns (to avoid errors with empty DataFrames)
        if data.shape[1] == 0:
            break
    return data

def clean_columns(data):
    """
    Cleans DataFrame columns by handling NaN headers and duplicate columns:
    - Removes columns where the header is NaN and the column is completely empty.
    - Retains columns with valid headers, even if they are empty.
    - Renames columns with NaN headers to their column index (1-based).
    - Ensures all column names are unique by appending a counter to duplicates.
    """
    # Step 1: Replace NaN headers with column index (1-based)
    data.columns = [
        str(i + 1) if pd.isna(col) else col for i, col in enumerate(data.columns)
    ]
    data.columns = data.columns.astype(str)

    # Step 2: Remove columns with NaN headers and completely empty data
    non_empty_columns = ~((data.isna().all() | (data == '').all() | (data == 0).all()) & data.columns.str.isnumeric())
    data = data.loc[:, non_empty_columns]

    # Step 3: Deduplicate column names
    new_columns = pd.Series(data.columns)
    for dup in new_columns[new_columns.duplicated()].unique():
        dups = new_columns[new_columns == dup].index
        new_columns[dups] = [f"{dup}.{i}" if i != 0 else dup for i in range(len(dups))]
    data.columns = new_columns

    # Step 4: Replace NaN values with empty strings
    data = data.fillna('')

    # step 5: Remove columns with headers like "Unnamed"

    data = data.loc[:, ~data.columns.str.contains('^Unnamed')]

    return data

def wrap_text(text, max_width=None):
        """
        Wrap text for labels to fit within a given width.

        Args:
            text (str): The text to wrap.
            max_width (int): The maximum number of characters per line. If None, it will calculate dynamically.

        Returns:
            str: The wrapped text with line breaks.
        """
        if max_width is None:
            # Dynamically calculate max_width based on widget size
            axes_width = self.figure.get_size_inches()[0] * self.figure.dpi  # Width in pixels
            button_width = axes_width * 0.12  # Fraction of the figure allocated to checkboxes (from `add_axes`)
            char_width = 8  # Approximate width of a character in pixels (adjust as needed)
            max_width = int(button_width / char_width)

        if len(text) <= max_width:
            return text
        else:
            return '\n'.join([text[i:i+max_width] for i in range(0, len(text), max_width)])

def autofit_columns_in_excel(file_path):
    """
    Adjusts the column widths in the Excel file to fit the content automatically.
    """
    try:
        workbook = load_workbook(file_path)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter  # Get the column letter
                for cell in column:
                    try:  # Ensure cells with no value don't cause issues
                        max_length = max(max_length, len(str(cell.value) if cell.value else ""))
                    except:
                        pass
                adjusted_width = max_length + 2  # Add padding for readability
                sheet.column_dimensions[column_letter].width = adjusted_width
        workbook.save(file_path)
    except Exception as e:
        print(f"Error while adjusting column widths: {e}")

def is_standard_file(file_path: str) -> bool:
    """
    Determine if the file is standard by checking for the presence of
    "sample id:" in the first four rows of the 6th column (zero-indexed column 5).
    
    Args:
        file_path (str): Path to the Excel file.
    
    Returns:
        bool: True if the file is standard (i.e. the cell contains "sample id:"), 
              False if it is legacy (non-standard).
    """
    try:
        sheets_dict = pd.read_excel(file_path, sheet_name=None, header = None, nrows=4)
        
        # Decide which sheet to use.
        if len(sheets_dict) > 1:
            if "Intense Test" in sheets_dict:
                df = sheets_dict["Intense Test"]
                print("Using sheet 'Intense Test'.")
            else:
                # If "Intense Test" isn't found, use the first sheet.
                first_sheet = list(sheets_dict.keys())[0]
                df = sheets_dict[first_sheet]
                print(f"'Intense Test' not found. Using first sheet: '{first_sheet}'.")
        else:
            # Only one sheet exists.
            df = list(sheets_dict.values())[0]
            print("Only one sheet found; using it.")
        
        # Check that there are at least 6 columns.
        if df.shape[1] > 5:
            # Check each of the first four rows in the 6th column (index 5)
            for row in range(min(4, df.shape[0])):
                cell_val = df.iat[row, 4]
                print(f"Row {row+1}, column 5 value: {cell_val}")
                if isinstance(cell_val, str) and cell_val.strip().lower() == "sample id:":
                    print("Found 'sample id:' -> File is standard.")
                    return True
        print("Did not find 'sample id:' in the expected location -> File is legacy.")
        return False
    except Exception as e:
        print(f"Error reading file: {e}")
        return False

def plotting_sheet_test(sheet_name, data):
    """
    Determine if a sheet is a plotting sheet by searching the first few rows for 'puffs' and 'tpm'.
    """
    try:
        if "legacy" in sheet_name.lower():
            return True

        if data.shape[0] < 3 or data.shape[1] < 2:
            return False

        # Dynamically search the first 5 rows for headers
        for i in range(min(5, len(data))):
            header_row = data.iloc[i].astype(str).str.lower()
            # print(f"Checking row {i} for headers: {header_row.values}")

            if header_row.str.contains("puffs").any() and header_row.str.contains("tpm").any():
                print(f"Found plotting headers in row {i}")
                return True

        print("No valid plotting headers found")
        return False

    except Exception as e:
        print(f"Error in plotting_sheet_test: {e}")
        return False

# Helper functions
def extract_metadata(worksheet, patterns):
    """Extract metadata from first 3 rows using regex patterns."""
    metadata = {}
    for row in range(1, 3):  # Rows 1-3 (1-indexed)
        for col in range(1, worksheet.max_column + 1):
            cell_value = str(worksheet.cell(row=row, column=col).value).lower()
            for key, pattern in patterns.items():
                if re.search(pattern, cell_value, re.IGNORECASE):
                    metadata[key] = worksheet.cell(
                        row=row, 
                        column=col+1
                    ).value  # Get value from next cell
                    break
    return metadata

def map_metadata_to_template(template_ws, metadata):
    """Map extracted metadata to template positions."""
    mapping = {
        "sample_name": (1, 6),   # Row 1, Column F
        "voltage": (4, 6),       # Row 4, Column F
        "viscosity": (4, 2),     # Row 4, Column B
        "resistance": (3, 4),    # Row 3, Column D
        "puff_regime": (3, 8),   # Row 3, Column H
        "initial_oil": (3, 8),   # Row 3, Column H
        "date": (1, 4),         # Row 1, Column D
        "media": (2, 2)         # Row 2, Column B
    }
    for key, (row, col) in mapping.items():
        if key in metadata:
            template_ws.cell(row=row, column=col, value=metadata[key])

def copy_data_rows(src_ws, dest_ws):
    """Copy data rows (4+) exactly, maintaining structure."""
    for row in src_ws.iter_rows(min_row=4):
        for cell in row:
            dest_ws.cell(
                row=cell.row,
                column=cell.column,
                value=cell.value
            )

def get_plot_sheet_names():
    """
    Get the names of plot sheets.

    Returns:
        list: List of plot sheet names.
    """
    return [
        "Quick Screening Test", "Lifetime Test", "Device Life Test", "Horizontal Puffing Test", "Extended Test", "Long Puff Test",
        "Rapid Puff Test", "Intense Test", "Big Headspace Low T Test", "Big Headspace High T Test",
        "Viscosity Compatibility", "Upside Down Test", "Big Headspace Pocket Test",
        "Low Temperature Stability","Vacuum Test", "Negative Pressure Test", "Viscosity Compatibility", "Various Oil Compatibility", "Sheet1"
    ]

def read_sheet_with_values(file_path: str, sheet_name: Optional[str] = None):
    # Load workbook with computed values
    wb = load_workbook(file_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    data = ws.values  # This is a generator over rows
    # Convert generator to DataFrame
    df = pd.DataFrame(data)
    return df

def read_sheet_with_values_standards(file_path: str, sheet_name: Optional[str] = None):
    """Read Excel sheet with values exactly as they appear."""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Assign the first row as the header
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    return df

def unmerge_all_cells(ws: Worksheet) -> None:
    """
    Unmerge all merged cells in the given worksheet.
    
    For each merged range, this function:
      - Retrieves the value from the top-left (master) cell.
      - Unmerges the range.
      - Sets every cell in the range to the retrieved value.
    
    Args:
        ws (Worksheet): The Openpyxl worksheet to process.
    """
    # Create a list of merged ranges to avoid modifying the collection while iterating.
    merged_ranges = list(ws.merged_cells.ranges)
    
    for merged_range in merged_ranges:
        # Get the bounds of the merged range.
        min_row, min_col, max_row, max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
        
        # Retrieve the value from the top-left (master) cell.
        master_value = ws.cell(row=min_row, column=min_col).value
        
        # Unmerge the cells.
        ws.unmerge_cells(str(merged_range))
        
        # Fill every cell in the previously merged range with the master value.
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row=row, column=col, value=master_value)
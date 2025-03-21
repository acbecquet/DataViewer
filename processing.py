﻿"""
Processing module for the DataViewer Application. Developed by Charlie Becquet.
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import os
from typing import Optional, List, Tuple, Any
from openpyxl import load_workbook
from pptx.util import Inches
from utils import (
    round_values,
    remove_empty_columns,
    wrap_text,
    plotting_sheet_test,
    extract_meta_data,
    map_meta_data_to_template,
    read_sheet_with_values,
    unmerge_all_cells,
    validate_sheet_data,
    header_matches,
    load_excel_file
)

#pd.set_option('future.no_silent_downcasting', True)

def get_valid_plot_options(plot_options: List[str], full_sample_data: pd.DataFrame) -> List[str]:
    """
    Check which plot options have valid, non-empty data and return the valid options.

    Args:
        plot_options (list): List of plot types to check.
        full_sample_data (pd.DataFrame): DataFrame containing the sample data.

    Returns:
        list: Valid plot options.
    """
    valid_options = []
    for plot_type in plot_options:
        y_data = get_y_data_for_plot_type(full_sample_data, plot_type)
        if y_data.dropna().astype(bool).any():  # Ensure there are non-NaN, non-zero values
            valid_options.append(plot_type)
    return valid_options


# ==================== PLOTTING FUNCTIONS ====================

def get_y_data_for_plot_type(sample_data, plot_type):
    """
    Extract y-data for the specified plot type.

    Args:
        sample_data (pd.DataFrame): DataFrame containing the sample data.
        plot_type (str): The type of plot to generate.

    Returns:
        pd.Series: y-data for plotting.
    """
    plot_type_mapping = {
        "TPM": sample_data.iloc[3:, 8],
        "Draw Pressure": sample_data.iloc[3:, 3],
        "Resistance": sample_data.iloc[3:, 4],
        "Power Efficiency": sample_data.iloc[3:, 9],
    }
    return plot_type_mapping.get(plot_type, sample_data.iloc[3:, 8])  # Default to TPM

def get_y_label_for_plot_type(plot_type):
    """
    Get the y-axis label for the specified plot type.

    Args:
        plot_type (str): The type of plot.

    Returns:
        str: y-axis label.
    """
    y_label_mapping = {
        "TPM": 'TPM (mg/puff)',
        "Draw Pressure": 'Draw Pressure (kPa)',
        "Resistance": 'Resistance (Ohms)',
        "Power Efficiency": 'Power Efficiency (mg/W)',
    }
    return y_label_mapping.get(plot_type, 'TPM (mg/puff)')  # Default to TPM

def plot_all_samples(full_sample_data: pd.DataFrame, num_columns_per_sample: int, plot_type: str) -> Tuple[plt.Figure, List[str]]:
    """Generate the plot for all samples and return the figure."""

    num_samples = full_sample_data.shape[1] // num_columns_per_sample
    full_sample_data = full_sample_data.replace(0, np.nan)
    fig, ax = plt.subplots(figsize=(8, 6))
    sample_names = []

    if plot_type == "TPM (Bar)":
        sample_names = plot_tpm_bar_chart(ax, full_sample_data, num_samples, num_columns_per_sample)
    else:
        y_max = 0
        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]

            x_data = sample_data.iloc[3:, 0].dropna()
            y_data = get_y_data_for_plot_type(sample_data, plot_type)
            y_data = pd.to_numeric(y_data, errors='coerce').dropna()

            common_index = x_data.index.intersection(y_data.index)
            x_data = x_data.loc[common_index]
            y_data = y_data.loc[common_index]

            if not x_data.empty and not y_data.empty:
                sample_name = sample_data.columns[5]
                ax.plot(x_data, y_data, marker='o', label=sample_name)
                sample_names.append(sample_name)
                y_max = max(y_max, y_data.max())
            #else:
                #print(f"Sample {i+1} - Missing Data for Plot")

        ax.set_xlabel('Puffs')
        ax.set_ylabel(get_y_label_for_plot_type(plot_type))
        ax.set_title(plot_type)
        ax.legend(loc='upper right')

        if y_max > 9 and y_max <= 50:
            ax.set_ylim(0, y_max)
        else:
            ax.set_ylim(0, 9)

    return fig, sample_names

def plot_tpm_bar_chart(ax, full_sample_data, num_samples, num_columns_per_sample):
    """
    Generate a bar chart of the average TPM for each sample with wrapped sample names.

    Args:
        ax (matplotlib.axes.Axes): Matplotlib axis to draw on.
        full_sample_data (pd.DataFrame): DataFrame of sample data.
        num_samples (int): Number of samples.
        num_columns_per_sample (int): Columns per sample.

    Returns:
        list: Sample names.
    """
    averages = []
    labels = []
    sample_names = []

    for i in range(num_samples):
        start_col = i * num_columns_per_sample
        end_col = start_col + num_columns_per_sample
        sample_data = full_sample_data.iloc[:, start_col:end_col]

        if pd.isna(sample_data.iloc[3, 1]):
            continue

        tpm_data = sample_data.iloc[3:, 8].dropna()  # TPM column (adjust if needed)
        average_tpm = tpm_data.mean()
        averages.append(average_tpm)

        sample_name = sample_data.columns[5]
        sample_names.append(sample_name)  # Add original sample name
        wrapped_name = wrap_text(text=sample_name, max_width=10)  # Use `wrap_text` to dynamically wrap names
        labels.append(wrapped_name)

    # Create a colormap with unique colors
    num_bars = len(averages)
    colors = cm.get_cmap('tab10', num_bars)(np.linspace(0, 1, num_bars))

    # Plot the bar chart with unique colors
    ax.bar(labels, averages, color=colors)
    ax.set_xlabel('Samples')
    ax.set_ylabel('Average TPM (mg/puff)')
    ax.set_title('Average TPM per Sample')
    ax.tick_params(axis='x', labelsize=8)  # Rotate x-axis labels for better readability

    return sample_names

def clean_presentation_tables(presentation):
    """
    Clean all tables in the given PowerPoint presentation by:
    - Removing rows where all cells are empty.
    - Removing columns where all cells, including the header, are empty.

    Args:
        presentation (pptx.Presentation): The PowerPoint presentation to process.
    """
    for slide in presentation.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue  # Skip if the shape is not a table
            
            table = shape.table

            # Get the number of rows and columns
            num_rows = len(table.rows)
            num_cols = len(table.columns)

            # Step 1: Remove empty rows
            rows_to_keep = []
            for row_idx in range(num_rows):
                is_empty_row = all(
                    not cell.text.strip() for cell in table.rows[row_idx].cells
                )
                if not is_empty_row:
                    rows_to_keep.append(row_idx)

            # Keep only the non-empty rows
            for row_idx in reversed(range(num_rows)):
                if row_idx not in rows_to_keep:
                    table._tbl.remove(table.rows[row_idx]._tr)  # Remove row from XML

            # Step 2: Remove empty columns
            num_rows = len(table.rows)  # Updated number of rows after cleaning
            cols_to_keep = []
            for col_idx in range(num_cols):
                # Check if all cells in the column (including the header) are empty or 'nan'
                is_empty_col = all(
                    not table.cell(row_idx, col_idx).text.strip() or table.cell(row_idx, col_idx).text.strip() == "nan"
                    for row_idx in range(num_rows)
                )
                if not is_empty_col:
                    cols_to_keep.append(col_idx)

            # Keep only the non-empty columns
            for col_idx in reversed(range(num_cols)):
                if col_idx not in cols_to_keep:
                    for row in table.rows:
                        row._tr.remove(row.cells[col_idx]._tc)  # Remove column cell from XML

# ==================== DATA PROCESSING FUNCTIONS ====================

def process_sheet(data, headers_row=3, data_start_row=4):
    """
    Process a sheet by extracting headers and data.
    
    Args:
        data (pd.DataFrame): Input data from the sheet.
        headers_row (int): Row index for headers.
        data_start_row (int): Row index where data starts.
        
    Returns:
        tuple: (processed_data, table_data)
    """
    data = remove_empty_columns(data)
    headers = data.iloc[headers_row, :].tolist()
    table_data = data.iloc[data_start_row:, :]
    table_data = table_data.replace(0, np.nan)
    processed_data = pd.DataFrame(table_data.values, columns=headers)
    return processed_data, table_data

def process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12, custom_extracted_data_fn=None):
    """
    Generic function to process sheets with plotting data.

    Args:
        data (pd.DataFrame): Input data from the sheet.
        headers_row (int): Row index for headers. Defaults to 3.
        data_start_row (int): Row index where data starts. Defaults to 4.
        num_columns_per_sample (int): Number of columns per sample. Defaults to 12.
        custom_extracted_data_fn (callable, optional): Custom function to extract additional data.

    Returns:
        tuple: (processed_data, sample_arrays, full_sample_data)
            processed_data (pd.DataFrame): Processed data for display.
            sample_arrays (dict): Extracted sample arrays for plotting.
            full_sample_data (pd.DataFrame): Concatenated data for all samples.
    """
    if not validate_sheet_data(data, required_rows=data_start_row + 1):
        return pd.DataFrame(), {}, pd.DataFrame()

    try:
        # Clean up the data
        data = remove_empty_columns(data).replace(0, np.nan)

        samples = []
        full_sample_data = []
        sample_arrays = {}

        # Calculate the number of samples
        num_samples = data.shape[1] // num_columns_per_sample

        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            end_col = start_col + num_columns_per_sample
            sample_data = data.iloc[:, start_col:end_col]

            if sample_data.empty:
                #print(f"Sample {i+1} is empty. Skipping.")
                continue

            # Extract plotting data
            try:
                sample_arrays[f"Sample_{i+1}_Puffs"] = sample_data.iloc[3:, 0].to_numpy()
                sample_arrays[f"Sample_{i+1}_TPM"] = sample_data.iloc[3:, 8].to_numpy()

                if custom_extracted_data_fn:
                    extracted_data = custom_extracted_data_fn(sample_data)
                else:
                    tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
                    avg_tpm = round_values(tpm_data.mean()) if not tpm_data.empty else None
                    std_tpm = round_values(tpm_data.std()) if not tpm_data.empty else None

                    extracted_data = {
                        "Sample Name": sample_data.columns[5],
                        "Media": sample_data.iloc[0, 1],
                        "Viscosity": sample_data.iloc[1, 1],
                        "Voltage, Resistance, Power": f"{sample_data.iloc[1, 5]} V, "
                                                       f"{round_values(sample_data.iloc[0, 3])} Ω, "
                                                       f"{round_values(sample_data.iloc[0, 5])} W",
                        "Average TPM": avg_tpm,
                        "Standard Deviation": std_tpm,
                        "Initial Oil Mass": round_values(sample_data.iloc[1,7]),
                        "Usage Efficiency": round_values(sample_data.iloc[1,8]),
                        "Burn?": sample_data.columns[10],
                        "Clog?": sample_data.iloc[0, 10],
                        "Leak?": sample_data.iloc[1, 10]
                    }

                samples.append(extracted_data)
                full_sample_data.append(sample_data)
            except IndexError as e:
                #print(f"Index error for sample {i+1}: {e}. Skipping.")
                continue

        if not samples:
            #print("No valid samples found.")
            return pd.DataFrame(), {}, pd.DataFrame()

        # Concatenate all sample data
        full_sample_data_df = pd.concat(full_sample_data, axis=1)

        # Create a processed data DataFrame
        processed_data = pd.DataFrame(samples)

        return processed_data, sample_arrays, full_sample_data_df
    except Exception as e:
        #print(f"Error processing plot sheet: {e}")
        return pd.DataFrame(), {}, pd.DataFrame()

def no_efficiency_extracted_data(sample_data):
    """
    Custom function to extract data without efficiency metrics.

    Args:
        sample_data (pd.DataFrame): The DataFrame containing the sample data.

    Returns:
        dict: Extracted data without efficiency metrics.
    """
    tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
    avg_tpm = round_values(tpm_data.mean()) if not tpm_data.empty else None
    std_tpm = round_values(tpm_data.std()) if not tpm_data.empty else None
    return {
        "Sample Name": sample_data.columns[5],
        "Media": sample_data.iloc[0, 1],
        "Viscosity": sample_data.iloc[1, 1],
        "Voltage, Resistance, Power": f"{sample_data.iloc[1, 5]} V, "
                                       f"{round_values(sample_data.iloc[0, 3])} Ω, "
                                       f"{round_values(sample_data.iloc[0, 5])} W",
        "Average TPM": avg_tpm,
        "Standard Deviation": std_tpm,
        "Burn?": sample_data.columns[10],
        "Clog?": sample_data.iloc[0, 10],
        "Leak?": sample_data.iloc[1, 10],
    }

def process_generic_sheet(data, headers_row=3, data_start_row=4):
    """
    Process data for any sheet that doesn't match predefined sheet names.
    
    Args:
        data (pd.DataFrame): Input data from the Excel sheet.
    
    Returns:
        tuple: (display_df, additional_output, full_sample_data)
            display_df (pd.DataFrame): Processed table data for display.
            additional_output (dict): Any additional information that might be needed.
            full_sample_data (pd.DataFrame): Full sample data concatenated for all samples.
    """
    # Check if data is empty or invalid
    if data.empty or data.isna().all().all():
        return pd.DataFrame(), {}, pd.DataFrame()  # Return empty data structures

    try:
        # Check for sufficient rows
        if data.shape[0] <= headers_row or data.shape[0] <= data_start_row:
            #print(f"Insufficient rows for processing (headers_row={headers_row}, data_start_row={data_start_row})")
            return pd.DataFrame(), {}, pd.DataFrame()

        # Extract headers and validate them
        column_names = data.iloc[headers_row].fillna("").tolist()
        
        # Slice the data starting from the data_start_row and convert to strings
        table_data = data.iloc[data_start_row:].copy()
        table_data.columns = column_names
        table_data = table_data.astype(str)  # Convert all data to strings

        # Create a DataFrame with the extracted headers and data
        display_df = table_data.copy()

        # Collect all data into a single DataFrame for further use
        full_sample_data_df = pd.concat([table_data], axis=1)
        
        return display_df, {}, full_sample_data_df
    except Exception as e:
        #print(f"Error processing generic sheet: {e}")
        return pd.DataFrame(), {}, pd.DataFrame()

# ==================== SHEET PROCESSING DISPATCHER ====================

def get_processing_function(sheet_name):
    """
    Get the appropriate processing function for a sheet based on its name.
    
    Args:
        sheet_name (str): Name of the sheet.
        
    Returns:
        callable: Function to process the sheet.
    """
    functions = get_processing_functions()
    if "legacy" in sheet_name.lower():
        return process_legacy_test
    return functions.get(sheet_name, functions["default"])

def get_processing_functions():
    """
    Get a dictionary mapping sheet names to their processing functions.
    
    Returns:
        dict: Map of sheet names to processing functions.
    """
    return {
        "Test Plan": process_test_plan,
        "Initial State Inspection": process_initial_state_inspection,
        "Quick Screening Test": process_quick_screening_test,
        "Lifetime Test": process_device_life_test,
        "Device Life Test": process_device_life_test,
        "Aerosol Temperature": process_aerosol_temp_test,
        "User Test - Full Cycle": process_user_test,
        "Horizontal Puffing Test": process_horizontal_test,
        "Extended Test": process_extended_test,
        "Long Puff Test": process_long_puff_test,
        "Rapid Puff Test": process_rapid_puff_test,
        "Intense Test": process_intense_test,
        "Big Headspace Low T Test": process_big_head_low_t_test,
        "Anti-Burn Protection Test": process_burn_protection_test,
        "Big Headspace High T Test": process_big_head_high_t_test,
        "Upside Down Test": process_upside_down_test,
        "Big Headspace Pocket Test": process_pocket_test,
        "Temperature Cycling Test": process_temperature_cycling_test,
        "High T High Humidity Test": process_high_t_high_humidity_test,
        "Low Temperature Stability": process_cold_storage_test,
        "Vacuum Test": process_vacuum_test,
        "Negative Pressure Test": process_vacuum_test,
        "Viscosity Compatibility": process_viscosity_compatibility_test,
        "Various Oil Compatibility": process_various_oil_test,
        "Quick Sensory Test": process_quick_sensory_test,
        "Off-odor Score": process_off_odor_score,
        "Sensory Consistency": process_sensory_consistency,
        "Heavy Metal Leaching Test": process_leaching_test,
        "Sheet1": process_sheet1,
        "default": process_generic_sheet
    }

# ==================== SHEET-SPECIFIC PROCESSING FUNCTIONS ====================
# These functions maintain compatibility with the rest of the codebase
# Many are now simplified wrappers around the more general processing functions

def process_test_plan(data):
    """Process data for the Test Plan sheet."""
    return process_generic_sheet(data, headers_row=5, data_start_row=6)

def process_initial_state_inspection(data):
    """Process data for the Initial State Inspection sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_quick_screening_test(data):
    """Process data for the Quick Screening Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_device_life_test(data):
    """Process data for the Device Life Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

def process_aerosol_temp_test(data):
    """Process data for the Aerosol Temperature sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_off_odor_score(data):
    """Process data for the Off-odor Score sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_sensory_consistency(data):
    """Process data for the Sensory Consistency sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_user_test(data):
    """Process data for the User Test sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_horizontal_test(data):
    """Process data for the Horizontal Puffing Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

def process_extended_test(data):
    """Process data for the Extended Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

def process_long_puff_test(data):
    """Process data for the Long Puff Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_rapid_puff_test(data):
    """Process data for the Rapid Puff Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_intense_test(data):
    """Process data for the Intense Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

def process_legacy_test(data, headers_row=3, data_start_row=4, num_columns_per_sample=12):
    """
    Process legacy test data (already converted to standard format) similarly to other plotting sheets.
    Legacy files (after conversion) contain only the key columns (e.g. puffs and TPM (mg/puff))
    and minimal meta_data. This function uses the same logic as process_plot_sheet to generate the
    processed data, sample arrays, and concatenated full_sample_data.
    
    Args:
        data (pd.DataFrame): The legacy file's data in standard format.
        headers_row (int): The row containing headers.
        data_start_row (int): The row at which the data starts.
        num_columns_per_sample (int): Number of columns per sample.
        
    Returns:
        tuple: (processed_data, sample_arrays, full_sample_data)
            processed_data (pd.DataFrame): Processed data for display.
            sample_arrays (dict): Sample arrays for plotting.
            full_sample_data (pd.DataFrame): Concatenated data for all samples.
    """
    processed_data, sample_arrays, full_sample_data = process_plot_sheet(data, headers_row, data_start_row, num_columns_per_sample)
    #print("Legacy data debug:")
    #print("Puffs array:", sample_arrays.get("Sample_1_Puffs"))
    #print("TPM array:", sample_arrays.get("Sample_1_TPM"))
    return processed_data, sample_arrays, full_sample_data

def process_big_head_low_t_test(data):
    """Process data for the Big Headspace Low T Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_burn_protection_test(data):
    """Process data for the Anti-Burn Protection Test sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_big_head_high_t_test(data):
    """Process data for the Big Headspace High T Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_upside_down_test(data):
    """Process data for the Upside Down Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_pocket_test(data):
    """Process data for the Big Headspace Pocket Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_temperature_cycling_test(data):
    """Process data for the Temperature Cycling Test sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_high_t_high_humidity_test(data):
    """Process data for the High T High Humidity Test sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_cold_storage_test(data):
    """Process data for the Low Temperature Stability sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_vacuum_test(data):
    """Process data for the Vacuum Test sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_viscosity_compatibility_test(data):
    """Process data for the Viscosity Compatibility sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_various_oil_test(data):
    """Process data for the Various Oil Compatibility sheet."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12,
        custom_extracted_data_fn=no_efficiency_extracted_data
    )

def process_quick_sensory_test(data):
    """Process data for the Quick Sensory Test sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_leaching_test(data):
    """Process data for the Heavy Metal Leaching Test sheet."""
    return process_generic_sheet(data, headers_row=3, data_start_row=4)

def process_sheet1(data):
    """Process data for 'Sheet1' similarly to 'process_extended_test'."""
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

# ==================== AGGREGATION FUNCTIONS ====================

def aggregate_sheet_metrics(full_sample_data: pd.DataFrame, num_columns_per_sample: int = 12) -> pd.DataFrame:
    """
    Given full_sample_data from a single sheet (which contains multiple samples laid out in groups of 12 columns),
    this function computes, for each sample, several metrics:
      - Average TPM (from column index 8, rows 3 onward)
      - TPM Standard Deviation (from column index 8, rows 3 onward)
      - Total Puffs: determined by locating the last non-empty row in column index 2 (weight) and then taking
        the value from the puffs column (index 0) at that row.
      - Draw Pressure: from the same last valid row in column index 3.
      - Smell: from the same last valid row in column index 5.
      - Date: from meta_data at row index 0, column index 3.
      - Puff Regime: from meta_data at row index 1, column index 7.
      - Initial Oil Mass: from meta_data at row index 2, column index 7.
      - Burn: from meta_data at row index 0, column index 9.
      - Clog: from meta_data at row index 1, column index 9.
      - Leak: from meta_data at row index 2, column index 9.
      - Sample Name: from column index 5 (assumed to contain the sample name).
    
    Returns a DataFrame with one row per sample.
    """
    num_samples = full_sample_data.shape[1] // num_columns_per_sample
    aggregates = []
    for i in range(num_samples):
        start_col = i * num_columns_per_sample
        end_col = start_col + num_columns_per_sample
        sample_data = full_sample_data.iloc[:, start_col:end_col]
        
        # Convert values to numeric; rows before index 3 are assumed header/info rows.
        try:
            tpm_series = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce')
            avg_tpm = tpm_series.mean()
            std_tpm = tpm_series.std()

        except Exception as e:
            avg_tpm = None
            std_tpm = None

        # For the total number of puffs, instead of summing, we locate the last non-empty row in the third column (index 2)
        # and then take the value in the puffs column (index 0) at that row.
        # for draw pressure, take the value corresponding to the last non-empty row in after weight as well. If empty set value to nan
        # for smell, do the same. If empty just set the value to nan
        try:
            # Select rows starting at index 3 for the third column
            weight_series = sample_data.iloc[3:, 2]
            # Find the last valid (non-NaN and non-empty) index
            last_valid_index = weight_series.last_valid_index()
            if last_valid_index is not None:
                total_puffs = pd.to_numeric(
                    sample_data.loc[last_valid_index, sample_data.columns[0]],
                    errors='coerce'
                )
                draw_pressure = pd.to_numeric(
                    sample_data.loc[last_valid_index, sample_data.columns[3]], 
                    errors='coerce'
                )
                smell = pd.to_numeric(
                    sample_data.loc[last_valid_index, sample_data.columns[5]], 
                    errors='coerce'
                )
            else:
                total_puffs = np.nan
                draw_pressure = np.nan
                smell = np.nan

        except Exception as e:
            total_puffs = None
            draw_pressure = None
            smell = None

        # Extract meta_data from the first three rows (indices 0, 1, and 2)
        try:
            date_val = sample_data.iloc[0, 3]  # Date from column 4 (index 3) at row 1 (index 0)
        except Exception:
            date_val = None

        try:
            puff_regime = sample_data.iloc[1, 7]  # Puff regime from column 8 (index 7) at row 2 (index 1)
        except Exception:
            puff_regime = None

        try:
            initial_oil_mass = sample_data.iloc[2, 7]  # Initial oil mass from column 8 (index 7) at row 3 (index 2)
        except Exception:
            initial_oil_mass = None

        try:
            burn = sample_data.iloc[0, 9]  # Burn from column 10 (index 9) at row 1 (index 0)
        except Exception:
            burn = None

        try:
            clog = sample_data.iloc[1, 9]  # Clog from column 10 (index 9) at row 2 (index 1)
        except Exception:
            clog = None

        try:
            leak = sample_data.iloc[2, 9]  # Leak from column 10 (index 9) at row 3 (index 2)
        except Exception:
            leak = None

        # Retrieve the sample name from column index 5
        sample_name = sample_data.columns[5] if len(sample_data.columns) > 5 else f"Sample {i+1}"
        
        aggregates.append({
            "Sample Name": sample_name,
            "Average TPM": avg_tpm,
            "TPM Std Dev": std_tpm,
            "Total Puffs": total_puffs,
            "Draw Pressure": draw_pressure,
            "Smell": smell,
            "Date": date_val,
            "Puff Regime": puff_regime,
            "Initial Oil Mass": initial_oil_mass,
            "Burn": burn,
            "Clog": clog,
            "Leak": leak
        })
        
    return pd.DataFrame(aggregates)

def plot_aggregate_trends(aggregate_df: pd.DataFrame) -> plt.Figure:
    """
    Given a DataFrame of aggregate metrics (with columns "Sample Name", "Average TPM" and "Total Puffs"),
    this function generates a plot that shows both the average TPM (as a bar chart) and the total number of puffs 
    (as a line plot on a secondary axis) across samples.
    """
    fig, ax1 = plt.subplots(figsize=(8, 6))
    samples = aggregate_df["Sample Name"]
    tpm = aggregate_df["Average TPM"]
    puffs = aggregate_df["Total Puffs"]

    # Plot average TPM as bars
    color_tpm = 'tab:blue'
    ax1.set_xlabel("Samples")
    ax1.set_ylabel("Average TPM (mg/puff)", color=color_tpm)
    bars = ax1.bar(samples, tpm, color=color_tpm, alpha=0.6, label="Average TPM")
    ax1.tick_params(axis='y', labelcolor=color_tpm)
    plt.xticks(rotation=45, ha='right')

    # Create a second y-axis for total puffs
    ax2 = ax1.twinx()
    color_puffs = 'tab:red'
    ax2.set_ylabel("Total Puffs", color=color_puffs)
    ax2.plot(samples, puffs, color=color_puffs, marker='o', label="Total Puffs")
    ax2.tick_params(axis='y', labelcolor=color_puffs)

    fig.tight_layout()
    plt.title("Aggregate Metrics per Sample")
    plt.legend(loc='upper left')
    return fig

# ==================== LEGACY FILE PROCESSING FUNCTIONS ====================



def extract_samples_from_old_file(file_path: str, sheet_name: Optional[str] = None) -> list:
    """
    Extract samples from an old Excel file by robustly locating the headers "puffs" and "TPM (mg/puff)" anywhere in the sheet.
    For each sample header found, a new set of meta_data (from up to 3 rows above) and column data is extracted.
    
    Returns a list of dictionaries, one per sample, each containing:
        - 'sample_name': The value found near the puffs header.
        - 'puffs': A pandas Series of the puffs column.
        - 'tpm': A pandas Series of the TPM column.
        - 'header_row': The index of the header row.
        - Additional meta_data if detected.
    """
    # Read the sheet without assuming a header row.
    df = read_sheet_with_values(file_path, sheet_name)
    samples = []
    nrows, ncols = df.shape

    # Regex patterns for meta_data and data headers.
    meta_data_patterns = {
        "sample_name": [
            r"(cart(ridge)?\s*#|sample\s*(name|id))",  # e.g. "Cart #", "Sample ID"
            r"puffing\s*data\s*for\s*:?\s*" # e.g puffing data for:
        ],
        "resistance": [
            r"\bri\s*\(?\s*ohms?\s*\)?\s*:?\s*",  # e.g. "Ri (Ohms)"
            r"resistance\s*\(?ohms?\)?\s*:?\s*"
        ],
        "voltage": [r"voltage\s*:?\s*"],
        "viscosity": [r"viscosity\b\s*:?\s*"],
        "puffing_regime": [r"\b(puff(ing)?\s*regime|puff\s*settings?)\s*:?\s*"],
        "initial_oil_mass": [r"initial\s*oil\s*mass\b\s*:?\s*"],
        "date": [r"date\s*:?\s*"],
        "media": [r"media\s*:?\s*"]
    }

    data_header_patterns = {
        "puffs": r"puffs",
        "tpm": r"tpm\s*\(mg\s*/\s*puff\)",
        "before_weight": r"before\s*weight/g",
        "after_weight": r"after\s*weight/g",
        "draw_pressure": r"pv1|draw\s*pressure\s*\(kpa\)",
        "smell": r"smell",
        "notes": r"notes"
    }

    numeric_columns = {"puffs", "tpm", "before_weight", "after_weight", "draw_pressure"}

    # Track processed columns for sample data and meta_data separately.
    processed_cols = {row: [] for row in range(nrows)}
    processed_meta_data = {row: [] for row in range(nrows)}
    # Set a threshold: cells within this number of columns are considered too close.
    proximity_threshold = 1

    # Loop over every cell in the sheet.
    for row in range(nrows):
        for col in range(ncols):
            # Skip this cell if it's near a previously processed sample header in the same row.
            if any(abs(col - proc_col) < proximity_threshold for proc_col in processed_cols[row]):
                continue

            cell_val = df.iat[row, col]
            if header_matches(cell_val, data_header_patterns["puffs"]):
                # New sample header found.
                sample = {"sample_name": str(cell_val).strip(), "header_row": row}

                # Mark this column as processed for sample data.
                processed_cols[row].append(col)

                # -----------------------------
                # Extract meta_data (up to 3 rows above)
                # -----------------------------
                start_search_row = max(0, row - 3)
                meta_data_found = {}
                for r in range(row - 1, start_search_row - 1, -1):
                    for c in range(ncols):
                        # Skip if this cell in meta_data row was already used.
                        if any(abs(c - pm) < proximity_threshold for pm in processed_meta_data[r]):
                            continue
                        cell_val_above = df.iat[r, c]
                        for key, patterns in meta_data_patterns.items():
                            if key in meta_data_found:
                                continue  # Already found this key for the current sample.
                            for pattern in patterns:
                                if header_matches(cell_val_above, pattern):
                                    value = df.iat[r, c + 1] if (c + 1 < ncols) else None
                                    meta_data_found[key] = value
                                    # Mark this column as used for meta_data in row r.
                                    processed_meta_data[r].append(c)
                                    break
                            if key in meta_data_found:
                                break
                sample.update(meta_data_found)

                # -----------------------------
                # Extract column data for this sample.
                # -----------------------------
                data_cols = {}
                for key, pattern in data_header_patterns.items():
                    # Look to the right from the current header cell.
                    for j in range(col, ncols):
                        if header_matches(df.iat[row, j], pattern):
                            raw_data = df.iloc[row + 1:, j]
                            if key in numeric_columns:
                                data_cols[key] = pd.to_numeric(raw_data, errors='coerce').dropna()
                            else:
                                data_cols[key] = raw_data.astype(str).replace("nan", "")
                            break
                # Only add the sample if both "puffs" and "tpm" data were found.
                if "puffs" in data_cols and "tpm" in data_cols:
                    sample.update(data_cols)
                    samples.append(sample)
    return samples

def convert_legacy_file_using_template(legacy_file_path: str, template_path: str = None) -> pd.DataFrame:
    """
    Converts a legacy Excel file to the standardized template format.
    - Uses 12-column blocks per sample.
    - Maps meta_data into exact template positions.
    - For each sample block, only rows up to the first empty 'After weight/g'
      are written and all cells below that row (within that block) are cleared.
    """
    # Determine the template path.
    if template_path is None:
        template_path = os.path.join(os.path.abspath("."), "resources", 
                                     "Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
    
    template_sheet = "Intense Test"
    wb = load_workbook(template_path)
    if template_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{template_sheet}' not found in template file.")
    ws = wb[template_sheet]

    # Load legacy samples using our extraction function.
    legacy_samples = extract_samples_from_old_file(legacy_file_path)
    if not legacy_samples:
        raise ValueError("No valid legacy sample data found.")

    # meta_data mapping: each key maps to a list of regex patterns and the target (row, col offset).
    meta_data_MAPPING = {
        "Sample ID:": (
            [r"cart\s*#:?", r"sample\s*(name|id):?"], 
            (1, 5)
        ),
        "Voltage:": ([r"voltage:?"], (3, 5)),
        "Viscosity:": ([r"viscosity:?"], (3, 1)),
        "Resistance (Ohms):": (
            [r"ri\s*\(\s*ohms?\s*\)", r"resistance\s*\(?ohms?\)?\s*:?\s*"], 
            (2, 3)
        ),
        "Puffing Regime:": ([r"puffing\s*regime:?", r"puff\s*regime:?"], (2, 7)),
        "Initial Oil Mass:": ([r"initial\s*oil\s*mass:?"], (3, 7)),
        "Date:": ([r"date:?"], (1, 3)),
        "Media:": ([r"media:?"], (2, 1))
    }

    # Data column mapping relative to the sample block start.
    # Note: "after_weight" is at column offset 2.
    DATA_COL_MAPPING = {
        "puffs": (0, True),
        "before_weight": (1, True),
        "after_weight": (2, True),
        "draw_pressure": (3, True),
        "smell": (5, False),
        "notes": (7, False),
        "tpm": (8, True)
    }

    # Process each legacy sample.
    for sample_idx, sample in enumerate(legacy_samples):
        col_offset = 1 + (sample_idx * 12)
        #print(f"\nProcessing sample {sample_idx + 1} at columns {col_offset} to {col_offset + 11}")

        # --- 1. meta_data Handling ---
        sample_name = sample.get("sample_name", f"Sample {sample_idx + 1}")
        ws.cell(row=1, column=col_offset, value=os.path.splitext(os.path.basename(legacy_file_path))[0])

        meta_data_values = {
            "Sample ID:": sample.get("sample_name", ""),
            "Voltage:": sample.get("voltage", ""),
            "Viscosity:": sample.get("viscosity", ""),
            "Resistance (Ohms):": sample.get("resistance", ""),
            "Puffing Regime:": sample.get("puffing_regime", ""),
            "Initial Oil Mass:": sample.get("initial_oil_mass", ""),
            "Date:": sample.get("date", ""),
            "Media:": sample.get("media", "")
        }

        for template_key, (patterns, (tpl_row, tpl_col_offset)) in meta_data_MAPPING.items():
            value = meta_data_values.get(template_key, "")
            ws.cell(row=tpl_row, column=col_offset + tpl_col_offset, value=value)

        # --- 2. Data Column Handling with Row Slicing ---
        # Determine the cutoff row for this sample block based on "after_weight".
        raw_after_weight = sample.get("after_weight", pd.Series(dtype=object))
        raw_after_weight = raw_after_weight.reset_index(drop=True)
        cutoff = len(raw_after_weight)
        for i, val in enumerate(raw_after_weight):
            if pd.isna(val) or str(val).strip() == "":
                cutoff = i
                break
        #print(f"Sample {sample_idx + 1}: Writing {cutoff} data rows based on 'After weight/g' column.")

        # Now loop over each key in DATA_COL_MAPPING and write only rows up to the cutoff.
        for key, (rel_col, is_numeric) in DATA_COL_MAPPING.items():
            raw_data = sample.get(key, pd.Series(dtype=object))
            raw_data = raw_data.reset_index(drop=True)
            sliced_data = raw_data.iloc[:cutoff]
            if is_numeric:
                data = pd.to_numeric(sliced_data, errors="coerce")
            else:
                data = sliced_data.astype(str).replace("nan", "").replace("None", "")
            for row_offset, value in enumerate(data):
                target_row = 5 + row_offset  # Data starts at row 5.
                target_col = col_offset + rel_col
                try:
                    ws.cell(row=target_row, column=target_col, 
                            value=float(value) if is_numeric else str(value))
                except Exception as e:
                    #print(f"Error writing {key} at ({target_row},{target_col}): {e}")
                    ws.cell(row=target_row, column=target_col, value="ERROR")

        # --- 3. Clearing Extra Rows in the Block ---
        # For rows below the cutoff (in the sample block), clear all cells in the block's columns.
        # We assume the block occupies 12 columns starting at col_offset.
        start_clear_row = 5 + cutoff  # first row to clear in the block.
        # Use the maximum row in the worksheet as the ending point.
        for row_clear in range(start_clear_row, ws.max_row + 1):
            for col_clear in range(col_offset, col_offset + 12):
                ws.cell(row=row_clear, column=col_clear).value = None

    # --- 4. Final Cleanup ---
    # Delete any extra columns beyond the last sample block.
    last_sample_col = len(legacy_samples) * 12
    if ws.max_column > last_sample_col:
        ws.delete_cols(last_sample_col + 1, ws.max_column - last_sample_col)

    # Delete all sheets except the template sheet.
    sheets_to_keep = [template_sheet]
    for sheet_name in list(wb.sheetnames):
        if sheet_name not in sheets_to_keep:
            del wb[sheet_name]

    # Rename the sheet based on the legacy file name (limited to 31 characters).
    base_name = os.path.splitext(os.path.basename(legacy_file_path))[0]
    new_sheet_name = f"{base_name} Data"[:31]
    ws.title = new_sheet_name
    folder_path = os.path.join(os.path.abspath("."), "legacy data")
    
    new_file_name = f"{base_name} Legacy.xlsx"
    new_file_path = os.path.join(folder_path, new_file_name)
    wb.save(new_file_path)
    #print(f"\nSaved processed file to: {new_file_path}")
    return load_excel_file(new_file_path)[new_sheet_name]

def convert_legacy_standards_using_template(legacy_file_path: str, template_path: str = None) -> dict:
    """
    Converts a legacy standards Excel file to the standardized template format.
    """
    from openpyxl.cell.cell import MergedCell
    
    if template_path is None:
        template_path = os.path.join(os.path.abspath("."), "resources", 
                                    "Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
    
    wb_template = load_workbook(template_path, read_only=False)
    legacy_wb = load_workbook(legacy_file_path, read_only=False)
    base_name = os.path.splitext(os.path.basename(legacy_file_path))[0]
    folder_path = os.path.join(os.path.abspath("."), "legacy data")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    new_file_name = f"{base_name} Legacy Standards.xlsx"
    new_file_path = os.path.join(folder_path, new_file_name)
    
    SIMPLE_meta_data_MAPPING = {
        "sample_name": r"(cart(ridge)?\s*#|sample\s*(name|id))",
        "voltage": r"voltage",
        "viscosity": r"viscosity",
        "resistance": r"(ri\s*\(?\s*ohms?\s*\)?|resistance)",
        "puff_regime": r"puff(ing)?\s*regime",
        "initial_oil": r"initial\s*oil\s*mass",
        "date": r"date",
        "media": r"media"
    }
    
    for sheet_name in wb_template.sheetnames:
        # Create a new sheet to avoid merged cell issues
        new_sheet_name = f"{sheet_name}_new"
        new_ws = wb_template.create_sheet(title=new_sheet_name)
        
        if sheet_name in legacy_wb.sheetnames:
            legacy_ws = legacy_wb[sheet_name]
            legacy_df = read_sheet_with_values(legacy_file_path, sheet_name)
            
            # Check if this is a plotting sheet
            is_plotting = plotting_sheet_test(sheet_name, legacy_df)
            
            # Process based on sheet type
            if is_plotting:
                # For plotting sheets: copy the entire legacy sheet
                processed_df = legacy_df.copy()
                meta_data = extract_meta_data(legacy_ws, SIMPLE_meta_data_MAPPING)
                
                # Write to the new sheet
                for i, row in processed_df.iterrows():
                    for j, value in enumerate(row):
                        new_ws.cell(row=i + 1, column=j + 1, value=value)
            else:
                # For non-plotting sheets: copy the legacy sheet
                processed_df = legacy_df.copy()
                for i, row in processed_df.iterrows():
                    for j, value in enumerate(row):
                        new_ws.cell(row=i + 1, column=j + 1, value=value)
        else:
            # If sheet not in legacy file, copy from template sheet
            ws = wb_template[sheet_name]
            for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                for col_idx, value in enumerate(row, 1):
                    new_ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Remove original sheets and rename new ones
    for sheet_name in list(wb_template.sheetnames):
        if not sheet_name.endswith('_new'):
            del wb_template[sheet_name]
    
    for sheet_name in list(wb_template.sheetnames):
        if sheet_name.endswith('_new'):
            original_name = sheet_name[:-4]  # Remove _new suffix
            wb_template[sheet_name].title = original_name
    
    wb_template.save(new_file_path)
    return load_excel_file(new_file_path)
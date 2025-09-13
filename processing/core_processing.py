"""
core_processing.py
Developed by Charlie Becquet
Core Processing module for the DataViewer application.
"""
import pandas as pd
import numpy as np
from typing import List, Optional, Tuple, Any
from utils import (
    debug_print,
    validate_sheet_data,
    remove_empty_columns,
    round_values
)

from .sheet_processors import create_empty_plot_structure, create_empty_user_test_simulation_structure

from .data_extraction import updated_extracted_data_function_with_raw_data

# Module constants
DEFAULT_HEADERS_ROW = 3
DEFAULT_DATA_START_ROW = 4
DEFAULT_COLUMNS_PER_SAMPLE = 12

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
    # Import here to avoid circular imports
    from . import sheet_processors as sp
    
    return {
        "Test Plan": sp.process_test_plan,
        "Initial State Inspection": sp.process_initial_state_inspection,
        "Quick Screening Test": sp.process_quick_screening_test,
        "Lifetime Test": sp.process_device_life_test,
        "Device Life Test": sp.process_device_life_test,
        "Aerosol Temperature": sp.process_aerosol_temp_test,
        "User Test - Full Cycle": sp.process_user_test,
        "User Test Simulation": sp.process_user_test_simulation,
        "User Simulation Test": sp.process_user_test_simulation,
        "Horizontal Puffing Test": sp.process_horizontal_test,
        "Extended Test": sp.process_extended_test,
        "Long Puff Test": sp.process_long_puff_test,
        "Rapid Puff Test": sp.process_rapid_puff_test,
        "Intense Test": sp.process_intense_test,
        "Big Headspace Low T Test": sp.process_big_head_low_t_test,
        "Big Headspace Serial Test": sp.process_big_head_serial_test,
        "Anti-Burn Protection Test": sp.process_burn_protection_test,
        "Big Headspace High T Test": sp.process_big_head_high_t_test,
        "Upside Down Test": sp.process_upside_down_test,
        "Big Headspace Pocket Test": sp.process_pocket_test,
        "Temperature Cycling Test": sp.process_temperature_cycling_test,
        "High T High Humidity Test": sp.process_high_t_high_humidity_test,
        "Low Temperature Stability": sp.process_cold_storage_test,
        "Vacuum Test": sp.process_vacuum_test,
        "Negative Pressure Test": sp.process_vacuum_test,
        "Viscosity Compatibility": sp.process_viscosity_compatibility_test,
        "Various Oil Compatibility": sp.process_various_oil_test,
        "Quick Sensory Test": sp.process_quick_sensory_test,
        "Off-odor Score": sp.process_off_odor_score,
        "Sensory Consistency": sp.process_sensory_consistency,
        "Heavy Metal Leaching Test": sp.process_leaching_test,
        "Sheet1": sp.process_sheet1,
        "default": process_generic_sheet
    }
        
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

def get_y_data_for_plot_type(sample_data, plot_type):
    """
    Extract y-data for the specified plot type.
    Always calculates TPM and Power Efficiency on the fly for consistency.

    Args:
        sample_data (pd.DataFrame): DataFrame containing the sample data.
        plot_type (str): The type of plot to generate.

    Returns:
        pd.Series: y-data for plotting.
    """
    if plot_type == "TPM":
        #debug_print("DEBUG: Calculating TPM from weight differences with puffing intervals")

        # Get puffs, before weights, and after weights
        puffs = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce')  # Column 0 for Extended Test
        before_weights = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce')
        after_weights = pd.to_numeric(sample_data.iloc[3:, 2], errors='coerce')

        # Calculate puffing intervals
        puffing_intervals = pd.Series(index=puffs.index, dtype=float)
        for i, idx in enumerate(puffs.index):
            if i == 0:  # First row: current_puffs - 0
                current_puffs = puffs.loc[idx] if pd.notna(puffs.loc[idx]) else 10  # Default to 10 if NaN
                puffing_intervals.loc[idx] = current_puffs
            else:  # Subsequent rows: current_puffs - previous_puffs
                prev_idx = puffs.index[i-1]
                current_puffs = puffs.loc[idx] if pd.notna(puffs.loc[idx]) else 0
                prev_puffs = puffs.loc[prev_idx] if pd.notna(puffs.loc[prev_idx]) else 0
                puff_interval = current_puffs - prev_puffs

                # If interval is negative, zero, or current puffs is NaN, use current puffs or default
                if puff_interval <= 0 or pd.isna(current_puffs):
                    if pd.isna(current_puffs) or current_puffs == 0:
                        # Both puffs and interval are invalid, default to 10
                        fallback_puffs = 10
                        #debug_print(f"DEBUG: Row {i} - Both puffs ({current_puffs}) and interval ({puff_interval}) invalid, using default {fallback_puffs}")
                        puffing_intervals.loc[idx] = fallback_puffs
                    else:
                        # Use current puffs value
                        #debug_print(f"DEBUG: Row {i} - Interval {puff_interval} <= 0, using current puffs {current_puffs}")
                        puffing_intervals.loc[idx] = current_puffs
                else:
                    puffing_intervals.loc[idx] = puff_interval

        # Calculate TPM: (before - after) / puffing_interval * 1000 (mg/puff)
        weight_diff = (before_weights - after_weights) * 1000  # Convert g to mg
        calculated_tpm = weight_diff / puffing_intervals
        # Only calculate TPM where puffing_intervals > 0
        calculated_tpm = pd.Series(index=weight_diff.index, dtype=float)
        for idx in weight_diff.index:
            if pd.notna(puffing_intervals.loc[idx]) and puffing_intervals.loc[idx] > 0:
                calculated_tpm.loc[idx] = weight_diff.loc[idx] / puffing_intervals.loc[idx]
            else:
                calculated_tpm.loc[idx] = np.nan
                #debug_print(f"DEBUG: Skipping TPM calculation for row with invalid interval: {puffing_intervals.loc[idx]}")
        return calculated_tpm

    elif plot_type == "Normalized TPM":
        debug_print("DEBUG: Calculating Normalized TPM by dividing TPM by puff time")

        # Get regular TPM data first
        tpm_data = get_y_data_for_plot_type(sample_data, "TPM")
        tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')
        debug_print(f"DEBUG: Got {len(tpm_numeric.dropna())} TPM values for normalization")

        # Extract puffing regime from this sample's position: row 1 (index 0), column 8 (index 7)
        puff_time = None
        puffing_regime = None

        try:
            puffing_regime_cell = sample_data.iloc[0, 7]  # Row 1, Column 8 (H) for this sample
            if pd.notna(puffing_regime_cell):
                puffing_regime = str(puffing_regime_cell).strip()
                debug_print(f"DEBUG: Found puffing regime at [0,7]: '{puffing_regime}'")

                # Extract puff time using regex pattern - FIXED for case sensitivity
                import re
                pattern = r'mL/(\d+(?:\.\d+)?)s/'
                match = re.search(pattern, puffing_regime, re.IGNORECASE)  # ADDED re.IGNORECASE
                if match:
                    puff_time = float(match.group(1))
                    debug_print(f"DEBUG: Extracted puff time: {puff_time}s from '{puffing_regime}'")
                else:
                    debug_print(f"DEBUG: Could not extract puff time from pattern: '{puffing_regime}'")
            else:
                debug_print("DEBUG: No puffing regime found at expected position [0,7]")

        except (ValueError, IndexError, TypeError, AttributeError) as e:
            debug_print(f"DEBUG: Error extracting puffing regime from [0,7]: {e}")

        # Apply normalization if puff time was found, otherwise use default
        if puff_time is not None and puff_time > 0:
            normalized_tpm = tpm_numeric / puff_time
            debug_print(f"DEBUG: Successfully normalized TPM by puff time {puff_time}s")
            debug_print(f"DEBUG: Normalized TPM values: {normalized_tpm.dropna().tolist()}")
            debug_print(f"DEBUG: Original TPM range: {tpm_numeric.min():.3f} - {tpm_numeric.max():.3f} mg/puff")
            debug_print(f"DEBUG: Normalized TPM range: {normalized_tpm.min():.3f} - {normalized_tpm.max():.3f} mg/s")
            return normalized_tpm
        else:
            debug_print("DEBUG: Cannot calculate normalized TPM - puff time not found or invalid")
            debug_print("DEBUG: Using default puff time of 3.0s for normalization")
            default_puff_time = 3.0
            normalized_tpm = tpm_numeric / default_puff_time
            debug_print(f"DEBUG: Normalized TPM with default {default_puff_time}s: {normalized_tpm.dropna().tolist()}")
            return normalized_tpm

    elif plot_type == "Power Efficiency":
        debug_print("DEBUG: Calculating Power Efficiency from TPM/Power")

        # Get TPM first
        tpm_data = get_y_data_for_plot_type(sample_data, "TPM")
        tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')

        # Extract voltage and resistance from metadata
        voltage = None
        resistance = None

        try:
            voltage_val = sample_data.iloc[1, 5]  # Adjust column index as needed
            if pd.notna(voltage_val):
                voltage = float(voltage_val)
                debug_print(f"DEBUG: Extracted voltage: {voltage}V")
        except (ValueError, IndexError, TypeError):
            debug_print("DEBUG: Could not extract voltage")

        try:
            resistance_val = sample_data.iloc[0, 3]  # Adjust column index as needed
            if pd.notna(resistance_val):
                resistance = float(resistance_val)
                debug_print(f"DEBUG: Extracted resistance: {resistance}Ω")
        except (ValueError, IndexError, TypeError):
            debug_print("DEBUG: Could not extract resistance")

        # Calculate power and power efficiency
        if voltage and resistance and voltage > 0 and resistance > 0:
            power = (voltage ** 2) / resistance
            debug_print(f"DEBUG: Calculated power: {power:.3f}W")
            calculated_power_eff = tpm_numeric / power
            debug_print(f"DEBUG: Calculated Power Efficiency values: {calculated_power_eff.dropna().tolist()}")
            return calculated_power_eff
        else:
            debug_print("DEBUG: Cannot calculate power efficiency - missing or invalid voltage/resistance")
            return pd.Series(dtype=float)

    elif plot_type == "Draw Pressure":
        return pd.to_numeric(sample_data.iloc[3:, 3], errors='coerce')

    elif plot_type == "Resistance":
        return pd.to_numeric(sample_data.iloc[3:, 4], errors='coerce')

    else:
        # Default to TPM
        return get_y_data_for_plot_type(sample_data, "TPM")

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
            #debug_print(f"Insufficient rows for processing (headers_row={headers_row}, data_start_row={data_start_row})")
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

def process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12, custom_extracted_data_fn=None):
    """
    Process plotting sheets with fixed burn/clog/leak extraction from raw data.
    """
    debug_print(f"DEBUG: process_plot_sheet_fixed called with data shape: {data.shape}")

    # Store reference to raw data before cleaning
    raw_data = data.copy()

    # For data collection, allow minimal data (less strict validation)
    min_required_rows = max(headers_row + 1, 3)
    if data.shape[0] < min_required_rows:
        debug_print(f"DEBUG: Data has {data.shape[0]} rows, minimum required is {min_required_rows}")
        return create_empty_plot_structure(data, headers_row, num_columns_per_sample)

    try:
        # Clean up the data for processing but keep raw_data unchanged
        data = remove_empty_columns(data).replace(0, np.nan)
        debug_print(f"DEBUG: Data after cleaning: {data.shape}")

        samples = []
        full_sample_data = []
        sample_arrays = {}

        # Calculate the number of samples
        num_samples = data.shape[1] // num_columns_per_sample
        debug_print(f"DEBUG: Calculated {num_samples} samples")

        if num_samples == 0:
            debug_print("DEBUG: No samples detected, creating empty structure")
            return create_empty_plot_structure(data, headers_row, num_columns_per_sample)

        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            end_col = start_col + num_columns_per_sample
            sample_data = data.iloc[:, start_col:end_col]

            if sample_data.empty:
                debug_print(f"DEBUG: Sample {i+1} is empty. Skipping.")
                continue

            # Extract plotting data with error handling
            try:
                sample_arrays[f"Sample_{i+1}_Puffs"] = sample_data.iloc[3:, 0].to_numpy()
                sample_arrays[f"Sample_{i+1}_TPM"] = sample_data.iloc[3:, 8].to_numpy()

                # Use custom function if provided, otherwise use our fixed function
                if custom_extracted_data_fn:
                    # For compatibility, try to pass raw data if the function supports it
                    try:
                        extracted_data = custom_extracted_data_fn(sample_data, raw_data, i)
                    except TypeError:
                        # Fallback to old signature
                        extracted_data = custom_extracted_data_fn(sample_data)
                else:
                    # Check if sample has sufficient data before processing
                    if sample_data.shape[0] < 4 or sample_data.shape[1] < 3:
                        debug_print(f"DEBUG: Sample {i+1} has insufficient data shape {sample_data.shape}, creating placeholder")
                        placeholder_data = {
                            "Sample Name": f"Sample {i+1}",
                            "Media": "",
                            "Viscosity": "",
                            "Voltage, Resistance, Power": "",
                            "Average TPM": "No data",
                            "Standard Deviation": "No data",
                            "Initial Oil Mass": "",
                            "Usage Efficiency": "",
                            "Burn?": "",
                            "Clog?": "",
                            "Leak?": ""
                        }
                        samples.append(placeholder_data)
                        full_sample_data.append(sample_data)
                        continue

                    # If we get here, sample has sufficient data
                    extracted_data = updated_extracted_data_function_with_raw_data(sample_data, raw_data, i)

                samples.append(extracted_data)
                full_sample_data.append(sample_data)
                debug_print(f"DEBUG: Successfully processed sample {i+1}")

            except IndexError as e:
                debug_print(f"DEBUG: Index error for sample {i+1}: {e}. Creating placeholder.")
                # Create placeholder data for data collection
                placeholder_data = {
                    "Sample Name": f"Sample {i+1}",
                    "Media": "",
                    "Viscosity": "",
                    "Voltage, Resistance, Power": "",
                    "Average TPM": "No data",
                    "Standard Deviation": "No data",
                    "Initial Oil Mass": "",
                    "Usage Efficiency": "",
                    "Burn?": "",
                    "Clog?": "",
                    "Leak?": ""
                }
                samples.append(placeholder_data)
                full_sample_data.append(sample_data)
                continue

        # Create processed data and full sample data
        if samples:
            processed_data = pd.DataFrame(samples)
            full_sample_data_df = pd.concat(full_sample_data, axis=1) if full_sample_data else data
        else:
            debug_print("DEBUG: No valid samples found, creating minimal structure")
            processed_data = pd.DataFrame([{
                "Sample Name": "Sample 1",
                "Media": "",
                "Viscosity": "",
                "Voltage, Resistance, Power": "",
                "Average TPM": "No data",
                "Standard Deviation": "No data",
                "Initial Oil Mass": "",
                "Usage Efficiency": "",
                "Burn?": "",
                "Clog?": "",
                "Leak?": ""
            }])
            full_sample_data_df = data

        debug_print(f"DEBUG: Final processed_data shape: {processed_data.shape}")
        debug_print(f"DEBUG: Final full_sample_data shape: {full_sample_data_df.shape}")
        return processed_data, sample_arrays, full_sample_data_df

    except Exception as e:
        debug_print(f"DEBUG: Error processing plot sheet: {e}")
        import traceback
        traceback.print_exc()
        return create_empty_plot_structure(data, headers_row, num_columns_per_sample)

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
    # data = remove_empty_columns(data) - for now, let's not remove empty columns for non-plotting sheets.
    headers = data.iloc[headers_row, :].tolist()
    table_data = data.iloc[data_start_row:, :]
    table_data = table_data.replace(0, np.nan)
    processed_data = pd.DataFrame(table_data.values, columns=headers)
    return processed_data, table_data

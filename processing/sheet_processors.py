"""
sheet_processors.py
Developed by Charlie Becquet
Sheet Processor module for the DataViewer application.
"""

import pandas as pd
import numpy as np
import traceback
from utils import (
    debug_print,
    validate_sheet_data, 
    remove_empty_columns,
    round_values
)

from .core_processing import (
    process_generic_sheet, 
    process_plot_sheet,
    create_empty_plot_structure,
    create_empty_user_test_simulation_structure
)

from .data_extraction import (
    no_efficiency_extracted_data,
    updated_extracted_data_function_with_raw_data
)

# Module constants for sheet processing
STANDARD_HEADERS_ROW = 3
STANDARD_DATA_START_ROW = 4
STANDARD_COLUMNS_PER_SAMPLE = 12
USER_TEST_SIM_COLUMNS_PER_SAMPLE = 8


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

def process_big_head_serial_test(data):
    """Process data for the Big Headspace Serial Test sheet."""
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

def process_user_test_simulation(data):
    """
    Process data for the User Test Simulation sheet.
    This test has unique characteristics:
    - 8 columns per sample instead of 12
    - Puffs column is in position 1 (second column) instead of 0
    - Data needs to be split for plotting (0-50 puffs and remaining puffs)
    - Sample ID taken from column header 5 (index 4)
    - Extracts metadata from specific header locations
    """
    debug_print(f"DEBUG: process_user_test_simulation called with data shape: {data.shape}")

    # For data collection, allow minimal data (less strict validation)
    min_required_rows = max(3 + 1, 3)  # At least header row + 1
    if data.shape[0] < min_required_rows:
        debug_print(f"DEBUG: Data has {data.shape[0]} rows, minimum required is {min_required_rows}")
        return create_empty_user_test_simulation_structure(data)

    if not validate_sheet_data(data, required_rows=min_required_rows):
        debug_print("DEBUG: Sheet validation failed, creating empty structure")
        return create_empty_user_test_simulation_structure(data)

    try:
        # Clean up the data
        data = remove_empty_columns(data).replace(0, np.nan)
        debug_print(f"DEBUG: Data after cleaning: {data.shape}")

        samples = []
        full_sample_data = []
        sample_arrays = {}

        # Calculate the number of potential samples (8 columns per sample)
        num_columns_per_sample = 8
        potential_samples = data.shape[1] // num_columns_per_sample
        debug_print(f"DEBUG: Potential samples based on columns: {potential_samples}")

        # Process each sample block
        for i in range(potential_samples):
            start_col = i * num_columns_per_sample
            end_col = start_col + num_columns_per_sample
            sample_data = data.iloc[:, start_col:end_col]

            debug_print(f"DEBUG: Processing sample {i+1} in columns {start_col} to {end_col-1}")

            # Check if this sample block has real measurement data
            # Look for numeric data in puffs column (index 1) starting from row 3
            has_real_data = False
            if sample_data.shape[0] > 3 and sample_data.shape[1] > 7:
                puffs_data = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce').dropna()

                if not puffs_data.empty and len(puffs_data) > 0:
                    has_real_data = True
                    debug_print(f"DEBUG: Sample {i+1} has real measurement data - {len(puffs_data)} puff values")
                else:
                    debug_print(f"DEBUG: Sample {i+1} has no real measurement data")

            if not has_real_data:
                debug_print(f"DEBUG: Skipping sample {i+1} - no measurement data found")
                continue

            # Extract sample name from row 0, column 5 (Sample ID location)
            sample_name = "Sample " + str(len(samples) + 1)  # Default name
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 5:
                sample_id_value = sample_data.columns[5]  # Row 0, Column 5
                if sample_id_value and str(sample_id_value).strip() and str(sample_id_value).strip().lower() != 'nan':
                    sample_name = str(sample_id_value).strip()
                    debug_print(f"DEBUG: Extracted sample name '{sample_name}' from Sample ID location (row 0, col 5)")
                else:
                    debug_print(f"DEBUG: Sample ID location empty or invalid, using default name '{sample_name}'")

            # Extract metadata from specific header locations
            media = ""
            voltage = ""
            resistance = ""
            power = ""
            initial_oil_mass = ""

            try:
                # Media: Row 1, Column 1
                if sample_data.shape[0] > 1 and sample_data.shape[1] > 1:
                    media_val = sample_data.iloc[0, 1]
                    if media_val and str(media_val).strip().lower() != 'nan':
                        media = str(media_val).strip()
                        debug_print(f"DEBUG: Extracted media: '{media}'")

                # Voltage: Row 1, Column 5
                if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
                    voltage_val = sample_data.iloc[0, 5]
                    if voltage_val and str(voltage_val).strip().lower() != 'nan':
                        voltage = str(voltage_val).strip()
                        debug_print(f"DEBUG: Extracted voltage: '{voltage}'")

                # Initial Oil Mass: Row 0, Column 7
                if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
                    oil_mass_val = sample_data.columns[7]
                    if oil_mass_val and str(oil_mass_val).strip().lower() != 'nan':
                        initial_oil_mass = str(oil_mass_val).strip()
                        debug_print(f"DEBUG: Extracted initial oil mass: '{initial_oil_mass}'")

                # Power: Row 1, Column 7
                if sample_data.shape[0] > 1 and sample_data.shape[1] > 7:
                    power_val = sample_data.iloc[0, 7]
                    if power_val and str(power_val).strip().lower() != 'nan' and str(power_val).strip() != '#DIV/0!':
                        power = str(power_val).strip()
                        debug_print(f"DEBUG: Extracted power: '{power}'")


                # Resistance: Row 0, Column 3
                if sample_data.shape[0] > 1 and sample_data.shape[1] > 3:
                    resistance_val = sample_data.columns[3]
                    if resistance_val and str(resistance_val).strip().lower() != 'nan' and str(resistance_val).strip() != '#DIV/0!':
                        resistance = str(resistance_val).strip()
                        debug_print(f"DEBUG: Extracted resistance: '{resistance}'")


            except Exception as e:
                debug_print(f"DEBUG: Error extracting metadata for sample {i+1}: {e}")

            try:
                # Extract plotting data for User Test Simulation
                sample_arrays[f"Sample_{len(samples)+1}_Puffs"] = sample_data.iloc[3:, 1].to_numpy()  # Column 1 for puffs
                sample_arrays[f"Sample_{len(samples)+1}_TPM"] = sample_data.iloc[3:, 7].to_numpy()   # Column 7 for TPM

                # Calculate TPM statistics from all available data
                tpm_data = pd.to_numeric(sample_data.iloc[3:, 7], errors='coerce')
                valid_tpm = tpm_data.dropna()

                if len(valid_tpm) > 0:
                    avg_tpm = round_values(valid_tpm.mean())
                    std_tpm = round_values(valid_tpm.std()) if len(valid_tpm) > 1 else 0
                else:
                    avg_tpm = "No data"
                    std_tpm = "No data"

                debug_print(f"DEBUG: Sample '{sample_name}' TPM stats - Avg: {avg_tpm}, Std: {std_tpm}")

                usage_efficiency = ""
                try:
                    if sample_data.shape[0] > 1 and sample_data.shape[1] > 8:
                        efficiency_val = sample_data.iloc[1, 8]  # I3 = row 1, column 8 with -1 row indexing
                        if not pd.isna(efficiency_val) and str(efficiency_val).strip().lower() != 'nan':
                            # Handle both percentage and decimal formats
                            efficiency_str = str(efficiency_val).strip()
                            if '%' not in efficiency_str:
                                # If it's a decimal, convert to percentage
                                try:
                                    efficiency_num = float(efficiency_str)
                                    usage_efficiency = f"{round_values(efficiency_num * 100):.1f}%"
                                except:
                                    usage_efficiency = efficiency_str
                            else:
                                usage_efficiency = efficiency_str
                            debug_print(f"DEBUG: User Test Simulation - Extracted usage efficiency from I3: {usage_efficiency}")
                except Exception as e:
                    debug_print(f"DEBUG: User Test Simulation - Error extracting usage efficiency from I3: {e}")

                # Create voltage, resistance, power combined string
                voltage_resistance_power = ""
                components = []
                if voltage:
                    components.append(f"V: {voltage}")
                if resistance:
                    components.append(f"R: {resistance}")
                if power:
                    components.append(f"P: {power}")
                voltage_resistance_power = ", ".join(components)

                # Create extracted data for this sample
                extracted_data = {
                    "Sample Name": sample_name,
                    "Media": media,
                    "Viscosity": "",  # Not visible in the header image
                    "Voltage, Resistance, Power": voltage_resistance_power,
                    "Average TPM": avg_tpm,
                    "Standard Deviation": std_tpm,
                    "Initial Oil Mass": initial_oil_mass,
                    "Usage Efficiency": usage_efficiency,
                    "Test Type": "User Test Simulation"
                }

                samples.append(extracted_data)
                full_sample_data.append(sample_data)
                debug_print(f"DEBUG: Successfully processed User Test Simulation sample {len(samples)}: '{sample_name}'")
                debug_print(f"DEBUG: Sample data - Media: '{media}', Voltage: '{voltage}', Power: '{power}', Oil Mass: '{initial_oil_mass}'")

            except Exception as e:
                debug_print(f"DEBUG: Error processing sample {i+1}: {e}")
                continue

        # Create processed data and full sample data
        if samples:
            processed_data = pd.DataFrame(samples)
            full_sample_data_df = pd.concat(full_sample_data, axis=1) if full_sample_data else pd.DataFrame()
        else:
            debug_print("DEBUG: No valid samples processed, creating minimal structure")
            processed_data = pd.DataFrame([{
                "Sample Name": "Sample 1",
                "Media": "",
                "Viscosity": "",
                "Voltage, Resistance, Power": "",
                "Average TPM": "No data",
                "Standard Deviation": "No data",
                "Initial Oil Mass": "",
                "Usage Efficiency": "",
                "Test Type": "User Test Simulation"
            }])
            full_sample_data_df = data

        debug_print(f"DEBUG: Final User Test Simulation processed_data shape: {processed_data.shape}")
        debug_print(f"DEBUG: Final User Test Simulation full_sample_data shape: {full_sample_data_df.shape}")
        debug_print(f"DEBUG: process_plot_sheet - using concatenated data: {bool(samples)}")
        debug_print(f"DEBUG: process_plot_sheet - samples count: {len(samples) if samples else 0}")
        return processed_data, sample_arrays, full_sample_data_df

    except Exception as e:
        debug_print(f"DEBUG: Error processing User Test Simulation sheet: {e}")
        debug_print(f"DEBUG: Error traceback: {traceback.format_exc()}")
        debug_print(f"DEBUG: process_plot_sheet - using concatenated data: {bool(samples)}")
        debug_print(f"DEBUG: process_plot_sheet - samples count: {len(samples) if samples else 0}")
        return create_empty_user_test_simulation_structure(data)

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
    
    return {
        "Test Plan": process_test_plan,
        "Initial State Inspection": process_initial_state_inspection,
        "Quick Screening Test": process_quick_screening_test,
        "Lifetime Test": process_device_life_test,
        "Device Life Test": process_device_life_test,
        "Aerosol Temperature": process_aerosol_temp_test,
        "User Test - Full Cycle": process_user_test,
        "User Test Simulation": process_user_test_simulation,
        "User Simulation Test": process_user_test_simulation,
        "Horizontal Puffing Test": process_horizontal_test,
        "Extended Test": process_extended_test,
        "Long Puff Test": process_long_puff_test,
        "Rapid Puff Test": process_rapid_puff_test,
        "Intense Test": process_intense_test,
        "Big Headspace Low T Test": process_big_head_low_t_test,
        "Big Headspace Serial Test": process_big_head_serial_test,
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
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

def process_long_puff_lifetime_test(data):
    """Process data for the Long Puff Lifetime Test sheet."""
    # Similar to Long Puff Test but for lifetime assessment
    debug_print("DEBUG: Processing Long Puff Lifetime Test")
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

def process_rapid_puff_lifetime_test(data):
    """Process data for the Rapid Puff Lifetime Test sheet."""
    # Similar to Rapid Puff Test but for lifetime assessment
    debug_print("DEBUG: Processing Rapid Puff Lifetime Test")
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
    )

def process_temperature_cycling_test_2(data):
    """Process data for Temperature Cycling Test #2 sheet."""
    # Similar processing to other temperature cycling tests
    debug_print("DEBUG: Processing Temperature Cycling Test #2")
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12
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
    "Process data for User Simulation Test"
    return process_plot_sheet(
        data,
        headers_row=3,
        data_start_row=4,
        num_columns_per_sample=12)

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
        "Long Puff Lifetime Test": process_long_puff_lifetime_test,
        "Rapid Puff Test": process_rapid_puff_test,
        "Rapid Puff Lifetime Test": process_rapid_puff_lifetime_test,
        "Intense Test": process_intense_test,
        "Big Headspace Low T Test": process_big_head_low_t_test,
        "Big Headspace Serial Test": process_big_head_serial_test,
        "Anti-Burn Protection Test": process_burn_protection_test,
        "Big Headspace High T Test": process_big_head_high_t_test,
        "Upside Down Test": process_upside_down_test,
        "Big Headspace Pocket Test": process_pocket_test,
        "Temperature Cycling Test": process_temperature_cycling_test,
        "Temperature Cycling Test #1": process_temperature_cycling_test,
        "Temperature Cycling Test #2": process_temperature_cycling_test_2,
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
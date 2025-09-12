"""
data_extraction.py
Developed by Charlie Becquet
Data Extraction module for the DataViewer application.
"""

import re
import pandas as pd
import numpy as np
from typing import Tuple, Dict, Any
from utils import debug_print, round_values

# Module constants for data extraction
BURN_CLOG_LEAK_MAPPING = {
    0: "burn",
    1: "clog", 
    2: "leak"
}

DEFAULT_PUFF_TIME = 3.0  # Default puff time in seconds for normalization


def extract_burn_clog_leak_from_raw_data(data, sample_index, num_columns_per_sample=12):
    """
    Simple, direct extraction of burn/clog/leak data.
    - Burn: ALWAYS from column header at K (column 10 + offset)
    - Clog: ALWAYS from row 0, column K (column 10 + offset)
    - Leak: ALWAYS from row 1, column K (column 10 + offset)
    """
    col_offset = sample_index * num_columns_per_sample
    target_col = 10 + col_offset  # K column + offset

    print(f"DEBUG: Simple extraction for sample {sample_index + 1}, target column: {target_col}")

    def safe_get_cell(row, col, default=""):
        try:
            if row < len(data) and col < len(data.columns):
                val = data.iloc[row, col]
                if pd.isna(val):
                    return ""
                return str(val).strip()
            return default
        except:
            return default

    def safe_get_header(col, default=""):
        try:
            if col < len(data.columns):
                header = data.columns[col]
                if pd.isna(header):
                    return ""
                header_str = str(header).strip()
                # If it's "Unnamed", return empty - otherwise return the actual value
                if header_str.startswith('Unnamed'):
                    return ""
                return header_str
            return default
        except:
            return default

    # Direct extraction - no searching, no guessing
    burn = safe_get_header(target_col)     # Column header
    clog = safe_get_cell(0, target_col)    # Row 0, same column
    leak = safe_get_cell(1, target_col)    # Row 1, same column

    print(f"  - Column header: '{data.columns[target_col] if target_col < len(data.columns) else 'N/A'}'")
    print(f"  - Row 0 value: '{safe_get_cell(0, target_col)}'")
    print(f"  - Row 1 value: '{safe_get_cell(1, target_col)}'")
    print(f"  - Extracted burn (header): '{burn}'")
    print(f"  - Extracted clog (row 0): '{clog}'")
    print(f"  - Extracted leak (row 1): '{leak}'")

    return burn, clog, leak

def no_efficiency_extracted_data(sample_data, raw_data, sample_index):
    """
    Custom function to extract data without efficiency metrics, using enhanced sample name extraction with suffix support.
    """
    print(f"DEBUG: Processing sample {sample_index + 1} (no efficiency) with enhanced name extraction")

    # Extract sample name with old format support (same logic as above)
    sample_name = f"Sample {sample_index + 1}"  # Default fallback

    if sample_data.shape[1] > 8:  # Ensure we have enough columns
        headers = sample_data.columns.astype(str)

        # Look for old format patterns with suffix support
        project_value = None
        sample_value = None

        for i, header in enumerate(headers):
            header_lower = str(header).lower().strip()

            # Enhanced pattern matching to handle suffixed headers
            if header_lower.startswith("project:") and i + 1 < len(headers):
                project_value = str(headers[i + 1]).strip()
                # Remove the suffix from the project value if it exists
                if project_value.endswith(f'.{sample_index}'):
                    project_value = project_value[:-len(f'.{sample_index}')]
            elif header_lower.startswith("sample:") and i + 1 < len(headers):
                sample_value = str(headers[i + 1]).strip()

        # Combine project and sample values if found (old format)
        if project_value and sample_value:
            if project_value.lower() not in ['nan', 'none', ''] and sample_value.lower() not in ['nan', 'none', '']:
                sample_name = f"{project_value} {sample_value}".strip()
        elif project_value and project_value.lower() not in ['nan', 'none', '']:
            sample_name = project_value.strip()
        elif sample_value and sample_value.lower() not in ['nan', 'none', '']:
            sample_name = sample_value.strip()
        else:
            # Try new format - Sample ID at column 5
            if len(headers) > 5:
                sample_id_value = str(headers[5]).strip()
                if sample_id_value and sample_id_value.lower() not in ['nan', 'none', '', 'unnamed: 5']:
                    sample_name = sample_id_value

    tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
    avg_tpm = round_values(tpm_data.mean()) if not tpm_data.empty else None
    std_tpm = round_values(tpm_data.std()) if not tpm_data.empty else None

    # Extract burn/clog/leak using the new method
    burn, clog, leak = extract_burn_clog_leak_from_raw_data(raw_data, sample_index)

    return {
        "Sample Name": sample_name,  # Now uses enhanced extraction
        "Media": sample_data.iloc[0, 1] if sample_data.shape[0] > 0 else "",
        "Viscosity": sample_data.iloc[1, 1] if sample_data.shape[0] > 1 else "",
        "Voltage, Resistance, Power": f"{sample_data.iloc[1, 5]} V, "
                                        f"{round_values(sample_data.iloc[0, 3])} ohm, "
                                        f"{round_values(sample_data.iloc[0, 5])} W"
                                        if sample_data.shape[0] > 1 and sample_data.shape[1] > 5 else "",
        "Average TPM": avg_tpm,
        "Standard Deviation": std_tpm,
        "Burn?": burn,
        "Clog?": clog,
        "Leak?": leak
    }

def get_y_data_for_user_test_simulation_plot_type(sample_data, plot_type):
    """
    Extract y-data for the specified plot type for User Test Simulation.
    Always calculates TPM and Power Efficiency on the fly for consistency.
    Adjusted for 8-column layout instead of 12-column.
    """
    if plot_type == "TPM":
        debug_print("DEBUG: User Test Simulation - Calculating TPM from weight differences with puffing intervals")

        puffs = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce')  # Column 1 for User Test Simulation
        before_weights = pd.to_numeric(sample_data.iloc[3:, 2], errors='coerce')
        after_weights = pd.to_numeric(sample_data.iloc[3:, 3], errors='coerce')

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
                        debug_print(f"DEBUG: Row {i} - Both puffs ({current_puffs}) and interval ({puff_interval}) invalid, using default {fallback_puffs}")
                        puffing_intervals.loc[idx] = fallback_puffs
                    else:
                        # Use current puffs value
                        debug_print(f"DEBUG: Row {i} - Interval {puff_interval} <= 0, using current puffs {current_puffs}")
                        puffing_intervals.loc[idx] = current_puffs
                else:
                    puffing_intervals.loc[idx] = puff_interval

        # Calculate TPM: (before - after) / puffing_interval * 1000 (mg/puff)
        weight_diff = (before_weights - after_weights) * 1000  # Convert g to mg
        calculated_tpm = weight_diff / puffing_intervals

        for idx in weight_diff.index:
            if pd.notna(puffing_intervals.loc[idx]) and puffing_intervals.loc[idx] > 0:
                calculated_tpm.loc[idx] = weight_diff.loc[idx] / puffing_intervals.loc[idx]
            else:
                calculated_tpm.loc[idx] = np.nan
                debug_print(f"DEBUG: User Test Simulation - Skipping TPM calculation for row with invalid interval: {puffing_intervals.loc[idx]}")

        return calculated_tpm

    elif plot_type == "Power Efficiency":
        debug_print("DEBUG: User Test Simulation - Calculating Power Efficiency from TPM/Power")

        # Get TPM first
        tpm_data = get_y_data_for_user_test_simulation_plot_type(sample_data, "TPM")
        tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')

        # Extract voltage and resistance from metadata (adjust for User Test Simulation layout)
        voltage = None
        resistance = None

        try:
            voltage_val = sample_data.iloc[0, 5]  # Adjust as needed
            if pd.notna(voltage_val):
                voltage = float(voltage_val)
                debug_print(f"DEBUG: User Test Simulation - Extracted voltage: {voltage}V")
        except (ValueError, IndexError, TypeError):
            debug_print("DEBUG: User Test Simulation - Could not extract voltage")

        try:
            resistance_val = sample_data.columns[3]  # Adjust as needed
            if pd.notna(resistance_val):
                resistance = float(resistance_val)
                debug_print(f"DEBUG: User Test Simulation - Extracted resistance: {resistance}Ω")
        except (ValueError, IndexError, TypeError):
            debug_print("DEBUG: User Test Simulation - Could not extract resistance")

        # Calculate power and power efficiency
        if voltage and resistance and voltage > 0 and resistance > 0:
            power = (voltage ** 2) / resistance
            debug_print(f"DEBUG: User Test Simulation - Calculated power: {power:.3f}W")
            calculated_power_eff = tpm_numeric / power
            debug_print(f"DEBUG: User Test Simulation - Calculated Power Efficiency values: {calculated_power_eff.dropna().tolist()}")
            return calculated_power_eff
        else:
            debug_print("DEBUG: User Test Simulation - Cannot calculate power efficiency - missing or invalid voltage/resistance")
            return pd.Series(dtype=float)

    elif plot_type == "Draw Pressure":
        # For User Test Simulation, Draw Pressure is in column 4 (adjust as needed)
        return pd.to_numeric(sample_data.iloc[3:, 4], errors='coerce')

    else:
        # Default to TPM for any other plot type
        return get_y_data_for_user_test_simulation_plot_type(sample_data, "TPM")

def fix_x_axis_sequence(x_data):
    """
    Fix x-axis data to ensure it's always increasing.
    If x2-x1 < 0, set x2 = x1+10 to avoid plotting issues.
    This is critical for User Test Simulation where puff counts can reset.

    Args:
        x_data (pd.Series): The x-axis data (usually puffs)

    Returns:
        pd.Series: Fixed x-axis data with monotonic increasing values
    """
    try:
        if len(x_data) <= 1:
            return x_data

        # Convert to numeric first if it's not already
        if x_data.dtype == 'object':  # String data
            x_numeric = pd.to_numeric(x_data, errors='coerce').dropna()
            debug_print(f"DEBUG: Converted string x_data to numeric: {len(x_numeric)} values")
        else:
            x_numeric = x_data.copy()

        if len(x_numeric) <= 1:
            return x_numeric

        # Convert to list for easier manipulation
        x_values = x_numeric.values.copy()

        # Fix any decreasing sequences
        fixes_applied = 0
        for i in range(1, len(x_values)):
            if x_values[i] - x_values[i-1] < 0:
                # If current value is less than previous, set it to previous + 10
                old_value = x_values[i]
                x_values[i] = x_values[i-1] + 10
                fixes_applied += 1
                #debug_print(f"DEBUG: Fixed x-axis sequence at index {i}: {old_value} -> {x_values[i]} (was decreasing)")
            elif x_values[i] == x_values[i-1]:
                # If values are equal, add small increment to maintain increasing sequence
                old_value = x_values[i]
                x_values[i] = x_values[i-1] + 5
                fixes_applied += 1
                #debug_print(f"DEBUG: Fixed duplicate x-axis value at index {i}: {old_value} -> {x_values[i]}")

        if fixes_applied > 0:
            debug_print(f"DEBUG: Applied {fixes_applied} fixes to x-axis sequence")

        # Return as pandas Series with original index
        return pd.Series(x_values, index=x_numeric.index)

    except Exception as e:
        debug_print(f"DEBUG: Error fixing x-axis sequence: {e}")
        return x_data  # Return original data if fixing fails

def updated_extracted_data_function_with_raw_data(sample_data, raw_data, sample_index):
    """
    Updated extraction function with new header structure.
    Enhanced to handle both old and new template formats for sample names with suffix support.
    Now includes Usage Efficiency, Normalized TPM, and Initial Oil Mass.
    """
    print(f"DEBUG: Processing sample {sample_index + 1} with enhanced name extraction and new fields")
    print(f"DEBUG: Sample {sample_index + 1} data shape: {sample_data.shape}")

    # Check if sample has sufficient data before processing
    if sample_data.shape[0] < 4 or sample_data.shape[1] < 3:
        print(f"DEBUG: Sample {sample_index + 1} has insufficient data shape {sample_data.shape}, skipping")
        return {
            "Sample Name": f"Sample {sample_index + 1}",
            "Media": "",
            "Viscosity": "",
            "Puffing Regime": "",
            "Voltage, Resistance, Power": "",
            "Average TPM": "No data",
            "Standard Deviation": "No data",
            "Normalized TPM": "",
            "Draw Pressure": "",
            "Usage Efficiency": "",
            "Initial Oil Mass": "",
            "Burn": "",
            "Clog": "",
            "Notes": ""
        }

    # Extract sample name - use the same logic that works in plotting functions
    sample_name = f"Sample {sample_index + 1}"  # Default fallback
    try:
        if sample_data.shape[1] > 5:
            sample_name_candidate = sample_data.columns[5]
            if pd.notna(sample_name_candidate):
                sample_name_str = str(sample_name_candidate).strip()
                if sample_name_str and sample_name_str.lower() not in ['nan', 'none', '', 'unnamed: 5']:
                    sample_name = sample_name_str
                    print(f"DEBUG: Extracted sample name from columns[5]: '{sample_name}'")
                else:
                    print(f"DEBUG: Sample name at columns[5] was invalid: '{sample_name_str}', using default")
            else:
                print(f"DEBUG: Sample name at columns[5] was NaN, using default")
        else:
            print(f"DEBUG: Not enough columns for sample name extraction, using default")
    except Exception as e:
        print(f"DEBUG: Error extracting sample name for sample {sample_index + 1}: {e}")
        sample_name = f"Sample {sample_index + 1}"

    print(f"DEBUG: Final sample name for sample {sample_index + 1}: '{sample_name}'")


    # Extract TPM data for calculations
    tpm_data = pd.Series(dtype=float)
    if sample_data.shape[0] > 3 and sample_data.shape[1] > 8:
        tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
    avg_tpm = round_values(tpm_data.mean()) if not tpm_data.empty else None
    std_tpm = round_values(tpm_data.std()) if not tpm_data.empty else None

    # Extract draw pressure (average from column 3, similar to TPM)
    draw_pressure_data = pd.Series(dtype=float)
    if sample_data.shape[0] > 3 and sample_data.shape[1] > 3:
        draw_pressure_data = pd.to_numeric(sample_data.iloc[3:, 3], errors='coerce').dropna()
    avg_draw_pressure = round_values(draw_pressure_data.mean()) if not draw_pressure_data.empty else ""

    # Extract burn/clog/leak using existing method
    burn, clog, leak = extract_burn_clog_leak_from_raw_data(raw_data, sample_index)

    # Extract puffing regime from header data (row 1, column 7)
    puffing_regime = ""
    if sample_data.shape[0] > 1 and sample_data.shape[1] > 7:
        puffing_regime = str(sample_data.iloc[0, 7]) if pd.notna(sample_data.iloc[0, 7]) else ""

    # Extract notes from header data (check multiple locations)
    notes = ""
    if sample_data.shape[0] > 2 and sample_data.shape[1] > 7:
        # Check row 2, column 7 for notes
        notes_candidate = str(sample_data.iloc[2, 7]) if pd.notna(sample_data.iloc[2, 7]) else ""
        if notes_candidate and not notes_candidate.lower().startswith('unnamed'):
            notes = notes_candidate

    # Keep the existing combined format for voltage, resistance, power
    voltage_resistance_power = ""
    if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
        voltage = sample_data.iloc[1, 5]
        resistance = round_values(sample_data.iloc[0, 3]) if sample_data.shape[1] > 3 else ""
        power = round_values(sample_data.iloc[0, 5])
        voltage_resistance_power = f"{voltage} V, {resistance} ohm, {power} W"

    # NEW: Calculate the three missing fields
    initial_oil_mass = extract_initial_oil_mass(sample_data)
    usage_efficiency = calculate_usage_efficiency_for_sample(sample_data)
    normalized_tpm = calculate_normalized_tpm_for_sample(sample_data, tpm_data)

    print(f"DEBUG: Calculated new fields for sample {sample_index + 1}:")
    print(f"  - Initial Oil Mass: '{initial_oil_mass}'")
    print(f"  - Usage Efficiency: '{usage_efficiency}'")
    print(f"  - Normalized TPM: '{normalized_tpm}'")

    return {
        "Sample Name": sample_name,
        "Media": sample_data.iloc[0, 1] if sample_data.shape[0] > 0 else "",
        "Viscosity": sample_data.iloc[1, 1] if sample_data.shape[0] > 1 else "",
        "Puffing Regime": puffing_regime,
        "Voltage, Resistance, Power": voltage_resistance_power,  # Combined as requested
        "Average TPM": avg_tpm,
        "Standard Deviation": std_tpm,
        "Normalized TPM": normalized_tpm,  # NEW FIELD
        "Draw Pressure": avg_draw_pressure,
        "Usage Efficiency": usage_efficiency,  # NEW FIELD
        "Initial Oil Mass": initial_oil_mass,  # NEW FIELD
        "Burn": burn,
        "Clog": clog,
        "Notes": notes
    }

def updated_extracted_data_function_with_raw_data_old(sample_data, raw_data, sample_index):
    """
    Updated extraction function that gets burn/clog/leak from raw data.
    Enhanced to handle both old and new template formats for sample names with suffix support.
    """
    print(f"DEBUG: Processing sample {sample_index + 1} with enhanced name extraction")
    print(f"DEBUG: Sample {sample_index + 1} data shape: {sample_data.shape}")

    # Check if sample has sufficient data before processing
    if sample_data.shape[0] < 4 or sample_data.shape[1] < 3:
        print(f"DEBUG: Sample {sample_index + 1} has insufficient data shape {sample_data.shape}, skipping")
        return {
            "Sample Name": f"Sample {sample_index + 1}",
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

    # Extract sample name with enhanced suffix handling
    sample_name = f"Sample {sample_index + 1}"  # Default fallback

    if sample_data.shape[1] > 4:  # Ensure we have enough columns
        headers = sample_data.columns.astype(str)

        # Look for sample ID patterns with suffix support
        project_value = None
        sample_value = None

        for i, header in enumerate(headers):
            header_lower = str(header).lower().strip()

            # Enhanced pattern matching to handle suffixed headers
            # Check for "sample id:" patterns (handles "sample id:", "sample id:.1", etc.)
            if (header_lower == "sample id:" or
                (header_lower.startswith("sample id:.") and header_lower[11:].isdigit())):
                if i + 1 < len(headers):
                    sample_value = str(headers[i + 1]).strip()
                    break

            # Check for old format "project:" patterns
            elif (header_lower == "project:" or
                    (header_lower.startswith("project:.") and header_lower[9:].isdigit())):
                if i + 1 < len(headers):
                    project_value = str(headers[i + 1]).strip()
                    # Remove pandas suffix from the project value if it exists
                    if '.' in project_value and project_value.split('.')[-1].isdigit():
                        project_value = '.'.join(project_value.split('.')[:-1])

            # Check for old format "sample:" patterns
            elif (header_lower == "sample:" or
                    (header_lower.startswith("sample:.") and header_lower[8:].isdigit())):
                if i + 1 < len(headers):
                    temp_sample_value = str(headers[i + 1]).strip()
                    # Don't overwrite if we already found one from Sample ID
                    if not sample_value:
                        sample_value = temp_sample_value

        # Determine final sample name based on what we found
        if sample_value and sample_value.lower() not in ['nan', 'none', '', f'unnamed: {5}']:
            if project_value and project_value.lower() not in ['nan', 'none', '']:
                # Old format: combine project + sample
                sample_name = f"{project_value} {sample_value}".strip()
            else:
                # New format or standalone sample value
                sample_name = sample_value.strip()
        elif project_value and project_value.lower() not in ['nan', 'none', '']:
            sample_name = project_value.strip()
        else:
            # Fallback: try direct column 5 access (for new format)
            if len(headers) > 5:
                fallback_value = str(headers[5]).strip()
                if fallback_value and fallback_value.lower() not in ['nan', 'none', '', 'unnamed: 5']:
                    sample_name = fallback_value

    print(f"DEBUG: Final sample name for sample {sample_index + 1}: '{sample_name}'")

    # Extract burn/clog/leak from raw data with proper offsets
    burn, clog, leak = extract_burn_clog_leak_from_raw_data(raw_data, sample_index)

    # Calculate TPM statistics from sample_data with bounds checking
    tpm_data = pd.Series(dtype=float)
    if sample_data.shape[0] > 3 and sample_data.shape[1] > 8:
        tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
    avg_tpm = round_values(tpm_data.mean()) if not tpm_data.empty else None
    std_tpm = round_values(tpm_data.std()) if not tpm_data.empty else None

    # Extract media and viscosity with bounds checking
    media = ""
    viscosity = ""
    if sample_data.shape[0] > 1 and sample_data.shape[1] > 1:
        media = str(sample_data.iloc[0, 1]) if not pd.isna(sample_data.iloc[0, 1]) else ""
        viscosity = str(sample_data.iloc[1, 1]) if not pd.isna(sample_data.iloc[1, 1]) else ""

    # Extract voltage, resistance, power with bounds checking
    voltage = ""
    resistance = ""
    power = ""
    if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
        voltage = str(sample_data.iloc[1, 5]) if not pd.isna(sample_data.iloc[1, 5]) else ""
        if sample_data.shape[1] > 3:
            resistance = round_values(sample_data.iloc[0, 3]) if not pd.isna(sample_data.iloc[0, 3]) else ""
        power = round_values(sample_data.iloc[0, 5]) if not pd.isna(sample_data.iloc[0, 5]) else ""

    # Calculate usage efficiency using the actual Excel formula logic
    usage_efficiency = ""
    calculated_efficiency = None

    try:
        if sample_data.shape[0] > 3 and sample_data.shape[1] > 8:
            # Get initial oil mass from H3 (column 7, row 1 with -1 indexing)
            initial_oil_mass_val = None
            if sample_data.shape[1] > 7:
                initial_oil_mass_val = sample_data.iloc[1, 7]

            # Get puffs values from column A (column 0) starting from row 4 (row 3 with -1 indexing)
            puffs_values = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce').dropna()

            # Get TPM values from column I (column 8) starting from row 4 (row 3 with -1 indexing)
            tpm_values = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()

            if (initial_oil_mass_val is not None and not pd.isna(initial_oil_mass_val) and
                len(puffs_values) > 0 and len(tpm_values) > 0):

                try:
                    initial_oil_mass_mg = float(initial_oil_mass_val) * 1000  # Convert g to mg

                    # Calculate total aerosol mass following Excel formula logic
                    total_aerosol_mass_mg = 0

                    # Align the arrays to same length
                    min_length = min(len(puffs_values), len(tpm_values))
                    puffs_aligned = puffs_values.iloc[:min_length]
                    tpm_aligned = tpm_values.iloc[:min_length]

                    for i in range(min_length):
                        tpm_val = tpm_aligned.iloc[i]
                        puffs_val = puffs_aligned.iloc[i]

                        if not pd.isna(tpm_val) and not pd.isna(puffs_val):
                            if i == 0:
                                # First measurement: TPM * total puffs
                                aerosol_mass = tpm_val * puffs_val
                            else:
                                # Subsequent measurements: TPM * (current_puffs - previous_puffs)
                                previous_puffs = puffs_aligned.iloc[i-1]
                                if not pd.isna(previous_puffs):
                                    incremental_puffs = puffs_val - previous_puffs
                                    aerosol_mass = tpm_val * incremental_puffs
                                else:
                                    aerosol_mass = tpm_val * puffs_val

                            total_aerosol_mass_mg += aerosol_mass

                    # Calculate usage efficiency: (total aerosol mass / initial oil mass) * 100
                    if initial_oil_mass_mg > 0:
                        calculated_efficiency = (total_aerosol_mass_mg / initial_oil_mass_mg) * 100
                        usage_efficiency = f"{round_values(calculated_efficiency):.1f}%"

                        print(f"DEBUG: Usage efficiency calculation for sample {sample_index + 1}:")
                        print(f"  - Initial oil mass: {initial_oil_mass_val}g ({initial_oil_mass_mg}mg)")
                        print(f"  - Data points: {min_length}")
                        print(f"  - Total aerosol mass: {round_values(total_aerosol_mass_mg):.2f}mg")
                        print(f"  - Calculated usage efficiency: {usage_efficiency}")
                    else:
                        print(f"DEBUG: Invalid initial oil mass for sample {sample_index + 1}: {initial_oil_mass_val}")

                except (ValueError, TypeError) as e:
                    print(f"DEBUG: Error converting values for sample {sample_index + 1}: {e}")
            else:
                print(f"DEBUG: Missing data for sample {sample_index + 1}: oil_mass={initial_oil_mass_val}, puffs_count={len(puffs_values)}, tpm_count={len(tpm_values)}")

    except Exception as e:
        print(f"DEBUG: Error calculating usage efficiency for sample {sample_index + 1}: {e}")


    return {
        "Sample Name": sample_name,
        "Media": media,
        "Viscosity": viscosity,
        "Voltage, Resistance, Power": f"{voltage} V, {resistance} ohm, {power} W" if voltage or resistance or power else "",
        "Average TPM": avg_tpm,
        "Standard Deviation": std_tpm,
        "Initial Oil Mass": initial_oil_mass_val,
        "Usage Efficiency": usage_efficiency,
        "Burn?": burn,
        "Clog?": clog,
        "Leak?": leak
    }

def calculate_normalized_tpm_for_sample(sample_data, tpm_data):
    """
    Calculate normalized TPM by dividing TPM by puff time.
    Reuses the logic from the plotting functions.

    Args:
        sample_data: DataFrame containing sample data
        tpm_data: Series of TPM values

    Returns:
        str: Formatted normalized TPM value or empty string
    """
    debug_print("DEBUG: Calculating Normalized TPM for data extraction")

    try:
        # Convert TPM data to numeric
        tpm_numeric = pd.to_numeric(tpm_data, errors='coerce').dropna()
        if tpm_numeric.empty:
            debug_print("DEBUG: No valid TPM data for normalization")
            return ""

        # Extract puffing regime from row 1, column 8 (index 0, 7)
        puff_time = None
        puffing_regime = None

        if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
            puffing_regime_cell = sample_data.iloc[0, 7]  # Row 1, Column 8 (H)
            if pd.notna(puffing_regime_cell):
                puffing_regime = str(puffing_regime_cell).strip()
                debug_print(f"DEBUG: Found puffing regime: '{puffing_regime}'")

                # Extract puff time using regex pattern
                import re
                pattern = r'mL/(\d+(?:\.\d+)?)s/'
                match = re.search(pattern, puffing_regime, re.IGNORECASE)
                if match:
                    puff_time = float(match.group(1))
                    debug_print(f"DEBUG: Extracted puff time: {puff_time}s")
                else:
                    debug_print(f"DEBUG: Could not extract puff time from: '{puffing_regime}'")

        # Apply normalization
        if puff_time is not None and puff_time > 0:
            avg_normalized_tpm = (tpm_numeric / puff_time).mean()
            result = f"{round_values(avg_normalized_tpm):.2f}"
            debug_print(f"DEBUG: Calculated normalized TPM: {result} mg/s")
            return result
        else:
            debug_print("DEBUG: Using default puff time of 3.0s for normalization")
            default_puff_time = 3.0
            avg_normalized_tpm = (tmp_numeric / default_puff_time).mean()
            result = f"{round_values(avg_normalized_tpm):.2f}"
            debug_print(f"DEBUG: Calculated normalized TPM with default: {result} mg/s")
            return result

    except Exception as e:
        debug_print(f"DEBUG: Error calculating normalized TPM: {e}")
        return ""

def calculate_usage_efficiency_for_sample(sample_data):
    """
    Calculate usage efficiency using the Excel formula logic.
    Formula: ((first_tpm * first_puffs + sum(tpm * incremental_puffs)) / 1000) / initial_oil_mass * 100

    Args:
        sample_data: DataFrame containing sample data

    Returns:
        str: Formatted usage efficiency percentage or empty string
    """
    debug_print("DEBUG: Calculating usage efficiency for sample")

    try:
        if sample_data.shape[0] < 4 or sample_data.shape[1] < 9:
            debug_print(f"DEBUG: Insufficient data shape {sample_data.shape} for usage efficiency calculation")
            return ""

        # Get initial oil mass from H3 (column 7, row 1 with -1 indexing = row 2)
        initial_oil_mass_val = sample_data.iloc[1, 7]
        if pd.isna(initial_oil_mass_val) or initial_oil_mass_val == 0:
            debug_print(f"DEBUG: Invalid initial oil mass: {initial_oil_mass_val}")
            return ""

        # Get puffs values from column A (column 0) starting from row 4 (row 3 with -1 indexing)
        puffs_values = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce').dropna()

        # Get TPM values from column I (column 8) starting from row 4 (row 3 with -1 indexing)
        tpm_values = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()

        if len(puffs_values) == 0 or len(tpm_values) == 0:
            debug_print(f"DEBUG: No valid puffs ({len(puffs_values)}) or TPM ({len(tpm_values)}) data")
            return ""

        initial_oil_mass_mg = float(initial_oil_mass_val) * 1000  # Convert g to mg
        debug_print(f"DEBUG: Initial oil mass: {initial_oil_mass_val}g ({initial_oil_mass_mg}mg)")

        # Calculate total aerosol mass following Excel formula logic
        total_aerosol_mass_mg = 0

        # Align the arrays to same length
        min_length = min(len(puffs_values), len(tpm_values))
        puffs_aligned = puffs_values.iloc[:min_length]
        tpm_aligned = tpm_values.iloc[:min_length]

        debug_print(f"DEBUG: Processing {min_length} data points for efficiency calculation")

        for i in range(min_length):
            tpm_val = tpm_aligned.iloc[i]
            puffs_val = puffs_aligned.iloc[i]

            if not pd.isna(tpm_val) and not pd.isna(puffs_val):
                if i == 0:
                    # First measurement: TPM * total puffs
                    aerosol_mass = tpm_val * puffs_val
                    debug_print(f"DEBUG: Row {i}: First measurement - {tpm_val} * {puffs_val} = {aerosol_mass}")
                else:
                    # Subsequent measurements: TPM * (current_puffs - previous_puffs)
                    previous_puffs = puffs_aligned.iloc[i-1]
                    if not pd.isna(previous_puffs):
                        incremental_puffs = puffs_val - previous_puffs
                        aerosol_mass = tpm_val * incremental_puffs
                        debug_print(f"DEBUG: Row {i}: {tpm_val} * ({puffs_val} - {previous_puffs}) = {aerosol_mass}")
                    else:
                        aerosol_mass = tpm_val * puffs_val
                        debug_print(f"DEBUG: Row {i}: Previous puffs NaN, using {tpm_val} * {puffs_val} = {aerosol_mass}")

                total_aerosol_mass_mg += aerosol_mass

        # Calculate usage efficiency: (total aerosol mass / initial oil mass) * 100
        if initial_oil_mass_mg > 0:
            calculated_efficiency = (total_aerosol_mass_mg / initial_oil_mass_mg) * 100
            usage_efficiency = f"{round_values(calculated_efficiency):.1f}%"

            debug_print(f"DEBUG: Usage efficiency calculation complete:")
            debug_print(f"  - Total aerosol mass: {round_values(total_aerosol_mass_mg):.2f}mg")
            debug_print(f"  - Initial oil mass: {initial_oil_mass_mg}mg")
            debug_print(f"  - Calculated usage efficiency: {usage_efficiency}")

            return usage_efficiency
        else:
            debug_print(f"DEBUG: Invalid initial oil mass for efficiency calculation: {initial_oil_mass_mg}")
            return ""

    except Exception as e:
        debug_print(f"DEBUG: Error calculating usage efficiency: {e}")
        import traceback
        traceback.print_exc()
        return ""

def extract_initial_oil_mass(sample_data):
    """
    Extract initial oil mass from sample data.

    Args:
        sample_data: DataFrame containing sample data

    Returns:
        str: Initial oil mass value or empty string
    """
    try:
        if sample_data.shape[0] > 1 and sample_data.shape[1] > 7:
            initial_oil_mass_val = sample_data.iloc[1, 7]  # Row 2, Column 8 (H)
            if pd.notna(initial_oil_mass_val):
                debug_print(f"DEBUG: Extracted initial oil mass: {initial_oil_mass_val}")
                return str(initial_oil_mass_val)
        debug_print("DEBUG: Could not extract initial oil mass")
        return ""
    except Exception as e:
        debug_print(f"DEBUG: Error extracting initial oil mass: {e}")
        return ""

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
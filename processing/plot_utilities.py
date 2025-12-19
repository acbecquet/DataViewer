"""
plot_utilities.py
Developed by Charlie Becquet
Plot Utlity processing module for the DataViewer application.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from typing import List, Tuple, Optional
from utils import debug_print, wrap_text
from processing.data_extraction import (
    get_y_data_for_user_test_simulation_plot_type,
    fix_x_axis_sequence
)
from processing.core_processing import get_y_data_for_plot_type

# Module constants for plotting
DEFAULT_FIGURE_SIZE = (8, 6)
SPLIT_PLOT_FIGURE_SIZE = (16, 6)
BAR_CHART_FIGURE_SIZE = (14, 6)
MAX_LABEL_WIDTH = 12
DEFAULT_Y_LIMIT = 9


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
        "Normalized TPM": 'Normalized TPM (mg/s)',
        "Draw Pressure": 'Draw Pressure (kPa)',
        "Resistance": 'Resistance (Ohms)',
        "Power Efficiency": 'Power Efficiency (mg/W)',
    }
    return y_label_mapping.get(plot_type, 'TPM (mg/puff)')  # Default to TPM

def plot_user_test_simulation_samples(full_sample_data: pd.DataFrame, num_columns_per_sample: int, plot_type: str, sample_names: List[str] = None) -> Tuple[plt.Figure, List[str]]:
    """
    Generate split plots for User Test Simulation.
    Creates two separate plots: one for puffs 0-50 (first 5 rows) and one for remaining puffs (9th row onwards).

    Args:
        full_sample_data (pd.DataFrame): DataFrame containing sample data.
        num_columns_per_sample (int): Number of columns per sample.
        plot_type (str): Type of plot to generate.
        sample_names (List[str], optional): List of sample names to use in legend.
    """
    debug_print(f"DEBUG: plot_user_test_simulation_samples called with data shape: {full_sample_data.shape}")
    debug_print(f"DEBUG: Provided sample_names: {sample_names}")
    debug_print(f"DEBUG: Full sample data first few rows:")
    debug_print(full_sample_data.iloc[:5, :15].to_string())

    num_samples = full_sample_data.shape[1] // num_columns_per_sample
    debug_print(f"DEBUG: User Test Simulation - Number of samples: {num_samples}")

    # Check if this should be a bar chart
    if plot_type == "TPM (Bar)":
        debug_print("DEBUG: Creating User Test Simulation bar chart")
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        extracted_sample_names = plot_user_test_simulation_bar_chart(ax1, ax2, full_sample_data, num_samples, num_columns_per_sample, sample_names)

        # Mark this as a split plot with bar chart data
        fig.is_split_plot = True
        fig.is_bar_chart = True

        # Store bar references for checkbox functionality
        fig.phase1_bars = ax1.patches
        fig.phase2_bars = ax2.patches

        debug_print(f"DEBUG: Successfully created User Test Simulation bar chart with {len(extracted_sample_names)} samples")
        return fig, extracted_sample_names

    # Original line plot logic for other plot types
    debug_print(f"DEBUG: User Test Simulation - Number of samples: {num_samples}")

    # Replace 0 with NaN for cleaner plotting
    full_sample_data = full_sample_data.replace(0, np.nan)

    # Create subplots for split plotting
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    extracted_sample_names = []
    phase1_lines = []
    phase2_lines = []
    y_max = 0

    for i in range(num_samples):
        start_col = i * num_columns_per_sample
        sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]

        debug_print(f"DEBUG: Processing sample {i+1} columns {start_col} to {start_col + num_columns_per_sample - 1}")

        # Extract puffs data (column index 1 in User Test Simulation)
        x_data = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce').dropna()
        debug_print(f"DEBUG: Sample {i+1} puffs data length: {len(x_data)}, values: {x_data.head().tolist()}")

        # Extract y-data based on plot type
        y_data = get_y_data_for_user_test_simulation_plot_type(sample_data, plot_type)
        y_data = pd.to_numeric(y_data, errors='coerce').dropna()
        debug_print(f"DEBUG: Sample {i+1} y_data length: {len(y_data)}, values: {y_data.head().tolist()}")

        # Ensure x and y data have common indices
        common_index = x_data.index.intersection(y_data.index)
        if common_index.empty:
            debug_print(f"DEBUG: Sample {i+1} SKIPPED - no common data points")
            continue

        x_data = x_data.loc[common_index]
        y_data = y_data.loc[common_index]

        x_data = fix_x_axis_sequence(x_data)
        debug_print(f"DEBUG: Sample {i+1} fixed puffs data length: {len(x_data)}, values: {x_data.head().tolist()}")


        # Use provided sample names if available, otherwise use default
        if sample_names and i < len(sample_names):
            sample_name = sample_names[i]
            debug_print(f"DEBUG: Using provided sample name: '{sample_name}'")
        else:
            sample_name = f"Sample {i+1}"
            debug_print(f"DEBUG: Using default sample name: '{sample_name}'")

        extracted_sample_names.append(sample_name)

        # Split data into two phases
        data_start_index = 3  # Data starts at row 3

        # Phase 1: First 5 data rows (puffs 0-50)
        phase1_end_index = data_start_index + 5  # First 5 data rows
        phase1_mask = x_data.index < phase1_end_index
        phase1_x = x_data[phase1_mask]
        phase1_y = y_data[phase1_mask]

        # Phase 2: From 9th data row onwards (extended puffs)
        phase2_start_index = data_start_index + 8  # 9th data row (0-indexed: 8th)
        phase2_mask = x_data.index >= phase2_start_index
        phase2_x = x_data[phase2_mask]
        phase2_y = y_data[phase2_mask]

        debug_print(f"DEBUG: Sample {i+1} Phase 1 data points: {len(phase1_x)}, Phase 2 data points: {len(phase2_x)}")

        # Plot Phase 1 (0-50 puffs) and store line reference
        if not phase1_x.empty and not phase1_y.empty:
            line1 = ax1.plot(phase1_x, phase1_y, marker='o', label=sample_name)[0]
            phase1_lines.append(line1)
            y_max = max(y_max, phase1_y.max())
            debug_print(f"DEBUG: Plotted Phase 1 for {sample_name} with {len(phase1_x)} points")
        else:
            # Add placeholder line for consistency
            line1 = ax1.plot([], [], marker='o', label=sample_name)[0]
            phase1_lines.append(line1)
            debug_print(f"DEBUG: Added placeholder Phase 1 line for {sample_name}")

        # Plot Phase 2 (remaining puffs) and store line reference
        if not phase2_x.empty and not phase2_y.empty:
            line2 = ax2.plot(phase2_x, phase2_y, marker='o', label=sample_name)[0]
            phase2_lines.append(line2)
            y_max = max(y_max, phase2_y.max())
            debug_print(f"DEBUG: Plotted Phase 2 for {sample_name} with {len(phase2_x)} points")
        else:
            # Add placeholder line for consistency
            line2 = ax2.plot([], [], marker='o', label=sample_name)[0]
            phase2_lines.append(line2)
            debug_print(f"DEBUG: Added placeholder Phase 2 line for {sample_name}")

    debug_print(f"DEBUG: Final sample names for legend: {extracted_sample_names}")

    # Configure Phase 1 plot
    ax1.set_xlabel('Puffs')
    ax1.set_ylabel(get_y_label_for_plot_type(plot_type))
    ax1.set_title(f'{plot_type} - Phase 1 (Puffs 0-50)')
    ax1.legend(loc='upper right')
    ax1.set_xlim(0, 60)  # Slightly wider than 50 for better visualization
    prevent_x_label_overlap(ax1)

    # Configure Phase 2 plot
    ax2.set_xlabel('Puffs')
    ax2.set_ylabel(get_y_label_for_plot_type(plot_type))
    ax2.set_title(f'{plot_type} - Phase 2 (Extended Puffs)')
    ax2.legend(loc='upper right')
    prevent_x_label_overlap(ax2)

    # Set consistent y-axis limits
    if plot_type == "Normalized TPM":
        ax1.set_ylim(-0.2, 4)
        ax2.set_ylim(-0.2, 4)
        debug_print("DEBUG: Set User Test Simulation Normalized TPM y-limits to -0.2 to 4")
    elif y_max > 9 and y_max <= 50:
        ax1.set_ylim(0, y_max)
        ax2.set_ylim(0, y_max)
    else:
        ax1.set_ylim(0, 9)
        ax2.set_ylim(0, 9)

    plt.tight_layout()

    # Store both sets of lines for checkbox functionality
    fig.phase1_lines = phase1_lines
    fig.phase2_lines = phase2_lines
    fig.is_split_plot = True

    debug_print(f"DEBUG: Successfully created User Test Simulation split plot with {len(extracted_sample_names)} samples")
    return fig, extracted_sample_names

def plot_user_test_simulation_bar_chart(ax1, ax2, full_sample_data, num_samples, num_columns_per_sample, sample_names=None):
    """
    Generate split bar charts for User Test Simulation average TPM with error bars.
    Creates two bar charts: Phase 1 (0-50 puffs) and Phase 2 (extended puffs).
    Each bar shows average TPM with standard deviation as error bars.
    """
    debug_print(f"DEBUG: plot_user_test_simulation_bar_chart called with {num_samples} samples")

    phase1_averages = []
    phase1_std_devs = []
    phase2_averages = []
    phase2_std_devs = []
    labels = []
    extracted_sample_names = []

    for i in range(num_samples):
        start_col = i * num_columns_per_sample
        end_col = start_col + num_columns_per_sample
        sample_data = full_sample_data.iloc[:, start_col:end_col]

        debug_print(f"DEBUG: Processing sample {i+1} for bar chart, columns {start_col} to {end_col-1}")

        # Check if sample has valid data (check puffs column - column 1 for User Test Simulation)
        if sample_data.shape[0] <= 3 or pd.isna(sample_data.iloc[3, 1]):
            debug_print(f"DEBUG: Sample {i+1} has no valid data, skipping")
            continue

        # Calculate TPM values using the same method as the line plots
        tpm_data = get_y_data_for_user_test_simulation_plot_type(sample_data, "TPM")
        tpm_numeric = pd.to_numeric(tpm_data, errors='coerce').dropna()

        debug_print(f"DEBUG: Sample {i+1} TPM data length: {len(tpm_numeric)}, values: {tpm_numeric.head().tolist()}")

        # Extract puffs data to understand the sequence
        puffs_data = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce').dropna()
        if not puffs_data.empty:
            # Fix the puff sequence
            fixed_puffs = fix_x_axis_sequence(puffs_data)
            debug_print(f"DEBUG: Sample {i+1} fixed puffs sequence for phase splitting")

        if tpm_numeric.empty:
            debug_print(f"DEBUG: Sample {i+1} has no valid TPM data, skipping")
            continue

        # Split TPM data into phases based on the same logic as line plots
        # Phase 1: First 5 data points (corresponding to 0-50 puffs)
        phase1_tpm = tpm_numeric.iloc[:5] if len(tpm_numeric) >= 5 else tpm_numeric
        # Phase 2: From 9th data point onwards (index 8+, corresponding to extended puffs)
        phase2_tpm = tpm_numeric.iloc[8:] if len(tpm_numeric) > 8 else pd.Series(dtype=float)

        # Calculate averages and standard deviations
        phase1_avg = phase1_tpm.mean() if not phase1_tpm.empty else 0
        phase1_std = phase1_tpm.std() if len(phase1_tpm) > 1 else 0
        phase2_avg = phase2_tpm.mean() if not phase2_tpm.empty else 0
        phase2_std = phase2_tpm.std() if len(phase2_tpm) > 1 else 0

        phase1_averages.append(phase1_avg)
        phase1_std_devs.append(phase1_std)
        phase2_averages.append(phase2_avg)
        phase2_std_devs.append(phase2_std)


        if sample_names and i < len(sample_names):
            sample_name = sample_names[i]
            debug_print(f"DEBUG: Using provided sample name: '{sample_name}'")
        else:
            sample_name = f"Sample {i+1}"
            debug_print(f"DEBUG: Using default sample name: '{sample_name}'")

        extracted_sample_names.append(sample_name)
        wrapped_name = wrap_text(text=sample_name, max_width=10)
        labels.append(wrapped_name)

        debug_print(f"DEBUG: Sample {i+1} - Phase 1: avg={phase1_avg:.3f}, std={phase1_std:.3f}, Phase 2: avg={phase2_avg:.3f}, std={phase2_std:.3f}")

    if not phase1_averages:
        debug_print("DEBUG: No valid samples found for bar chart")
        return []

    # Create colormaps for unique colors
    num_bars = len(phase1_averages)
    colors = plt.cm.get_cmap('tab10', num_bars)(np.linspace(0, 1, num_bars))

    # Create numeric positions for bars to prevent overlapping
    x_positions = np.arange(len(phase1_averages))

    # Plot Phase 1 bar chart with error bars
    bars1 = ax1.bar(x_positions, phase1_averages, yerr=phase1_std_devs,
                        color=colors, capsize=5, error_kw={'elinewidth': 2, 'capthick': 2}, width=0.6)
    ax1.set_xlabel('Samples')
    ax1.set_ylabel('Average TPM (mg/puff)')
    ax1.set_title('Average TPM - Phase 1 (Puffs 0-50)')

    # Set x-axis ticks and labels properly
    ax1.set_xticks(x_positions)
    ax1.set_xticklabels(labels, rotation=45, ha='right')
    ax1.tick_params(axis='x', labelsize=8)

    # Set y-axis to start from 0 and add some padding
    ax1.set_ylim(0, max(phase1_averages) * 1.2 if phase1_averages else 1)

    # Plot Phase 2 bar chart with error bars
    bars2 = ax2.bar(x_positions, phase2_averages, yerr=phase2_std_devs,
                        color=colors, capsize=5, error_kw={'elinewidth': 2, 'capthick': 2}, width=0.6)
    ax2.set_xlabel('Samples')
    ax2.set_ylabel('Average TPM (mg/puff)')
    ax2.set_title('Average TPM - Phase 2 (Extended Puffs)')

    # Set x-axis ticks and labels properly
    ax2.set_xticks(x_positions)
    ax2.set_xticklabels(labels, rotation=45, ha='right')
    ax2.tick_params(axis='x', labelsize=8)

    # Set y-axis to start from 0 and add some padding
    ax2.set_ylim(0, max(phase2_averages) * 1.2 if phase2_averages else 1)

    debug_print(f"DEBUG: Created User Test Simulation bar charts with {len(extracted_sample_names)} samples")
    return extracted_sample_names

def plot_all_samples(full_sample_data: pd.DataFrame, plot_type: str, num_columns_per_sample: int = 12, sample_names: List[str] = None) -> Tuple[plt.Figure, List[str]]:
    """
    Generate plots for all samples in the provided data.

    Args:
        full_sample_data (pd.DataFrame): DataFrame containing sample data.
        plot_type (str): Type of plot to generate.
        num_columns_per_sample (int): Number of columns per sample.
        sample_names (List[str], optional): List of sample names to use in legend.

    Returns:
        tuple: (matplotlib.figure.Figure, list of sample names)
    """
    # ADD THESE DEBUG LINES:
    debug_print(f"DEBUG: plot_all_samples - full_sample_data shape: {full_sample_data.shape}")
    debug_print(f"DEBUG: plot_all_samples - provided sample_names: {sample_names}")
    debug_print(f"DEBUG: plot_all_samples - first 5 rows, first 15 columns:")
    debug_print(full_sample_data.iloc[:5, :15].to_string())
    debug_print("=" * 80)

    # ENSURE num_columns_per_sample is an integer
    try:
        num_columns_per_sample = int(num_columns_per_sample)
        debug_print(f"DEBUG:   num_columns_per_sample after int conversion: {num_columns_per_sample}")
    except (ValueError, TypeError) as e:
        print(f"ERROR: Could not convert num_columns_per_sample to int: {e}")
        raise ValueError(f"num_columns_per_sample must be convertible to int, got: {num_columns_per_sample} (type: {type(num_columns_per_sample)})")

    # Original logic for standard tests (12 columns per sample)
    num_samples = full_sample_data.shape[1] // num_columns_per_sample
    full_sample_data = full_sample_data.replace(0, np.nan)
    fig, ax = plt.subplots(figsize=(8, 6))
    extracted_sample_names = []

    if plot_type == "TPM (Bar)":
        extracted_sample_names = plot_tpm_bar_chart(ax, full_sample_data, num_samples, num_columns_per_sample, sample_names)
    else:
        y_max = 0
        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]

            #debug_print(f"DEBUG: Sample {i+1} data shape: {sample_data.shape}")
            #debug_print(f"DEBUG: Sample {i+1} columns: {sample_data.columns.tolist()}")

            x_data = sample_data.iloc[3:, 0].dropna()
            #debug_print(f"DEBUG: Sample {i+1} x_data (puffs) length: {len(x_data)}, first few values: {x_data.head().tolist()}")

            y_data = get_y_data_for_plot_type(sample_data, plot_type)
            #debug_print(f"DEBUG: Sample {i+1} raw y_data length: {len(y_data)}, first few values: {y_data.head().tolist()}")

            y_data = pd.to_numeric(y_data, errors='coerce').dropna()
            #debug_print(f"DEBUG: Sample {i+1} numeric y_data length: {len(y_data)}, first few values: {y_data.head().tolist()}")

            common_index = x_data.index.intersection(y_data.index)
            #debug_print(f"DEBUG: Sample {i+1} common_index length: {len(common_index)}")

            x_data = x_data.loc[common_index]
            y_data = y_data.loc[common_index]

            x_data = fix_x_axis_sequence(x_data)
            #debug_print(f"DEBUG: Sample {i+1} fixed x_data length: {len(x_data)}, first few values: {x_data.head().tolist()}")

            #debug_print(f"DEBUG: Sample {i+1} final x_data length: {len(x_data)}, y_data length: {len(y_data)}")

            if not x_data.empty and not y_data.empty:
                # Use provided sample name if available, otherwise extract from data
                if sample_names and i < len(sample_names):
                    sample_name = sample_names[i]
                    debug_print(f"DEBUG: Using provided sample name: '{sample_name}'")
                else:
                    # Extract sample name from column 5
                    if len(sample_data.columns) > 5:
                        sample_name_candidate = sample_data.columns[5]
            
                        # Check if valid: not empty, not NaN, and doesn't start with 'Unnamed'
                        if (pd.notna(sample_name_candidate) and 
                            str(sample_name_candidate).strip() and 
                            not str(sample_name_candidate).lower().startswith('unnamed')):
                            sample_name = str(sample_name_candidate).strip()
                            debug_print(f"DEBUG: Using extracted sample name: '{sample_name}'")
                        else:
                            sample_name = f"Sample {i+1}"
                            debug_print(f"DEBUG: Using default sample name: '{sample_name}'")
                    else:
                        sample_name = f"Sample {i+1}"
                        debug_print(f"DEBUG: Using default sample name (column index out of range): '{sample_name}'")

                ax.plot(x_data, y_data, marker='o', label=sample_name)
                extracted_sample_names.append(sample_name)
                y_max = max(y_max, y_data.max())
            else:
                debug_print(f"DEBUG: Sample {i+1} SKIPPED - x_data empty: {x_data.empty}, y_data empty: {y_data.empty}")

        ax.set_xlabel('Puffs')
        ax.set_ylabel(get_y_label_for_plot_type(plot_type))
        ax.set_title(plot_type)
        ax.legend(loc='upper right')

        prevent_x_label_overlap(ax)

        # Set y-axis limits based on plot type
        if plot_type == "Normalized TPM":
            ax.set_ylim(-0.2, 4)
            debug_print("DEBUG: Set Normalized TPM y-limits to -0.2 to 4")
        elif y_max > 9 and y_max <= 50:
            ax.set_ylim(0, y_max)
        else:
            ax.set_ylim(0, 9)

    return fig, extracted_sample_names

def plot_tpm_bar_chart(ax, full_sample_data, num_samples, num_columns_per_sample, sample_names=None):
    """
    Generate a bar chart of the average TPM for each sample with wrapped sample names.

    Args:
        ax (matplotlib.axes.Axes): Matplotlib axis to draw on.
        full_sample_data (pd.DataFrame): DataFrame of sample data.
        num_samples (int): Number of samples.
        num_columns_per_sample (int): Columns per sample.
        sample_names (list): Optional list of sample names.

    Returns:
        list: Sample names.
    """
    averages = []
    labels = []
    extracted_sample_names = []

    for i in range(num_samples):
        start_col = i * num_columns_per_sample
        end_col = start_col + num_columns_per_sample
        sample_data = full_sample_data.iloc[:, start_col:end_col]

        if pd.isna(sample_data.iloc[3, 1]):
            continue

        tpm_data = sample_data.iloc[3:, 8].dropna()  # TPM column (adjust if needed)
        average_tpm = tpm_data.mean()
        averages.append(average_tpm)

        # Use provided sample name if available, otherwise extract from data
        if sample_names and i < len(sample_names):
            sample_name = sample_names[i]
            debug_print(f"DEBUG: Using provided sample name: '{sample_name}'")
        else:
            sample_name = sample_data.columns[5] if len(sample_data.columns) > 5 else f"Sample {i+1}"
            debug_print(f"DEBUG: Using extracted sample name: '{sample_name}'")

        extracted_sample_names.append(sample_name)
        wrapped_name = wrap_text(text=sample_name, max_width=10)  # Use `wrap_text` to dynamically wrap names
        labels.append(wrapped_name)

    if not averages:
        debug_print("DEBUG: No valid samples found for bar chart")
        return []

    # Create numeric positions for bars
    x_positions = np.arange(len(averages))

    # Create a colormap with unique colors
    num_bars = len(averages)
    colors = plt.cm.get_cmap('tab10', num_bars)(np.linspace(0, 1, num_bars))

    # Plot the bar chart with proper positioning
    bars = ax.bar(x_positions, averages, color=colors, width=0.6)

    # Set the x-axis labels and ticks
    ax.set_xticks(x_positions)
    ax.set_xticklabels(labels, rotation=45, ha='right')

    ax.set_xlabel('Samples')
    ax.set_ylabel('Average TPM (mg/puff)')
    ax.set_title('Average TPM per Sample')
    ax.tick_params(axis='x', labelsize=8)

    # Set y-axis to start from 0 and add some padding
    ax.set_ylim(0, max(averages) * 1.2 if averages else 1)

    return extracted_sample_names

def prevent_x_label_overlap(ax):
    """
    Hide overlapping x-axis labels to improve readability.
    Simple solution that doesn't modify tick positions, just visibility.
    """
    try:
        # Get current labels and their positions
        labels = ax.get_xticklabels()
        if len(labels) <= 1:
            return

        # Calculate approximate label widths
        fig = ax.get_figure()
        if fig is None:
            return

        # Get the axis width in points
        bbox = ax.get_window_extent()
        axis_width_points = bbox.width

        # Estimate average label width (rough approximation)
        avg_label_length = sum(len(label.get_text()) for label in labels) / len(labels)
        estimated_label_width = avg_label_length * 12  # roughly 12 points per character

        # Calculate how many labels can fit without overlap
        if estimated_label_width > 0:
            max_labels = max(1, int(axis_width_points / (estimated_label_width * 1.3)))  # 1.2 for spacing

            if len(labels) > max_labels:
                # Calculate step to show evenly distributed labels
                step = max(1, len(labels) // max_labels)

                # Hide labels that would cause overlap
                for i, label in enumerate(labels):
                    if i % step != 0:
                        label.set_visible(False)

                debug_print(f"DEBUG: Prevented x-label overlap - showing every {step} labels ({max_labels} total)")
            else:
                debug_print(f"DEBUG: No x-label overlap detected - showing all {len(labels)} labels")
    except Exception as e:
        debug_print(f"DEBUG: Error in prevent_x_label_overlap: {e}")
        # If anything goes wrong, just leave labels as they are
        pass

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



"""
services/plot_service.py
Consolidated plotting service for chart generation and management.
This consolidates all plotting functionality from plot_manager.py and processing.py plotting functions.
"""

# Standard library imports
import os
import tempfile
import threading
from typing import Optional, Dict, Any, List, Tuple, Union, Callable
from dataclasses import dataclass

# Third party imports (lazy loaded)
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox


def debug_print(message: str):
    """Debug print function for plotting operations."""
    print(f"DEBUG: PlotService - {message}")


def wrap_text(text: str, max_width: int = 12) -> str:
    """Wrap text for display in checkboxes."""
    if len(text) <= max_width:
        return text
    
    words = text.split()
    lines = []
    current_line = ""
    
    for word in words:
        if len(current_line + " " + word) <= max_width:
            current_line += (" " + word) if current_line else word
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    
    if current_line:
        lines.append(current_line)
    
    return "\n".join(lines)


@dataclass
class PlotConfiguration:
    """Configuration for plot generation."""
    plot_type: str = "TPM"
    figure_size: Tuple[float, float] = (8, 6)
    num_columns_per_sample: int = 12
    enable_checkboxes: bool = True
    enable_zoom: bool = True
    enable_toolbar: bool = True
    y_limits: Optional[Tuple[float, float]] = None
    x_limits: Optional[Tuple[float, float]] = None


class PlotService:
    """
    Consolidated service for plot generation and management.
    Handles matplotlib integration, data visualization, and interactive plot features.
    """
    
    def __init__(self, calculation_service=None):
        """Initialize the plot service."""
        debug_print("Initializing PlotService")
        
        # Core configuration
        self.supported_plot_types = [
            "TPM", "Draw Pressure", "Resistance", "Power Efficiency", 
            "TPM (Bar)", "Normalized TPM"
        ]
        
        # Service dependencies
        self.calculation_service = calculation_service
        
        # Plot state management
        self.current_figure = None
        self.current_axes = None
        self.current_canvas = None
        self.current_lines = None
        self.check_buttons = None
        self.checkbox_cid = None
        
        # Interactive features
        self.line_labels = []
        self.original_lines_data = []
        self.label_mapping = {}
        self.enable_interactions = True
        
        # Threading locks
        self.plot_lock = threading.Lock()
        
        # Lazy-loaded dependencies
        self._matplotlib_components = None
        self._pandas = None
        
        # Y-axis label mapping
        self.y_label_mapping = {
            "TPM": 'TPM (mg/puff)',
            "Normalized TPM": 'Normalized TPM (mg/s)',
            "Draw Pressure": 'Draw Pressure (kPa)',
            "Resistance": 'Resistance (Ohms)',
            "Power Efficiency": 'Power Efficiency (mg/W)',
        }
        
        debug_print("PlotService initialized successfully")
        debug_print(f"Supported plot types: {', '.join(self.supported_plot_types)}")
    
    # ===================== DEPENDENCY MANAGEMENT =====================
    
    def _lazy_import_matplotlib(self):
        """Lazy import matplotlib components."""
        if self._matplotlib_components is None:
            try:
                import matplotlib.pyplot as plt
                import matplotlib
                matplotlib.use('TkAgg')
                
                from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
                from matplotlib.widgets import CheckButtons
                
                self._matplotlib_components = (plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons)
                debug_print("Matplotlib components loaded successfully")
            except ImportError as e:
                debug_print(f"WARNING: Matplotlib not available: {e}")
                self._matplotlib_components = (None, None, None, None)
        return self._matplotlib_components
    
    def _lazy_import_pandas(self):
        """Lazy import pandas."""
        if self._pandas is None:
            try:
                import pandas as pd
                self._pandas = pd
                debug_print("Pandas loaded successfully")
            except ImportError:
                debug_print("WARNING: Pandas not available")
                self._pandas = None
        return self._pandas
    
    def get_matplotlib_components(self):
        """Get matplotlib components with error handling."""
        return self._lazy_import_matplotlib()
    
    # ===================== CORE PLOTTING FUNCTIONS =====================
    
    def plot_all_samples(self, frame: tk.Widget, full_sample_data: pd.DataFrame, 
                        num_columns_per_sample: int = 12, plot_type: str = "TPM",
                        sample_names: List[str] = None) -> Tuple[bool, Optional[Any], str]:
        """
        Main plotting function that generates plots for all samples.
        Handles both standard plots and User Test Simulation split plots.
        """
        debug_print(f"Plotting all samples: plot_type={plot_type}, columns_per_sample={num_columns_per_sample}")
        debug_print(f"Data shape: {full_sample_data.shape}")
        
        try:
            with self.plot_lock:
                # Clear existing plot
                self.clear_plot_area(frame)
                
                # Validate inputs
                if full_sample_data.empty:
                    return False, None, "No data provided for plotting"
                
                if plot_type not in self.supported_plot_types:
                    return False, None, f"Unsupported plot type: {plot_type}"
                
                # Ensure num_columns_per_sample is integer
                try:
                    num_columns_per_sample = int(num_columns_per_sample)
                except (ValueError, TypeError):
                    return False, None, f"Invalid columns per sample: {num_columns_per_sample}"
                
                # Check if this is User Test Simulation (8 columns per sample)
                if num_columns_per_sample == 8:
                    debug_print("Detected User Test Simulation - using split plotting")
                    fig, extracted_sample_names = self._plot_user_test_simulation_samples(
                        full_sample_data, num_columns_per_sample, plot_type, sample_names
                    )
                else:
                    debug_print("Using standard plotting")
                    fig, extracted_sample_names = self._plot_standard_samples(
                        full_sample_data, num_columns_per_sample, plot_type, sample_names
                    )
                
                if fig is None:
                    return False, None, "Failed to generate plot"
                
                # Embed plot in frame
                canvas = self._embed_plot_in_frame(fig, frame)
                if canvas is None:
                    return False, None, "Failed to embed plot in frame"
                
                # Store plot state
                self.current_figure = fig
                self.current_canvas = canvas
                self.line_labels = extracted_sample_names
                
                # Add interactive features
                if self.enable_interactions:
                    self._add_checkboxes(sample_names=extracted_sample_names)
                    self._bind_zoom_events()
                
                debug_print(f"Successfully created plot with {len(extracted_sample_names)} samples")
                return True, fig, "Success"
                
        except Exception as e:
            error_msg = f"Failed to plot samples: {e}"
            debug_print(f"ERROR: {error_msg}")
            import traceback
            traceback.print_exc()
            return False, None, error_msg
    
    def _plot_standard_samples(self, full_sample_data: pd.DataFrame, num_columns_per_sample: int,
                              plot_type: str, sample_names: List[str] = None) -> Tuple[Any, List[str]]:
        """Generate standard plots for regular test data."""
        plt, _, _, _ = self.get_matplotlib_components()
        if not plt:
            raise ImportError("Matplotlib not available")
        
        num_samples = full_sample_data.shape[1] // num_columns_per_sample
        full_sample_data = full_sample_data.replace(0, np.nan)
        
        if plot_type == "TPM (Bar)":
            fig, ax = plt.subplots(figsize=(8, 6))
            extracted_sample_names = self._plot_bar_chart(
                ax, full_sample_data, num_samples, num_columns_per_sample, sample_names
            )
        else:
            fig, ax = plt.subplots(figsize=(8, 6))
            extracted_sample_names = self._plot_line_chart(
                ax, full_sample_data, num_samples, num_columns_per_sample, plot_type, sample_names
            )
        
        # Configure plot
        ax.set_xlabel('Puffs')
        ax.set_ylabel(self._get_y_label_for_plot_type(plot_type))
        ax.set_title(f'{plot_type} Plot')
        ax.legend(loc='upper right')
        ax.grid(True, alpha=0.3)
        
        # Set appropriate y-limits
        self._set_plot_limits(ax, plot_type)
        
        plt.tight_layout()
        return fig, extracted_sample_names
    
    def _plot_user_test_simulation_samples(self, full_sample_data: pd.DataFrame, 
                                          num_columns_per_sample: int, plot_type: str,
                                          sample_names: List[str] = None) -> Tuple[Any, List[str]]:
        """
        Generate split plots for User Test Simulation.
        Creates two separate plots: one for puffs 0-50 and one for remaining puffs.
        """
        plt, _, _, _ = self.get_matplotlib_components()
        if not plt:
            raise ImportError("Matplotlib not available")
        
        num_samples = full_sample_data.shape[1] // num_columns_per_sample
        debug_print(f"User Test Simulation - Number of samples: {num_samples}")
        
        # Check if this should be a bar chart
        if plot_type == "TPM (Bar)":
            debug_print("Creating User Test Simulation bar chart")
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            extracted_sample_names = self._plot_user_test_simulation_bar_chart(
                ax1, ax2, full_sample_data, num_samples, num_columns_per_sample, sample_names
            )
            
            # Mark this as a split plot with bar chart data
            fig.is_split_plot = True
            fig.is_bar_chart = True
            fig.phase1_bars = ax1.patches
            fig.phase2_bars = ax2.patches
        else:
            # Line plot for other types
            full_sample_data = full_sample_data.replace(0, np.nan)
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
            
            extracted_sample_names = []
            phase1_lines = []
            phase2_lines = []
            y_max = 0
            
            for i in range(num_samples):
                start_col = i * num_columns_per_sample
                sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]
                
                # Extract sample name
                sample_name = self._extract_sample_name(sample_data, i, sample_names)
                extracted_sample_names.append(sample_name)
                
                # Get data for both phases
                phase1_data, phase2_data = self._split_user_test_simulation_data(sample_data, plot_type)
                
                # Plot Phase 1 (puffs 0-50)
                if not phase1_data[0].empty and not phase1_data[1].empty:
                    line1 = ax1.plot(phase1_data[0], phase1_data[1], marker='o', label=sample_name)[0]
                    phase1_lines.append(line1)
                    y_max = max(y_max, phase1_data[1].max())
                
                # Plot Phase 2 (extended puffs)
                if not phase2_data[0].empty and not phase2_data[1].empty:
                    line2 = ax2.plot(phase2_data[0], phase2_data[1], marker='o', label=sample_name)[0]
                    phase2_lines.append(line2)
                    y_max = max(y_max, phase2_data[1].max())
                else:
                    # Add placeholder line for consistency
                    line2 = ax2.plot([], [], marker='o', label=sample_name)[0]
                    phase2_lines.append(line2)
            
            # Configure plots
            self._configure_split_plots(ax1, ax2, plot_type, y_max)
            
            # Store line references for checkbox functionality
            fig.phase1_lines = phase1_lines
            fig.phase2_lines = phase2_lines
            fig.is_split_plot = True
        
        plt.tight_layout()
        debug_print(f"Successfully created User Test Simulation split plot with {len(extracted_sample_names)} samples")
        return fig, extracted_sample_names
    
    def _plot_line_chart(self, ax, full_sample_data: pd.DataFrame, num_samples: int,
                        num_columns_per_sample: int, plot_type: str, 
                        sample_names: List[str] = None) -> List[str]:
        """Plot line chart for standard data."""
        extracted_sample_names = []
        
        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]
            
            # Extract sample name
            sample_name = self._extract_sample_name(sample_data, i, sample_names)
            extracted_sample_names.append(sample_name)
            
            # Get X and Y data
            x_data = sample_data.iloc[3:, 0].dropna()
            y_data = self._get_y_data_for_plot_type(sample_data, plot_type)
            y_data = pd.to_numeric(y_data, errors='coerce').dropna()
            
            # Align data on common index
            common_index = x_data.index.intersection(y_data.index)
            x_data = x_data.loc[common_index]
            y_data = y_data.loc[common_index]
            
            # Fix X-axis sequence and plot
            x_data = self._fix_x_axis_sequence(x_data)
            
            if not x_data.empty and not y_data.empty:
                ax.plot(x_data, y_data, marker='o', label=sample_name)
        
        return extracted_sample_names
    
    def _plot_bar_chart(self, ax, full_sample_data: pd.DataFrame, num_samples: int,
                       num_columns_per_sample: int, sample_names: List[str] = None) -> List[str]:
        """Plot bar chart for TPM data."""
        extracted_sample_names = []
        avg_tpm_values = []
        std_tpm_values = []
        
        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]
            
            # Extract sample name
            sample_name = self._extract_sample_name(sample_data, i, sample_names)
            extracted_sample_names.append(sample_name)
            
            # Get TPM data
            tpm_data = self._get_y_data_for_plot_type(sample_data, "TPM")
            tpm_numeric = pd.to_numeric(tpm_data, errors='coerce').dropna()
            
            if not tpm_numeric.empty:
                avg_tpm = tpm_numeric.mean()
                std_tpm = tpm_numeric.std()
                avg_tpm_values.append(avg_tpm)
                std_tpm_values.append(std_tpm)
            else:
                avg_tpm_values.append(0)
                std_tpm_values.append(0)
        
        # Create bar chart
        x_positions = np.arange(len(extracted_sample_names))
        bars = ax.bar(x_positions, avg_tpm_values, yerr=std_tpm_values, 
                     capsize=5, alpha=0.7)
        
        # Set labels
        ax.set_xticks(x_positions)
        ax.set_xticklabels([wrap_text(name, 10) for name in extracted_sample_names])
        
        return extracted_sample_names
    
    # ===================== DATA PROCESSING FUNCTIONS =====================
    
    def _get_y_data_for_plot_type(self, sample_data: pd.DataFrame, plot_type: str) -> pd.Series:
        """Extract Y-data for the specified plot type."""
        pd = self._lazy_import_pandas()
        if not pd:
            return pd.Series() if pd else []
        
        try:
            if plot_type == "TPM":
                # Calculate TPM from weight differences
                puffs = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce')
                before_weights = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce')
                after_weights = pd.to_numeric(sample_data.iloc[3:, 2], errors='coerce')
                
                if self.calculation_service:
                    return self.calculation_service.calculate_tpm_from_weights(
                        puffs, before_weights, after_weights
                    )
                else:
                    # Simple TPM calculation
                    weight_diff = before_weights - after_weights
                    return (weight_diff * 1000) / puffs  # Convert to mg and normalize by puffs
                    
            elif plot_type == "Normalized TPM":
                # Get TPM first, then normalize by puff time
                tpm_data = self._get_y_data_for_plot_type(sample_data, "TPM")
                puff_time = self._extract_puff_time(sample_data)
                return tmp_data / puff_time if puff_time > 0 else tpm_data / 3.0
                
            elif plot_type == "Power Efficiency":
                # Calculate TPM/Power ratio
                tpm_data = self._get_y_data_for_plot_type(sample_data, "TPM")
                voltage, resistance = self._extract_electrical_parameters(sample_data)
                
                if voltage and resistance and voltage > 0 and resistance > 0:
                    power = (voltage ** 2) / resistance
                    return tpm_data / power
                else:
                    return pd.Series(dtype=float)
                    
            elif plot_type == "Draw Pressure":
                return pd.to_numeric(sample_data.iloc[3:, 3], errors='coerce')
                
            elif plot_type == "Resistance":
                return pd.to_numeric(sample_data.iloc[3:, 4], errors='coerce')
                
            else:
                # Default to TPM
                return self._get_y_data_for_plot_type(sample_data, "TPM")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to get Y-data for {plot_type}: {e}")
            return pd.Series(dtype=float)
    
    def _extract_puff_time(self, sample_data: pd.DataFrame) -> float:
        """Extract puff time from sample data."""
        try:
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
                puffing_regime_cell = sample_data.iloc[0, 7]
                if pd.notna(puffing_regime_cell):
                    puffing_regime = str(puffing_regime_cell).strip()
                    
                    # Extract puff time using regex
                    import re
                    pattern = r'mL/(\d+(?:\.\d+)?)'
                    match = re.search(pattern, puffing_regime)
                    if match:
                        return float(match.group(1))
            
            return 3.0  # Default puff time
            
        except Exception as e:
            debug_print(f"WARNING: Failed to extract puff time: {e}")
            return 3.0
    
    def _extract_electrical_parameters(self, sample_data: pd.DataFrame) -> Tuple[Optional[float], Optional[float]]:
        """Extract voltage and resistance from sample data."""
        voltage = None
        resistance = None
        
        try:
            # Extract voltage from row 1, column 5
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
                voltage_val = sample_data.iloc[1, 5]
                if pd.notna(voltage_val):
                    voltage = float(voltage_val)
            
            # Extract resistance from row 0, column 3
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 3:
                resistance_val = sample_data.iloc[0, 3]
                if pd.notna(resistance_val):
                    resistance = float(resistance_val)
                    
        except (ValueError, IndexError, TypeError) as e:
            debug_print(f"WARNING: Failed to extract electrical parameters: {e}")
        
        return voltage, resistance
    
    def _extract_sample_name(self, sample_data: pd.DataFrame, sample_index: int, 
                           sample_names: List[str] = None) -> str:
        """Extract sample name from data or use provided names."""
        if sample_names and sample_index < len(sample_names):
            return sample_names[sample_index]
        
        # Try to extract from data
        try:
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 0:
                # Check first column, second row for sample name
                sample_name_cell = sample_data.iloc[1, 0]
                if pd.notna(sample_name_cell) and str(sample_name_cell).strip():
                    return str(sample_name_cell).strip()
        except:
            pass
        
        return f"Sample {sample_index + 1}"
    
    def _fix_x_axis_sequence(self, x_data: pd.Series) -> pd.Series:
        """Fix X-axis sequence to ensure proper ordering."""
        try:
            # Sort by index to maintain proper sequence
            return x_data.sort_index()
        except Exception as e:
            debug_print(f"WARNING: Failed to fix X-axis sequence: {e}")
            return x_data
    
    def _split_user_test_simulation_data(self, sample_data: pd.DataFrame, 
                                       plot_type: str) -> Tuple[Tuple[pd.Series, pd.Series], 
                                                               Tuple[pd.Series, pd.Series]]:
        """Split User Test Simulation data into Phase 1 and Phase 2."""
        # Get full X and Y data
        x_data = sample_data.iloc[3:, 0].dropna()
        y_data = self._get_y_data_for_plot_type(sample_data, plot_type)
        y_data = pd.to_numeric(y_data, errors='coerce').dropna()
        
        # Align data on common index
        common_index = x_data.index.intersection(y_data.index)
        x_data = x_data.loc[common_index]
        y_data = y_data.loc[common_index]
        
        # Split into phases
        # Phase 1: first 5 rows (puffs 0-50)
        phase1_mask = x_data.index < (x_data.index.min() + 5)
        phase1_x = x_data[phase1_mask]
        phase1_y = y_data[phase1_mask]
        
        # Phase 2: from row 9 onwards (extended puffs)
        phase2_start_idx = x_data.index.min() + 8
        phase2_mask = x_data.index >= phase2_start_idx
        phase2_x = x_data[phase2_mask]
        phase2_y = y_data[phase2_mask]
        
        return (phase1_x, phase1_y), (phase2_x, phase2_y)
    
    def _plot_user_test_simulation_bar_chart(self, ax1, ax2, full_sample_data: pd.DataFrame,
                                            num_samples: int, num_columns_per_sample: int,
                                            sample_names: List[str] = None) -> List[str]:
        """Generate split bar charts for User Test Simulation average TPM."""
        extracted_sample_names = []
        phase1_averages = []
        phase1_stds = []
        phase2_averages = []
        phase2_stds = []
        
        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            sample_data = full_sample_data.iloc[:, start_col:start_col + num_columns_per_sample]
            
            # Extract sample name
            sample_name = self._extract_sample_name(sample_data, i, sample_names)
            extracted_sample_names.append(sample_name)
            
            # Get phase data
            phase1_data, phase2_data = self._split_user_test_simulation_data(sample_data, "TPM")
            
            # Calculate averages and standard deviations
            if not phase1_data[1].empty:
                phase1_avg = phase1_data[1].mean()
                phase1_std = phase1_data[1].std()
            else:
                phase1_avg = 0
                phase1_std = 0
            
            if not phase2_data[1].empty:
                phase2_avg = phase2_data[1].mean()
                phase2_std = phase2_data[1].std()
            else:
                phase2_avg = 0
                phase2_std = 0
            
            phase1_averages.append(phase1_avg)
            phase1_stds.append(phase1_std)
            phase2_averages.append(phase2_avg)
            phase2_stds.append(phase2_std)
        
        # Create bar charts
        x_positions = np.arange(len(extracted_sample_names))
        
        # Phase 1 bar chart
        ax1.bar(x_positions, phase1_averages, yerr=phase1_stds, 
               capsize=5, alpha=0.7, color='skyblue')
        ax1.set_title('Phase 1 - Average TPM (Puffs 0-50)')
        ax1.set_ylabel('TPM (mg/puff)')
        ax1.set_xticks(x_positions)
        ax1.set_xticklabels([wrap_text(name, 10) for name in extracted_sample_names])
        
        # Phase 2 bar chart
        ax2.bar(x_positions, phase2_averages, yerr=phase2_stds, 
               capsize=5, alpha=0.7, color='lightcoral')
        ax2.set_title('Phase 2 - Average TPM (Extended Puffs)')
        ax2.set_ylabel('TPM (mg/puff)')
        ax2.set_xticks(x_positions)
        ax2.set_xticklabels([wrap_text(name, 10) for name in extracted_sample_names])
        
        return extracted_sample_names
    
    def _configure_split_plots(self, ax1, ax2, plot_type: str, y_max: float) -> None:
        """Configure split plots with appropriate labels and limits."""
        y_label = self._get_y_label_for_plot_type(plot_type)
        
        # Configure Phase 1 plot
        ax1.set_xlabel('Puffs')
        ax1.set_ylabel(y_label)
        ax1.set_title(f'{plot_type} - Phase 1 (Puffs 0-50)')
        ax1.legend(loc='upper right')
        ax1.set_xlim(0, 60)
        ax1.grid(True, alpha=0.3)
        
        # Configure Phase 2 plot
        ax2.set_xlabel('Puffs')
        ax2.set_ylabel(y_label)
        ax2.set_title(f'{plot_type} - Phase 2 (Extended Puffs)')
        ax2.legend(loc='upper right')
        ax2.grid(True, alpha=0.3)
        
        # Set consistent y-axis limits
        if plot_type == "Normalized TPM":
            ax1.set_ylim(-0.2, 4)
            ax2.set_ylim(-0.2, 4)
        elif y_max > 9 and y_max <= 50:
            ax1.set_ylim(0, y_max)
            ax2.set_ylim(0, y_max)
        else:
            ax1.set_ylim(0, 9)
            ax2.set_ylim(0, 9)
    
    def _get_y_label_for_plot_type(self, plot_type: str) -> str:
        """Get the appropriate Y-axis label for the plot type."""
        return self.y_label_mapping.get(plot_type, 'TPM (mg/puff)')
    
    def _set_plot_limits(self, ax, plot_type: str) -> None:
        """Set appropriate plot limits based on plot type."""
        if plot_type == "Normalized TPM":
            ax.set_ylim(-0.2, 4)
        elif plot_type == "Draw Pressure":
            ax.set_ylim(0, 10)
        elif plot_type == "Resistance":
            ax.set_ylim(0, 5)
        # For other types, let matplotlib auto-scale
    
    # ===================== MATPLOTLIB INTEGRATION =====================
    
    def _embed_plot_in_frame(self, fig, frame: tk.Widget) -> Optional[Any]:
        """Embed a matplotlib figure into a tkinter frame."""
        plt, FigureCanvasTkAgg, NavigationToolbar2Tk, _ = self.get_matplotlib_components()
        
        if not plt or not FigureCanvasTkAgg or not NavigationToolbar2Tk:
            debug_print("ERROR: Could not load matplotlib components")
            return None
        
        try:
            # Clear existing widgets
            for widget in frame.winfo_children():
                widget.destroy()
            
            # Create dropdown frame
            dropdown_frame = ttk.Frame(frame)
            dropdown_frame.pack(side='bottom', fill='x', pady=10)
            
            # Create plot container
            plot_container = ttk.Frame(frame)
            plot_container.pack(fill='both', expand=True, pady=(0, 0))
            
            # Adjust margins for checkboxes
            is_split_plot = hasattr(fig, 'is_split_plot') and fig.is_split_plot
            if is_split_plot:
                fig.subplots_adjust(right=0.80)
            else:
                fig.subplots_adjust(right=0.82)
            
            # Embed figure
            canvas = FigureCanvasTkAgg(fig, master=plot_container)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
            # Add toolbar
            toolbar = NavigationToolbar2Tk(canvas, plot_container)
            toolbar.update()
            
            # Store references
            if is_split_plot:
                self.current_axes = fig.axes
                self.current_lines = []
            else:
                self.current_axes = fig.gca()
                self.current_lines = self.current_axes.lines
            
            debug_print("Successfully embedded plot in frame")
            return canvas
            
        except Exception as e:
            debug_print(f"ERROR: Failed to embed plot: {e}")
            return None
    
    def clear_plot_area(self, frame: tk.Widget) -> None:
        """Clear the plot area and release matplotlib resources."""
        plt, _, _, _ = self.get_matplotlib_components()
        
        try:
            # Clear frame widgets
            if frame and hasattr(frame, 'winfo_exists') and frame.winfo_exists():
                for widget in frame.winfo_children():
                    widget.destroy()
            
            # Close matplotlib figure
            if self.current_figure and plt:
                plt.close(self.current_figure)
            
            # Reset references
            self.current_figure = None
            self.current_canvas = None
            self.current_axes = None
            self.current_lines = None
            self.line_labels = []
            self.original_lines_data = []
            self.check_buttons = None
            self.checkbox_cid = None
            
            debug_print("Plot area cleared successfully")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to clear plot area: {e}")
    
    # ===================== INTERACTIVE FEATURES =====================
    
    def _add_checkboxes(self, sample_names: List[str] = None) -> None:
        """Add checkboxes to toggle visibility of plot elements."""
        plt, _, _, CheckButtons = self.get_matplotlib_components()
        
        if not plt or not CheckButtons:
            debug_print("ERROR: Could not load matplotlib components for checkboxes")
            return
        
        try:
            if not sample_names:
                sample_names = self.line_labels
            
            if not sample_names:
                debug_print("No sample names available for checkboxes")
                return
            
            # Clear existing checkboxes
            if self.check_buttons:
                self.check_buttons.ax.clear()
                self.check_buttons = None
            
            # Determine plot characteristics
            is_split_plot = hasattr(self.current_figure, 'is_split_plot') and self.current_figure.is_split_plot
            is_bar_chart = hasattr(self.current_figure, 'is_bar_chart') and self.current_figure.is_bar_chart
            
            # Store data for restoration
            self._store_original_plot_data(sample_names, is_split_plot, is_bar_chart)
            
            # Create wrapped labels for display
            wrapped_labels = []
            self.label_mapping = {}
            
            for label in sample_names:
                wrapped_label = wrap_text(label, max_width=12)
                self.label_mapping[wrapped_label] = label
                wrapped_labels.append(wrapped_label)
            
            # Calculate checkbox position and size
            num_checkboxes = len(wrapped_labels)
            checkbox_height = min(0.8, 0.05 * num_checkboxes)
            checkbox_bottom = 0.5 - checkbox_height / 2
            
            # Create checkbox axes
            checkbox_ax = plt.axes([0.83, checkbox_bottom, 0.15, checkbox_height])
            
            # Create CheckButtons
            self.check_buttons = CheckButtons(
                checkbox_ax, wrapped_labels, [True] * len(wrapped_labels)
            )
            
            # Style checkboxes
            for rect in self.check_buttons.rectangles:
                rect.set_facecolor('lightblue')
                rect.set_edgecolor('black')
                rect.set_linewidth(1)
            
            for line in self.check_buttons.lines:
                line.set_color('red')
                line.set_linewidth(2)
                line.set_zorder(2)
            
            # Bind appropriate callback
            if is_split_plot and is_bar_chart:
                self.checkbox_cid = self.check_buttons.on_clicked(self._on_split_bar_checkbox_click)
            elif is_split_plot:
                self.checkbox_cid = self.check_buttons.on_clicked(self._on_split_plot_checkbox_click)
            elif is_bar_chart:
                self.checkbox_cid = self.check_buttons.on_clicked(self._on_bar_checkbox_click)
            else:
                self.checkbox_cid = self.check_buttons.on_clicked(self._on_line_checkbox_click)
            
            if self.current_canvas:
                self.current_canvas.draw_idle()
            
            debug_print(f"Added checkboxes for {len(sample_names)} samples")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to add checkboxes: {e}")
    
    def _store_original_plot_data(self, sample_names: List[str], is_split_plot: bool, is_bar_chart: bool) -> None:
        """Store original plot data for restoration."""
        try:
            self.line_labels = sample_names
            
            if is_split_plot and hasattr(self.current_figure, 'phase1_bars') and is_bar_chart:
                # Split bar chart data
                phase1_data = [(bar.get_x(), bar.get_height()) for bar in self.current_figure.phase1_bars]
                phase2_data = [(bar.get_x(), bar.get_height()) for bar in self.current_figure.phase2_bars]
                self.original_lines_data = list(zip(phase1_data, phase2_data))
            elif is_split_plot and hasattr(self.current_figure, 'phase1_lines'):
                # Split line plot data
                phase1_data = [(line.get_xdata(), line.get_ydata()) for line in self.current_figure.phase1_lines]
                phase2_data = [(line.get_xdata(), line.get_ydata()) for line in self.current_figure.phase2_lines]
                self.original_lines_data = list(zip(phase1_data, phase2_data))
            elif is_bar_chart and self.current_axes:
                # Regular bar chart data
                self.original_lines_data = [(patch.get_x(), patch.get_height()) for patch in self.current_axes.patches]
            elif self.current_lines:
                # Regular line plot data
                self.original_lines_data = [(line.get_xdata(), line.get_ydata()) for line in self.current_lines]
            
            debug_print(f"Stored original data for {len(self.line_labels)} elements")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to store original plot data: {e}")
    
    def _bind_zoom_events(self) -> None:
        """Bind mouse scroll events for zooming."""
        if self.current_canvas:
            self.current_canvas.mpl_connect("scroll_event", self._on_zoom)
    
    def _on_zoom(self, event) -> None:
        """Handle zoom events on the plot."""
        try:
            if not self.current_axes or event.inaxes not in (
                self.current_axes if isinstance(self.current_axes, list) else [self.current_axes]
            ):
                return
            
            # Determine which axes was scrolled
            target_ax = event.inaxes
            
            x_min, x_max = target_ax.get_xlim()
            y_min, y_max = target_ax.get_ylim()
            cursor_x, cursor_y = event.xdata, event.ydata
            zoom_scale = 0.9 if event.button == 'up' else 1.1
            
            if cursor_x is None or cursor_y is None:
                return
            
            new_x_min = cursor_x - (cursor_x - x_min) * zoom_scale
            new_x_max = cursor_x + (x_max - cursor_x) * zoom_scale
            new_y_min = cursor_y - (cursor_y - y_min) * zoom_scale
            new_y_max = cursor_y + (y_max - cursor_y) * zoom_scale
            
            target_ax.set_xlim(new_x_min, new_x_max)
            target_ax.set_ylim(new_y_min, new_y_max)
            
            if self.current_canvas:
                self.current_canvas.draw_idle()
                
        except Exception as e:
            debug_print(f"ERROR: Zoom event failed: {e}")
    
    # ===================== CHECKBOX CALLBACKS =====================
    
    def _on_line_checkbox_click(self, wrapped_label: str) -> None:
        """Handle checkbox clicks for line plots."""
        try:
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None or original_label not in self.line_labels:
                return
            
            index = self.line_labels.index(original_label)
            if index < len(self.current_lines):
                line = self.current_lines[index]
                line.set_visible(not line.get_visible())
                
                if self.current_canvas:
                    self.current_canvas.draw_idle()
                    
        except Exception as e:
            debug_print(f"ERROR: Line checkbox callback failed: {e}")
    
    def _on_bar_checkbox_click(self, wrapped_label: str) -> None:
        """Handle checkbox clicks for bar charts."""
        try:
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None or original_label not in self.line_labels:
                return
            
            index = self.line_labels.index(original_label)
            if self.current_axes and hasattr(self.current_axes, 'patches') and index < len(self.current_axes.patches):
                bar = self.current_axes.patches[index]
                bar.set_visible(not bar.get_visible())
                
                if self.current_canvas:
                    self.current_canvas.draw_idle()
                    
        except Exception as e:
            debug_print(f"ERROR: Bar checkbox callback failed: {e}")
    
    def _on_split_plot_checkbox_click(self, wrapped_label: str) -> None:
        """Handle checkbox clicks for split line plots."""
        try:
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None or original_label not in self.line_labels:
                return
            
            index = self.line_labels.index(original_label)
            
            # Toggle visibility for both phases
            if (hasattr(self.current_figure, 'phase1_lines') and 
                hasattr(self.current_figure, 'phase2_lines')):
                
                if index < len(self.current_figure.phase1_lines):
                    line1 = self.current_figure.phase1_lines[index]
                    line1.set_visible(not line1.get_visible())
                
                if index < len(self.current_figure.phase2_lines):
                    line2 = self.current_figure.phase2_lines[index]
                    line2.set_visible(not line2.get_visible())
                
                if self.current_canvas:
                    self.current_canvas.draw_idle()
                    
        except Exception as e:
            debug_print(f"ERROR: Split plot checkbox callback failed: {e}")
    
    def _on_split_bar_checkbox_click(self, wrapped_label: str) -> None:
        """Handle checkbox clicks for split bar charts."""
        try:
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None or original_label not in self.line_labels:
                return
            
            index = self.line_labels.index(original_label)
            
            # Toggle visibility for both phases
            if (hasattr(self.current_figure, 'phase1_bars') and 
                hasattr(self.current_figure, 'phase2_bars')):
                
                if index < len(self.current_figure.phase1_bars):
                    bar1 = self.current_figure.phase1_bars[index]
                    bar1.set_visible(not bar1.get_visible())
                
                if index < len(self.current_figure.phase2_bars):
                    bar2 = self.current_figure.phase2_bars[index]
                    bar2.set_visible(not bar2.get_visible())
                
                if self.current_canvas:
                    self.current_canvas.draw_idle()
                    
        except Exception as e:
            debug_print(f"ERROR: Split bar checkbox callback failed: {e}")
    
    # ===================== PLOT MANAGEMENT =====================
    
    def restore_all_plot_elements(self) -> None:
        """Restore all plot elements to their original visible state."""
        debug_print("Restoring all plot elements")
        
        try:
            if self.check_buttons:
                self.check_buttons.disconnect(self.checkbox_cid)
            
            # Determine plot type
            is_split_plot = hasattr(self.current_figure, 'is_split_plot') and self.current_figure.is_split_plot
            is_bar_chart = hasattr(self.current_figure, 'is_bar_chart') and self.current_figure.is_bar_chart
            
            if is_split_plot and is_bar_chart:
                # Restore split bar chart
                if (hasattr(self.current_figure, 'phase1_bars') and 
                    hasattr(self.current_figure, 'phase2_bars') and 
                    self.original_lines_data):
                    
                    for i, (phase1_data, phase2_data) in enumerate(self.original_lines_data):
                        if i < len(self.current_figure.phase1_bars):
                            bar1 = self.current_figure.phase1_bars[i]
                            bar1.set_x(phase1_data[0])
                            bar1.set_height(phase1_data[1])
                            bar1.set_visible(True)
                        
                        if i < len(self.current_figure.phase2_bars):
                            bar2 = self.current_figure.phase2_bars[i]
                            bar2.set_x(phase2_data[0])
                            bar2.set_height(phase2_data[1])
                            bar2.set_visible(True)
                            
            elif is_split_plot:
                # Restore split line plot
                if (hasattr(self.current_figure, 'phase1_lines') and 
                    hasattr(self.current_figure, 'phase2_lines') and 
                    self.original_lines_data):
                    
                    for i, (phase1_data, phase2_data) in enumerate(self.original_lines_data):
                        if i < len(self.current_figure.phase1_lines):
                            line1 = self.current_figure.phase1_lines[i]
                            line1.set_xdata(phase1_data[0])
                            line1.set_ydata(phase1_data[1])
                            line1.set_visible(True)
                        
                        if i < len(self.current_figure.phase2_lines):
                            line2 = self.current_figure.phase2_lines[i]
                            line2.set_xdata(phase2_data[0])
                            line2.set_ydata(phase2_data[1])
                            line2.set_visible(True)
                            
            elif is_bar_chart:
                # Restore regular bar chart
                if self.current_axes and hasattr(self.current_axes, 'patches') and self.original_lines_data:
                    for patch, original_data in zip(self.current_axes.patches, self.original_lines_data):
                        patch.set_x(original_data[0])
                        patch.set_height(original_data[1])
                        patch.set_visible(True)
                        
            else:
                # Restore regular line plot
                if self.current_lines and self.original_lines_data:
                    for line, original_data in zip(self.current_lines, self.original_lines_data):
                        line.set_xdata(original_data[0])
                        line.set_ydata(original_data[1])
                        line.set_visible(True)
            
            # Reset checkboxes to checked state
            if self.check_buttons and self.line_labels:
                for i in range(len(self.line_labels)):
                    if not self.check_buttons.get_status()[i]:
                        self.check_buttons.set_active(i)
            
            # Reconnect checkbox callback
            if self.check_buttons:
                if is_split_plot and is_bar_chart:
                    self.checkbox_cid = self.check_buttons.on_clicked(self._on_split_bar_checkbox_click)
                elif is_split_plot:
                    self.checkbox_cid = self.check_buttons.on_clicked(self._on_split_plot_checkbox_click)
                elif is_bar_chart:
                    self.checkbox_cid = self.check_buttons.on_clicked(self._on_bar_checkbox_click)
                else:
                    self.checkbox_cid = self.check_buttons.on_clicked(self._on_line_checkbox_click)
            
            # Redraw canvas
            if self.current_canvas:
                self.current_canvas.draw_idle()
            
            debug_print("All plot elements restored successfully")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to restore plot elements: {e}")
    
    def get_valid_plot_options(self, data: pd.DataFrame) -> List[str]:
        """Get valid plot options based on available data."""
        if data.empty:
            return []
        
        valid_options = ["TPM", "TPM (Bar)"]  # Always available
        
        # Check for additional data columns
        if data.shape[1] > 3:
            valid_options.append("Draw Pressure")
        
        if data.shape[1] > 4:
            valid_options.append("Resistance")
        
        # Power efficiency requires electrical parameters
        if data.shape[1] > 5:
            valid_options.append("Power Efficiency")
        
        # Normalized TPM requires puffing regime data
        if data.shape[1] > 7:
            valid_options.append("Normalized TPM")
        
        return valid_options
    
    def get_service_status(self) -> Dict[str, Any]:
        """Get comprehensive service status information."""
        return {
            'supported_plot_types': self.supported_plot_types,
            'current_figure_available': self.current_figure is not None,
            'current_canvas_available': self.current_canvas is not None,
            'checkboxes_enabled': self.check_buttons is not None,
            'sample_count': len(self.line_labels),
            'interactions_enabled': self.enable_interactions,
            'calculation_service_available': self.calculation_service is not None,
            'matplotlib_available': self._lazy_import_matplotlib()[0] is not None,
            'pandas_available': self._lazy_import_pandas() is not None
        }
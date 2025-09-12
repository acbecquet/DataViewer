# models/plot_model.py
"""
models/plot_model.py
Plot model for managing plot configurations, calculations, and state.
Consolidated from plot_manager.py, processing.py plotting functions, and viscosity plotting functionality.
"""

import os
import copy
import math
import tempfile
import threading
from typing import Optional, Dict, Any, List, Tuple, Union, Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox


def debug_print(message: str):
    """Debug print function for plot operations."""
    print(f"DEBUG: PlotModel - {message}")


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
class PlotSettings:
    """Settings for plot appearance and behavior."""
    
    # Visual settings
    figure_size: Tuple[float, float] = (8, 6)
    dpi: int = 100
    font_size: int = 10
    line_width: float = 2.0
    marker_size: float = 6.0
    transparency: float = 0.8
    
    # Interactive features
    enable_checkboxes: bool = True
    enable_zoom: bool = True
    enable_toolbar: bool = True
    enable_grid: bool = True
    
    # Layout settings
    tight_layout: bool = True
    legend_location: str = "upper right"
    max_checkbox_width: int = 12
    
    # Axis settings
    x_limits: Optional[Tuple[float, float]] = None
    y_limits: Optional[Tuple[float, float]] = None
    
    # Color and style
    color_palette: List[str] = field(default_factory=lambda: [
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
        '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
    ])
    
    def get_color(self, index: int) -> str:
        """Get color from palette by index."""
        return self.color_palette[index % len(self.color_palette)]


@dataclass
class PlotConfig:
    """Configuration for specific plot types and data processing."""
    
    # Plot type configuration
    plot_type: str = "TPM"
    supported_plot_types: List[str] = field(default_factory=lambda: [
        "TPM", "Draw Pressure", "Resistance", "Power Efficiency", 
        "TPM (Bar)", "Normalized TPM"
    ])
    
    # Data processing settings
    num_columns_per_sample: int = 12
    default_puff_time: float = 3.0
    default_puffs: int = 10
    precision_digits: int = 3
    
    # Y-axis label mapping
    y_label_mapping: Dict[str, str] = field(default_factory=lambda: {
        "TPM": 'TPM (mg/puff)',
        "Normalized TPM": 'Normalized TPM (mg/s)',
        "Draw Pressure": 'Draw Pressure (kPa)',
        "Resistance": 'Resistance (Ohms)',
        "Power Efficiency": 'Power Efficiency (mg/W)',
        "TPM (Bar)": 'TPM (mg/puff)'
    })
    
    # Plot-specific limits
    y_limits_by_type: Dict[str, Tuple[float, float]] = field(default_factory=lambda: {
        "TPM": (0, 9),
        "Normalized TPM": (-0.2, 4),
        "Draw Pressure": (0, 10),
        "Resistance": (0, 50),
        "Power Efficiency": (0, 20),
        "TPM (Bar)": (0, 9)
    })
    
    def get_y_label(self, plot_type: str) -> str:
        """Get y-axis label for plot type."""
        return self.y_label_mapping.get(plot_type, f"{plot_type} Values")
    
    def get_y_limits(self, plot_type: str) -> Tuple[float, float]:
        """Get y-axis limits for plot type."""
        return self.y_limits_by_type.get(plot_type, (0, 10))


@dataclass
class PlotState:
    """Current state of plot display and interaction."""
    
    # Matplotlib objects
    figure: Optional[Any] = None
    canvas: Optional[Any] = None
    axes: Optional[Any] = None
    lines: Optional[List[Any]] = None
    
    # Interactive elements
    check_buttons: Optional[Any] = None
    checkbox_cid: Optional[Any] = None
    plot_dropdown: Optional[Any] = None
    
    # Plot data and labels
    line_labels: List[str] = field(default_factory=list)
    original_lines_data: List[Any] = field(default_factory=list)
    label_mapping: Dict[str, str] = field(default_factory=dict)
    wrapped_labels: List[str] = field(default_factory=list)
    
    # State tracking
    is_zoomed: bool = False
    current_plot_type: str = "TPM"
    last_update: Optional[datetime] = None
    
    def clear_state(self):
        """Clear all plot state."""
        self.figure = None
        self.canvas = None
        self.axes = None
        self.lines = None
        self.check_buttons = None
        self.checkbox_cid = None
        self.line_labels.clear()
        self.original_lines_data.clear()
        self.label_mapping.clear()
        self.wrapped_labels.clear()
        self.is_zoomed = False
        self.last_update = None
        debug_print("Plot state cleared")


class PlotModel:
    """
    Main plot model that manages plot configurations, calculations, and state.
    Consolidates plotting functionality from plot_manager.py, processing.py, and viscosity calculator.
    """
    
    def __init__(self):
        """Initialize the plot model."""
        # Configuration and settings
        self.config = PlotConfig()
        self.settings = PlotSettings()
        self.state = PlotState()
        
        # Lazy loading infrastructure
        self._matplotlib_loaded = False
        self._pandas_loaded = False
        
        # Matplotlib components (loaded lazily)
        self.plt = None
        self.FigureCanvasTkAgg = None
        self.NavigationToolbar2Tk = None
        self.CheckButtons = None
        
        # Threading support
        self.plot_lock = threading.Lock()
        
        # Cache for plot data
        self.plot_cache: Dict[str, Any] = {}
        self.cache_duration_hours = 1
        
        debug_print("PlotModel initialized")
        debug_print(f"Supported plot types: {', '.join(self.config.supported_plot_types)}")
    
    # ===================== LAZY LOADING INFRASTRUCTURE =====================
    
    def _lazy_import_matplotlib_components(self):
        """Lazy import matplotlib components."""
        if self._matplotlib_loaded:
            return self.plt, self.FigureCanvasTkAgg, self.NavigationToolbar2Tk, self.CheckButtons
            
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
            from matplotlib.widgets import CheckButtons
            
            self.plt = plt
            self.FigureCanvasTkAgg = FigureCanvasTkAgg
            self.NavigationToolbar2Tk = NavigationToolbar2Tk
            self.CheckButtons = CheckButtons
            self._matplotlib_loaded = True
            
            debug_print("Matplotlib components loaded successfully")
            return plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons
            
        except ImportError as e:
            debug_print(f"Error importing matplotlib: {e}")
            return None, None, None, None
    
    def get_matplotlib_components(self):
        """Get matplotlib components with lazy loading."""
        return self._lazy_import_matplotlib_components()
    
    # ===================== PLOT CALCULATION METHODS =====================
    
    def calculate_tpm_from_weights(self, puffs: pd.Series, before_weights: pd.Series, 
                                   after_weights: pd.Series) -> pd.Series:
        """
        Calculate TPM from weight differences.
        Consolidated from processing.py TPM calculation logic.
        """
        try:
            # Convert to numeric, handling any string values
            puffs_numeric = pd.to_numeric(puffs, errors='coerce')
            before_numeric = pd.to_numeric(before_weights, errors='coerce')
            after_numeric = pd.to_numeric(after_weights, errors='coerce')
            
            # Calculate weight differences
            weight_diff = before_numeric - after_numeric
            
            # Calculate TPM (weight difference / number of puffs)
            # Handle division by zero
            tpm = weight_diff / puffs_numeric.replace(0, np.nan)
            
            debug_print(f"Calculated TPM from weights - samples: {len(tpm)}")
            debug_print(f"TPM range: {tpm.min():.3f} to {tpm.max():.3f}")
            
            return tpm.round(self.config.precision_digits)
            
        except Exception as e:
            debug_print(f"Error calculating TPM from weights: {e}")
            return pd.Series(dtype=float)
    
    def calculate_normalized_tpm(self, tpm_data: pd.Series, puff_time: Optional[float] = None) -> pd.Series:
        """
        Calculate normalized TPM (TPM per second).
        Consolidated from processing.py normalization logic.
        """
        try:
            tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')
            
            if puff_time is None or puff_time <= 0:
                debug_print("Using default puff time for normalization")
                puff_time = self.config.default_puff_time
            
            normalized_tpm = tpm_numeric / puff_time
            
            debug_print(f"Normalized TPM with puff time {puff_time}s")
            debug_print(f"Normalized TPM range: {normalized_tpm.min():.3f} to {normalized_tpm.max():.3f}")
            
            return normalized_tpm.round(self.config.precision_digits)
            
        except Exception as e:
            debug_print(f"Error calculating normalized TPM: {e}")
            return pd.Series(dtype=float)
    
    def calculate_power_efficiency(self, tpm_data: pd.Series, voltage: float, 
                                   resistance: float) -> pd.Series:
        """
        Calculate power efficiency (TPM per watt).
        Consolidated from processing.py power efficiency calculation.
        """
        try:
            tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')
            
            if voltage <= 0 or resistance <= 0:
                debug_print("Invalid voltage or resistance for power efficiency")
                return pd.Series(dtype=float)
            
            # Calculate power (P = V²/R)
            power = (voltage ** 2) / resistance
            
            # Calculate power efficiency (TPM per watt)
            power_efficiency = tpm_numeric / power
            
            debug_print(f"Calculated power efficiency with V={voltage}V, R={resistance}Ω, P={power:.3f}W")
            debug_print(f"Power efficiency range: {power_efficiency.min():.3f} to {power_efficiency.max():.3f}")
            
            return power_efficiency.round(self.config.precision_digits)
            
        except Exception as e:
            debug_print(f"Error calculating power efficiency: {e}")
            return pd.Series(dtype=float)
    
    def get_y_data_for_plot_type(self, sample_data: pd.DataFrame, plot_type: str) -> pd.Series:
        """
        Extract y-axis data for the specified plot type.
        Consolidated from processing.py get_y_data_for_plot_type function.
        """
        debug_print(f"Extracting y-data for plot type: {plot_type}")
        
        try:
            if plot_type == "TPM":
                # Extract TPM data from before/after weights
                if len(sample_data) >= 4:
                    puffs = sample_data.iloc[3:, 1]  # Column B (puffs)
                    before_weights = sample_data.iloc[3:, 6]  # Column G (before weights)
                    after_weights = sample_data.iloc[3:, 7]  # Column H (after weights)
                    
                    return self.calculate_tpm_from_weights(puffs, before_weights, after_weights)
                else:
                    debug_print("Insufficient data rows for TPM calculation")
                    return pd.Series(dtype=float)
            
            elif plot_type == "Normalized TPM":
                # Get TPM first, then normalize
                tpm_data = self.get_y_data_for_plot_type(sample_data, "TPM")
                
                # Try to extract puff time from metadata
                puff_time = None
                try:
                    if len(sample_data) >= 2:
                        puff_time_val = sample_data.iloc[1, 6]  # Adjust column as needed
                        if pd.notna(puff_time_val):
                            puff_time = float(puff_time_val)
                except (ValueError, IndexError, TypeError):
                    pass
                
                return self.calculate_normalized_tpm(tpm_data, puff_time)
            
            elif plot_type == "Power Efficiency":
                # Get TPM and calculate power efficiency
                tpm_data = self.get_y_data_for_plot_type(sample_data, "TPM")
                
                # Extract voltage and resistance from metadata
                voltage = None
                resistance = None
                
                try:
                    if len(sample_data) >= 2:
                        voltage_val = sample_data.iloc[1, 5]  # Adjust column index as needed
                        if pd.notna(voltage_val):
                            voltage = float(voltage_val)
                except (ValueError, IndexError, TypeError):
                    pass
                    
                try:
                    if len(sample_data) >= 1:
                        resistance_val = sample_data.iloc[0, 3]  # Adjust column index as needed
                        if pd.notna(resistance_val):
                            resistance = float(resistance_val)
                except (ValueError, IndexError, TypeError):
                    pass
                
                if voltage and resistance:
                    return self.calculate_power_efficiency(tpm_data, voltage, resistance)
                else:
                    debug_print("Cannot calculate power efficiency - missing voltage/resistance")
                    return pd.Series(dtype=float)
            
            elif plot_type == "Draw Pressure":
                if len(sample_data) >= 4:
                    return pd.to_numeric(sample_data.iloc[3:, 3], errors='coerce')  # Column D
                else:
                    return pd.Series(dtype=float)
            
            elif plot_type == "Resistance":
                if len(sample_data) >= 4:
                    return pd.to_numeric(sample_data.iloc[3:, 4], errors='coerce')  # Column E
                else:
                    return pd.Series(dtype=float)
            
            elif plot_type == "TPM (Bar)":
                # Same as TPM but for bar chart display
                return self.get_y_data_for_plot_type(sample_data, "TPM")
            
            else:
                debug_print(f"Unknown plot type: {plot_type}, defaulting to TPM")
                return self.get_y_data_for_plot_type(sample_data, "TPM")
                
        except Exception as e:
            debug_print(f"Error extracting y-data for {plot_type}: {e}")
            return pd.Series(dtype=float)
    
    def get_x_data_for_plot(self, sample_data: pd.DataFrame) -> pd.Series:
        """
        Extract x-axis data (puff numbers) from sample data.
        Consolidated from processing.py x-data extraction logic.
        """
        try:
            if len(sample_data) >= 4:
                x_data = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce')  # Column B (puffs)
                debug_print(f"Extracted x-data with {len(x_data)} points")
                return x_data
            else:
                debug_print("Insufficient data for x-axis extraction")
                return pd.Series(dtype=float)
                
        except Exception as e:
            debug_print(f"Error extracting x-data: {e}")
            return pd.Series(dtype=float)
    
    # ===================== VISCOSITY CALCULATION METHODS =====================
    
    def calculate_viscosity_temperature_curve(self, viscosity_model: Dict[str, Any], 
                                               temperature_range: List[float]) -> Tuple[List[float], List[float]]:
        """
        Calculate viscosity values across a temperature range using a viscosity model.
        Consolidated from viscosity_calculator functionality.
        """
        try:
            if not viscosity_model or 'model' not in viscosity_model:
                debug_print("Invalid viscosity model provided")
                return [], []
            
            model = viscosity_model['model']
            features = viscosity_model.get('features', [])
            
            viscosities = []
            temperatures_kelvin = []
            
            for temp_c in temperature_range:
                temp_k = temp_c + 273.15
                temperatures_kelvin.append(temp_k)
                
                # Create feature vector for prediction
                # This would need to be customized based on your specific model features
                feature_vector = [temp_k, 1/temp_k]  # Basic Arrhenius features
                
                try:
                    # Predict viscosity using the model
                    viscosity = model.predict([feature_vector])[0]
                    viscosities.append(max(0, viscosity))  # Ensure non-negative
                except Exception as e:
                    debug_print(f"Error predicting viscosity at {temp_c}°C: {e}")
                    viscosities.append(0)
            
            debug_print(f"Calculated viscosity curve for {len(temperature_range)} temperature points")
            return temperature_range, viscosities
            
        except Exception as e:
            debug_print(f"Error calculating viscosity temperature curve: {e}")
            return [], []
    
    def calculate_arrhenius_activation_energy(self, temperatures: List[float], 
                                               viscosities: List[float]) -> Optional[float]:
        """
        Calculate activation energy from Arrhenius plot of viscosity vs temperature.
        Consolidated from viscosity_calculator Arrhenius analysis.
        """
        try:
            if len(temperatures) < 3 or len(viscosities) < 3:
                debug_print("Insufficient data points for Arrhenius analysis")
                return None
            
            # Convert to numpy arrays
            temps_k = np.array(temperatures) + 273.15  # Convert to Kelvin
            visc_array = np.array(viscosities)
            
            # Filter out zero or negative viscosities
            valid_mask = visc_array > 0
            temps_k = temps_k[valid_mask]
            visc_array = visc_array[valid_mask]
            
            if len(temps_k) < 3:
                debug_print("Insufficient valid data points after filtering")
                return None
            
            # Calculate Arrhenius variables
            inverse_temp = 1 / temps_k
            ln_viscosity = np.log(visc_array)
            
            # Linear regression: ln(η) = ln(A) + Ea/(RT)
            # Slope = Ea/R, where R = 8.314 J/(mol·K)
            coefficients = np.polyfit(inverse_temp, ln_viscosity, 1)
            slope = coefficients[0]
            
            # Calculate activation energy in kJ/mol
            R = 8.314  # J/(mol·K)
            activation_energy = slope * R / 1000  # Convert to kJ/mol
            
            debug_print(f"Calculated activation energy: {activation_energy:.2f} kJ/mol")
            return activation_energy
            
        except Exception as e:
            debug_print(f"Error calculating activation energy: {e}")
            return None
    
    # ===================== PLOT GENERATION METHODS =====================
    
    def create_line_plot(self, full_sample_data: pd.DataFrame, 
                         sample_names: Optional[List[str]] = None) -> Tuple[Any, List[str]]:
        """
        Create a line plot for multiple samples.
        Consolidated from processing.py plot_samples_line function.
        """
        plt, _, _, _ = self.get_matplotlib_components()
        if not plt:
            debug_print("Could not load matplotlib")
            return None, []
        
        try:
            fig, ax = plt.subplots(figsize=self.settings.figure_size, dpi=self.settings.dpi)
            
            extracted_sample_names = []
            y_max = 0
            
            # Calculate number of samples
            num_samples = len(full_sample_data.columns) // self.config.num_columns_per_sample
            debug_print(f"Creating line plot for {num_samples} samples")
            
            for i in range(num_samples):
                start_col = i * self.config.num_columns_per_sample
                end_col = start_col + self.config.num_columns_per_sample
                sample_data = full_sample_data.iloc[:, start_col:end_col]
                
                # Get x and y data
                x_data = self.get_x_data_for_plot(sample_data)
                y_data = self.get_y_data_for_plot_type(sample_data, self.state.current_plot_type)
                
                # Remove NaN values
                valid_mask = ~(pd.isna(x_data) | pd.isna(y_data))
                x_data = x_data[valid_mask]
                y_data = y_data[valid_mask]
                
                debug_print(f"Sample {i+1}: {len(x_data)} valid data points")
                
                if not x_data.empty and not y_data.empty:
                    # Get sample name
                    if sample_names and i < len(sample_names):
                        sample_name = sample_names[i]
                    else:
                        sample_name = sample_data.columns[5] if len(sample_data.columns) > 5 else f"Sample {i+1}"
                    
                    # Plot the line
                    color = self.settings.get_color(i)
                    ax.plot(x_data, y_data, marker='o', label=sample_name, 
                           color=color, linewidth=self.settings.line_width,
                           markersize=self.settings.marker_size, 
                           alpha=self.settings.transparency)
                    
                    extracted_sample_names.append(sample_name)
                    y_max = max(y_max, y_data.max())
            
            # Configure plot
            ax.set_xlabel('Puffs')
            ax.set_ylabel(self.config.get_y_label(self.state.current_plot_type))
            ax.set_title(self.state.current_plot_type)
            ax.legend(loc=self.settings.legend_location)
            
            if self.settings.enable_grid:
                ax.grid(True, alpha=0.3)
            
            # Set y-axis limits
            if self.state.current_plot_type in self.config.y_limits_by_type:
                y_limits = self.config.get_y_limits(self.state.current_plot_type)
                if self.state.current_plot_type == "Normalized TPM":
                    ax.set_ylim(y_limits)
                elif y_max > 9 and y_max <= 50:
                    ax.set_ylim(0, y_max)
                else:
                    ax.set_ylim(y_limits)
            
            if self.settings.tight_layout:
                plt.tight_layout()
            
            debug_print(f"Line plot created with {len(extracted_sample_names)} samples")
            return fig, extracted_sample_names
            
        except Exception as e:
            debug_print(f"Error creating line plot: {e}")
            return None, []
    
    def create_bar_plot(self, full_sample_data: pd.DataFrame, 
                        sample_names: Optional[List[str]] = None) -> Tuple[Any, List[str]]:
        """
        Create a bar plot showing average values for each sample.
        Consolidated from processing.py plot_tpm_bar_chart function.
        """
        plt, _, _, _ = self.get_matplotlib_components()
        if not plt:
            debug_print("Could not load matplotlib")
            return None, []
        
        try:
            fig, ax = plt.subplots(figsize=self.settings.figure_size, dpi=self.settings.dpi)
            
            # Calculate number of samples
            num_samples = len(full_sample_data.columns) // self.config.num_columns_per_sample
            debug_print(f"Creating bar plot for {num_samples} samples")
            
            sample_averages = []
            sample_stds = []
            extracted_sample_names = []
            
            for i in range(num_samples):
                start_col = i * self.config.num_columns_per_sample
                end_col = start_col + self.config.num_columns_per_sample
                sample_data = full_sample_data.iloc[:, start_col:end_col]
                
                # Get y data for averaging
                y_data = self.get_y_data_for_plot_type(sample_data, self.state.current_plot_type)
                y_data_clean = pd.to_numeric(y_data, errors='coerce').dropna()
                
                if not y_data_clean.empty:
                    avg_value = y_data_clean.mean()
                    std_value = y_data_clean.std()
                    
                    sample_averages.append(avg_value)
                    sample_stds.append(std_value if not pd.isna(std_value) else 0)
                    
                    # Get sample name
                    if sample_names and i < len(sample_names):
                        sample_name = sample_names[i]
                    else:
                        sample_name = sample_data.columns[5] if len(sample_data.columns) > 5 else f"Sample {i+1}"
                    
                    # Wrap long sample names
                    wrapped_name = wrap_text(sample_name, self.settings.max_checkbox_width)
                    extracted_sample_names.append(wrapped_name)
                else:
                    debug_print(f"Sample {i+1}: No valid data for averaging")
            
            if sample_averages:
                # Create bar plot
                x_positions = range(len(sample_averages))
                bars = ax.bar(x_positions, sample_averages, yerr=sample_stds, 
                             capsize=5, alpha=self.settings.transparency,
                             color=[self.settings.get_color(i) for i in range(len(sample_averages))])
                
                # Configure plot
                ax.set_xlabel('Samples')
                ax.set_ylabel(self.config.get_y_label(self.state.current_plot_type))
                ax.set_title(f"Average {self.state.current_plot_type}")
                ax.set_xticks(x_positions)
                ax.set_xticklabels(extracted_sample_names, rotation=45, ha='right')
                
                if self.settings.enable_grid:
                    ax.grid(True, alpha=0.3, axis='y')
                
                if self.settings.tight_layout:
                    plt.tight_layout()
                
                debug_print(f"Bar plot created with {len(sample_averages)} samples")
            else:
                debug_print("No valid data for bar plot")
            
            return fig, extracted_sample_names
            
        except Exception as e:
            debug_print(f"Error creating bar plot: {e}")
            return None, []
    
    # ===================== PLOT STATE MANAGEMENT =====================
    
    def set_plot_type(self, plot_type: str):
        """Set the current plot type."""
        if plot_type in self.config.supported_plot_types:
            self.state.current_plot_type = plot_type
            debug_print(f"Plot type set to: {plot_type}")
        else:
            debug_print(f"Unsupported plot type: {plot_type}")
    
    def get_plot_type(self) -> str:
        """Get the current plot type."""
        return self.state.current_plot_type
    
    def clear_plot_state(self):
        """Clear current plot state and release resources."""
        plt, _, _, _ = self.get_matplotlib_components()
        
        if plt and self.state.figure:
            plt.close(self.state.figure)
        
        self.state.clear_state()
        debug_print("Plot state and resources cleared")
    
    def cache_plot_data(self, key: str, data: Any):
        """Cache plot data for reuse."""
        self.plot_cache[key] = {
            'data': data,
            'timestamp': datetime.now()
        }
        debug_print(f"Plot data cached with key: {key}")
    
    def get_cached_plot_data(self, key: str) -> Optional[Any]:
        """Get cached plot data if still valid."""
        if key in self.plot_cache:
            cached_item = self.plot_cache[key]
            age_hours = (datetime.now() - cached_item['timestamp']).total_seconds() / 3600
            
            if age_hours < self.cache_duration_hours:
                debug_print(f"Retrieved cached plot data for key: {key}")
                return cached_item['data']
            else:
                del self.plot_cache[key]
                debug_print(f"Expired cache data removed for key: {key}")
        
        return None
    
    def clear_plot_cache(self):
        """Clear all cached plot data."""
        cache_count = len(self.plot_cache)
        self.plot_cache.clear()
        debug_print(f"Cleared {cache_count} cached plot items")
    
    # ===================== UTILITY METHODS =====================
    
    def get_supported_plot_types(self) -> List[str]:
        """Get list of supported plot types."""
        return self.config.supported_plot_types.copy()
    
    def update_settings(self, **kwargs):
        """Update plot settings."""
        for key, value in kwargs.items():
            if hasattr(self.settings, key):
                setattr(self.settings, key, value)
                debug_print(f"Updated setting {key} = {value}")
            else:
                debug_print(f"Unknown setting: {key}")
    
    def update_config(self, **kwargs):
        """Update plot configuration."""
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
                debug_print(f"Updated config {key} = {value}")
            else:
                debug_print(f"Unknown config: {key}")
    
    def get_model_stats(self) -> Dict[str, Any]:
        """Get statistics about the plot model."""
        return {
            'current_plot_type': self.state.current_plot_type,
            'matplotlib_loaded': self._matplotlib_loaded,
            'pandas_loaded': self._pandas_loaded,
            'cached_items': len(self.plot_cache),
            'supported_plot_types': len(self.config.supported_plot_types),
            'last_update': self.state.last_update,
            'figure_active': self.state.figure is not None,
            'canvas_active': self.state.canvas is not None,
            'plot_settings': {
                'figure_size': self.settings.figure_size,
                'enable_checkboxes': self.settings.enable_checkboxes,
                'enable_zoom': self.settings.enable_zoom,
                'enable_grid': self.settings.enable_grid
            }
        }


# Export the main classes
__all__ = ['PlotModel', 'PlotConfig', 'PlotSettings', 'PlotState']

# Debug output for model initialization
debug_print("PlotModel module loaded successfully")
debug_print("Available classes: PlotModel, PlotConfig, PlotSettings, PlotState")
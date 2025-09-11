# views/widgets/plot_widget.py
"""
views/widgets/plot_widget.py
Plot widget for displaying matplotlib plots.
This contains plotting UI functionality from plot_manager.py.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List, Callable, Tuple
import threading


class PlotWidget:
    """Widget for displaying matplotlib plots."""
    
    def __init__(self, parent: tk.Widget, controller: Optional[Any] = None):
        """Initialize the plot widget."""
        self.parent = parent
        self.controller = controller
        
        # Plot components
        self.figure: Optional[Any] = None
        self.canvas: Optional[Any] = None
        self.axes: Optional[Any] = None
        self.toolbar: Optional[Any] = None
        self.lines: Optional[List] = None
        self.check_buttons: Optional[Any] = None
        self.checkbox_cid: Optional[Any] = None
        
        # Plot variables
        self.plot_dropdown: Optional[ttk.Combobox] = None
        self.selected_plot_type = tk.StringVar(value="TPM")
        self.line_labels: List[str] = []
        self.original_lines_data: List = []
        
        # Plotting state
        self.plot_options = ["TPM", "Normalized TPM", "Draw Pressure", "Resistance", 
                           "Power Efficiency", "TPM (Bar)"]
        self.is_user_test_simulation = False
        self.num_columns_per_sample = 12
        
        # Callbacks
        self.on_plot_type_changed: Optional[Callable] = None
        self.on_checkbox_click: Optional[Callable] = None
        self.on_bar_checkbox_click: Optional[Callable] = None
        
        print("DEBUG: PlotWidget initialized")
    
    def get_matplotlib_components(self):
        """Lazy load matplotlib components."""
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
            from matplotlib.widgets import CheckButtons
            print("DEBUG: PlotWidget - matplotlib components loaded")
            return plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons
        except ImportError as e:
            print(f"ERROR: PlotWidget - could not load matplotlib: {e}")
            return None, None, None, None
    
    def create_plot_frame(self, parent_frame: tk.Widget) -> tk.Frame:
        """Create a frame for the plot."""
        plot_frame = tk.Frame(parent_frame, bg='white')
        plot_frame.pack(fill="both", expand=True)
        print("DEBUG: PlotWidget - plot frame created")
        return plot_frame
    
    def plot_all_samples(self, frame: tk.Widget, full_sample_data: Any, num_columns: int = 12):
        """Plot all samples with the current plot type."""
        plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()
        if not plt:
            print("ERROR: PlotWidget - matplotlib not available")
            return
        
        try:
            # Clear existing plot
            self.clear_plot_area(frame)
            
            # Get current plot type
            selected_plot = self.selected_plot_type.get()
            
            # Determine if this is a User Test Simulation
            self.is_user_test_simulation = self._detect_user_test_simulation(full_sample_data)
            
            if self.is_user_test_simulation:
                self._plot_user_test_simulation(frame, full_sample_data, selected_plot, plt, FigureCanvasTkAgg)
            else:
                self._plot_standard_samples(frame, full_sample_data, selected_plot, num_columns, plt, FigureCanvasTkAgg, CheckButtons)
            
            # Add plot dropdown
            self.add_plot_dropdown(frame)
            
            print(f"DEBUG: PlotWidget - plotted {selected_plot} for {'User Test Simulation' if self.is_user_test_simulation else 'standard'} data")
            
        except Exception as e:
            print(f"ERROR: PlotWidget - error in plot_all_samples: {e}")
            self._show_plot_error(frame, str(e))
    
    def _detect_user_test_simulation(self, full_sample_data: Any) -> bool:
        """Detect if this is User Test Simulation data."""
        try:
            pd = self._get_pandas()
            if not pd or not isinstance(full_sample_data, pd.DataFrame):
                return False
            
            # Check for User Test Simulation indicators
            columns = [str(col).lower() for col in full_sample_data.columns]
            uts_indicators = ['user test simulation', 'uts', 'simulation', 'user_test']
            
            return any(indicator in ' '.join(columns) for indicator in uts_indicators)
        except Exception:
            return False
    
    def _plot_standard_samples(self, frame: tk.Widget, full_sample_data: Any, selected_plot: str, 
                             num_columns: int, plt: Any, FigureCanvasTkAgg: Any, CheckButtons: Any):
        """Plot standard sample data."""
        pd = self._get_pandas()
        if not pd or full_sample_data.empty:
            self._show_empty_plot_message(frame)
            return
        
        # Create figure
        self.figure = plt.figure(figsize=(10, 6))
        self.axes = self.figure.add_subplot(111)
        
        # Process data and create plots
        sample_names = self._extract_sample_names(full_sample_data, num_columns)
        plot_data = self._prepare_plot_data(full_sample_data, selected_plot, num_columns)
        
        if not plot_data:
            self._show_empty_plot_message(frame)
            return
        
        # Create plots
        self.lines = []
        self.line_labels = []
        
        for i, (sample_name, data) in enumerate(plot_data.items()):
            if data is not None and len(data) > 0:
                line, = self.axes.plot(data, marker='o', label=sample_name, linewidth=2)
                self.lines.append(line)
                self.line_labels.append(sample_name)
        
        # Configure plot
        self._configure_plot_appearance(selected_plot)
        
        # Embed plot in frame
        self._embed_plot_in_frame(frame, FigureCanvasTkAgg)
        
        # Add checkboxes for line visibility
        if self.lines and len(self.lines) > 1:
            self._add_plot_checkboxes(CheckButtons)
        
        print(f"DEBUG: PlotWidget - standard plot created with {len(self.lines)} lines")
    
    def _plot_user_test_simulation(self, frame: tk.Widget, full_sample_data: Any, selected_plot: str, 
                                 plt: Any, FigureCanvasTkAgg: Any):
        """Plot User Test Simulation data with split layout."""
        pd = self._get_pandas()
        if not pd or full_sample_data.empty:
            self._show_empty_plot_message(frame)
            return
        
        # Create figure with subplots
        self.figure = plt.figure(figsize=(12, 8))
        
        # Split data into two categories
        condition_data, baseline_data = self._split_simulation_data(full_sample_data)
        
        if condition_data.empty and baseline_data.empty:
            self._show_empty_plot_message(frame)
            return
        
        # Create subplots
        if not condition_data.empty and not baseline_data.empty:
            ax1 = self.figure.add_subplot(121)
            ax2 = self.figure.add_subplot(122)
            self._plot_simulation_subplot(ax1, condition_data, selected_plot, "Condition")
            self._plot_simulation_subplot(ax2, baseline_data, selected_plot, "Baseline")
        else:
            # Single plot if only one category has data
            self.axes = self.figure.add_subplot(111)
            data_to_plot = condition_data if not condition_data.empty else baseline_data
            title = "Condition" if not condition_data.empty else "Baseline"
            self._plot_simulation_subplot(self.axes, data_to_plot, selected_plot, title)
        
        # Embed plot
        self._embed_plot_in_frame(frame, FigureCanvasTkAgg)
        
        print("DEBUG: PlotWidget - User Test Simulation plot created")
    
    def _split_simulation_data(self, full_sample_data: Any) -> Tuple[Any, Any]:
        """Split simulation data into condition and baseline."""
        pd = self._get_pandas()
        
        # Split based on sample names or data patterns
        condition_data = pd.DataFrame()
        baseline_data = pd.DataFrame()
        
        try:
            # Implementation depends on specific data structure
            # This is a simplified version
            half_point = len(full_sample_data.columns) // 2
            condition_data = full_sample_data.iloc[:, :half_point]
            baseline_data = full_sample_data.iloc[:, half_point:]
        except Exception as e:
            print(f"DEBUG: PlotWidget - error splitting simulation data: {e}")
        
        return condition_data, baseline_data
    
    def _plot_simulation_subplot(self, ax: Any, data: Any, plot_type: str, title: str):
        """Plot a subplot for simulation data."""
        try:
            # Extract and plot data for this subplot
            sample_names = self._extract_sample_names(data, 6)  # Assume 6 columns per sample for simulation
            plot_data = self._prepare_plot_data(data, plot_type, 6)
            
            for sample_name, values in plot_data.items():
                if values is not None and len(values) > 0:
                    ax.plot(values, marker='o', label=sample_name, linewidth=2)
            
            ax.set_title(f"{title} - {plot_type}")
            ax.set_xlabel("Puff Number")
            ax.set_ylabel(self._get_y_label(plot_type))
            ax.legend()
            ax.grid(True, alpha=0.3)
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error plotting simulation subplot: {e}")
    
    def _extract_sample_names(self, data: Any, num_columns: int) -> List[str]:
        """Extract sample names from data columns."""
        try:
            sample_names = []
            for i in range(0, len(data.columns), num_columns):
                column_name = str(data.columns[i])
                if 'Sample' in column_name:
                    # Extract sample name from column header
                    parts = column_name.split()
                    if len(parts) >= 2:
                        sample_names.append(parts[1])
                    else:
                        sample_names.append(f"Sample {i//num_columns + 1}")
                else:
                    sample_names.append(f"Sample {i//num_columns + 1}")
            
            return sample_names
        except Exception as e:
            print(f"DEBUG: PlotWidget - error extracting sample names: {e}")
            return []
    
    def _prepare_plot_data(self, data: Any, plot_type: str, num_columns: int) -> Dict[str, List]:
        """Prepare data for plotting based on plot type."""
        try:
            plot_data = {}
            sample_names = self._extract_sample_names(data, num_columns)
            
            for i, sample_name in enumerate(sample_names):
                start_col = i * num_columns
                end_col = start_col + num_columns
                
                if end_col <= len(data.columns):
                    sample_data = data.iloc[:, start_col:end_col]
                    values = self._extract_values_for_plot_type(sample_data, plot_type)
                    plot_data[sample_name] = values
            
            return plot_data
        except Exception as e:
            print(f"DEBUG: PlotWidget - error preparing plot data: {e}")
            return {}
    
    def _extract_values_for_plot_type(self, sample_data: Any, plot_type: str) -> List[float]:
        """Extract values for specific plot type from sample data."""
        try:
            pd = self._get_pandas()
            import numpy as np
            
            # Map plot types to column patterns
            column_mapping = {
                "TPM": ["tpm", "total particulate matter"],
                "Normalized TPM": ["tpm", "total particulate matter"], 
                "Draw Pressure": ["draw pressure", "pressure"],
                "Resistance": ["resistance"],
                "Power Efficiency": ["power efficiency", "efficiency"],
                "TPM (Bar)": ["tpm", "total particulate matter"]
            }
            
            # Find relevant columns
            target_patterns = column_mapping.get(plot_type, [plot_type.lower()])
            relevant_cols = []
            
            for col in sample_data.columns:
                col_str = str(col).lower()
                if any(pattern in col_str for pattern in target_patterns):
                    relevant_cols.append(col)
            
            if not relevant_cols:
                # Fallback to first numeric column
                for col in sample_data.columns:
                    try:
                        pd.to_numeric(sample_data[col], errors='coerce')
                        relevant_cols.append(col)
                        break
                    except:
                        continue
            
            # Extract values
            values = []
            for col in relevant_cols[:10]:  # Limit to first 10 puffs
                try:
                    val = pd.to_numeric(sample_data[col].iloc[0] if len(sample_data) > 0 else 0, errors='coerce')
                    if not pd.isna(val) and val > 0:
                        values.append(float(val))
                except:
                    continue
            
            # Apply normalization for Normalized TPM
            if plot_type == "Normalized TPM" and values:
                first_val = values[0] if values[0] != 0 else 1
                values = [v / first_val for v in values]
            
            return values
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error extracting values for {plot_type}: {e}")
            return []
    
    def _configure_plot_appearance(self, plot_type: str):
        """Configure plot appearance and labels."""
        if not self.axes:
            return
        
        try:
            self.axes.set_xlabel("Puff Number")
            self.axes.set_ylabel(self._get_y_label(plot_type))
            self.axes.set_title(f"{plot_type} vs Puff Number")
            self.axes.legend()
            self.axes.grid(True, alpha=0.3)
            
            # Tight layout for better appearance
            if self.figure:
                self.figure.tight_layout()
                
        except Exception as e:
            print(f"DEBUG: PlotWidget - error configuring plot appearance: {e}")
    
    def _get_y_label(self, plot_type: str) -> str:
        """Get appropriate Y-axis label for plot type."""
        labels = {
            "TPM": "TPM (mg)",
            "Normalized TPM": "Normalized TPM",
            "Draw Pressure": "Draw Pressure (Pa)",
            "Resistance": "Resistance (Pa·s/mL)",
            "Power Efficiency": "Power Efficiency (%)",
            "TPM (Bar)": "TPM (mg)"
        }
        return labels.get(plot_type, plot_type)
    
    def _embed_plot_in_frame(self, frame: tk.Widget, FigureCanvasTkAgg: Any):
        """Embed matplotlib figure in tkinter frame."""
        try:
            # Create canvas
            self.canvas = FigureCanvasTkAgg(self.figure, frame)
            self.canvas.draw()
            
            # Pack canvas
            canvas_widget = self.canvas.get_tk_widget()
            canvas_widget.pack(fill="both", expand=True)
            
            # Add toolbar
            toolbar_frame = tk.Frame(frame)
            toolbar_frame.pack(side="bottom", fill="x")
            
            self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
            self.toolbar.update()
            
            # Bind events
            self.canvas.mpl_connect('scroll_event', self._on_scroll)
            self.canvas.mpl_connect('button_press_event', self._on_click)
            
            print("DEBUG: PlotWidget - plot embedded successfully")
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error embedding plot: {e}")
    
    def _add_plot_checkboxes(self, CheckButtons: Any):
        """Add checkboxes for toggling line visibility."""
        try:
            if not self.lines or not self.figure:
                return
            
            # Create checkbox axes
            checkbox_ax = self.figure.add_axes([0.85, 0.4, 0.1, 0.15])
            
            # Get line visibility
            line_visibility = [line.get_visible() for line in self.lines]
            
            # Create checkboxes
            self.check_buttons = CheckButtons(checkbox_ax, self.line_labels, line_visibility)
            
            # Connect callback
            if self.on_checkbox_click:
                self.checkbox_cid = self.check_buttons.on_clicked(self.on_checkbox_click)
            else:
                self.checkbox_cid = self.check_buttons.on_clicked(self._default_checkbox_click)
            
            print("DEBUG: PlotWidget - checkboxes added")
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error adding checkboxes: {e}")
    
    def _default_checkbox_click(self, label: str):
        """Default checkbox click handler."""
        try:
            if label in self.line_labels:
                index = self.line_labels.index(label)
                if index < len(self.lines):
                    line = self.lines[index]
                    line.set_visible(not line.get_visible())
                    if self.canvas:
                        self.canvas.draw_idle()
                    print(f"DEBUG: PlotWidget - toggled visibility for {label}")
        except Exception as e:
            print(f"DEBUG: PlotWidget - error in checkbox click: {e}")
    
    def add_plot_dropdown(self, frame: tk.Widget):
        """Add plot type dropdown to frame."""
        try:
            # Create dropdown frame
            dropdown_frame = tk.Frame(frame)
            dropdown_frame.pack(side="top", fill="x", pady=5)
            
            # Label
            tk.Label(dropdown_frame, text="Plot Type:").pack(side="left", padx=5)
            
            # Dropdown
            plot_options = (self.user_test_simulation_plot_options if self.is_user_test_simulation 
                          else self.plot_options)
            
            self.plot_dropdown = ttk.Combobox(dropdown_frame, textvariable=self.selected_plot_type,
                                             values=plot_options, state="readonly", width=20)
            self.plot_dropdown.pack(side="left", padx=5)
            self.plot_dropdown.bind('<<ComboboxSelected>>', self._on_plot_dropdown_change)
            
            print("DEBUG: PlotWidget - plot dropdown added")
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error adding plot dropdown: {e}")
    
    def clear_plot_area(self, frame: tk.Widget = None):
        """Clear the plot area and release resources."""
        try:
            plt, _, _, _ = self.get_matplotlib_components()
            
            if frame:
                for widget in frame.winfo_children():
                    widget.destroy()
            
            if self.figure and plt:
                plt.close(self.figure)
                self.figure = None
            
            if self.canvas:
                try:
                    self.canvas.get_tk_widget().destroy()
                except:
                    pass
                self.canvas = None
            
            self.axes = None
            self.lines = None
            self.line_labels = []
            self.original_lines_data = []
            self.check_buttons = None
            self.checkbox_cid = None
            
            print("DEBUG: PlotWidget - plot area cleared")
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error clearing plot area: {e}")
    
    def _show_empty_plot_message(self, frame: tk.Widget):
        """Show message when no data is available for plotting."""
        try:
            for widget in frame.winfo_children():
                widget.destroy()
            
            message_label = tk.Label(frame, 
                                   text="No data to plot yet.\nUse 'Collect Data' to add measurements.",
                                   font=("Arial", 12), fg="gray", justify="center")
            message_label.pack(expand=True)
            
            print("DEBUG: PlotWidget - empty plot message shown")
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error showing empty message: {e}")
    
    def _show_plot_error(self, frame: tk.Widget, error_msg: str):
        """Show error message when plotting fails."""
        try:
            for widget in frame.winfo_children():
                widget.destroy()
            
            error_label = tk.Label(frame,
                                 text=f"Error creating plot:\n{error_msg}",
                                 font=("Arial", 12), fg="red", justify="center")
            error_label.pack(expand=True, pady=20)
            
            print(f"DEBUG: PlotWidget - error message shown: {error_msg}")
            
        except Exception as e:
            print(f"DEBUG: PlotWidget - error showing error message: {e}")
    
    # Event handlers
    def _on_plot_dropdown_change(self, event=None):
        """Handle plot dropdown selection change."""
        if self.on_plot_type_changed:
            self.on_plot_type_changed(self.selected_plot_type.get())
        print(f"DEBUG: PlotWidget - plot type changed to: {self.selected_plot_type.get()}")
    
    def _on_scroll(self, event):
        """Handle mouse scroll for zooming."""
        try:
            if not self.axes or not event.inaxes:
                return
            
            # Zoom functionality
            base_scale = 1.1
            if event.button == 'up':
                scale_factor = 1 / base_scale
            elif event.button == 'down':
                scale_factor = base_scale
            else:
                return
            
            # Get current axis limits
            cur_xlim = self.axes.get_xlim()
            cur_ylim = self.axes.get_ylim()
            
            # Calculate new limits
            xdata = event.xdata
            ydata = event.ydata
            
            new_xlim = [xdata - (xdata - cur_xlim[0]) * scale_factor,
                       xdata + (cur_xlim[1] - xdata) * scale_factor]
            new_ylim = [ydata - (ydata - cur_ylim[0]) * scale_factor,
                       ydata + (cur_ylim[1] - ydata) * scale_factor]
            
            self.axes.set_xlim(new_xlim)
            self.axes.set_ylim(new_ylim)
            
            if self.canvas:
                self.canvas.draw_idle()
                
        except Exception as e:
            print(f"DEBUG: PlotWidget - error in scroll event: {e}")
    
    def _on_click(self, event):
        """Handle mouse click events."""
        try:
            if event.button == 3 and event.inaxes:  # Right click
                # Could add context menu here
                pass
        except Exception as e:
            print(f"DEBUG: PlotWidget - error in click event: {e}")
    
    # Utility methods
    def _get_pandas(self):
        """Get pandas with lazy loading."""
        try:
            import pandas as pd
            return pd
        except ImportError:
            return None
    
    def update_plot_options(self, options: List[str]):
        """Update available plot options."""
        self.plot_options = options
        if self.plot_dropdown:
            self.plot_dropdown['values'] = options
        print(f"DEBUG: PlotWidget - plot options updated: {options}")
    
    def set_plot_type(self, plot_type: str):
        """Set the current plot type."""
        if plot_type in self.plot_options:
            self.selected_plot_type.set(plot_type)
            if self.plot_dropdown:
                self.plot_dropdown.set(plot_type)
            print(f"DEBUG: PlotWidget - plot type set to: {plot_type}")
    
    def get_current_plot_type(self) -> str:
        """Get the current plot type."""
        return self.selected_plot_type.get()
    
    def export_plot(self, filename: str):
        """Export current plot to file."""
        try:
            if self.figure:
                self.figure.savefig(filename, dpi=300, bbox_inches='tight')
                print(f"DEBUG: PlotWidget - plot exported to: {filename}")
                return True
        except Exception as e:
            print(f"DEBUG: PlotWidget - error exporting plot: {e}")
        return False
    
    def refresh_plot(self):
        """Refresh the current plot."""
        try:
            if self.canvas:
                self.canvas.draw()
                print("DEBUG: PlotWidget - plot refreshed")
        except Exception as e:
            print(f"DEBUG: PlotWidget - error refreshing plot: {e}")
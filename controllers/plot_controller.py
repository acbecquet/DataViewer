# controllers/plot_controller.py
"""
controllers/plot_controller.py
Plot generation controller that coordinates between plot service and models.
This consolidates functionality from plot_manager.py and related plotting code.
"""

import os
import re
import copy
import math
import threading
import traceback
from typing import Optional, Dict, Any, List, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, StringVar

# Local imports
from utils import debug_print, APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT, wrap_text
from models.plot_model import PlotModel
from models.data_model import DataModel
import processing


class PlotController:
    """Controller for plot generation and display operations."""
    
    def __init__(self, plot_model: PlotModel, data_model: DataModel, plot_service: Any):
        """Initialize the plot controller."""
        self.plot_model = plot_model
        self.data_model = data_model
        self.plot_service = plot_service
        
        # Cross-controller references
        self.data_controller: Optional['DataController'] = None
        
        # GUI reference will be set when connected
        self.gui = None
        
        # Lazy loading infrastructure
        self._matplotlib_loaded = False
        self._pandas_loaded = False
        
        # Matplotlib components (loaded lazily)
        self.plt = None
        self.FigureCanvasTkAgg = None
        self.NavigationToolbar2Tk = None
        self.CheckButtons = None
        self.pd = None
        
        # Plot state management
        self.figure = None
        self.canvas = None
        self.axes = None
        self.lines = None
        self.check_buttons = None
        self.checkbox_cid = None
        self.plot_dropdown = None
        self.dropdown_frame = None
        
        # Plot interaction state
        self.label_mapping = {}
        self.wrapped_labels = []
        
        print("DEBUG: PlotController initialized")
        print(f"DEBUG: Connected to PlotModel and DataModel")
        print("DEBUG: Lazy loading system initialized for matplotlib components")
    
    def set_gui_reference(self, gui):
        """Set reference to main GUI for UI operations."""
        self.gui = gui
        print("DEBUG: PlotController connected to GUI")
    
    def set_data_controller(self, data_controller: 'DataController'):
        """Set reference to data controller."""
        self.data_controller = data_controller
        print("DEBUG: PlotController connected to DataController")
    
    def _lazy_import_matplotlib_components(self):
        """Lazy import matplotlib components."""
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
            from matplotlib.widgets import CheckButtons
            print("TIMING: Lazy loaded matplotlib components in PlotController")
            return plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons
        except ImportError as e:
            print(f"ERROR: importing matplotlib: {e}")
            return None, None, None, None
    
    def _lazy_import_pandas(self):
        """Lazy import pandas."""
        try:
            import pandas as pd
            print("TIMING: Lazy loaded pandas in PlotController")
            return pd
        except ImportError as e:
            print(f"ERROR: importing pandas: {e}")
            return None
    
    def get_matplotlib_components(self):
        """Get matplotlib components with lazy loading."""
        if not self._matplotlib_loaded:
            self.plt, self.FigureCanvasTkAgg, self.NavigationToolbar2Tk, self.CheckButtons = self._lazy_import_matplotlib_components()
            self._matplotlib_loaded = True
        return self.plt, self.FigureCanvasTkAgg, self.NavigationToolbar2Tk, self.CheckButtons
    
    def get_pandas(self):
        """Get pandas with lazy loading."""
        if not self._pandas_loaded:
            self.pd = self._lazy_import_pandas()
            self._pandas_loaded = True
        return self.pd
    
    def generate_plot(self, plot_type: str, frame: ttk.Frame, full_sample_data, 
                     num_columns_per_sample: int = 12) -> bool:
        """Generate a plot of the specified type."""
        debug_print(f"DEBUG: PlotController generating {plot_type} plot")
        debug_print(f"DEBUG: Data shape: {full_sample_data.shape if hasattr(full_sample_data, 'shape') else 'No shape available'}")
        
        try:
            # Set plot type in model
            self.plot_model.set_plot_type(plot_type)
            
            # Update GUI plot type if available
            if self.gui and hasattr(self.gui, 'selected_plot_type'):
                self.gui.selected_plot_type.set(plot_type)
            
            # Generate the plot
            success = self.plot_all_samples(frame, full_sample_data, num_columns_per_sample)
            
            if success:
                debug_print(f"DEBUG: PlotController successfully generated {plot_type} plot")
                return True
            else:
                debug_print(f"ERROR: PlotController failed to generate {plot_type} plot")
                return False
                
        except Exception as e:
            debug_print(f"ERROR: PlotController failed to generate plot: {e}")
            traceback.print_exc()
            return False
    
    def plot_all_samples(self, frame: ttk.Frame, full_sample_data, num_columns_per_sample: int) -> bool:
        """Plot all samples in the given frame with enhanced empty data handling."""
        debug_print(f"DEBUG: PlotController plot_all_samples called with data shape: {full_sample_data.shape if hasattr(full_sample_data, 'shape') else 'No shape'}")
        
        try:
            # Get components
            plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()
            pd = self.get_pandas()
            
            # Error handling
            if not pd:
                debug_print("ERROR: Could not load pandas in plot_all_samples")
                return False
            
            if not plt:
                debug_print("ERROR: Could not load matplotlib in plot_all_samples")
                return False
            
            # Validate data
            if not hasattr(full_sample_data, 'empty') or not hasattr(full_sample_data, 'shape'):
                debug_print("ERROR: full_sample_data is not a valid DataFrame-like object")
                return False
            
            # Clear frame contents
            for widget in frame.winfo_children():
                widget.pack_forget()
            
            # Check if data is empty or invalid
            if full_sample_data.empty:
                debug_print("DEBUG: Data is completely empty - showing placeholder")
                self.show_empty_plot_placeholder(frame, "No data loaded yet.\nUse 'Collect Data' to add measurements.")
                return True
            
            # Check if data contains only NaN values
            if full_sample_data.isna().all().all():
                debug_print("DEBUG: Data contains only NaN values - showing placeholder")
                self.show_empty_plot_placeholder(frame, "No measurement data available yet.\nUse 'Collect Data' to add measurements.")
                return True
            
            # Check if there's any numeric data for plotting
            numeric_data = full_sample_data.apply(pd.to_numeric, errors='coerce')
            if numeric_data.isna().all().all():
                debug_print("DEBUG: No numeric data available for plotting - showing placeholder")
                self.show_empty_plot_placeholder(frame, "No numeric data available for plotting.\nUse 'Collect Data' to add measurement values.")
                return True
            
            # Get current plot type
            plot_type = "TPM"  # Default
            if self.gui and hasattr(self.gui, 'selected_plot_type'):
                plot_type = self.gui.selected_plot_type.get()
            elif hasattr(self.plot_model, 'current_plot_type'):
                plot_type = self.plot_model.current_plot_type
            
            debug_print(f"DEBUG: Using plot type: {plot_type}")
            
            # Get sample names if available
            sample_names = None
            if self.gui and hasattr(self.gui, 'line_labels'):
                sample_names = self.gui.line_labels
            
            # Check for User Test Simulation
            current_sheet = ""
            if self.gui and hasattr(self.gui, 'selected_sheet'):
                current_sheet = self.gui.selected_sheet.get()
            
            is_user_test_simulation = current_sheet in ["User Test Simulation", "User Simulation Test"]
            
            if is_user_test_simulation:
                debug_print("DEBUG: Generating User Test Simulation plot")
                fig, extracted_sample_names = processing.plot_user_test_simulation_samples(
                    full_sample_data, num_columns_per_sample, plot_type, sample_names
                )
            else:
                debug_print("DEBUG: Generating standard plot")
                fig, extracted_sample_names = processing.plot_all_samples(
                    full_sample_data, num_columns_per_sample, plot_type, sample_names
                )
            
            if fig is None:
                debug_print("ERROR: Failed to generate figure")
                return False
            
            # Embed the plot
            canvas = self.embed_plot_in_frame(fig, frame)
            
            if canvas:
                # Update line labels
                if self.gui and hasattr(self.gui, 'line_labels'):
                    self.gui.line_labels = extracted_sample_names
                
                debug_print(f"DEBUG: Plot embedded successfully with {len(extracted_sample_names)} samples")
                return True
            else:
                debug_print("ERROR: Failed to embed plot in frame")
                return False
                
        except Exception as e:
            debug_print(f"ERROR: Exception in plot_all_samples: {e}")
            traceback.print_exc()
            return False
    
    def embed_plot_in_frame(self, fig, frame: ttk.Frame):
        """Embed a Matplotlib figure into a Tkinter frame with proper layout control."""
        debug_print("DEBUG: PlotController embedding plot in frame")
        
        try:
            plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()
            
            # Error handling
            if not plt or not FigureCanvasTkAgg or not NavigationToolbar2Tk:
                debug_print("ERROR: Could not load matplotlib components in embed_plot_in_frame")
                return None
            
            # Clear existing widgets
            for widget in frame.winfo_children():
                widget.destroy()
            
            # Create a separate frame for the dropdown - ensure it's visible
            self.dropdown_frame = ttk.Frame(frame)
            self.dropdown_frame.pack(side='bottom', fill='x', pady=10)
            
            # Create a container frame for the plot
            plot_container = ttk.Frame(frame)
            plot_container.pack(fill='both', expand=True, pady=(0, 0))
            
            if self.figure:
                plt.close(self.figure)
            
            # Check if this is a split plot and adjust margins accordingly
            is_split_plot = hasattr(fig, 'is_split_plot') and fig.is_split_plot
            
            if is_split_plot:
                # For split plots, adjust margins to accommodate checkboxes for both plots
                fig.subplots_adjust(right=0.80)
                debug_print("DEBUG: Split plot detected, adjusted margins")
            else:
                # Standard right margin - don't make it too small
                fig.subplots_adjust(right=0.82)
            
            # Embed figure in the plot container
            self.canvas = FigureCanvasTkAgg(fig, master=plot_container)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill='both', expand=True)
            
            # Add toolbar to the plot container
            toolbar = NavigationToolbar2Tk(self.canvas, plot_container)
            toolbar.update()
            
            # Save references to the figure and its axes
            self.figure = fig
            
            if is_split_plot:
                # For split plots, we have multiple axes
                self.axes = fig.axes  # List of axes
                self.lines = []  # Will be handled by the split plot logic
                debug_print(f"DEBUG: Split plot - found {len(self.axes)} axes")
            else:
                # Standard single plot
                self.axes = fig.gca()
                self.lines = self.axes.lines
            
            # Bind scroll event for zooming
            self.canvas.mpl_connect("scroll_event", lambda event: self.zoom(event))
            
            # Add checkboxes for toggling plot elements
            self.add_checkboxes()
            
            debug_print("DEBUG: Plot embedded successfully")
            return self.canvas
            
        except Exception as e:
            debug_print(f"ERROR: Failed to embed plot in frame: {e}")
            traceback.print_exc()
            return None
    
    def show_empty_plot_placeholder(self, frame: ttk.Frame, message: str):
        """Show a placeholder message when no plot data is available."""
        debug_print(f"DEBUG: Showing empty plot placeholder: {message}")
        
        try:
            # Clear frame
            for widget in frame.winfo_children():
                widget.destroy()
            
            # Create placeholder label
            placeholder_label = ttk.Label(
                frame,
                text=message,
                font=('Arial', 14),
                foreground='gray',
                anchor='center',
                justify='center'
            )
            placeholder_label.pack(expand=True, fill='both')
            
            # Update the frame
            frame.update_idletasks()
            
        except Exception as e:
            debug_print(f"ERROR: Failed to show empty plot placeholder: {e}")
    
    def add_checkboxes(self, sample_names=None):
        """Add checkboxes to toggle visibility of plot elements."""
        debug_print("DEBUG: PlotController adding checkboxes")
        
        try:
            plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()
            
            if not plt:
                debug_print("ERROR: Could not load matplotlib")
                return None
            
            # Initialize tracking variables
            if self.gui:
                if not hasattr(self.gui, 'line_labels'):
                    self.gui.line_labels = []
                if not hasattr(self.gui, 'original_lines_data'):
                    self.gui.original_lines_data = []
            
            self.label_mapping = {}
            wrapped_labels = []
            
            if self.check_buttons:
                self.check_buttons.ax.clear()
                self.check_buttons = None
            
            # Determine plot type
            plot_type = "TPM"
            if self.gui and hasattr(self.gui, 'selected_plot_type'):
                plot_type = self.gui.selected_plot_type.get()
            
            is_bar_chart = plot_type == "TPM (Bar)"
            
            # Check if this is a User Test Simulation split plot
            is_split_plot = hasattr(self.figure, 'is_split_plot') and self.figure.is_split_plot
            is_split_bar_chart = is_split_plot and hasattr(self.figure, 'is_bar_chart') and self.figure.is_bar_chart
            debug_print(f"DEBUG: add_checkboxes - is_split_plot: {is_split_plot}, is_bar_chart: {is_bar_chart}, is_split_bar_chart: {is_split_bar_chart}")
            
            # Get sample names
            if sample_names is None and self.gui:
                sample_names = getattr(self.gui, 'line_labels', [])
            
            if sample_names and self.gui:
                self.gui.line_labels = sample_names
                debug_print(f"DEBUG: Using sample_names: {sample_names}")
            
            # Create wrapped labels for checkboxes
            for i, label in enumerate(sample_names or []):
                wrapped_label = wrap_text(label, 15)  # Wrap at 15 characters
                wrapped_labels.append(wrapped_label)
                self.label_mapping[wrapped_label] = label
            
            if not wrapped_labels:
                debug_print("DEBUG: No sample names available for checkboxes")
                return
            
            # Calculate checkbox positioning
            num_labels = len(wrapped_labels)
            spacing = 0.04
            
            if is_split_plot:
                # For split plots, use more spacing
                total_height = num_labels * spacing
                top_position = 0.95
                bottom_position = max(0.05, top_position - total_height)
                checkbox_height = top_position - bottom_position
            else:
                # Standard single plot
                total_height = num_labels * spacing
                top_position = min(0.95, 0.5 + total_height / 2)
                bottom_position = max(0.05, top_position - total_height)
                checkbox_height = top_position - bottom_position
            
            # Create checkbox axes
            checkbox_ax = self.figure.add_axes([0.83, bottom_position, 0.15, checkbox_height])
            
            # Create checkboxes
            self.check_buttons = CheckButtons(
                checkbox_ax,
                wrapped_labels,
                [True] * len(wrapped_labels)  # All initially checked
            )
            
            # Style the checkboxes
            for i, (rect, line1, line2) in enumerate(zip(
                self.check_buttons.rectangles,
                self.check_buttons.lines[0::2],
                self.check_buttons.lines[1::2]
            )):
                rect.set_facecolor('white')
                rect.set_edgecolor('black')
                rect.set_linewidth(1)
                line1.set_color('green')
                line1.set_linewidth(3)
                line2.set_color('green')
                line2.set_linewidth(3)
                line1.set_zorder(2)
                line2.set_zorder(2)
            
            # Bind callbacks based on plot type
            if is_split_bar_chart:
                debug_print("DEBUG: Using User Test Simulation split bar chart checkbox callback")
                self.checkbox_cid = self.check_buttons.on_clicked(self.on_user_test_simulation_bar_checkbox_click)
            elif is_split_plot:
                debug_print("DEBUG: Using User Test Simulation checkbox callback")
                self.checkbox_cid = self.check_buttons.on_clicked(self.on_user_test_simulation_checkbox_click)
            elif is_bar_chart:
                debug_print("DEBUG: Using bar chart checkbox callback")
                self.checkbox_cid = self.check_buttons.on_clicked(self.on_bar_checkbox_click)
            else:
                debug_print("DEBUG: Using standard line plot checkbox callback")
                self.checkbox_cid = self.check_buttons.on_clicked(self.on_checkbox_click)
            
            if self.canvas:
                self.canvas.draw_idle()
            
            debug_print(f"DEBUG: Added {len(wrapped_labels)} checkboxes")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to add checkboxes: {e}")
            traceback.print_exc()
    
    def on_checkbox_click(self, wrapped_label):
        """Toggle visibility of a line when its checkbox is clicked."""
        debug_print(f"DEBUG: Standard checkbox clicked: {wrapped_label}")
        
        try:
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None:
                debug_print(f"WARNING: Could not find original label for: {wrapped_label}")
                return
            
            if not self.gui or not hasattr(self.gui, 'line_labels'):
                debug_print("ERROR: No line labels available")
                return
            
            index = list(self.label_mapping.values()).index(original_label)
            
            if hasattr(self.gui, 'selected_plot_type') and self.gui.selected_plot_type.get() == "TPM (Bar)":
                if self.lines and index < len(self.lines):
                    bar = self.lines[index]
                    bar.set_visible(not bar.get_visible())
            else:
                if self.lines and index < len(self.lines):
                    line = self.lines[index]
                    line.set_visible(not line.get_visible())
            
            if self.canvas:
                self.canvas.draw_idle()
                
            debug_print(f"DEBUG: Toggled visibility for {original_label}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to handle checkbox click: {e}")
    
    def on_bar_checkbox_click(self, wrapped_label):
        """Toggle visibility of a bar when its checkbox is clicked."""
        debug_print(f"DEBUG: Bar chart checkbox clicked: {wrapped_label}")
        
        try:
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None:
                debug_print(f"WARNING: Could not find original label for: {wrapped_label}")
                return
            
            if not self.gui or not hasattr(self.gui, 'line_labels'):
                debug_print("ERROR: No line labels available")
                return
            
            index = self.gui.line_labels.index(original_label)
            
            if self.axes and hasattr(self.axes, 'patches') and index < len(self.axes.patches):
                bar = self.axes.patches[index]
                bar.set_visible(not bar.get_visible())
                
                if self.canvas:
                    self.canvas.draw_idle()
                    
                debug_print(f"DEBUG: Toggled bar visibility for {original_label}")
            else:
                debug_print(f"ERROR: Could not find bar at index {index}")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to handle bar checkbox click: {e}")
    
    def on_user_test_simulation_checkbox_click(self, wrapped_label):
        """Handle checkbox clicks for User Test Simulation split plots."""
        debug_print(f"DEBUG: User Test Simulation checkbox clicked: {wrapped_label}")
        
        try:
            if not hasattr(self.figure, 'phase1_lines') or not hasattr(self.figure, 'phase2_lines'):
                debug_print("DEBUG: No phase lines found on figure")
                return
            
            # Get the original label from the wrapped label
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None:
                debug_print(f"DEBUG: Could not find original label for: {wrapped_label}")
                return
            
            # Find the index of the clicked sample
            if not self.gui or not hasattr(self.gui, 'line_labels'):
                debug_print("ERROR: No line labels available")
                return
            
            try:
                index = self.gui.line_labels.index(original_label)
                debug_print(f"DEBUG: Found sample index: {index} for label: {original_label}")
            except ValueError:
                debug_print(f"DEBUG: Could not find index for label: {original_label}")
                return
            
            # Get the corresponding lines from both plots
            if index < len(self.figure.phase1_lines) and index < len(self.figure.phase2_lines):
                phase1_line = self.figure.phase1_lines[index]
                phase2_line = self.figure.phase2_lines[index]
                
                # Toggle visibility for both lines
                is_visible = phase1_line.get_visible()
                new_visibility = not is_visible
                
                phase1_line.set_visible(new_visibility)
                phase2_line.set_visible(new_visibility)
                
                debug_print(f"DEBUG: Toggled sample '{original_label}' visibility to {new_visibility} in both plots")
                
                # Redraw the canvas
                if self.canvas:
                    self.canvas.draw_idle()
            else:
                debug_print(f"DEBUG: Index {index} out of range for phase lines")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to handle User Test Simulation checkbox click: {e}")
    
    def on_user_test_simulation_bar_checkbox_click(self, wrapped_label):
        """Handle checkbox clicks for User Test Simulation split bar charts."""
        debug_print(f"DEBUG: User Test Simulation bar chart checkbox clicked: {wrapped_label}")
        
        try:
            if not hasattr(self.figure, 'phase1_bars') or not hasattr(self.figure, 'phase2_bars'):
                debug_print("DEBUG: No phase bars found on figure")
                return
            
            # Get the original label from the wrapped label
            original_label = self.label_mapping.get(wrapped_label)
            if original_label is None:
                debug_print(f"DEBUG: Could not find original label for: {wrapped_label}")
                return
            
            # Find the index of the clicked sample
            if not self.gui or not hasattr(self.gui, 'line_labels'):
                debug_print("ERROR: No line labels available")
                return
            
            try:
                index = self.gui.line_labels.index(original_label)
                debug_print(f"DEBUG: Found sample index: {index} for label: {original_label}")
            except ValueError:
                debug_print(f"DEBUG: Could not find index for label: {original_label}")
                return
            
            # Get the corresponding bars from both plots
            if index < len(self.figure.phase1_bars) and index < len(self.figure.phase2_bars):
                phase1_bar = self.figure.phase1_bars[index]
                phase2_bar = self.figure.phase2_bars[index]
                
                # Toggle visibility for both bars
                is_visible = phase1_bar.get_visible()
                new_visibility = not is_visible
                
                phase1_bar.set_visible(new_visibility)
                phase2_bar.set_visible(new_visibility)
                
                debug_print(f"DEBUG: Toggled sample '{original_label}' bar visibility to {new_visibility} in both plots")
                
                # Redraw the canvas
                if self.canvas:
                    self.canvas.draw_idle()
            else:
                debug_print(f"DEBUG: Index {index} out of range for phase bars")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to handle User Test Simulation bar checkbox click: {e}")
    
    def zoom(self, event):
        """Zoom in/out on the plot with the cursor as the zoom origin."""
        debug_print(f"DEBUG: Zoom event triggered: {event.button}")
        
        try:
            if not self.axes or event.inaxes != self.axes:
                return
            
            x_min, x_max = self.axes.get_xlim()
            y_min, y_max = self.axes.get_ylim()
            cursor_x, cursor_y = event.xdata, event.ydata
            zoom_scale = 0.9 if event.button == 'up' else 1.1
            
            if cursor_x is None or cursor_y is None:
                return
            
            new_x_min = cursor_x - (cursor_x - x_min) * zoom_scale
            new_x_max = cursor_x + (x_max - cursor_x) * zoom_scale
            new_y_min = cursor_y - (cursor_y - y_min) * zoom_scale
            new_y_max = cursor_y + (y_max - cursor_y) * zoom_scale
            
            self.axes.set_xlim(new_x_min, new_x_max)
            self.axes.set_ylim(new_y_min, new_y_max)
            
            if self.canvas:
                self.canvas.draw_idle()
                
            debug_print("DEBUG: Zoom applied successfully")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to apply zoom: {e}")
    
    def add_plot_dropdown(self, frame: ttk.Frame) -> None:
        """Create or update the plot type dropdown, centered with no background color."""
        debug_print("DEBUG: PlotController adding plot dropdown")
        
        try:
            # Use existing dropdown frame or create a new one
            if not hasattr(self, 'dropdown_frame') or not self.dropdown_frame.winfo_exists():
                self.dropdown_frame = ttk.Frame(frame)
                self.dropdown_frame.pack(side='bottom', fill='x', pady=5)
            
            # Clear existing widgets
            for widget in self.dropdown_frame.winfo_children():
                widget.destroy()
            
            # Create a container frame to center the dropdown components
            center_frame = ttk.Frame(self.dropdown_frame)
            center_frame.pack(side="top", fill="x")
            
            # Use pack with expand=True to center the content
            label_frame = ttk.Frame(center_frame)
            label_frame.pack(side="top", fill="none", expand=True)
            
            # Create the dropdown label
            dropdown_label = ttk.Label(
                label_frame,
                text="Plot Type:",
                font=('Arial', 10), 
                foreground='white'
            )
            dropdown_label.pack(side="left", padx=(0, 5))
            
            # Get plot options
            plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
            if self.gui and hasattr(self.gui, 'plot_options'):
                plot_options = self.gui.plot_options
            
            # Get selected plot type variable
            selected_plot_type = None
            if self.gui and hasattr(self.gui, 'selected_plot_type'):
                selected_plot_type = self.gui.selected_plot_type
            else:
                selected_plot_type = StringVar(value="TPM")
            
            # Create the dropdown with no background styling
            self.plot_dropdown = ttk.Combobox(
                label_frame,
                textvariable=selected_plot_type,
                values=plot_options,
                state="readonly",
                font=('Arial', 10),
                width=15
            )
            self.plot_dropdown.pack(side="left")
            
            # Remove ALL background styling
            style = ttk.Style()
            style.map('TCombobox', fieldbackground=[('readonly', 'white')])
            style.map('TCombobox', selectbackground=[('readonly', 'white')])
            style.configure('TCombobox', background='white')
            
            # Additional styling to ensure no blue background
            self.plot_dropdown.configure(background='white')
            
            # Bind event handler
            self.plot_dropdown.bind("<<ComboboxSelected>>", self.update_plot_from_dropdown)
            
            # Set default value if needed
            if selected_plot_type.get() not in plot_options:
                selected_plot_type.set(plot_options[0])
            
            debug_print(f"DEBUG: Plot dropdown created with {len(plot_options)} options")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to add plot dropdown: {e}")
            traceback.print_exc()
    
    def update_plot_from_dropdown(self, event) -> None:
        """Callback for when the plot type dropdown changes."""
        debug_print("DEBUG: PlotController updating plot from dropdown")
        
        try:
            selected_plot = ""
            if self.gui and hasattr(self.gui, 'selected_plot_type'):
                selected_plot = self.gui.selected_plot_type.get()
            
            if not selected_plot:
                messagebox.showerror("Error", "No plot type selected.")
                return
            
            debug_print(f"DEBUG: Selected plot type: {selected_plot}")
            
            # Get current sheet and data
            if not self.gui:
                debug_print("ERROR: No GUI reference available")
                return
            
            current_sheet_name = self.gui.selected_sheet.get()
            if current_sheet_name not in self.gui.filtered_sheets:
                messagebox.showerror("Error", f"Sheet '{current_sheet_name}' not found.")
                return
            
            sheet_info = self.gui.filtered_sheets[current_sheet_name]
            sheet_data = sheet_info["data"]
            
            if sheet_data.empty:
                messagebox.showwarning("Warning", "The selected sheet has no data.")
                return
            
            # Process data for the selected plot option
            process_function = processing.get_processing_function(current_sheet_name)
            processed_data, _, full_sample_data = process_function(sheet_data)
            
            # Use the correct number of columns per sample
            num_columns = getattr(self.gui, 'num_columns_per_sample', 12)
            debug_print(f"DEBUG: update_plot_from_dropdown using {num_columns} columns per sample")
            
            # Update the plot
            if hasattr(self.gui, 'plot_frame'):
                self.plot_all_samples(self.gui.plot_frame, full_sample_data, num_columns)
            
            # Refresh the dropdown in case it was recreated
            if not self.plot_dropdown or not self.plot_dropdown.winfo_exists():
                if hasattr(self.gui, 'plot_frame'):
                    self.add_plot_dropdown(self.gui.plot_frame)
            else:
                plot_options = getattr(self.gui, 'plot_options', ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"])
                self.plot_dropdown["values"] = plot_options
                self.plot_dropdown.set(selected_plot)
            
            debug_print("DEBUG: Plot updated from dropdown successfully")
            
        except Exception as e:
            debug_print(f"ERROR: Error in update_plot_from_dropdown: {e}")
            messagebox.showerror("Error", f"An error occurred while updating the plot: {e}")
    
    def clear_plot_area(self):
        """Clear the plot area and release Matplotlib resources."""
        debug_print("DEBUG: PlotController clearing plot area")
        
        try:
            plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()
            
            if not plt:
                debug_print("ERROR: Could not load matplotlib")
                return None
            
            # Clear plot frame if available
            if self.gui and hasattr(self.gui, 'plot_frame') and self.gui.plot_frame.winfo_exists():
                for widget in self.gui.plot_frame.winfo_children():
                    widget.destroy()
            
            # Close figure
            if self.figure:
                plt.close(self.figure)
                self.figure = None
            
            # Destroy canvas
            if self.canvas:
                self.canvas.get_tk_widget().destroy()
                self.canvas = None
            
            # Reset references
            self.axes = None
            self.lines = None
            self.check_buttons = None
            self.checkbox_cid = None
            
            # Clear GUI state
            if self.gui:
                if hasattr(self.gui, 'line_labels'):
                    self.gui.line_labels = []
                if hasattr(self.gui, 'original_lines_data'):
                    self.gui.original_lines_data = []
            
            debug_print("DEBUG: Plot area cleared successfully")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to clear plot area: {e}")
    
    def restore_all_elements(self):
        """Restore all lines or bars to their original state."""
        debug_print("DEBUG: PlotController restoring all plot elements")
        
        try:
            if self.check_buttons:
                self.check_buttons.disconnect(self.checkbox_cid)
            
            # Determine if this is a bar chart
            is_bar_chart = False
            if self.gui and hasattr(self.gui, 'selected_plot_type'):
                is_bar_chart = self.gui.selected_plot_type.get() == "TPM (Bar)"
            
            if is_bar_chart:
                # Restore bars
                if self.axes and hasattr(self.axes, 'patches') and self.gui and hasattr(self.gui, 'original_lines_data'):
                    for patch, original_data in zip(self.axes.patches, self.gui.original_lines_data):
                        patch.set_x(original_data[0])
                        patch.set_height(original_data[1])
                        patch.set_visible(True)
            else:
                # Restore lines
                if self.lines and self.gui and hasattr(self.gui, 'original_lines_data'):
                    for line, original_data in zip(self.lines, self.gui.original_lines_data):
                        line.set_xdata(original_data[0])
                        line.set_ydata(original_data[1])
                        line.set_visible(True)
            
            # Reset all checkboxes to the default "checked" state
            if self.check_buttons and self.gui and hasattr(self.gui, 'line_labels'):
                for i in range(len(self.gui.line_labels)):
                    if not self.check_buttons.get_status()[i]:  # Only activate if not already active
                        self.check_buttons.set_active(i)  # Reset each checkbox to "checked"
            
            # Reconnect the callback for CheckButtons after restoring
            if self.check_buttons:
                is_bar_chart = False
                if self.gui and hasattr(self.gui, 'selected_plot_type'):
                    is_bar_chart = self.gui.selected_plot_type.get() == "TPM (Bar)"
                
                self.checkbox_cid = self.check_buttons.on_clicked(
                    self.on_bar_checkbox_click if is_bar_chart else self.on_checkbox_click
                )
            
            # Redraw the canvas to reflect the restored state
            if self.canvas:
                self.canvas.draw_idle()
            
            debug_print("DEBUG: All plot elements restored successfully")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to restore plot elements: {e}")
    
    def update_plot_options(self, plot_options: List[str]) -> None:
        """Update available plot options and refresh dropdown."""
        debug_print(f"DEBUG: PlotController updating plot options: {plot_options}")
        
        try:
            # Update GUI plot options
            if self.gui and hasattr(self.gui, 'plot_options'):
                self.gui.plot_options = plot_options
            
            # Update dropdown if it exists
            if self.plot_dropdown and self.plot_dropdown.winfo_exists():
                self.plot_dropdown['values'] = plot_options
                
                # Ensure current selection is valid
                current_selection = ""
                if self.gui and hasattr(self.gui, 'selected_plot_type'):
                    current_selection = self.gui.selected_plot_type.get()
                
                if current_selection not in plot_options and plot_options:
                    if self.gui:
                        self.gui.selected_plot_type.set(plot_options[0])
            
            debug_print(f"DEBUG: Plot options updated to: {plot_options}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to update plot options: {e}")
    
    def on_data_changed(self):
        """Handle notification that data has changed."""
        debug_print("DEBUG: PlotController received data change notification")
        
        try:
            # Refresh plot if auto-refresh is enabled and we have an active plot
            if (self.plot_model.settings.auto_refresh if hasattr(self.plot_model, 'settings') else True):
                if self.gui and hasattr(self.gui, 'plot_frame') and hasattr(self.gui, 'filtered_sheets'):
                    current_sheet = self.gui.selected_sheet.get()
                    
                    if current_sheet in self.gui.filtered_sheets:
                        sheet_info = self.gui.filtered_sheets[current_sheet]
                        sheet_data = sheet_info["data"]
                        
                        if not sheet_data.empty:
                            # Process the data
                            process_function = processing.get_processing_function(current_sheet)
                            processed_data, _, full_sample_data = process_function(sheet_data)
                            
                            # Refresh the plot
                            num_columns = getattr(self.gui, 'num_columns_per_sample', 12)
                            self.plot_all_samples(self.gui.plot_frame, full_sample_data, num_columns)
                            
                            debug_print("DEBUG: Plot refreshed due to data change")
        
        except Exception as e:
            debug_print(f"ERROR: Failed to handle data change: {e}")
    
    def toggle_data_visibility(self, data_id: str):
        """Toggle visibility of plot data."""
        debug_print(f"DEBUG: PlotController toggling visibility for {data_id}")
        
        try:
            self.plot_model.toggle_data_visibility(data_id)
            
            # Refresh plot display
            if self.canvas:
                self.canvas.draw_idle()
            
        except Exception as e:
            debug_print(f"ERROR: Failed to toggle data visibility: {e}")
    
    def get_valid_plot_options(self, full_sample_data) -> List[str]:
        """Get valid plot options for the current data."""
        debug_print("DEBUG: PlotController getting valid plot options")
        
        try:
            pd = self.get_pandas()
            if not pd:
                debug_print("ERROR: Pandas not available")
                return []
            
            plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
            if self.gui and hasattr(self.gui, 'plot_options'):
                plot_options = self.gui.plot_options
            
            valid_options = processing.get_valid_plot_options(plot_options, full_sample_data)
            
            debug_print(f"DEBUG: Valid plot options: {valid_options}")
            return valid_options
            
        except Exception as e:
            debug_print(f"ERROR: Failed to get valid plot options: {e}")
            return []
    
    # Getter methods for external access
    def get_current_plot_type(self) -> str:
        """Get the currently selected plot type."""
        if self.gui and hasattr(self.gui, 'selected_plot_type'):
            return self.gui.selected_plot_type.get()
        return self.plot_model.current_plot_type if hasattr(self.plot_model, 'current_plot_type') else "TPM"
    
    def get_plot_canvas(self):
        """Get the current plot canvas."""
        return self.canvas
    
    def get_plot_figure(self):
        """Get the current plot figure."""
        return self.figure
    
    def is_plot_available(self) -> bool:
        """Check if a plot is currently available."""
        return self.figure is not None and self.canvas is not None
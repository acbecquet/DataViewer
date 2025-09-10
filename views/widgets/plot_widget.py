# views/widgets/plot_widget.py
"""
views/widgets/plot_widget.py
Plot display widget.
This will contain the plot display logic from main_gui.py and plot_manager.py.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, Any, List, Callable
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import pandas as pd


class PlotWidget:
    """Widget for displaying plots and plot controls."""
    
    def __init__(self, parent: tk.Widget):
        """Initialize the plot widget."""
        self.parent = parent
        self.frame = ttk.Frame(parent)
        
        # Matplotlib components
        self.figure: Optional[Figure] = None
        self.canvas: Optional[FigureCanvasTkAgg] = None
        self.toolbar: Optional[NavigationToolbar2Tk] = None
        self.axes: Optional[Any] = None
        
        # Plot data and controls
        self.line_objects: Dict[str, Any] = {}
        self.checkboxes: Dict[str, tk.BooleanVar] = {}
        self.checkbox_frame: Optional[ttk.Frame] = None
        
        # Callbacks
        self.on_line_toggled: Optional[Callable] = None
        
        print("DEBUG: PlotWidget initialized")
    
    def setup_widget(self):
        """Set up the plot widget layout."""
        print("DEBUG: PlotWidget setting up layout")
        
        # Create main layout
        self._create_plot_area()
        self._create_checkbox_area()
        
        print("DEBUG: PlotWidget layout complete")
    
    def _create_plot_area(self):
        """Create the matplotlib plot area."""
        # Create figure
        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.axes = self.figure.add_subplot(111)
        
        # Create canvas
        self.canvas = FigureCanvasTkAgg(self.figure, self.frame)
        self.canvas.get_tk_widget().pack(side="left", fill="both", expand=True)
        
        # Create toolbar
        toolbar_frame = ttk.Frame(self.frame)
        toolbar_frame.pack(side="bottom", fill="x")
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()
        
        print("DEBUG: PlotWidget plot area created")
    
    def _create_checkbox_area(self):
        """Create the checkbox control area."""
        # Create checkbox frame on the right side
        self.checkbox_frame = ttk.LabelFrame(self.frame, text="Plot Lines", padding=5)
        self.checkbox_frame.pack(side="right", fill="y", padx=(5, 0))
        
        # Add scrollable area for many checkboxes
        canvas = tk.Canvas(self.checkbox_frame, width=150)
        scrollbar = ttk.Scrollbar(self.checkbox_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Store reference to scrollable frame for adding checkboxes
        self.scrollable_checkbox_frame = scrollable_frame
        
        print("DEBUG: PlotWidget checkbox area created")
    
    def plot_data(self, data: pd.DataFrame, plot_type: str, title: str = ""):
        """Plot data on the widget."""
        print(f"DEBUG: PlotWidget plotting {plot_type} data")
        
        if not self.axes or data.empty:
            print("WARNING: PlotWidget - no axes or empty data")
            return
        
        try:
            # Clear previous plot
            self.axes.clear()
            self.line_objects.clear()
            
            # Plot based on data structure
            # This is a simplified version - real implementation would be more sophisticated
            if len(data.columns) >= 2:
                x_data = data.iloc[:, 0]
                
                # Plot each numeric column
                for i, col in enumerate(data.columns[1:]):
                    if pd.api.types.is_numeric_dtype(data[col]):
                        y_data = data[col]
                        line, = self.axes.plot(x_data, y_data, marker='o', label=str(col))
                        self.line_objects[str(col)] = line
            
            # Set labels and title
            self.axes.set_title(title or f"{plot_type} Plot")
            self.axes.set_xlabel("X Axis")
            self.axes.set_ylabel(plot_type)
            self.axes.grid(True, alpha=0.3)
            self.axes.legend()
            
            # Update checkboxes
            self._update_checkboxes()
            
            # Refresh display
            self.canvas.draw()
            
            print(f"DEBUG: PlotWidget plotted {len(self.line_objects)} lines")
            
        except Exception as e:
            print(f"ERROR: PlotWidget failed to plot data: {e}")
    
    def _update_checkboxes(self):
        """Update checkboxes for plot lines."""
        if not self.scrollable_checkbox_frame:
            return
        
        # Clear existing checkboxes
        for widget in self.scrollable_checkbox_frame.winfo_children():
            widget.destroy()
        self.checkboxes.clear()
        
        # Create checkboxes for each line
        for line_id, line_obj in self.line_objects.items():
            var = tk.BooleanVar(value=True)  # Start with all lines visible
            self.checkboxes[line_id] = var
            
            checkbox = ttk.Checkbutton(
                self.scrollable_checkbox_frame,
                text=line_id,
                variable=var,
                command=lambda lid=line_id: self._on_checkbox_toggled(lid)
            )
            checkbox.pack(anchor="w", pady=2)
        
        print(f"DEBUG: PlotWidget updated {len(self.checkboxes)} checkboxes")
    
    def _on_checkbox_toggled(self, line_id: str):
        """Handle checkbox toggle."""
        if line_id not in self.checkboxes or line_id not in self.line_objects:
            return
        
        is_visible = self.checkboxes[line_id].get()
        line_obj = self.line_objects[line_id]
        
        # Toggle line visibility
        line_obj.set_visible(is_visible)
        
        # Update legend
        self.axes.legend()
        
        # Refresh display
        self.canvas.draw()
        
        print(f"DEBUG: PlotWidget toggled {line_id} visibility: {is_visible}")
        
        # Notify controller if callback is set
        if self.on_line_toggled:
            self.on_line_toggled(line_id, is_visible)
    
    def clear_plot(self):
        """Clear the plot."""
        if self.axes:
            self.axes.clear()
            self.line_objects.clear()
            
            # Clear checkboxes
            if self.scrollable_checkbox_frame:
                for widget in self.scrollable_checkbox_frame.winfo_children():
                    widget.destroy()
            self.checkboxes.clear()
            
            # Refresh display
            if self.canvas:
                self.canvas.draw()
            
            print("DEBUG: PlotWidget cleared")
    
    def export_plot(self, file_path: str, dpi: int = 300):
        """Export the plot to file."""
        if not self.figure:
            return False
        
        try:
            self.figure.savefig(file_path, dpi=dpi, bbox_inches='tight')
            print(f"DEBUG: PlotWidget exported plot to {file_path}")
            return True
        except Exception as e:
            print(f"ERROR: PlotWidget failed to export plot: {e}")
            return False
    
    def set_line_toggle_callback(self, callback: Callable):
        """Set callback for line visibility toggle."""
        self.on_line_toggled = callback
        print("DEBUG: PlotWidget line toggle callback set")
    
    def get_widget(self) -> ttk.Frame:
        """Get the main widget frame."""
        return self.frame
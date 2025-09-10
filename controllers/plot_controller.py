# controllers/plot_controller.py
"""
controllers/plot_controller.py
Plot generation controller that coordinates between plot service and models.
This replaces the coordination logic currently in plot_manager.py.
"""

from typing import Optional, Dict, Any, List
from models.plot_model import PlotModel
from models.data_model import DataModel


class PlotController:
    """Controller for plot generation and display."""
    
    def __init__(self, plot_model: PlotModel, data_model: DataModel, plot_service: Any):
        """Initialize the plot controller."""
        self.plot_model = plot_model
        self.data_model = data_model
        self.plot_service = plot_service
        
        # Cross-controller references
        self.data_controller: Optional['DataController'] = None
        
        print("DEBUG: PlotController initialized")
        print(f"DEBUG: Connected to PlotModel and DataModel")
    
    def set_data_controller(self, data_controller: 'DataController'):
        """Set reference to data controller."""
        self.data_controller = data_controller
        print("DEBUG: PlotController connected to DataController")
    
    def generate_plot(self, plot_type: str) -> bool:
        """Generate a plot of the specified type."""
        print(f"DEBUG: PlotController generating {plot_type} plot")
        
        try:
            # Set plot type in model
            self.plot_model.set_plot_type(plot_type)
            
            # Get data from data model
            current_sheet = self.data_model.get_selected_sheet()
            if not current_sheet:
                print("WARNING: No sheet selected for plotting")
                return False
            
            # Generate plot through service (placeholder)
            # plot_data = self.plot_service.create_plot(current_sheet.data, plot_type)
            
            print(f"DEBUG: PlotController successfully generated {plot_type} plot")
            return True
            
        except Exception as e:
            print(f"ERROR: PlotController failed to generate plot: {e}")
            return False
    
    def on_data_changed(self):
        """Handle notification that data has changed."""
        print("DEBUG: PlotController received data change notification")
        
        # Refresh plot if auto-refresh is enabled
        if self.plot_model.settings.auto_refresh:
            current_plot_type = self.plot_model.get_plot_type()
            self.generate_plot(current_plot_type)
    
    def toggle_data_visibility(self, data_id: str):
        """Toggle visibility of plot data."""
        self.plot_model.toggle_data_visibility(data_id)
        print(f"DEBUG: PlotController toggled visibility for {data_id}")
# services/file_service.py
"""
services/file_service.py
File I/O operations service.
This will contain the core file operations from file_manager.py.
"""

from typing import Optional, Dict, Any, List, Tuple
import os
import pandas as pd
from pathlib import Path


class FileService:
    """Service for file I/O operations."""
    
    def __init__(self):
        """Initialize the file service."""
        self.supported_formats = ['.xlsx', '.xls', '.csv', '.vap3']
        print("DEBUG: FileService initialized")
        print(f"DEBUG: Supported formats: {', '.join(self.supported_formats)}")
    
    def load_excel_file(self, file_path: str) -> Tuple[bool, Dict[str, pd.DataFrame], str]:
        """Load an Excel file and return sheets data."""
        print(f"DEBUG: FileService loading Excel file: {file_path}")
        
        try:
            if not Path(file_path).exists():
                return False, {}, f"File not found: {file_path}"
            
            # Load Excel file
            # In real implementation, would use pandas.read_excel with proper error handling
            # sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            
            # Placeholder implementation
            sheets = {"Sheet1": pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})}
            
            print(f"DEBUG: FileService loaded {len(sheets)} sheets from {file_path}")
            return True, sheets, "Success"
            
        except Exception as e:
            error_msg = f"Failed to load Excel file: {e}"
            print(f"ERROR: FileService - {error_msg}")
            return False, {}, error_msg
    
    def save_excel_file(self, file_path: str, sheets_data: Dict[str, pd.DataFrame]) -> Tuple[bool, str]:
        """Save data to an Excel file."""
        print(f"DEBUG: FileService saving Excel file: {file_path}")
        
        try:
            # Create directory if it doesn't exist
            Path(file_path).parent.mkdir(parents=True, exist_ok=True)
            
            # Save Excel file
            # In real implementation, would use pandas ExcelWriter
            # with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            #     for sheet_name, data in sheets_data.items():
            #         data.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"DEBUG: FileService saved {len(sheets_data)} sheets to {file_path}")
            return True, "Success"
            
        except Exception as e:
            error_msg = f"Failed to save Excel file: {e}"
            print(f"ERROR: FileService - {error_msg}")
            return False, error_msg
    
    def load_vap3_file(self, file_path: str) -> Tuple[bool, Dict[str, Any], str]:
        """Load a VAP3 file."""
        print(f"DEBUG: FileService loading VAP3 file: {file_path}")
        
        try:
            # Placeholder implementation
            # In real implementation, would load VAP3 format
            vap3_data = {
                'filtered_sheets': {},
                'plot_options': [],
                'metadata': {}
            }
            
            print(f"DEBUG: FileService loaded VAP3 file: {file_path}")
            return True, vap3_data, "Success"
            
        except Exception as e:
            error_msg = f"Failed to load VAP3 file: {e}"
            print(f"ERROR: FileService - {error_msg}")
            return False, {}, error_msg
    
    def save_vap3_file(self, file_path: str, data: Dict[str, Any]) -> Tuple[bool, str]:
        """Save data as a VAP3 file."""
        print(f"DEBUG: FileService saving VAP3 file: {file_path}")
        
        try:
            # Placeholder implementation
            # In real implementation, would save in VAP3 format
            
            print(f"DEBUG: FileService saved VAP3 file: {file_path}")
            return True, "Success"
            
        except Exception as e:
            error_msg = f"Failed to save VAP3 file: {e}"
            print(f"ERROR: FileService - {error_msg}")
            return False, error_msg
    
    def validate_file_path(self, file_path: str) -> Tuple[bool, str]:
        """Validate if a file path is accessible and supported."""
        if not Path(file_path).exists():
            return False, f"File does not exist: {file_path}"
        
        file_ext = Path(file_path).suffix.lower()
        if file_ext not in self.supported_formats:
            return False, f"Unsupported file format: {file_ext}"
        
        return True, "Valid file"
    
    def get_file_info(self, file_path: str) -> Dict[str, Any]:
        """Get information about a file."""
        try:
            path = Path(file_path)
            stat = path.stat()
            
            return {
                'filename': path.name,
                'size_bytes': stat.st_size,
                'modified_time': stat.st_mtime,
                'extension': path.suffix.lower(),
                'exists': path.exists()
            }
        except Exception as e:
            print(f"ERROR: FileService failed to get file info: {e}")
            return {}


# services/plot_service.py
"""
services/plot_service.py
Plot generation service.
This will contain the core plotting logic from plot_manager.py.
"""

import matplotlib.pyplot as plt
import matplotlib.figure as mpl_figure
from typing import Optional, Dict, Any, List, Tuple
import pandas as pd
import numpy as np


class PlotService:
    """Service for plot generation and management."""
    
    def __init__(self):
        """Initialize the plot service."""
        self.plot_types = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        self.current_figure: Optional[mpl_figure.Figure] = None
        self.current_axes: Optional[Any] = None
        
        print("DEBUG: PlotService initialized")
        print(f"DEBUG: Available plot types: {', '.join(self.plot_types)}")
    
    def create_plot(self, data: pd.DataFrame, plot_type: str, 
                   title: str = "", xlabel: str = "", ylabel: str = "") -> Tuple[bool, Optional[mpl_figure.Figure], str]:
        """Create a plot from data."""
        print(f"DEBUG: PlotService creating {plot_type} plot")
        
        try:
            if data.empty:
                return False, None, "No data provided for plotting"
            
            if plot_type not in self.plot_types:
                return False, None, f"Unsupported plot type: {plot_type}"
            
            # Create figure and axes
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # Plot data based on type
            success = self._plot_data_by_type(ax, data, plot_type)
            if not success:
                plt.close(fig)
                return False, None, f"Failed to plot {plot_type} data"
            
            # Set labels and title
            ax.set_title(title or f"{plot_type} Plot")
            ax.set_xlabel(xlabel or "X Axis")
            ax.set_ylabel(ylabel or plot_type)
            ax.grid(True, alpha=0.3)
            ax.legend()
            
            # Store references
            self.current_figure = fig
            self.current_axes = ax
            
            print(f"DEBUG: PlotService successfully created {plot_type} plot")
            return True, fig, "Success"
            
        except Exception as e:
            error_msg = f"Failed to create plot: {e}"
            print(f"ERROR: PlotService - {error_msg}")
            return False, None, error_msg
    
    def _plot_data_by_type(self, ax: Any, data: pd.DataFrame, plot_type: str) -> bool:
        """Plot data based on the specified type."""
        try:
            # Placeholder plotting logic
            # In real implementation, would extract appropriate columns based on plot_type
            
            if len(data.columns) < 2:
                print("WARNING: Insufficient columns for plotting")
                return False
            
            # Simple plotting - would be more sophisticated in real implementation
            x_data = data.iloc[:, 0] if len(data) > 0 else []
            y_data = data.iloc[:, 1] if len(data) > 0 else []
            
            ax.plot(x_data, y_data, marker='o', linewidth=2, label=plot_type)
            
            print(f"DEBUG: PlotService plotted {len(x_data)} points for {plot_type}")
            return True
            
        except Exception as e:
            print(f"ERROR: PlotService failed to plot {plot_type}: {e}")
            return False
    
    def update_plot_visibility(self, line_id: str, visible: bool) -> bool:
        """Update visibility of a plot line."""
        print(f"DEBUG: PlotService updating visibility for {line_id}: {visible}")
        
        try:
            # Placeholder - would interact with matplotlib line objects
            print(f"DEBUG: PlotService updated visibility for {line_id}")
            return True
            
        except Exception as e:
            print(f"ERROR: PlotService failed to update visibility: {e}")
            return False
    
    def export_plot(self, file_path: str, format: str = "png", dpi: int = 300) -> Tuple[bool, str]:
        """Export the current plot to file."""
        print(f"DEBUG: PlotService exporting plot to {file_path}")
        
        try:
            if not self.current_figure:
                return False, "No plot to export"
            
            self.current_figure.savefig(file_path, format=format, dpi=dpi, bbox_inches='tight')
            
            print(f"DEBUG: PlotService exported plot to {file_path}")
            return True, "Success"
            
        except Exception as e:
            error_msg = f"Failed to export plot: {e}"
            print(f"ERROR: PlotService - {error_msg}")
            return False, error_msg
    
    def clear_plot(self):
        """Clear the current plot."""
        if self.current_axes:
            self.current_axes.clear()
        print("DEBUG: PlotService cleared current plot")
    
    def get_plot_data_ranges(self) -> Dict[str, Tuple[float, float]]:
        """Get the data ranges of the current plot."""
        if not self.current_axes:
            return {}
        
        try:
            xlim = self.current_axes.get_xlim()
            ylim = self.current_axes.get_ylim()
            
            return {
                'x_range': xlim,
                'y_range': ylim
            }
        except Exception as e:
            print(f"ERROR: PlotService failed to get data ranges: {e}")
            return {}
# controllers/data_controller.py
"""
controllers/data_controller.py
Data processing controller that coordinates data operations.
This replaces the coordination logic currently in processing.py integration.
"""

from typing import Optional, Dict, Any, List
from models.data_model import DataModel, SheetData


class DataController:
    """Controller for data processing operations."""
    
    def __init__(self, data_model: DataModel, processing_service: Any):
        """Initialize the data controller."""
        self.data_model = data_model
        self.processing_service = processing_service
        
        print("DEBUG: DataController initialized")
        print(f"DEBUG: Connected to DataModel and ProcessingService")
    
    def process_sheet_data(self, sheet_name: str, raw_data: Any) -> bool:
        """Process raw sheet data through processing service."""
        print(f"DEBUG: DataController processing sheet data: {sheet_name}")
        
        try:
            # Process data through service (placeholder)
            # processed_data, metadata, full_sample_data = self.processing_service.process_sheet(raw_data, sheet_name)
            
            # For now, create placeholder processed data
            # Create SheetData object
            sheet_data = SheetData(
                name=sheet_name,
                data=raw_data,  # Placeholder - would be processed data
                is_plotting_sheet=self._is_plotting_sheet(sheet_name)
            )
            
            # Add to data model
            current_file = self.data_model.get_current_file()
            if current_file:
                current_file.add_filtered_sheet(sheet_name, sheet_data)
            
            print(f"DEBUG: DataController successfully processed {sheet_name}")
            return True
            
        except Exception as e:
            print(f"ERROR: DataController failed to process {sheet_name}: {e}")
            return False
    
    def _is_plotting_sheet(self, sheet_name: str) -> bool:
        """Determine if a sheet should have plotting capabilities."""
        # Placeholder logic - would use processing.py logic
        plotting_keywords = ["test", "data", "measurement"]
        return any(keyword in sheet_name.lower() for keyword in plotting_keywords)
    
    def get_valid_plot_options(self, sheet_name: str) -> List[str]:
        """Get valid plot options for a sheet."""
        # Placeholder - would delegate to processing service
        return ["TPM", "Draw Pressure", "Resistance", "Power Efficiency"]
    
    def validate_sheet_data(self, sheet_data: SheetData) -> List[str]:
        """Validate sheet data and return list of issues."""
        issues = []
        
        if sheet_data.is_empty():
            issues.append(f"Sheet '{sheet_data.name}' is empty")
        
        if sheet_data.get_row_count() < 2:
            issues.append(f"Sheet '{sheet_data.name}' has insufficient data")
        
        print(f"DEBUG: DataController validated {sheet_data.name} - {len(issues)} issues found")
        return issues

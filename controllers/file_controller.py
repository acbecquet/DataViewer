# controllers/file_controller.py
"""
controllers/file_controller.py
File operations controller that coordinates between file service and models.
This replaces the coordination logic currently in file_manager.py.
"""

from typing import Optional, Dict, Any, List
from models.file_model import FileModel, FileState
from models.data_model import DataModel


class FileController:
    """Controller for file operations and data loading."""
    
    def __init__(self, file_model: FileModel, data_model: DataModel, 
                 file_service: Any, database_service: Any):
        """Initialize the file controller."""
        self.file_model = file_model
        self.data_model = data_model
        self.file_service = file_service
        self.database_service = database_service
        
        # Cross-controller references (set later)
        self.plot_controller: Optional['PlotController'] = None
        
        print("DEBUG: FileController initialized")
        print(f"DEBUG: Connected to FileModel and DataModel")
    
    def set_plot_controller(self, plot_controller: 'PlotController'):
        """Set reference to plot controller for notifications."""
        self.plot_controller = plot_controller
        print("DEBUG: FileController connected to PlotController")
    
    def load_file(self, file_path: str) -> bool:
        """Load a file through the service layer."""
        print(f"DEBUG: FileController loading file: {file_path}")
        
        try:
            # Check if file should be reloaded
            if not self.file_model.should_reload_file(file_path):
                print(f"DEBUG: File {file_path} is up to date")
                return True
            
            # Create or update file state
            filename = file_path.split('/')[-1]  # Simple filename extraction
            file_state = self.file_model.add_file_state(file_path, filename)
            file_state.set_loading()
            
            # Load through service (placeholder for now)
            # result = self.file_service.load_file(file_path)
            
            # For now, just mark as loaded
            file_state.set_loaded(1.0)  # 1 second placeholder
            
            # Notify plot controller of data change
            if self.plot_controller:
                self.plot_controller.on_data_changed()
            
            print(f"DEBUG: FileController successfully loaded {filename}")
            return True
            
        except Exception as e:
            print(f"ERROR: FileController failed to load {file_path}: {e}")
            return False
    
    def save_file(self, file_path: str, data: Any) -> bool:
        """Save file through the service layer."""
        print(f"DEBUG: FileController saving file: {file_path}")
        
        try:
            # Save through service (placeholder)
            # result = self.file_service.save_file(file_path, data)
            
            print(f"DEBUG: FileController successfully saved {file_path}")
            return True
            
        except Exception as e:
            print(f"ERROR: FileController failed to save {file_path}: {e}")
            return False
    
    def get_file_list(self) -> List[str]:
        """Get list of available files."""
        # Placeholder implementation
        return list(self.file_model.file_states.keys())
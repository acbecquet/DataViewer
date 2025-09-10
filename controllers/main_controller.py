"""
controllers/main_controller.py
Main application controller that coordinates all other controllers and services.
This replaces the coordination logic currently in main_gui.py.
"""

from typing import Optional, Dict, Any, List
import threading
import queue
from datetime import datetime

# Model imports
from models.data_model import DataModel
from models.file_model import FileModel
from models.plot_model import PlotModel

# Service imports will be added as we create them
# from services.file_service import FileService
# from services.plot_service import PlotService
# from services.report_service import ReportService
# from services.processing_service import ProcessingService
# from services.database_service import DatabaseService


class MainController:
    """Main application controller that coordinates between all components."""
    
    def __init__(self):
        """Initialize the main controller."""
        # Initialize models
        self.data_model = DataModel()
        self.file_model = FileModel()
        self.plot_model = PlotModel()
        
        # Controller instances (will be initialized later)
        self.file_controller: Optional['FileController'] = None
        self.plot_controller: Optional['PlotController'] = None
        self.report_controller: Optional['ReportController'] = None
        self.data_controller: Optional['DataController'] = None
        self.image_controller: Optional['ImageController'] = None
        self.calculation_controller: Optional['CalculationController'] = None
        
        # Service instances (will be initialized later)
        self.file_service: Optional[Any] = None
        self.plot_service: Optional[Any] = None
        self.report_service: Optional[Any] = None
        self.processing_service: Optional[Any] = None
        self.database_service: Optional[Any] = None
        self.image_service: Optional[Any] = None
        
        # Application state
        self.is_initialized = False
        self.initialization_error: Optional[str] = None
        self.active_threads: List[threading.Thread] = []
        self.event_queue = queue.Queue()
        
        # Application settings (from current main_gui.py)
        self.threads = []  # Active threads tracking
        self.lock = threading.Lock()
        
        print("DEBUG: MainController initialized with core models")
        print(f"DEBUG: Data model: {type(self.data_model).__name__}")
        print(f"DEBUG: File model: {type(self.file_model).__name__}")
        print(f"DEBUG: Plot model: {type(self.plot_model).__name__}")
    
    def initialize_services(self):
        """Initialize all service instances."""
        try:
            print("DEBUG: Initializing services...")
            
            # Import and initialize services
            # Note: These imports will work once we create the service files
            # For now, we'll use placeholders that can be replaced
            
            # Placeholder service initialization
            # TODO: Replace with actual service imports once created
            self.file_service = self._create_placeholder_service("FileService")
            self.plot_service = self._create_placeholder_service("PlotService")
            self.report_service = self._create_placeholder_service("ReportService")
            self.processing_service = self._create_placeholder_service("ProcessingService")
            self.database_service = self._create_placeholder_service("DatabaseService")
            self.image_service = self._create_placeholder_service("ImageService")
            
            print("DEBUG: All services initialized successfully")
            return True
            
        except Exception as e:
            error_msg = f"Failed to initialize services: {e}"
            print(f"ERROR: {error_msg}")
            self.initialization_error = error_msg
            return False
    
    def _create_placeholder_service(self, service_name: str):
        """Create a placeholder service until real services are implemented."""
        class PlaceholderService:
            def __init__(self, name):
                self.name = name
                print(f"DEBUG: Created placeholder {name}")
        
        return PlaceholderService(service_name)
    
    def initialize_controllers(self):
        """Initialize all controller instances."""
        try:
            print("DEBUG: Initializing controllers...")
            
            # Import controller classes
            from .file_controller import FileController
            from .plot_controller import PlotController
            from .report_controller import ReportController
            from .data_controller import DataController
            from .image_controller import ImageController
            from .calculation_controller import CalculationController
            
            # Initialize controllers with models and services
            self.file_controller = FileController(
                file_model=self.file_model,
                data_model=self.data_model,
                file_service=self.file_service,
                database_service=self.database_service
            )
            
            self.plot_controller = PlotController(
                plot_model=self.plot_model,
                data_model=self.data_model,
                plot_service=self.plot_service
            )
            
            self.report_controller = ReportController(
                data_model=self.data_model,
                report_service=self.report_service
            )
            
            self.data_controller = DataController(
                data_model=self.data_model,
                processing_service=self.processing_service
            )
            
            self.image_controller = ImageController(
                data_model=self.data_model,
                image_service=self.image_service
            )
            
            self.calculation_controller = CalculationController(
                data_model=self.data_model
            )
            
            print("DEBUG: All controllers initialized successfully")
            print(f"DEBUG: File controller: {type(self.file_controller).__name__}")
            print(f"DEBUG: Plot controller: {type(self.plot_controller).__name__}")
            print(f"DEBUG: Report controller: {type(self.report_controller).__name__}")
            print(f"DEBUG: Data controller: {type(self.data_controller).__name__}")
            print(f"DEBUG: Image controller: {type(self.image_controller).__name__}")
            print(f"DEBUG: Calculation controller: {type(self.calculation_controller).__name__}")
            
            return True
            
        except Exception as e:
            error_msg = f"Failed to initialize controllers: {e}"
            print(f"ERROR: {error_msg}")
            self.initialization_error = error_msg
            return False
    
    def initialize_application(self):
        """Initialize the complete application stack."""
        try:
            print("DEBUG: Starting application initialization...")
            
            # Step 1: Initialize services
            if not self.initialize_services():
                return False
            
            # Step 2: Initialize controllers  
            if not self.initialize_controllers():
                return False
            
            # Step 3: Set up cross-controller communication
            self._setup_controller_communication()
            
            # Step 4: Mark as initialized
            self.is_initialized = True
            print("DEBUG: Application initialization completed successfully")
            return True
            
        except Exception as e:
            error_msg = f"Application initialization failed: {e}"
            print(f"ERROR: {error_msg}")
            self.initialization_error = error_msg
            import traceback
            traceback.print_exc()
            return False
    
    def _setup_controller_communication(self):
        """Set up communication between controllers."""
        print("DEBUG: Setting up controller communication...")
        
        # File controller needs to notify plot controller when data changes
        if self.file_controller and self.plot_controller:
            self.file_controller.set_plot_controller(self.plot_controller)
            print("DEBUG: Connected file controller to plot controller")
        
        # Plot controller needs data controller for processing
        if self.plot_controller and self.data_controller:
            self.plot_controller.set_data_controller(self.data_controller)
            print("DEBUG: Connected plot controller to data controller")
        
        # Report controller needs access to all data
        if self.report_controller:
            self.report_controller.set_file_controller(self.file_controller)
            self.report_controller.set_plot_controller(self.plot_controller)
            print("DEBUG: Connected report controller to other controllers")
    
    # Getter methods for controllers (for views to access)
    def get_file_controller(self):
        """Get the file controller instance."""
        if not self.file_controller:
            print("WARNING: File controller not initialized")
        return self.file_controller
    
    def get_plot_controller(self):
        """Get the plot controller instance."""
        if not self.plot_controller:
            print("WARNING: Plot controller not initialized")
        return self.plot_controller
    
    def get_report_controller(self):
        """Get the report controller instance."""
        if not self.report_controller:
            print("WARNING: Report controller not initialized")
        return self.report_controller
    
    def get_data_controller(self):
        """Get the data controller instance."""
        if not self.data_controller:
            print("WARNING: Data controller not initialized")
        return self.data_controller
    
    def get_image_controller(self):
        """Get the image controller instance."""
        if not self.image_controller:
            print("WARNING: Image controller not initialized")
        return self.image_controller
    
    def get_calculation_controller(self):
        """Get the calculation controller instance."""
        if not self.calculation_controller:
            print("WARNING: Calculation controller not initialized")
        return self.calculation_controller
    
    # Getter methods for models (for views and controllers to access)
    def get_data_model(self):
        """Get the data model instance."""
        return self.data_model
    
    def get_file_model(self):
        """Get the file model instance."""
        return self.file_model
    
    def get_plot_model(self):
        """Get the plot model instance."""
        return self.plot_model
    
    # Application lifecycle methods
    def start_application(self):
        """Start the application after initialization."""
        if not self.is_initialized:
            print("ERROR: Cannot start application - not initialized")
            return False
        
        print("DEBUG: Starting application...")
        # Additional startup logic can go here
        return True
    
    def shutdown_application(self):
        """Shutdown the application and clean up resources."""
        print("DEBUG: Shutting down application...")
        
        try:
            # Stop all active threads
            for thread in self.active_threads:
                if thread.is_alive():
                    print(f"DEBUG: Waiting for thread {thread.name} to finish...")
                    thread.join(timeout=2.0)
            
            # Clean up models
            if hasattr(self.data_model, 'clear_all_files'):
                self.data_model.clear_all_files()
            
            if hasattr(self.file_model, 'clear_all_caches'):
                self.file_model.clear_all_caches()
            
            if hasattr(self.plot_model, 'clear_all_data'):
                self.plot_model.clear_all_data()
            
            print("DEBUG: Application shutdown completed")
            return True
            
        except Exception as e:
            print(f"ERROR: Error during application shutdown: {e}")
            return False
    
    def get_application_status(self) -> Dict[str, Any]:
        """Get the current status of the application."""
        status = {
            'initialized': self.is_initialized,
            'initialization_error': self.initialization_error,
            'active_threads': len(self.active_threads),
            'controllers_initialized': {
                'file_controller': self.file_controller is not None,
                'plot_controller': self.plot_controller is not None,
                'report_controller': self.report_controller is not None,
                'data_controller': self.data_controller is not None,
                'image_controller': self.image_controller is not None,
                'calculation_controller': self.calculation_controller is not None
            },
            'models_status': {
                'data_model': self.data_model.get_stats() if self.data_model else None,
                'file_model': self.file_model.get_model_stats() if self.file_model else None,
                'plot_model': self.plot_model.get_data_summary() if self.plot_model else None
            }
        }
        
        print(f"DEBUG: Application status: {status['initialized']}")
        return status
    
    def add_thread(self, thread: threading.Thread):
        """Add a thread to the active threads list."""
        self.active_threads.append(thread)
        print(f"DEBUG: Added thread '{thread.name}' to active threads")
    
    def remove_thread(self, thread: threading.Thread):
        """Remove a thread from the active threads list."""
        if thread in self.active_threads:
            self.active_threads.remove(thread)
            print(f"DEBUG: Removed thread '{thread.name}' from active threads")
    
    def handle_error(self, error: Exception, context: str = ""):
        """Handle application errors centrally."""
        error_msg = f"Error in {context}: {error}" if context else f"Application error: {error}"
        print(f"ERROR: {error_msg}")
        
        # Additional error handling logic can go here
        # For example: logging, user notification, recovery attempts
        
        return error_msg
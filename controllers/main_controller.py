# controllers/main_controller.py
"""
controllers/main_controller.py
Main application controller that coordinates all other controllers and services.
This consolidates functionality from main_gui.py and application management code.
"""

import os
import sys
import threading
import queue
import traceback
import tkinter as tk
from tkinter import messagebox
from typing import Optional, Dict, Any, List, Callable
from datetime import datetime

# Local imports
from utils import debug_print, get_resource_path, APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT

# Model imports
from models.data_model import DataModel
from models.file_model import FileModel
from models.plot_model import PlotModel


class MainController:
    """Main application controller that coordinates between all components."""
    
    def __init__(self):
        """Initialize the main controller."""
        # Core application state
        self.is_initialized = False
        self.initialization_error: Optional[str] = None
        self.startup_time = datetime.now()
        
        # Threading and event management
        self.active_threads: List[threading.Thread] = []
        self.event_queue = queue.Queue()
        self.lock = threading.Lock()
        self.report_thread = None
        self.report_queue = queue.Queue()
        
        # Application configuration
        self.app_name = "SDR DataViewer"
        self.app_version = "0.0.13"
        self.app_title = f"{self.app_name} Beta Version {self.app_version}"
        
        # GUI reference (set when GUI is created)
        self.gui = None
        self.root = None
        
        # Initialize core models
        self.data_model = DataModel()
        self.file_model = FileModel()
        self.plot_model = PlotModel()
        
        # Controller instances (initialized later)
        self.file_controller: Optional['FileController'] = None
        self.plot_controller: Optional['PlotController'] = None
        self.report_controller: Optional['ReportController'] = None
        self.data_controller: Optional['DataController'] = None
        self.image_controller: Optional['ImageController'] = None
        self.calculation_controller: Optional['CalculationController'] = None
        
        # Service instances (initialized later)
        self.file_service: Optional[Any] = None
        self.plot_service: Optional[Any] = None
        self.report_service: Optional[Any] = None
        self.processing_service: Optional[Any] = None
        self.database_service: Optional[Any] = None
        self.image_service: Optional[Any] = None
        
        # Application state variables (from main_gui.py)
        self.current_file: Optional[str] = None
        self.file_path: Optional[str] = None
        self.sheets: Dict[str, Any] = {}
        self.filtered_sheets: Dict[str, Any] = {}
        self.all_filtered_sheets: List[Dict[str, Any]] = []
        self.sheet_images: Dict[str, Dict[str, List[str]]] = {}
        
        # Plot and UI state
        self.selected_sheet = None  # Will be StringVar when GUI is created
        self.selected_plot_type = None  # Will be StringVar when GUI is created
        self.plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        self.standard_plot_options = ["TPM", "Normalized TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        self.user_test_simulation_plot_options = ["TPM", "Normalized TPM", "Draw Pressure", "Power Efficiency", "TPM (Bar)"]
        self.is_user_test_simulation = False
        self.num_columns_per_sample = 12
        
        # UI components (initialized when GUI is created)
        self.crop_enabled = None  # Will be BooleanVar when GUI is created
        self.file_dropdown_var = None  # Will be StringVar when GUI is created
        
        print("DEBUG: MainController initialized")
        print(f"DEBUG: Application: {self.app_title}")
        print(f"DEBUG: Startup time: {self.startup_time}")
        print(f"DEBUG: Core models initialized: DataModel, FileModel, PlotModel")
    
    def set_gui_reference(self, gui, root):
        """Set reference to main GUI and root window."""
        self.gui = gui
        self.root = root
        
        # Initialize Tkinter variables now that we have root
        self.selected_sheet = tk.StringVar()
        self.selected_plot_type = tk.StringVar(value="TPM")
        self.crop_enabled = tk.BooleanVar(value=False)
        self.file_dropdown_var = tk.StringVar()
        
        # Set up GUI event handling
        self.root.protocol("WM_DELETE_WINDOW", self.on_application_close)
        
        print("DEBUG: MainController connected to GUI")
        print(f"DEBUG: Root window: {root}")
        print("DEBUG: Tkinter variables initialized")
    
    def initialize_application(self) -> bool:
        """Initialize the complete application stack."""
        debug_print("DEBUG: Starting complete application initialization...")
        
        try:
            # Step 1: Initialize services
            if not self.initialize_services():
                return False
            
            # Step 2: Initialize controllers  
            if not self.initialize_controllers():
                return False
            
            # Step 3: Set up cross-controller communication
            self._setup_controller_communication()
            
            # Step 4: Connect controllers to GUI
            self._connect_controllers_to_gui()
            
            # Step 5: Mark as initialized
            self.is_initialized = True
            
            # Step 6: Configure application
            self._configure_application()
            
            print("DEBUG: Complete application initialization successful")
            print(f"DEBUG: Initialization time: {(datetime.now() - self.startup_time).total_seconds():.2f} seconds")
            return True
            
        except Exception as e:
            error_msg = f"Application initialization failed: {e}"
            print(f"ERROR: {error_msg}")
            self.initialization_error = error_msg
            traceback.print_exc()
            return False
    
    def initialize_services(self) -> bool:
        """Initialize all service instances."""
        try:
            debug_print("DEBUG: Initializing services...")
            
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
            print(f"DEBUG: Services created: {[s.name for s in [self.file_service, self.plot_service, self.report_service, self.processing_service, self.database_service, self.image_service]]}")
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
                debug_print(f"DEBUG: Created placeholder {name}")
                
            def __getattr__(self, name):
                """Handle any method calls to placeholder service."""
                def placeholder_method(*args, **kwargs):
                    debug_print(f"DEBUG: Placeholder {self.name}.{name} called with args={args}, kwargs={kwargs}")
                    return None
                return placeholder_method
        
        return PlaceholderService(service_name)
    
    def initialize_controllers(self) -> bool:
        """Initialize all controller instances."""
        try:
            debug_print("DEBUG: Initializing controllers...")
            
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
            traceback.print_exc()
            return False
    
    def _setup_controller_communication(self):
        """Set up communication between controllers."""
        debug_print("DEBUG: Setting up controller communication...")
        
        try:
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
            
            # Image controller needs file controller for file operations
            if self.image_controller and self.file_controller:
                self.image_controller.set_file_controller(self.file_controller)
                print("DEBUG: Connected image controller to file controller")
            
            # Calculation controller needs data controller for data access
            if self.calculation_controller and self.data_controller:
                self.calculation_controller.set_data_controller(self.data_controller)
                print("DEBUG: Connected calculation controller to data controller")
            
            print("DEBUG: Controller communication setup completed")
            
        except Exception as e:
            print(f"ERROR: Failed to setup controller communication: {e}")
            traceback.print_exc()
    
    def _connect_controllers_to_gui(self):
        """Connect controllers to GUI for UI operations."""
        debug_print("DEBUG: Connecting controllers to GUI...")
        
        try:
            if self.gui:
                # Connect all controllers to GUI
                controllers = [
                    self.file_controller,
                    self.plot_controller, 
                    self.report_controller,
                    self.data_controller,
                    self.image_controller,
                    self.calculation_controller
                ]
                
                for controller in controllers:
                    if controller and hasattr(controller, 'set_gui_reference'):
                        controller.set_gui_reference(self.gui)
                        debug_print(f"DEBUG: Connected {type(controller).__name__} to GUI")
                
                # Special setup for specific controllers
                if self.plot_controller:
                    # Connect plot controller to GUI variables
                    if hasattr(self.gui, 'selected_plot_type'):
                        self.plot_controller.selected_plot_type = self.gui.selected_plot_type
                    if hasattr(self.gui, 'plot_options'):
                        self.plot_controller.plot_options = self.gui.plot_options
                
                print("DEBUG: All controllers connected to GUI")
            else:
                print("WARNING: No GUI reference available for controller connection")
                
        except Exception as e:
            print(f"ERROR: Failed to connect controllers to GUI: {e}")
            traceback.print_exc()
    
    def _configure_application(self):
        """Configure application settings and appearance."""
        debug_print("DEBUG: Configuring application...")
        
        try:
            if self.root:
                # Set application title
                self.root.title(self.app_title)
                
                # Set application icon
                try:
                    icon_path = get_resource_path('resources/ccell_icon.png')
                    if os.path.exists(icon_path):
                        self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
                        debug_print("DEBUG: Application icon set")
                except Exception as e:
                    debug_print(f"WARNING: Could not set application icon: {e}")
                
                # Configure UI colors and style
                self._configure_ui_style()
                
                # Set initial window properties
                self.root.minsize(1200, 800)
                self._set_window_size(0.8, 0.8)
                self._center_window(self.root)
                
                debug_print("DEBUG: Application configuration completed")
            else:
                print("WARNING: No root window available for configuration")
                
        except Exception as e:
            print(f"ERROR: Failed to configure application: {e}")
            traceback.print_exc()
    
    def _configure_ui_style(self):
        """Configure UI appearance and style."""
        try:
            import tkinter.ttk as ttk
            
            # Configure root window
            self.root.configure(bg=APP_BACKGROUND_COLOR)
            
            # Configure ttk styles
            style = ttk.Style()
            style.configure('TLabel', background=APP_BACKGROUND_COLOR, font=FONT)
            style.configure('TButton', background=BUTTON_COLOR, font=FONT, padding=6)
            style.configure('TCombobox', font=FONT)
            style.map('TCombobox', background=[('readonly', APP_BACKGROUND_COLOR)])
            
            # Configure widget colors
            for widget in self.root.winfo_children():
                try:
                    # Only works for standard tkinter widgets
                    widget.configure(bg='#EFEFEF')
                except Exception:
                    # Skip ttk widgets that don't support bg
                    continue
            
            debug_print("DEBUG: UI style configured")
            
        except Exception as e:
            debug_print(f"WARNING: Failed to configure UI style: {e}")
    
    def _set_window_size(self, width_ratio: float, height_ratio: float):
        """Set window size as a ratio of screen size."""
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            window_width = int(screen_width * width_ratio)
            window_height = int(screen_height * height_ratio)
            
            self.root.geometry(f"{window_width}x{window_height}")
            debug_print(f"DEBUG: Window size set to {window_width}x{window_height}")
            
        except Exception as e:
            debug_print(f"WARNING: Failed to set window size: {e}")
    
    def _center_window(self, window):
        """Center a window on the screen."""
        try:
            window.update_idletasks()
            
            # Get window dimensions
            width = window.winfo_width()
            height = window.winfo_height()
            
            # Calculate center position
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            # Set window position
            window.geometry(f"{width}x{height}+{x}+{y}")
            debug_print(f"DEBUG: Window centered at {x},{y}")
            
        except Exception as e:
            debug_print(f"WARNING: Failed to center window: {e}")
    
    def start_application(self) -> bool:
        """Start the application after initialization."""
        if not self.is_initialized:
            print("ERROR: Cannot start application - not initialized")
            return False
        
        debug_print("DEBUG: Starting application...")
        
        try:
            # Show startup menu if GUI is available
            if self.gui and hasattr(self.gui, 'show_startup_menu'):
                self.gui.show_startup_menu()
            
            # Start any background services
            self._start_background_services()
            
            # Mark application as started
            self.is_started = True
            
            print("DEBUG: Application started successfully")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to start application: {e}")
            traceback.print_exc()
            return False
    
    def _start_background_services(self):
        """Start any background services or threads."""
        debug_print("DEBUG: Starting background services...")
        
        try:
            # Start database service if available
            if hasattr(self.database_service, 'start'):
                self.database_service.start()
                debug_print("DEBUG: Database service started")
            
            # Start any other background services here
            
            debug_print("DEBUG: Background services started")
            
        except Exception as e:
            debug_print(f"WARNING: Failed to start some background services: {e}")
    
    def on_application_close(self):
        """Handle application close event."""
        debug_print("DEBUG: Application close requested")
        
        try:
            # Show confirmation dialog for unsaved changes
            if self._check_unsaved_changes():
                result = messagebox.askyesnocancel(
                    "Unsaved Changes",
                    "You have unsaved changes. Do you want to save before closing?"
                )
                if result is None:  # Cancel
                    return
                elif result:  # Yes - save
                    self._save_application_state()
            
            # Shutdown application
            self.shutdown_application()
            
        except Exception as e:
            print(f"ERROR: Error during application close: {e}")
            traceback.print_exc()
            # Force exit on error
            os._exit(1)
    
    def _check_unsaved_changes(self) -> bool:
        """Check if there are unsaved changes."""
        try:
            # Check each controller for unsaved changes
            if self.file_controller and hasattr(self.file_controller, 'has_unsaved_changes'):
                if self.file_controller.has_unsaved_changes():
                    return True
            
            if self.data_controller and hasattr(self.data_controller, 'has_unsaved_changes'):
                if self.data_controller.has_unsaved_changes():
                    return True
            
            return False
            
        except Exception as e:
            debug_print(f"WARNING: Error checking unsaved changes: {e}")
            return False
    
    def _save_application_state(self):
        """Save current application state."""
        debug_print("DEBUG: Saving application state...")
        
        try:
            # Save through file controller
            if self.file_controller and hasattr(self.file_controller, 'save_current_state'):
                self.file_controller.save_current_state()
            
            # Save configuration
            self._save_configuration()
            
            debug_print("DEBUG: Application state saved")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to save application state: {e}")
    
    def _save_configuration(self):
        """Save application configuration."""
        try:
            config = {
                'window_geometry': self.root.geometry() if self.root else None,
                'selected_plot_type': self.selected_plot_type.get() if self.selected_plot_type else "TPM",
                'plot_options': self.plot_options,
                'num_columns_per_sample': self.num_columns_per_sample,
                'last_saved': datetime.now().isoformat()
            }
            
            # Save configuration to file
            config_file = 'application_config.json'
            import json
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=2)
            
            debug_print(f"DEBUG: Configuration saved to {config_file}")
            
        except Exception as e:
            debug_print(f"WARNING: Failed to save configuration: {e}")
    
    def shutdown_application(self) -> bool:
        """Shutdown the application and clean up resources."""
        debug_print("DEBUG: Shutting down application...")
        
        try:
            # Stop all active threads
            self._stop_all_threads()
            
            # Shutdown controllers
            self._shutdown_controllers()
            
            # Shutdown services
            self._shutdown_services()
            
            # Clean up models
            self._cleanup_models()
            
            # Close GUI
            self._close_gui()
            
            print("DEBUG: Application shutdown completed successfully")
            return True
            
        except Exception as e:
            print(f"ERROR: Error during application shutdown: {e}")
            traceback.print_exc()
            return False
        finally:
            # Force exit
            try:
                if self.root:
                    self.root.quit()
                    self.root.destroy()
            except:
                pass
            os._exit(0)
    
    def _stop_all_threads(self):
        """Stop all active threads."""
        debug_print(f"DEBUG: Stopping {len(self.active_threads)} active threads...")
        
        try:
            for thread in self.active_threads[:]:  # Copy list to avoid modification during iteration
                if thread.is_alive():
                    debug_print(f"DEBUG: Waiting for thread {thread.name} to finish...")
                    thread.join(timeout=2.0)
                    if thread.is_alive():
                        debug_print(f"WARNING: Thread {thread.name} did not finish gracefully")
                self.active_threads.remove(thread)
            
            # Stop report thread if active
            if self.report_thread and self.report_thread.is_alive():
                debug_print("DEBUG: Stopping report thread...")
                self.report_thread.join(timeout=2.0)
            
            debug_print("DEBUG: All threads stopped")
            
        except Exception as e:
            debug_print(f"WARNING: Error stopping threads: {e}")
    
    def _shutdown_controllers(self):
        """Shutdown all controllers."""
        debug_print("DEBUG: Shutting down controllers...")
        
        try:
            controllers = [
                ('file_controller', self.file_controller),
                ('plot_controller', self.plot_controller),
                ('report_controller', self.report_controller),
                ('data_controller', self.data_controller),
                ('image_controller', self.image_controller),
                ('calculation_controller', self.calculation_controller)
            ]
            
            for name, controller in controllers:
                if controller and hasattr(controller, 'shutdown'):
                    try:
                        controller.shutdown()
                        debug_print(f"DEBUG: {name} shutdown completed")
                    except Exception as e:
                        debug_print(f"WARNING: Error shutting down {name}: {e}")
            
            debug_print("DEBUG: Controllers shutdown completed")
            
        except Exception as e:
            debug_print(f"WARNING: Error shutting down controllers: {e}")
    
    def _shutdown_services(self):
        """Shutdown all services."""
        debug_print("DEBUG: Shutting down services...")
        
        try:
            services = [
                ('file_service', self.file_service),
                ('plot_service', self.plot_service),
                ('report_service', self.report_service),
                ('processing_service', self.processing_service),
                ('database_service', self.database_service),
                ('image_service', self.image_service)
            ]
            
            for name, service in services:
                if service and hasattr(service, 'shutdown'):
                    try:
                        service.shutdown()
                        debug_print(f"DEBUG: {name} shutdown completed")
                    except Exception as e:
                        debug_print(f"WARNING: Error shutting down {name}: {e}")
            
            debug_print("DEBUG: Services shutdown completed")
            
        except Exception as e:
            debug_print(f"WARNING: Error shutting down services: {e}")
    
    def _cleanup_models(self):
        """Clean up models and release resources."""
        debug_print("DEBUG: Cleaning up models...")
        
        try:
            if hasattr(self.data_model, 'clear_all_files'):
                self.data_model.clear_all_files()
            
            if hasattr(self.file_model, 'clear_all_caches'):
                self.file_model.clear_all_caches()
            
            if hasattr(self.plot_model, 'clear_all_data'):
                self.plot_model.clear_all_data()
            
            debug_print("DEBUG: Models cleanup completed")
            
        except Exception as e:
            debug_print(f"WARNING: Error cleaning up models: {e}")
    
    def _close_gui(self):
        """Close GUI components."""
        debug_print("DEBUG: Closing GUI...")
        
        try:
            if self.root:
                self.root.quit()
                self.root.destroy()
            
            debug_print("DEBUG: GUI closed")
            
        except Exception as e:
            debug_print(f"WARNING: Error closing GUI: {e}")
    
    def add_thread(self, thread: threading.Thread):
        """Add a thread to the active threads list."""
        with self.lock:
            self.active_threads.append(thread)
            debug_print(f"DEBUG: Added thread {thread.name} to active threads list")
    
    def remove_thread(self, thread: threading.Thread):
        """Remove a thread from the active threads list."""
        with self.lock:
            if thread in self.active_threads:
                self.active_threads.remove(thread)
                debug_print(f"DEBUG: Removed thread {thread.name} from active threads list")
    
    def post_event(self, event_type: str, data: Any = None):
        """Post an event to the event queue for processing."""
        try:
            event = {
                'type': event_type,
                'data': data,
                'timestamp': datetime.now()
            }
            self.event_queue.put(event)
            debug_print(f"DEBUG: Posted event: {event_type}")
        except Exception as e:
            debug_print(f"ERROR: Failed to post event {event_type}: {e}")
    
    def process_events(self):
        """Process events from the event queue."""
        try:
            while not self.event_queue.empty():
                event = self.event_queue.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        except Exception as e:
            debug_print(f"ERROR: Error processing events: {e}")
    
    def _handle_event(self, event: Dict[str, Any]):
        """Handle a specific event."""
        try:
            event_type = event.get('type')
            data = event.get('data')
            
            debug_print(f"DEBUG: Handling event: {event_type}")
            
            # Route events to appropriate handlers
            if event_type == 'file_loaded':
                self._handle_file_loaded_event(data)
            elif event_type == 'plot_generated':
                self._handle_plot_generated_event(data)
            elif event_type == 'report_completed':
                self._handle_report_completed_event(data)
            elif event_type == 'error_occurred':
                self._handle_error_event(data)
            else:
                debug_print(f"WARNING: Unknown event type: {event_type}")
        
        except Exception as e:
            debug_print(f"ERROR: Error handling event: {e}")
    
    def _handle_file_loaded_event(self, data):
        """Handle file loaded event."""
        debug_print("DEBUG: Handling file loaded event")
        # Notify relevant controllers
        if self.plot_controller:
            self.plot_controller.on_data_changed()
    
    def _handle_plot_generated_event(self, data):
        """Handle plot generated event."""
        debug_print("DEBUG: Handling plot generated event")
        # Update UI or notify other components
    
    def _handle_report_completed_event(self, data):
        """Handle report completed event."""
        debug_print("DEBUG: Handling report completed event")
        # Show completion message or update status
    
    def _handle_error_event(self, data):
        """Handle error event."""
        debug_print(f"DEBUG: Handling error event: {data}")
        # Log error and potentially show user notification
    
    # Getter methods for controllers (for views to access)
    def get_file_controller(self):
        """Get the file controller instance."""
        if not self.file_controller:
            debug_print("WARNING: File controller not initialized")
        return self.file_controller
    
    def get_plot_controller(self):
        """Get the plot controller instance."""
        if not self.plot_controller:
            debug_print("WARNING: Plot controller not initialized")
        return self.plot_controller
    
    def get_report_controller(self):
        """Get the report controller instance."""
        if not self.report_controller:
            debug_print("WARNING: Report controller not initialized")
        return self.report_controller
    
    def get_data_controller(self):
        """Get the data controller instance."""
        if not self.data_controller:
            debug_print("WARNING: Data controller not initialized")
        return self.data_controller
    
    def get_image_controller(self):
        """Get the image controller instance."""
        if not self.image_controller:
            debug_print("WARNING: Image controller not initialized")
        return self.image_controller
    
    def get_calculation_controller(self):
        """Get the calculation controller instance."""
        if not self.calculation_controller:
            debug_print("WARNING: Calculation controller not initialized")
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
    
    # Getter methods for services (for controllers to access)
    def get_file_service(self):
        """Get the file service instance."""
        return self.file_service
    
    def get_plot_service(self):
        """Get the plot service instance."""
        return self.plot_service
    
    def get_report_service(self):
        """Get the report service instance."""
        return self.report_service
    
    def get_processing_service(self):
        """Get the processing service instance."""
        return self.processing_service
    
    def get_database_service(self):
        """Get the database service instance."""
        return self.database_service
    
    def get_image_service(self):
        """Get the image service instance."""
        return self.image_service
    
    def get_application_status(self) -> Dict[str, Any]:
        """Get the current status of the application."""
        status = {
            'initialized': self.is_initialized,
            'initialization_error': self.initialization_error,
            'startup_time': self.startup_time.isoformat(),
            'active_threads': len(self.active_threads),
            'event_queue_size': self.event_queue.qsize(),
            'current_file': self.current_file,
            'loaded_sheets': len(self.filtered_sheets),
            'total_files': len(self.all_filtered_sheets),
            'controllers_status': {
                'file_controller': self.file_controller is not None,
                'plot_controller': self.plot_controller is not None,
                'report_controller': self.report_controller is not None,
                'data_controller': self.data_controller is not None,
                'image_controller': self.image_controller is not None,
                'calculation_controller': self.calculation_controller is not None
            },
            'services_status': {
                'file_service': self.file_service is not None,
                'plot_service': self.plot_service is not None,
                'report_service': self.report_service is not None,
                'processing_service': self.processing_service is not None,
                'database_service': self.database_service is not None,
                'image_service': self.image_service is not None
            },
            'models_status': {
                'data_model': self.data_model.get_stats() if self.data_model and hasattr(self.data_model, 'get_stats') else None,
                'file_model': self.file_model.get_model_stats() if self.file_model and hasattr(self.file_model, 'get_model_stats') else None,
                'plot_model': self.plot_model.get_data_summary() if self.plot_model and hasattr(self.plot_model, 'get_data_summary') else None
            }
        }
        
        debug_print(f"DEBUG: Application status requested - initialized: {status['initialized']}")
        return status
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get performance statistics."""
        import psutil
        import gc
        
        try:
            # Get memory usage
            process = psutil.Process()
            memory_info = process.memory_info()
            
            # Get garbage collection stats
            gc_stats = gc.get_stats()
            
            stats = {
                'memory_usage_mb': memory_info.rss / 1024 / 1024,
                'cpu_percent': process.cpu_percent(),
                'thread_count': threading.active_count(),
                'gc_collections': sum(stat['collections'] for stat in gc_stats),
                'uptime_seconds': (datetime.now() - self.startup_time).total_seconds(),
                'active_threads': len(self.active_threads),
                'event_queue_size': self.event_queue.qsize()
            }
            
            return stats
            
        except Exception as e:
            debug_print(f"WARNING: Could not get performance stats: {e}")
            return {'error': str(e)}
    
    def create_menu_structure(self) -> Dict[str, List[Dict[str, Any]]]:
        """Create the main application menu structure."""
        menu_structure = {
            'File': [
                {'label': 'New', 'command': 'show_new_template_dialog'},
                {'label': 'Load Excel', 'command': 'load_initial_file'},
                {'label': 'Load VAP3', 'command': 'open_vap3_file'},
                {'separator': True},
                {'label': 'Batch Load Folder', 'command': 'batch_load_folder'},
                {'separator': True},
                {'label': 'Save As VAP3', 'command': 'save_as_vap3'},
                {'separator': True},
                {'label': 'Update Database', 'accelerator': 'Ctrl+U', 'command': 'update_database'},
                {'separator': True},
                {'label': 'Exit', 'command': 'on_app_close'}
            ],
            'View': [
                {'label': 'View Raw Data', 'command': 'view_raw_data'},
                {'label': 'Trend Analysis', 'command': 'open_trend_analysis_window'}
            ],
            'Database': [
                {'label': 'Browse Database', 'command': 'show_database_browser'},
                {'label': 'Load from Database', 'command': 'load_from_database'}
            ],
            'Calculate': [
                {'label': 'Viscosity', 'command': 'open_viscosity_calculator'}
            ],
            'Compare': [
                {'label': 'Sample Comparison', 'command': 'show_sample_comparison'}
            ],
            'Reports': [
                {'label': 'Generate Test Report', 'command': 'generate_test_report'},
                {'label': 'Generate Full Report', 'command': 'generate_full_report'}
            ],
            'Help': [
                {'label': 'Help', 'command': 'show_help'},
                {'separator': True},
                {'label': 'About', 'command': 'show_about'}
            ]
        }
        
        debug_print("DEBUG: Menu structure created")
        return menu_structure
    
    def handle_command(self, command: str, *args, **kwargs) -> Any:
        """Handle application commands."""
        debug_print(f"DEBUG: Handling command: {command}")
        
        try:
            # Route commands to appropriate controllers
            if command.startswith('file_'):
                return self._handle_file_command(command, *args, **kwargs)
            elif command.startswith('plot_'):
                return self._handle_plot_command(command, *args, **kwargs)
            elif command.startswith('report_'):
                return self._handle_report_command(command, *args, **kwargs)
            elif command.startswith('data_'):
                return self._handle_data_command(command, *args, **kwargs)
            else:
                return self._handle_general_command(command, *args, **kwargs)
                
        except Exception as e:
            debug_print(f"ERROR: Error handling command {command}: {e}")
            return False
    
    def _handle_file_command(self, command: str, *args, **kwargs):
        """Handle file-related commands."""
        if self.file_controller:
            method_name = command.replace('file_', '')
            if hasattr(self.file_controller, method_name):
                return getattr(self.file_controller, method_name)(*args, **kwargs)
        return False
    
    def _handle_plot_command(self, command: str, *args, **kwargs):
        """Handle plot-related commands."""
        if self.plot_controller:
            method_name = command.replace('plot_', '')
            if hasattr(self.plot_controller, method_name):
                return getattr(self.plot_controller, method_name)(*args, **kwargs)
        return False
    
    def _handle_report_command(self, command: str, *args, **kwargs):
        """Handle report-related commands."""
        if self.report_controller:
            method_name = command.replace('report_', '')
            if hasattr(self.report_controller, method_name):
                return getattr(self.report_controller, method_name)(*args, **kwargs)
        return False
    
    def _handle_data_command(self, command: str, *args, **kwargs):
        """Handle data-related commands."""
        if self.data_controller:
            method_name = command.replace('data_', '')
            if hasattr(self.data_controller, method_name):
                return getattr(self.data_controller, method_name)(*args, **kwargs)
        return False
    
    def _handle_general_command(self, command: str, *args, **kwargs):
        """Handle general application commands."""
        if command == 'show_about':
            if self.gui and hasattr(self.gui, 'show_about'):
                return self.gui.show_about()
        elif command == 'show_help':
            if self.gui and hasattr(self.gui, 'show_help'):
                return self.gui.show_help()
        elif command == 'on_app_close':
            return self.on_application_close()
        else:
            debug_print(f"WARNING: Unknown general command: {command}")
        return False
    
    def get_version_info(self) -> Dict[str, str]:
        """Get application version information."""
        return {
            'app_name': self.app_name,
            'version': self.app_version,
            'title': self.app_title,
            'startup_time': self.startup_time.isoformat()
        }
    
    def is_ready(self) -> bool:
        """Check if application is ready for use."""
        return (self.is_initialized and 
                self.file_controller is not None and 
                self.plot_controller is not None and 
                self.data_controller is not None)
    
    def get_debug_info(self) -> Dict[str, Any]:
        """Get comprehensive debug information."""
        debug_info = {
            'application': self.get_version_info(),
            'status': self.get_application_status(),
            'performance': self.get_performance_stats(),
            'initialization_error': self.initialization_error,
            'thread_details': [
                {
                    'name': thread.name,
                    'alive': thread.is_alive(),
                    'daemon': thread.daemon
                } for thread in self.active_threads
            ]
        }
        
        debug_print("DEBUG: Debug information compiled")
        return debug_info
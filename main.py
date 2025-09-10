# main.py
"""
main.py
New architecture entry point for the DataViewer application.
This replaces the direct instantiation approach and uses the clean architecture.
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox

# Add current directory to Python path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the new architecture components
from controllers.main_controller import MainController
from views.main_window import MainWindow
from utils.debug import debug_print, error_print, set_debug_enabled
from utils.config import get_config


class DataViewerApplication:
    """Main application class using the new architecture."""
    
    def __init__(self):
        """Initialize the DataViewer application."""
        print("=" * 60)
        print("🚀 STARTING DATAVIEWER APPLICATION")
        print("=" * 60)
        
        # Initialize core components
        self.root: tk.Tk = None
        self.main_controller: MainController = None
        self.main_window: MainWindow = None
        self.config = None
        
        # Application state
        self.is_initialized = False
        self.initialization_error = None
        
        debug_print("DataViewerApplication instance created")
    
    def initialize(self) -> bool:
        """Initialize the complete application."""
        try:
            debug_print("Starting application initialization...")
            
            # Step 1: Load configuration
            if not self._initialize_config():
                return False
            
            # Step 2: Create Tkinter root
            if not self._initialize_tkinter():
                return False
            
            # Step 3: Initialize controller layer
            if not self._initialize_controllers():
                return False
            
            # Step 4: Initialize view layer
            if not self._initialize_views():
                return False
            
            # Step 5: Connect components
            if not self._connect_components():
                return False
            
            # Step 6: Finalize setup
            if not self._finalize_initialization():
                return False
            
            self.is_initialized = True
            debug_print("Application initialization completed successfully!")
            print("✅ DataViewer initialized successfully")
            
            return True
            
        except Exception as e:
            error_message = f"Application initialization failed: {e}"
            error_print(error_message, e)
            self.initialization_error = error_message
            return False
    
    def _initialize_config(self) -> bool:
        """Initialize application configuration."""
        try:
            debug_print("Initializing configuration...")
            
            self.config = get_config()
            
            # Set debug mode from config
            debug_enabled = True  # Could be read from config
            set_debug_enabled(debug_enabled)
            
            debug_print("Configuration initialized successfully")
            return True
            
        except Exception as e:
            error_print("Failed to initialize configuration", e)
            return False
    
    def _initialize_tkinter(self) -> bool:
        """Initialize the Tkinter root window."""
        try:
            debug_print("Initializing Tkinter root...")
            
            self.root = tk.Tk()
            self.root.withdraw()  # Hide until we're ready to show
            
            # Configure root window properties
            self.root.title("DataViewer - Initializing...")
            
            debug_print("Tkinter root initialized successfully")
            return True
            
        except Exception as e:
            error_print("Failed to initialize Tkinter", e)
            return False
    
    def _initialize_controllers(self) -> bool:
        """Initialize the controller layer."""
        try:
            debug_print("Initializing controller layer...")
            
            # Create main controller
            self.main_controller = MainController()
            
            # Initialize the complete controller stack
            if not self.main_controller.initialize_application():
                error_message = self.main_controller.initialization_error or "Unknown controller initialization error"
                error_print(f"Controller initialization failed: {error_message}")
                return False
            
            debug_print("Controller layer initialized successfully")
            return True
            
        except Exception as e:
            error_print("Failed to initialize controllers", e)
            return False
    
    def _initialize_views(self) -> bool:
        """Initialize the view layer."""
        try:
            debug_print("Initializing view layer...")
            
            # Create main window
            self.main_window = MainWindow(self.main_controller)
            
            # Setup the window
            self.main_window.setup_window()
            self.main_window.create_layout()
            
            debug_print("View layer initialized successfully")
            return True
            
        except Exception as e:
            error_print("Failed to initialize views", e)
            return False
    
    def _connect_components(self) -> bool:
        """Connect controllers and views."""
        try:
            debug_print("Connecting application components...")
            
            # Set up callbacks from view to controllers
            self._setup_view_callbacks()
            
            # Initialize any cross-component communication
            self._setup_component_communication()
            
            debug_print("Components connected successfully")
            return True
            
        except Exception as e:
            error_print("Failed to connect components", e)
            return False
    
    def _setup_view_callbacks(self):
        """Set up callbacks from views to controllers."""
        debug_print("Setting up view callbacks...")
        
        # File operations callbacks
        def on_file_selected(filename):
            debug_print(f"File selected callback: {filename}")
            file_controller = self.main_controller.get_file_controller()
            if file_controller:
                # file_controller.load_file(filename)  # Will be implemented
                pass
        
        # Sheet selection callbacks
        def on_sheet_selected(sheet_name):
            debug_print(f"Sheet selected callback: {sheet_name}")
            data_controller = self.main_controller.get_data_controller()
            if data_controller:
                # data_controller.set_active_sheet(sheet_name)  # Will be implemented
                pass
        
        # Plot type callbacks
        def on_plot_type_changed(plot_type):
            debug_print(f"Plot type changed callback: {plot_type}")
            plot_controller = self.main_controller.get_plot_controller()
            if plot_controller:
                # plot_controller.set_plot_type(plot_type)  # Will be implemented
                pass
        
        # Set callbacks on the main window
        self.main_window.set_callbacks(
            file_callback=on_file_selected,
            sheet_callback=on_sheet_selected,
            plot_callback=on_plot_type_changed
        )
        
        debug_print("View callbacks configured")
    
    def _setup_component_communication(self):
        """Set up communication between components."""
        debug_print("Setting up component communication...")
        
        # Controllers can communicate through the main controller
        # Views communicate through callbacks
        # Services are called by controllers
        
        debug_print("Component communication configured")
    
    def _finalize_initialization(self) -> bool:
        """Finalize the application initialization."""
        try:
            debug_print("Finalizing application initialization...")
            
            # Start the controller application
            if not self.main_controller.start_application():
                return False
            
            # Show the main window
            self.root.deiconify()  # Show the window
            self.root.title("DataViewer v3.0")
            
            # Bind close event
            self.root.protocol("WM_DELETE_WINDOW", self._on_application_close)
            
            debug_print("Application finalization completed")
            return True
            
        except Exception as e:
            error_print("Failed to finalize initialization", e)
            return False
    
    def _on_application_close(self):
        """Handle application close event."""
        debug_print("Application close requested")
        
        try:
            # Shutdown controllers
            if self.main_controller:
                self.main_controller.shutdown_application()
            
            # Save configuration
            if self.config:
                self.config.save_config()
            
            # Close Tkinter
            if self.root:
                self.root.quit()
                self.root.destroy()
            
            debug_print("Application shutdown completed")
            
        except Exception as e:
            error_print("Error during application shutdown", e)
        finally:
            # Force exit
            sys.exit(0)
    
    def run(self):
        """Run the application main loop."""
        if not self.is_initialized:
            error_print("Cannot run application - not initialized")
            if self.initialization_error:
                messagebox.showerror("Initialization Error", self.initialization_error)
            return False
        
        try:
            debug_print("Starting application main loop...")
            print("🎯 DataViewer is ready! Starting main loop...")
            
            # Start the Tkinter main loop
            self.main_window.run()
            
            debug_print("Application main loop ended")
            return True
            
        except Exception as e:
            error_print("Error in application main loop", e)
            messagebox.showerror("Application Error", f"A critical error occurred: {e}")
            return False
    
    def get_status(self) -> dict:
        """Get current application status."""
        status = {
            'initialized': self.is_initialized,
            'initialization_error': self.initialization_error,
            'tkinter_ready': self.root is not None,
            'controller_ready': self.main_controller is not None,
            'view_ready': self.main_window is not None,
        }
        
        if self.main_controller:
            controller_status = self.main_controller.get_application_status()
            status.update(controller_status)
        
        return status


def main():
    """Main entry point for the DataViewer application."""
    print("🔥 DataViewer Starting Up...")
    print(f"🐍 Python {sys.version}")
    print(f"📁 Working Directory: {os.getcwd()}")
    
    try:
        # Create application instance
        app = DataViewerApplication()
        
        # Initialize the application
        if not app.initialize():
            print("❌ Failed to initialize DataViewer")
            status = app.get_status()
            print(f"📊 Application Status: {status}")
            return 1
        
        # Run the application
        if not app.run():
            print("❌ DataViewer encountered an error during execution")
            return 1
        
        print("✅ DataViewer closed successfully")
        return 0
        
    except KeyboardInterrupt:
        print("\n⚠️ DataViewer interrupted by user")
        return 130
    except Exception as e:
        print(f"💥 Critical error in DataViewer: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    """Entry point when run directly."""
    sys.exit(main())
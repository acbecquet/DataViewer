"""
Main Sensory Data Collection Window
Lightweight coordinator class that manages all sensory evaluation operations
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
import math
from datetime import datetime
import json
import os
from utils import APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT, debug_print, show_success_message

from .sensory_utils import _lazy_import_cv2, _lazy_import_pil, _lazy_import_sklearn, _lazy_import_pytesseract
from .session_manager import SessionManager
from .sample_manager import SampleManager
from .ui_layout import UILayoutManager
from .plot_manager import PlotManager
from .mode_manager import ModeManager
from .file_io import FileIOManager
from .export_manager import ExportManager

# Import ML components if they exist
try:
    from sensory_ml_training import SensoryMLTrainer, SensoryAIProcessor
    ML_AVAILABLE = True
except ImportError:
    ML_AVAILABLE = False
    debug_print("DEBUG: ML components not available")


class SensoryDataCollectionWindow:
    """Main window for sensory data collection and visualization - Modular Version."""

    def __init__(self, parent, close_callback=None):
        """Initialize the sensory data collection window."""
        self.parent = parent
        self.close_callback = close_callback
        self.window = None
    
        # Core data structures - these remain in main class as they're shared across modules
        self.data = {}
        self.samples = {}
        self.current_sample = None
    
        # Sensory metrics
        self.metrics = [
            "Burnt Taste",
            "Vapor Volume", 
            "Overall Flavor",
            "Smoothness",
            "Overall Liking"
        ]

        # Mode and session data
        self.current_mode = "collection"
        self.all_sessions_data = {}
        self.average_samples = {}

        # Header data fields
        self.header_fields = [
            "Assessor Name",
            "Media", 
            "Puff Length",
            "Date"
        ]

        # SOP text
        self.sop_text = """
    SENSORY EVALUATION STANDARD OPERATING PROCEDURE

    1. PREPARATION:
       - Ensure all cartridges are at room temperature
       - Use clean, odor-free environment, ideally in a fume hood

    2. EVALUATION:
       - Take 2-3 moderate puffs per sample
       - Be sure to compare flavor with original oil if available
       - Rate each attribute on 1-9 scale (1=poor, 9=excellent)

    3. SCALE INTERPRETATION:
       - Burnt Taste: 1=Very Burnt, 9=No Burnt Taste
       - Vapor Volume: 1=Very Low, 9=Very High
       - Overall Flavor: 1=Poor, 9=Excellent
       - Smoothness: 1=Very Harsh, 9=Very Smooth
       - Overall Liking: 1=Dislike Extremely, 9=Like Extremely

    4. NOTES:
       - Record any unusual observations
       - Note any technical issues with samples
            """

        # Initialize base manager classes (no cross-dependencies)
        self.session_manager = SessionManager(self)
        self.sample_manager = SampleManager(self)
        self.ui_layout = UILayoutManager(self)
        self.plot_manager = PlotManager(self)

        # Add cross-references for base managers
        self.session_manager.sample_manager = self.sample_manager
        self.session_manager.plot_manager = self.plot_manager
        self.sample_manager.plot_manager = self.plot_manager
        self.ui_layout.session_manager = self.session_manager
        self.ui_layout.plot_manager = self.plot_manager

        # Initialize managers with cross-dependencies
        self.mode_manager = ModeManager(self, self.session_manager, self.plot_manager)
        self.file_io = FileIOManager(self, self.session_manager, self.sample_manager)
        self.export_manager = ExportManager(self, self.sample_manager, self.plot_manager)

        # Final cross-reference
        self.plot_manager.mode_manager = self.mode_manager

        # Initialize ML components if available
        if ML_AVAILABLE:
            self.ml_trainer = SensoryMLTrainer(self)
            self.ai_processor = SensoryAIProcessor(self)

        debug_print("DEBUG: Initialized modular sensory data collection window")

    def show(self):
        """Create and display the sensory data collection window with optimized sizing."""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Sensory Data Collection")
        self.window.resizable(True, True)

        self.window.state('zoomed')
        self.window.minsize(1000,600)


        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)

        self.setup_layout()
        self.setup_menu()
        self.center_window()

    def on_window_close(self):
        """Handle window close event and call the callback if provided."""
        debug_print("DEBUG: Sensory window close event triggered")

        # Check for unsaved changes and handle auto-save if needed
        try:
            # Need to add save logic on window close
            debug_print("DEBUG: Performing cleanup before window close")

            # Close the window
            self.window.destroy()
            debug_print("DEBUG: Sensory window destroyed")

            # Call the callback to restore main window
            if self.close_callback:
                debug_print("DEBUG: Calling close callback to restore main window")
                self.close_callback()
            else:
                debug_print("DEBUG: No close callback provided")

        except Exception as e:
            debug_print(f"DEBUG: Error during window close: {e}")
            # Still try to close and call callback even if there's an error
            try:
                self.window.destroy()
                if self.close_callback:
                    self.close_callback()
            except:
                pass

    def setup_menu(self):
        """Create the menu bar."""
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Session", command=self.add_new_session)
        file_menu.add_command(label="Load Session", command=self.load_session)
        file_menu.add_command(label="Save Session", command=self.save_session)
        file_menu.add_separator()
        file_menu.add_command(label="Merge Sessions from Files", command=self.merge_sessions_from_files)
        file_menu.add_separator()
        file_menu.add_command(label="Load from Image (ML)", command=self.ai_processor.load_from_image_enhanced) #add this back once new function added
        file_menu.add_command(label="Load with AI (Claude)", command=self.ai_processor.load_from_image_with_ai)
        file_menu.add_command(label="Batch Process with AI", command=self.ai_processor.batch_process_with_ai)
        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Close", command=self.window.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        # Sample menu
        sample_menu = tk.Menu(menubar, tearoff=0)
        sample_menu.add_command(label="Add Sample", command=self.add_sample)
        sample_menu.add_command(label="Remove Sample", command=self.remove_sample)
        sample_menu.add_command(label="Clear Sample Data", command=self.clear_current_sample)
        sample_menu.add_separator()  # NEW
        sample_menu.add_command(label="Rename Current Sample", command=self.rename_current_sample)  # NEW
        sample_menu.add_command(label="Batch Rename Samples", command=self.batch_rename_samples)  # NEW
        menubar.add_cascade(label="Sample", menu=sample_menu)

        # Export menu
        export_menu = tk.Menu(menubar, tearoff=0)
        export_menu.add_command(label="Save Plot as Image", command=self.save_plot_as_image)
        export_menu.add_command(label="Save Table as Image", command=self.save_table_as_image)
        export_menu.add_separator()
        export_menu.add_command(label="Generate PowerPoint Report", command=self.generate_powerpoint_report)
        menubar.add_cascade(label="Export", menu=export_menu)

        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Show SOP", command=self.show_sop)
        view_menu.add_command(label="Refresh Plot", command=self.update_plot)
        menubar.add_cascade(label="View", menu=view_menu)

        ml_menu = tk.Menu(menubar, tearoff=0)

        # Core ML workflow - no training structure creation (separate scripts handle this)
        ml_menu.add_command(label="Check Enhanced Data Balance", command=self.ml_trainer.check_enhanced_data_balance)
        ml_menu.add_command(label="Train Enhanced Model", command=self.ml_trainer.train_enhanced_model)
        ml_menu.add_separator()

        # Testing and validation
        ml_menu.add_command(label="Test Enhanced Model", command=self.ml_trainer.test_enhanced_model)
        ml_menu.add_command(label="Validate Model Performance", command=self.ml_trainer.validate_enhanced_performance)
        ml_menu.add_separator()

        # Configuration management
        ml_menu.add_command(label="Update Processor Configuration", command=self.ml_trainer.update_processor_config)

        menubar.add_cascade(label="Enhanced ML", menu=ml_menu)

    def show_sop(self):
        """Display the Standard Operating Procedure."""
        sop_window = tk.Toplevel(self.window)
        sop_window.title("Sensory Evaluation SOP")
        sop_window.geometry("600x500")
        sop_window.configure(bg='white')

        text_widget = tk.Text(sop_window, wrap='word', font=('Arial', 11),
                             bg='white', fg='black', padx=20, pady=20)
        text_widget.pack(fill='both', expand=True)
        text_widget.insert('1.0', self.sop_text)
        text_widget.config(state='disabled')

        # Center the SOP window
        sop_window.update_idletasks()
        x = (sop_window.winfo_screenwidth() // 2) - (300)
        y = (sop_window.winfo_screenheight() // 2) - (250)
        sop_window.geometry(f"600x500+{x}+{y}")

    # Properties to maintain compatibility with existing code
    @property
    def sessions(self):
        """Access sessions through session manager."""
        return self.session_manager.sessions
    
    @property
    def current_session_id(self):
        """Access current session ID through session manager."""
        return self.session_manager.current_session_id
    
    @current_session_id.setter
    def current_session_id(self, value):
        """Set current session ID through session manager."""
        self.session_manager.current_session_id = value

    # Session manager delegations
    def add_new_session(self):
        return self.session_manager.add_new_session()

    def merge_sessions_from_files(self):
        return self.session_manager.merge_sessions_from_files()

    # File I/O delegations
    def load_session(self):
        return self.file_io.load_session()

    def save_session(self):
        return self.file_io.save_session()

    def export_to_excel(self):
        return self.file_io.export_to_excel()

    # Sample manager delegations
    def add_sample(self):
        return self.sample_manager.add_sample()

    def remove_sample(self):
        return self.sample_manager.remove_sample()

    def clear_current_sample(self):
        return self.sample_manager.clear_current_sample()

    def rename_current_sample(self):
        return self.sample_manager.rename_current_sample()

    def batch_rename_samples(self):
        return self.sample_manager.batch_rename_samples()

    def on_sample_changed(self, event=None):
        return self.sample_manager.on_sample_changed(event)

    def auto_save_and_update(self):
        return self.sample_manager.auto_save_and_update()

    # Export manager delegations
    def save_plot_as_image(self):
        return self.export_manager.save_plot_as_image()

    def save_table_as_image(self):
        return self.export_manager.save_table_as_image()

    def generate_powerpoint_report(self):
        return self.export_manager.generate_powerpoint_report()

    # Plot manager delegations
    def update_plot(self):
        return self.plot_manager.update_plot()

    # Mode manager delegations  
    def toggle_mode(self):
        return self.mode_manager.toggle_mode()

    # UI Layout delegations
    def setup_layout(self):
        return self.ui_layout.setup_layout()

    def center_window(self):
        return self.ui_layout.center_window()
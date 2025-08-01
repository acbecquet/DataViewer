"""
data_collection_window.py
Developed by Charlie Becquet.
Interface for rapid test data collection with enhanced saving and menu functionality.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import numpy as np
import os
import copy
import time
import openpyxl
import logging
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import statistics
from openpyxl.styles import PatternFill
from utils import FONT, debug_print, show_success_message
import threading
import subprocess

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("datacollection.log"),
        logging.StreamHandler()
    ]
)

# Standard Operating Procedures for each test
TEST_SOPS = {
    "User Test Simulation": """
SOP - User Test Simulation:

Day 1: Collect Initial Draw Resistance, TPM for 50 puffs, then use puff per day calculator to determine number of puffs. Make sure to enter initial oil mass.

Days 2-4: 10 puffs TPM after 

Day 5: Extended Test, every 20 puffs until device is empty

Day 6: Full Disassembly + Photos

Be sure to take detailed notes while conducting this test.

Key Points:
- Split testing: Phase 1 (0-50 puffs) and Phase 2 (extended puffs)
- No resistance measurements during extended testing
    """,
    
    "Intense Test": """
SOP - Intense Test:

Conduct intense testing at a draw pressure of 5kPa. Use either 200mL/3s/30s or 160mL/3s/30s regime.
Record TPM, draw pressure, and resistance at regular intervals.
Monitor for device overheating or failure.
Document any unusual observations in notes.

Key Points:
- 5kPa draw pressure target
- Standard 12-column data collection
- Record all measurements consistently
    """,

"User Simulation Test": """
SOP - User Test Simulation:

Day 1: Collect Initial Draw Resistance, TPM for 50 puffs, then use puff per day calculator to determine number of puffs. Make sure to enter initial oil mass.
Days 2-4: 10 puffs with TPM, then complete the daily number of puffs.
Day 5: Extended Test, every 10 puffs until device is empty
Day 6: Full Disassembly + Photos

Be sure to take detailed notes while conducting this test.

Key Points:
- Split testing: Phase 1 (0-50 puffs) and Phase 2 (extended puffs)
- No resistance measurements during extended testing
    """,
    
    "Big Headspace High Temperature": """
SOP - Big Headspace High Temperature Test:

Drain device to 30% remaining and then place in the oven at 40C.
After 1 hour, collect 10 puffs, tracking TPM and draw resistance for each puff. Repeat this 3 times, for 30 total puffs.
Be sure to take detailed notes on bubbling and any failure modes.

Key Points:
- 40C big headspace test
- Monitor for thermal effects
- Document temperature-related observations
    """,
    
    "Extended Test": """
SOP - Extended Test:

Long-duration testing to assess device lifetime and performance degradation.
Sessions of 10 puffs with a 60mL/3s/30s puffing regime. Rest 15 minutes between sessions.
Monitor for performance and consistency over time. Measure initial and final draw resistance, and measure TPM every 10 puffs.

Key Points:
- Extended duration testing
- Regular measurement intervals
- Performance degradation monitoring
- Comprehensive data collection
    """,
    
    "Quick Screening Test": """
SOP - Quick Screening Test:

Rapid assessment of device basic performance.
Focus on key performance indicators.
Streamlined testing for initial evaluation.

Key Points:
- Rapid testing protocol
- Key performance metrics only
- Initial device assessment
- Basic functionality verification
    """,
    
    "default": """
SOP - Standard Test Protocol:

Follow standard testing procedures for this test type.
Record all required measurements accurately.
Document any observations or anomalies.
Ensure proper test conditions throughout.

Key Points:
- Standard measurement protocol
- Accurate data recording
- Proper documentation
- Consistent test conditions
    """
}

class DataCollectionWindow:
    def __init__(self, parent, file_path, test_name, header_data, original_filename=None):
        """
        Initialize the data collection window.
    
        Args:
            parent (tk.Tk): The parent window.
            file_path (str): Path to the Excel file.
            test_name (str): Name of the test being conducted.
            header_data (dict): Dictionary containing header data.
            original_filename (str): original filename for .vap3 files
        """
        # Debug flag - set to False to disable all debug output
        self.DEBUG = False
        self.parent = parent
        self.root = parent.root
        self.file_path = file_path
        self.test_name = test_name
        self.header_data = header_data
        self.num_samples = header_data["num_samples"]
        self.result = None

        # Store the original filename for saving purposes
        self.original_filename = original_filename
        if self.original_filename:
            debug_print(f"DEBUG: DataCollectionWindow initialized with original filename: {self.original_filename}")
    

        # Validate num_samples
        if self.num_samples <= 0:
            self.num_samples = len(header_data.get('samples', []))
            if self.num_samples <= 0:
                self.num_samples = 1

        # Auto-save settings
        self.auto_save_interval = 5 * 60 * 1000  # 5 minutes in milliseconds
        self.auto_save_timer = None
        self.has_unsaved_changes = False
        self.last_save_time = None
    
        # Create the window
        self.window = tk.Toplevel(self.root)
        self.window.title(f"Data Collection - {test_name}")
        self.window.geometry("1375x750")
        self.window.minsize(1250, 625)
        
    
        self.window.transient(self.root)  # Make it a child of the main window
        self.window.grab_set()  # Make it modal and bring to front
        self.window.lift()  # Bring to top of window stack
        self.window.focus_force()  # Force focus to this window
    
        # Make it stay on top temporarily
        self.window.attributes('-topmost', True)
        self.window.after(100, lambda: self.window.attributes('-topmost', False)) 

        # Default puff interval
        self.puff_interval = 10  # Default to 10
    
        # Tracking variables for cell editing
        self.editing = False
        self.current_edit_widget = None
        self.current_edit = None
        self.current_item = None
        self.current_column = None
    
        # Set up keyboard shortcut flags
        self.hotkeys_enabled = True
        self.hotkey_bindings = {}
    
        # Create the style for ttk widgets
        self.style = ttk.Style()
        self.setup_styles()

        # Click tracking variables
        # Simplified click tracking
        self.last_click_time = 0
        self.last_clicked_item = None
        self.last_clicked_column = None


        # Initialize data structures
        self.initialize_data()
    
        # Create the menu bar first
        self.create_menu_bar()
    
        # Create the UI
        self.create_widgets()

        self.load_existing_data_from_file()
    
        # Center the window
        self.center_window()
    
        # Set up event handlers
        self.setup_event_handlers()
    
        # Start auto-save timer
        self.start_auto_save_timer()

        self.ensure_initial_tpm_calculation()
    
        self.log(f"DataCollectionWindow initialized for {test_name} with {self.num_samples} samples", "debug")

    def create_menu_bar(self):
        """Create a comprehensive menu bar for the data collection window."""
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
    
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Save", accelerator="Ctrl+S", 
                            command=lambda: self.save_data(show_confirmation=False))
        file_menu.add_command(label="Save and Continue", 
                            command=lambda: self.save_data(show_confirmation=True))
        file_menu.add_command(label="Save and Exit", 
                            command=lambda: self.save_data(exit_after=True))
        file_menu.add_separator()
        file_menu.add_separator()
        file_menu.add_command(label="Export CSV", command=self.export_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit without Saving", command=self.exit_without_saving)
        menubar.add_cascade(label="File", menu=file_menu)
    
        # Edit menu  
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Edit Headers", command=self.edit_headers)
        edit_menu.add_command(label="Clear Current Sample", command=self.clear_current_sample)
        edit_menu.add_command(label="Clear All Data", command=self.clear_all_data)
        edit_menu.add_separator()
        edit_menu.add_command(label="Add Row", command=self.add_row)
        edit_menu.add_command(label="Remove Last Row", command=self.remove_last_row)
        edit_menu.add_separator()
        edit_menu.add_command(label="Recalculate TPM", command=self.recalculate_all_tpm)
        menubar.add_cascade(label="Edit", menu=edit_menu)
    
        # Navigate menu
        navigate_menu = tk.Menu(menubar, tearoff=0)
        navigate_menu.add_command(label="Previous Sample", accelerator="Ctrl+Left", 
                                command=self.go_to_previous_sample)
        navigate_menu.add_command(label="Next Sample", accelerator="Ctrl+Right", 
                                command=self.go_to_next_sample)
        navigate_menu.add_separator()
        navigate_menu.add_command(label="Go to Sample...", command=self.go_to_sample_dialog)
        menubar.add_cascade(label="Navigate", menu=navigate_menu)
    
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Change Puff Interval", command=self.change_puff_interval_dialog)
        tools_menu.add_command(label="Auto-Save Settings", command=self.auto_save_settings_dialog)
        tools_menu.add_separator()
        menubar.add_cascade(label="Tools", menu=tools_menu)
    
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_keyboard_shortcuts)
        help_menu.add_command(label="About Data Collection", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
    
        self.log("Menu bar created successfully", "debug")
    
    def setup_styles(self):
        """Set up styles for ttk widgets with enhanced visual separation."""
        # Configure ttk styles to use system defaults
        self.style.configure('TFrame', background='SystemButtonFace')
        self.style.configure('TLabel', background='SystemButtonFace')
        self.style.configure('TLabelframe', background='SystemButtonFace')
        self.style.configure('TLabelframe.Label', background='SystemButtonFace')
        self.style.configure('TNotebook', background='SystemButtonFace')
        self.style.configure('TNotebook.Tab', background='SystemButtonFace')

        # Create special styles for headers
        self.style.configure('Header.TLabel', font=("Arial", 14, "bold"), background='SystemButtonFace')
        self.style.configure('SubHeader.TLabel', font=("Arial", 12), background='SystemButtonFace')
        self.style.configure('Stats.TLabel', font=("Arial", 14, "bold"), background='SystemButtonFace')
        self.style.configure('SampleInfo.TLabel', font=("Arial", 11), background='SystemButtonFace')
                     
        # Enhanced Treeview styling with better visual separation
        self.style.configure('Treeview', 
                            background='white',
                            fieldbackground='white',
                            borderwidth=1,
                            rowheight=28,  # Taller rows for better separation
                            font=('Arial', 10))
    
        # Configure selection colors
        self.style.map('Treeview', 
                      background=[('selected', '#E8F4FD')],
                      foreground=[('selected', 'black')])

        # Style headers with enhanced borders
        self.style.configure('Treeview.Heading',
                            background='#D0D0D0',
                            foreground='black',
                            font=('Arial', 10, 'bold'),
                            borderwidth=2,
                            relief='raised')
    
        # Add hover effect
        self.style.map('Treeview.Heading',
                      background=[('active', '#C0C0C0')])

        debug_print("DEBUG: Enhanced styles configured with better visual separation")
    
    def center_window(self):
        """Center the window on the screen."""
        # Force the window to update and calculate its actual size
        self.window.update()
    
        # Get the actual window dimensions
        width = self.window.winfo_width()
        height = self.window.winfo_height()
    
        # If width/height are still too small, use the requested size
        if width < 100 or height < 100:
            self.window.update_idletasks()
            width = self.window.winfo_reqwidth()
            height = self.window.winfo_reqheight()
    
        # Calculate center position
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
    
        # Ensure the window isn't positioned off-screen
        x = max(0, min(x, screen_width - width))
        y = max(0, min(y, screen_height - height))
    
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
        debug_print(f"DEBUG: Window centered at {x},{y} with size {width}x{height}")
    
    def on_canvas_resize(self, event=None):
        """Handle canvas resize events to update plot size."""
        if hasattr(self, 'tpm_canvas') and hasattr(self, 'tpm_figure'):
            # Get the current canvas size
            canvas_widget = self.tpm_canvas.get_tk_widget()
            width = canvas_widget.winfo_width()
            height = canvas_widget.winfo_height()
        
            if width > 1 and height > 1:  # Valid size
                # Convert to inches at 80dpi
                fig_width = width / 80.0
                fig_height = height / 80.0
            
                # Update figure size
                self.tpm_figure.set_size_inches(fig_width, fig_height)
                self.tpm_canvas.draw()

    def create_widgets(self):
        """Create the data collection UI with a cleaner structure."""
        # Configure window
        self.window.configure(background='SystemButtonFace')
    
        # Create a main frame with padding
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)
    
        # Create header section
        self.create_header_section(main_frame)
    
        # Create SOP section
        self.create_sop_section(main_frame)
    
        # Create main content area with horizontal split
        content_frame = ttk.PanedWindow(main_frame, orient="horizontal")
        content_frame.pack(fill="both", expand=True)
    
        # Create data entry area (left side)
        data_frame = ttk.Frame(content_frame)
        content_frame.add(data_frame, weight=3)  # 75% width
    
        # Create stats panel (right side)
        self.stats_frame = ttk.Frame(content_frame)
        content_frame.add(self.stats_frame, weight=1)  # 25% width
    
        # Create notebook with sample tabs
        self.notebook = ttk.Notebook(data_frame)
        self.notebook.pack(fill="both", expand=True)
    
        # Create sample tabs
        self.create_sample_tabs()
    
        # Create stats panel
        self.create_tpm_stats_panel()
    
        # Create control buttons at bottom
        self.create_control_buttons(main_frame)
    
        # Bind notebook tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self.update_stats_panel)
   
    def log(self, message, level="info"):
        """Log a message with the specified level."""
        logger = logging.getLogger("DataCollectionWindow")
    
        if level.lower() == "debug":
            logger.debug(message)
        elif level.lower() == "info":
            logger.info(message)
        elif level.lower() == "warning":
            logger.warning(message)
        elif level.lower() == "error":
            logger.error(message)
        elif level.lower() == "critical":
            logger.critical(message)

    def create_sop_section(self, parent_frame):
        """Create the SOP (Standard Operating Procedure) section."""
        
        # Create a collapsible SOP frame
        sop_frame = ttk.LabelFrame(parent_frame, text="", style='TLabelframe')
        sop_frame.pack(fill="x", pady=(0, 10))
    
        # Create a header frame for button and title
        header_frame = ttk.Frame(sop_frame, style='TFrame')
        header_frame.pack(fill="x", padx=5, pady=5)
    
        # Add small toggle button on the left with minimal size
        self.sop_visible = False
        self.toggle_sop_button = ttk.Button(
            header_frame,
            text="Show",
            width=6,  # Even smaller width
            command=self.toggle_sop_visibility
        )
        self.toggle_sop_button.pack(side="left", padx=(0, 5))  # Small gap between button and title
    
        # Add title label after the button
        title_label = ttk.Label(header_frame, text="Standard Operating Procedure (SOP)", 
                               font=("Arial", 10, "bold"), style='TLabel')
        title_label.pack(side="left")
    
        # Get the SOP text for the current test
        sop_text = TEST_SOPS.get(self.test_name, TEST_SOPS["default"])
    
        # Create a text widget to display the SOP with centered text
        self.sop_text_widget = tk.Text(
            sop_frame, 
            height=6,  
            wrap=tk.WORD,
            font=("Arial", 9),
            bg="SystemButtonFace",
            fg="black",
            relief="flat",
            borderwidth=0,
            padx=10,
            pady=5
        )

        # Insert the SOP text and center it
        self.sop_text_widget.insert("1.0", sop_text.strip())
    
        # Center the text
        self.sop_text_widget.tag_configure("center", justify='center')
        self.sop_text_widget.tag_add("center", "1.0", "end")
    
        # Make it read-only
        self.sop_text_widget.config(state="disabled")
    
        debug_print("DEBUG: SOP section created successfully")

    def toggle_sop_visibility(self):
        """Toggle the visibility of the SOP text."""
        if self.sop_visible:
            # Hide the SOP
            self.sop_text_widget.pack_forget()
            self.toggle_sop_button.config(text="Show")
            self.sop_visible = False
            debug_print("DEBUG: SOP section hidden")
        else:
            # Show the SOP
            self.sop_text_widget.pack(fill="x", padx=10, pady=(0, 5))
            self.toggle_sop_button.config(text="Hide")
            self.sop_visible = True
            debug_print("DEBUG: SOP section shown")

    # Auto-save functionality
    def start_auto_save_timer(self):
        """Start the auto-save timer."""
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)
        
        self.auto_save_timer = self.window.after(self.auto_save_interval, self.auto_save)
        
    def auto_save(self):
        """Perform automatic save without user confirmation."""
        if self.has_unsaved_changes:            
            try:
                self.save_data(show_confirmation=False, auto_save=True)
                self.update_save_status(False)  # Mark as saved
                debug_print("DEBUG: Auto-save completed successfully")
            except Exception as e:
                debug_print(f"DEBUG: Auto-save failed: {e}")
                # Continue even if auto-save fails
        
        # Restart the timer
        self.start_auto_save_timer()
    
    def mark_unsaved_changes(self):
        """Mark that there are unsaved changes."""
        if not self.has_unsaved_changes:
            self.has_unsaved_changes = True
            self.update_save_status(True)
            
    def update_save_status(self, has_changes):
        """Update the save status indicator in the UI."""
        if has_changes:
            self.save_status_label.config(foreground="red", text="●")
            self.save_status_text.config(text="Unsaved changes")
        else:
            self.save_status_label.config(foreground="green", text="●")
            self.save_status_text.config(text="All changes saved")
            import datetime
            self.last_save_time = datetime.datetime.now()
    
    def exit_without_saving(self):
        """Exit without saving (with confirmation)."""
        if self.has_unsaved_changes:
            if not messagebox.askyesno("Confirm Exit", 
                                     "You have unsaved changes. Are you sure you want to exit without saving?"):
                return
        
        debug_print("DEBUG: Exiting without saving")
        self.result = "cancel"
        self.window.destroy()
    
    def export_csv(self):
        """Export current data to CSV files."""
        
        from tkinter import filedialog
        
        # Ask for directory to save CSV files
        directory = filedialog.askdirectory(title="Select directory to save CSV files")
        if not directory:
            return
        
        try:
            for i in range(self.num_samples):
                sample_id = f"Sample {i+1}"
                sample_name = self.header_data["samples"][i]["id"]
                
                # Create DataFrame for this sample
                df_data = {
                    "Puffs": self.data[sample_id]["puffs"],
                    "Before Weight (g)": self.data[sample_id]["before_weight"],
                    "After Weight (g)": self.data[sample_id]["after_weight"],
                    "Draw Pressure (kPa)": self.data[sample_id]["draw_pressure"],
                    "Smell": self.data[sample_id]["smell"],
                    "Notes": self.data[sample_id]["notes"],
                    "TPM (mg/puff)": self.data[sample_id]["tpm"]
                }
                
                df = pd.DataFrame(df_data)
                
                # Remove empty rows
                df = df.dropna(how='all', subset=["Before Weight (g)", "After Weight (g)"])
                
                csv_filename = f"{self.test_name}_{sample_name}_data.csv"
                csv_path = os.path.join(directory, csv_filename)
                df.to_csv(csv_path, index=False)
                
            show_success_message("Export Complete", f"Exported {self.num_samples} CSV files to {directory}", self.window)
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV files: {e}")
    
    def edit_headers(self):
        """Open header editing dialog from within data collection."""
        debug_print("DEBUG: Edit Headers requested from data collection window")
        
        # Show the header data dialog in edit mode
        from header_data_dialog import HeaderDataDialog
        header_dialog = HeaderDataDialog(
            self.window, 
            self.file_path, 
            self.test_name, 
            edit_mode=True,
            current_data=self.header_data
        )
        result, new_header_data = header_dialog.show()
        
        if result:
            debug_print("DEBUG: Header data updated successfully")
            # Update our internal header data
            old_header_data = self.header_data.copy()
            self.header_data = new_header_data
            
            # Check if number of samples changed
            old_num_samples = old_header_data.get("num_samples", 0)
            new_num_samples = new_header_data.get("num_samples", 0)
            
            if old_num_samples != new_num_samples:
                debug_print(f"DEBUG: Number of samples changed from {old_num_samples} to {new_num_samples}")
                # Handle sample count change
                self.handle_sample_count_change(old_num_samples, new_num_samples)
            
            # Update the header display in all tabs
            self.update_header_display()
            
            # Apply changes to the Excel file
            self.apply_header_changes_to_file()
            
            # Mark as having unsaved changes
            self.mark_unsaved_changes()
            
            # Show success message
            self.show_temp_status_message("Headers updated successfully", 3000)
        else:
            debug_print("DEBUG: Header editing was cancelled")

    def handle_sample_count_change(self, old_count, new_count):
        """Handle changes in the number of samples."""
        debug_print(f"DEBUG: Handling sample count change: {old_count} -> {new_count}")
    
        if new_count > old_count:
            # Add new samples
            for i in range(old_count, new_count):
                sample_id = f"Sample {i+1}"
                self.data[sample_id] = {
                    "puffs": [],
                    "before_weight": [],
                    "after_weight": [],
                    "draw_pressure": [],
                    "smell": [],
                    "notes": [],
                    "tpm": [],
                    "current_row_index": 0,
                    "avg_tpm": 0.0
                }
            
                # Add chronography for User Test Simulation
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    self.data[sample_id]["chronography"] = []
            
                # Pre-initialize 50 rows for new sample
                for j in range(50):
                    puff = (j + 1) * self.puff_interval
                    self.data[sample_id]["puffs"].append(puff)
                    self.data[sample_id]["before_weight"].append("")
                    self.data[sample_id]["after_weight"].append("")
                    self.data[sample_id]["draw_pressure"].append("")
                    self.data[sample_id]["smell"].append("")
                    self.data[sample_id]["notes"].append("")
                    self.data[sample_id]["tpm"].append(None)
                
                    # Add chronography initialization for User Test Simulation
                    if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                        self.data[sample_id]["chronography"].append("")
        
            debug_print(f"DEBUG: Added {new_count - old_count} new samples")
        
        elif new_count < old_count:
            # Remove excess samples
            for i in range(new_count, old_count):
                sample_id = f"Sample {i+1}"
                if sample_id in self.data:
                    del self.data[sample_id]
        
            debug_print(f"DEBUG: Removed {old_count - new_count} samples")
    
        # Update the number of samples
        self.num_samples = new_count
    
        # Recreate the UI to reflect the new sample count
        self.recreate_sample_tabs()

    def recreate_sample_tabs(self):
        """Recreate all sample tabs when sample count changes."""
        debug_print("DEBUG: Recreating sample tabs")
        
        # Clear existing tabs
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        
        # Clear the lists
        self.sample_frames.clear()
        self.sample_trees.clear()
        
        # Recreate tabs for the new number of samples
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            sample_frame = ttk.Frame(self.notebook, padding=10, style='TFrame')
            
            # Get sample name safely
            sample_name = "New Sample"
            if i < len(self.header_data.get('samples', [])):
                sample_name = self.header_data['samples'][i].get('id', f"Sample {i+1}")
            
            self.notebook.add(sample_frame, text=f"Sample {i+1} - {sample_name}")
            self.sample_frames.append(sample_frame)
            
            # Create the sample tab content
            tree = self.create_sample_tab(sample_frame, sample_id, i)
            self.sample_trees.append(tree)
        
        # Recreate the TPM stats panel
        self.create_tpm_stats_panel()
        
        debug_print(f"DEBUG: Recreated {self.num_samples} sample tabs")

    def update_header_display(self):
        """Update the header information displayed in the UI."""
        
        # Update the header labels (if they exist)
        try:
            # Update any header text that shows tester name, etc.
            for widget in self.window.winfo_children():
                if hasattr(widget, 'winfo_children'):
                    self.update_header_labels_recursive(widget)
            
            # Update notebook tab labels
            for i in range(min(self.num_samples, len(self.sample_frames))):
                if i < len(self.header_data.get('samples', [])):
                    sample_name = self.header_data['samples'][i].get('id', f"Sample {i+1}")
                    self.notebook.tab(i, text=f"Sample {i+1} - {sample_name}")
                    
        except Exception as e:
            debug_print(f"DEBUG: Error updating header display: {e}")

    def update_header_labels_recursive(self, widget):
        """Recursively update header labels in the widget tree."""
        try:
            if isinstance(widget, ttk.Label):
                text = widget.cget('text')
                if 'Tester:' in text:
                    new_text = f"Tester: {self.header_data['common']['tester']}"
                    widget.config(text=new_text)
            
            # Recurse through children
            for child in widget.winfo_children():
                self.update_header_labels_recursive(child)
                
        except Exception as e:
            debug_print(f"DEBUG: Error updating label: {e}")

    def apply_header_changes_to_file(self):
        """Apply header changes to the Excel file."""
        debug_print("DEBUG: Applying header changes to Excel file")
        try:
            import openpyxl
            wb = openpyxl.load_workbook(self.file_path)
            
            if self.test_name not in wb.sheetnames:
                debug_print(f"DEBUG: Sheet {self.test_name} not found")
                return
                
            ws = wb[self.test_name]
            
            # Apply common data
            common_data = self.header_data.get('common', {})
            ws.cell(row=2, column=2, value=common_data.get('media', ''))
            ws.cell(row=3, column=2, value=common_data.get('viscosity', ''))
            ws.cell(row=2, column=6, value=common_data.get('voltage', ''))
            ws.cell(row=2, column=8, value=common_data.get('oil_mass', ''))
            
            # Apply sample-specific data
            for i, sample_data in enumerate(self.header_data.get('samples', [])):
                col_offset = i * 12
                ws.cell(row=1, column=6 + col_offset, value=sample_data.get('id', ''))
                ws.cell(row=3, column=4 + col_offset, value=sample_data.get('resistance', ''))
            
            # Save the workbook
            wb.save(self.file_path)
            debug_print("DEBUG: Header changes applied to Excel file")
            
        except Exception as e:
            debug_print(f"DEBUG: Error applying header changes to file: {e}")

    def show_temp_status_message(self, message, duration_ms):
        """Show a temporary status message."""
        # Create a temporary label
        temp_label = ttk.Label(
            self.window,
            text=message,
            font=("Arial", 10),
            foreground="green",
            background="SystemButtonFace"
        )
        temp_label.pack(side="bottom", fill="x")
        
        # Remove it after the specified duration
        self.window.after(duration_ms, lambda: temp_label.destroy() if temp_label.winfo_exists() else None)

    def clear_current_sample(self):
        """Clear data for the currently selected sample."""
        current_tab = self.notebook.index(self.notebook.select())
        sample_id = f"Sample {current_tab + 1}"

        if messagebox.askyesno("Confirm Clear", f"Are you sure you want to clear all data for {sample_id}?"):
            # Clear all data except puffs
            clear_keys = ["before_weight", "after_weight", "draw_pressure", "smell", "notes", "tpm"]
        
            # Add chronography for User Test Simulation
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                clear_keys.append("chronography")
            else:
                # Add resistance and clog for standard tests
                clear_keys.extend(["resistance", "clog"])
        
            for key in clear_keys:
                for i in range(len(self.data[sample_id][key])):
                    if key == "tpm":
                        self.data[sample_id][key][i] = None
                    else:
                        self.data[sample_id][key][i] = ""
        
            # Update the display
            tree = self.sample_trees[current_tab]
            self.update_treeview(tree, sample_id)
            self.update_stats_panel()
            self.mark_unsaved_changes()
        
            debug_print(f"DEBUG: Cleared data for {sample_id}")
    
    def clear_all_data(self):
        """Clear all data for all samples."""
        if messagebox.askyesno("Confirm Clear All", 
                             "Are you sure you want to clear ALL data for ALL samples?"):
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                clear_keys = ["before_weight", "after_weight", "draw_pressure", "smell", "notes", "tpm"]
            
                # Add chronography for User Test Simulation
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    clear_keys.append("chronography")
            
                for key in clear_keys:
                    for j in range(len(self.data[sample_id][key])):
                        if key == "tpm":
                            self.data[sample_id][key][j] = None
                        else:
                            self.data[sample_id][key][j] = ""
            
                # Update the display for this sample
                tree = self.sample_trees[i]
                self.update_treeview(tree, sample_id)
        
            self.update_stats_panel()
            self.mark_unsaved_changes()
            debug_print("DEBUG: Cleared all data for all samples")
    
    def add_row(self):
        """Add a new row to all samples."""
        new_puff = max([max(self.data[f"Sample {i+1}"]["puffs"]) for i in range(self.num_samples)]) + self.puff_interval
    
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"
            self.data[sample_id]["puffs"].append(new_puff)
            self.data[sample_id]["before_weight"].append("")
            self.data[sample_id]["after_weight"].append("")
            self.data[sample_id]["draw_pressure"].append("")
            self.data[sample_id]["smell"].append("")
            self.data[sample_id]["notes"].append("")
            self.data[sample_id]["tpm"].append(None)
        
            # Add chronography for User Test Simulation
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                self.data[sample_id]["chronography"].append("")
        
            # Update treeview
            tree = self.sample_trees[i]
            self.update_treeview(tree, sample_id)
    
        self.mark_unsaved_changes()
        debug_print(f"DEBUG: Added new row with puff count {new_puff}")
    
    def remove_last_row(self):
        """Remove the last row from all samples."""
        if messagebox.askyesno("Confirm Remove", "Remove the last row from all samples?"):
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                for key in self.data[sample_id]:
                    if self.data[sample_id][key]:  # Only remove if list is not empty
                        self.data[sample_id][key].pop()
                
                # Update treeview
                tree = self.sample_trees[i]
                self.update_treeview(tree, sample_id)
            
            self.update_stats_panel()
            self.mark_unsaved_changes()
            debug_print("DEBUG: Removed last row from all samples")
    
    def recalculate_all_tpm(self):
        """Recalculate TPM for all samples."""
        debug_print("DEBUG: Recalculating TPM for all samples")
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"
            self.calculate_tpm(sample_id)
        
        self.update_stats_panel()
        self.mark_unsaved_changes()
        show_success_message("Recalculation Complete", "TPM values have been recalculated for all samples.", self.window)
    
    def go_to_sample_dialog(self):
        """Show dialog to jump to a specific sample."""
        sample_names = [f"Sample {i+1} - {self.header_data['samples'][i]['id']}" for i in range(self.num_samples)]
        
        dialog = tk.Toplevel(self.window)
        dialog.title("Go to Sample")
        dialog.geometry("300x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Select sample:").pack(pady=10)
        
        selected_sample = tk.StringVar()
        combo = ttk.Combobox(dialog, textvariable=selected_sample, values=sample_names, state="readonly")
        combo.pack(pady=5)
        combo.set(sample_names[0])
        
        def go_to_selected():
            index = sample_names.index(selected_sample.get())
            self.notebook.select(index)
            dialog.destroy()
        
        ttk.Button(dialog, text="Go", command=go_to_selected).pack(pady=10)
    
    def change_puff_interval_dialog(self):
        """Show dialog to change the puff interval."""
        new_interval = simpledialog.askinteger(
            "Change Puff Interval",
            f"Current puff interval: {self.puff_interval}\nEnter new interval:",
            initialvalue=self.puff_interval,
            minvalue=1,
            maxvalue=1000
        )
        
        if new_interval and new_interval != self.puff_interval:
            self.puff_interval = new_interval
            self.puff_interval_var.set(new_interval)
            
            # Update puff values for future rows
            debug_print(f"DEBUG: Changed puff interval to {new_interval}")
    
    def auto_save_settings_dialog(self):
        """Show dialog to configure auto-save settings."""
        current_minutes = self.auto_save_interval / 60 / 1000
        
        new_minutes = simpledialog.askfloat(
            "Auto-Save Settings",
            f"Current auto-save interval: {current_minutes} minutes\nEnter new interval (minutes):",
            initialvalue=current_minutes,
            minvalue=0.5,
            maxvalue=60
        )
        
        if new_minutes:
            self.auto_save_interval = int(new_minutes * 60 * 1000)
            self.start_auto_save_timer()  # Restart with new interval
            debug_print(f"DEBUG: Changed auto-save interval to {new_minutes} minutes")
    
    def show_keyboard_shortcuts(self):
        """Show keyboard shortcuts help."""
        shortcuts = """
Keyboard Shortcuts:

Navigation:
• Tab - Move to next cell
• Shift+Tab - Move to previous cell
• Arrow Keys - Navigate between cells
• Ctrl+Left/Right - Switch samples
• Enter - Move down one row

Editing:
• Double-click - Edit cell
• Type - Start editing selected cell
• Enter - Confirm edit and move down
• Escape - Cancel edit

File Operations:
• Ctrl+S - Quick save

General:
• F1 - Show this help
        """
        
        help_window = tk.Toplevel(self.window)
        help_window.title("Keyboard Shortcuts")
        help_window.geometry("400x500")
        help_window.transient(self.window)
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill="both", expand=True)
        text_widget.insert("1.0", shortcuts)
        text_widget.config(state="disabled")
    
    def show_about(self):
        """Show about dialog."""
        about_text = f"""
Data Collection Window
Version 3.0

Test: {self.test_name}
Samples: {self.num_samples}
Auto-save: Every {self.auto_save_interval/60/1000:.1f} minutes

Features:
• Real-time TPM calculation
• Auto-save functionality
• Excel and CSV export
• Keyboard navigation
• Comprehensive data validation

Developed by Charlie Becquet
        """
        show_success_message("About Data Collection", about_text, self.window)
    
    def _save_to_excel(self):
        """Save data to the appropriate file format."""
        debug_print(f"DEBUG: _save_to_excel() starting - file: {self.file_path}")
    
        # Check if this is a .vap3 file or temporary file
        if self.file_path.endswith('.vap3') or not os.path.exists(self.file_path):
            debug_print("DEBUG: Detected .vap3 file or non-existent file, saving to loaded sheets")
            self._save_to_loaded_sheets()
        else:
            debug_print("DEBUG: Detected Excel file, saving using openpyxl")
            self._save_to_excel_file()

        debug_print("DEBUG: Excel save completed, updating main GUI...")

        # For .vap3 files, the data is already updated in memory via _save_to_loaded_sheets
        if self.file_path.endswith('.vap3') or not os.path.exists(self.file_path):
            debug_print("DEBUG: .vap3 file detected, skipping file-based update")
            # The _save_to_loaded_sheets method should have already updated the main GUI data
        else:
            debug_print("DEBUG: Excel file detected, updating from file")
            self._update_excel_data_in_main_gui()

    
    def _save_to_excel_file(self):
        """Save data to the Excel file."""
        debug_print(f"DEBUG: _save_to_excel() starting - file: {self.file_path}")

        # Load the workbook
        wb = openpyxl.load_workbook(self.file_path)
        debug_print(f"DEBUG: Loaded workbook, sheets: {wb.sheetnames}")

        # Get the sheet for this test
        if self.test_name not in wb.sheetnames:
            raise Exception(f"Sheet '{self.test_name}' not found in the file.")
    
        ws = wb[self.test_name]
        debug_print(f"DEBUG: Opened sheet '{self.test_name}'")

        # Define green fill for TPM cells
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

        # Determine column layout based on test type
        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
            columns_per_sample = 8  # Including chronography
            debug_print(f"DEBUG: Using User Test Simulation format with 8 columns per sample")
        else:
            columns_per_sample = 12  # Standard format
            debug_print(f"DEBUG: Using standard format with 12 columns per sample")

        # Track how much data we're actually writing
        total_data_written = 0

        # For each sample, write the data
        for sample_idx in range(self.num_samples):
            sample_id = f"Sample {sample_idx+1}"
    
            # Calculate column offset based on test type
            col_offset = sample_idx * columns_per_sample
    
            debug_print(f"DEBUG: Writing data for {sample_id} at column offset {col_offset}")
    
            sample_data_written = 0
    
            # Write the data starting at row 5
            for i, puff in enumerate(self.data[sample_id]["puffs"]):
                row = i + 5  # Row 5 is the first data row
        
                # Only write if we have actual data (not just empty rows)
                has_data = (
                    self.data[sample_id]["before_weight"][i] or
                    self.data[sample_id]["after_weight"][i] or
                    self.data[sample_id]["draw_pressure"][i] or
                    self.data[sample_id]["smell"][i] or
                    self.data[sample_id]["notes"][i] or
                    self.data[sample_id]["tpm"][i] is not None
                )
            
                # For User Test Simulation, also check chronography
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    has_data = has_data or (i < len(self.data[sample_id]["chronography"]) and self.data[sample_id]["chronography"][i])
        
                if not has_data:
                    continue  # Skip empty rows
            
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    # User Test Simulation column layout (8 columns)
                    # Chronography column (A + offset)
                    if i < len(self.data[sample_id]["chronography"]) and self.data[sample_id]["chronography"][i]:
                        ws.cell(row=row, column=1 + col_offset, value=str(self.data[sample_id]["chronography"][i]))
                
                    # Puffs column (B + offset) 
                    ws.cell(row=row, column=2 + col_offset, value=puff)
                
                    # Before weight column (C + offset)
                    if self.data[sample_id]["before_weight"][i]:
                        try:
                            ws.cell(row=row, column=3 + col_offset, value=float(self.data[sample_id]["before_weight"][i]))
                        except:
                            ws.cell(row=row, column=3 + col_offset, value=self.data[sample_id]["before_weight"][i])
                
                    # After weight column (D + offset)
                    if self.data[sample_id]["after_weight"][i]:
                        try:
                            ws.cell(row=row, column=4 + col_offset, value=float(self.data[sample_id]["after_weight"][i]))
                        except:
                            ws.cell(row=row, column=4 + col_offset, value=self.data[sample_id]["after_weight"][i])
                
                    # Draw pressure column (E + offset)
                    if self.data[sample_id]["draw_pressure"][i]:
                        try:
                            ws.cell(row=row, column=5 + col_offset, value=float(self.data[sample_id]["draw_pressure"][i]))
                        except:
                            ws.cell(row=row, column=5 + col_offset, value=self.data[sample_id]["draw_pressure"][i])
                
                    # Skip resistance column (F + offset) - not used in User Test Simulation
                
                    # Failure column (G + offset)
                    if self.data[sample_id]["smell"][i]:
                        try:
                            ws.cell(row=row, column=6 + col_offset, value=float(self.data[sample_id]["smell"][i]))
                        except:
                            ws.cell(row=row, column=6 + col_offset, value=self.data[sample_id]["smell"][i])
                
                    # Notes column (H + offset)
                    if self.data[sample_id]["notes"][i]:
                        ws.cell(row=row, column=7 + col_offset, value=str(self.data[sample_id]["notes"][i]))
                    
                    debug_print(f"DEBUG: Saved User Test Simulation row {i} for {sample_id}")
                
                else:
                    # Standard column layout (12 columns)
                    # Puffs column (A + offset)
                    ws.cell(row=row, column=1 + col_offset, value=puff)
        
                    # Before weight column (B + offset)
                    if self.data[sample_id]["before_weight"][i]:
                        try:
                            ws.cell(row=row, column=2 + col_offset, value=float(self.data[sample_id]["before_weight"][i]))
                        except:
                            ws.cell(row=row, column=2 + col_offset, value=self.data[sample_id]["before_weight"][i])
        
                    # After weight column (C + offset)
                    if self.data[sample_id]["after_weight"][i]:
                        try:
                            ws.cell(row=row, column=3 + col_offset, value=float(self.data[sample_id]["after_weight"][i]))
                        except:
                            ws.cell(row=row, column=3 + col_offset, value=self.data[sample_id]["after_weight"][i])
        
                    # Draw pressure column (D + offset)
                    if self.data[sample_id]["draw_pressure"][i]:
                        try:
                            ws.cell(row=row, column=4 + col_offset, value=float(self.data[sample_id]["draw_pressure"][i]))
                        except:
                            ws.cell(row=row, column=4 + col_offset, value=self.data[sample_id]["draw_pressure"][i])
        
                    if self.data[sample_id]["resistance"][i]:
                        try:
                            ws.cell(row=row, column=5 + col_offset, value=float(self.data[sample_id]["resistance"][i]))
                        except:
                            ws.cell(row=row, column=5 + col_offset, value=self.data[sample_id]["resistance"][i])

                    # Smell column (F + offset)
                    if self.data[sample_id]["smell"][i]:
                        try:
                            ws.cell(row=row, column=6 + col_offset, value=float(self.data[sample_id]["smell"][i]))
                        except:
                            ws.cell(row=row, column=6 + col_offset, value=self.data[sample_id]["smell"][i])

                    if self.data[sample_id]["clog"][i]:
                        try:
                            ws.cell(row=row, column=7 + col_offset, value=float(self.data[sample_id]["clog"][i]))
                        except:
                            ws.cell(row=row, column=7 + col_offset, value=self.data[sample_id]["clog"][i])
        
                    # Notes column (H + offset)
                    if self.data[sample_id]["notes"][i]:
                        ws.cell(row=row, column=8 + col_offset, value=str(self.data[sample_id]["notes"][i]))
        
                    # TPM column (I + offset) - if calculated
                    if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None:
                        tpm_cell = ws.cell(row=row, column=9 + col_offset, value=float(self.data[sample_id]["tpm"][i]))
                        tpm_cell.fill = green_fill
        
                sample_data_written += 1
        
            total_data_written += sample_data_written
            debug_print(f"DEBUG: Wrote {sample_data_written} data rows for {sample_id}")

        # Save the workbook
        debug_print(f"DEBUG: Saving workbook with {total_data_written} total data rows written...")

        debug_print("DEBUG: About to save Excel file - verifying data to be saved:")
        for sample_idx in range(min(2, self.num_samples)):  # Check first 2 samples
            sample_id = f"Sample {sample_idx+1}"
            col_offset = sample_idx * columns_per_sample
        
            debug_print(f"DEBUG: Sample {sample_idx+1} data preview:")
            for i in range(min(3, len(self.data[sample_id]["puffs"]))):  # First 3 rows
                row = i + 5
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    chrono_val = ws.cell(row=row, column=1 + col_offset).value
                    puff_val = ws.cell(row=row, column=2 + col_offset).value
                    before_val = ws.cell(row=row, column=3 + col_offset).value
                    after_val = ws.cell(row=row, column=4 + col_offset).value
                    debug_print(f"DEBUG:   Row {i}: Chrono={chrono_val}, Puff={puff_val}, Before={before_val}, After={after_val}")
                else:
                    puff_val = ws.cell(row=row, column=1 + col_offset).value
                    before_val = ws.cell(row=row, column=2 + col_offset).value
                    after_val = ws.cell(row=row, column=3 + col_offset).value
                    tpm_val = ws.cell(row=row, column=9 + col_offset).value
                    debug_print(f"DEBUG:   Row {i}: Puff={puff_val}, Before={before_val}, After={after_val}, TPM={tpm_val}")

        wb.save(self.file_path)
        debug_print(f"DEBUG: Excel file saved successfully to {self.file_path}")
    
    def _save_to_loaded_sheets(self):
        """Save data to the loaded sheets in memory."""
        debug_print("DEBUG: _save_to_loaded_sheets() starting")
    
        # Use original filename if available, otherwise fall back to file_path
        display_filename = self.original_filename if self.original_filename else os.path.basename(self.file_path)
    
        debug_print(f"DEBUG: Saving to loaded sheets with display filename: {display_filename}")
    
        try:
            # Find the correct file data in all_filtered_sheets
            
            current_file_data = None
        
            # For .vap3 files, the matching logic needs to be more flexible
            if self.file_path.endswith('.vap3'):
                debug_print("DEBUG: .vap3 file detected, using flexible matching")
            
                # Try multiple matching strategies for .vap3 files
                for file_data in self.parent.all_filtered_sheets:
                    # First try original filename match
                    if (self.original_filename and 
                        (file_data.get("original_filename") == self.original_filename or
                         file_data.get("database_filename") == self.original_filename or
                         file_data.get("file_name") == self.original_filename)):
                        current_file_data = file_data
                        debug_print(f"DEBUG: Found matching .vap3 file by original filename: {self.original_filename}")
                        break
                
                    # Then try display filename match
                    if (display_filename and 
                        (file_data.get("file_name") == display_filename or
                         file_data.get("display_filename") == display_filename)):
                        current_file_data = file_data
                        debug_print(f"DEBUG: Found matching .vap3 file by display filename: {display_filename}")
                        break
            
                # If still no match, just use the first loaded file for .vap3 files
                if not current_file_data and self.parent.all_filtered_sheets:
                    current_file_data = self.parent.all_filtered_sheets[0]
                    debug_print("DEBUG: Using first loaded file as fallback for .vap3 file")
            else:
                # Regular Excel file matching
                for file_data in self.parent.all_filtered_sheets:
                    if file_data.get("file_path") == self.file_path:
                        current_file_data = file_data
                        debug_print(f"DEBUG: Found matching Excel file: {self.file_path}")
                        break

            if not current_file_data:
                debug_print(f"ERROR: Could not find file data for {display_filename}")
                return

            debug_print(f"DEBUG: Saving data to loaded sheets for test: {self.test_name}")

            # Check if the sheet exists in loaded data
            if not hasattr(self.parent, 'filtered_sheets') or self.test_name not in self.parent.filtered_sheets:
                debug_print(f"ERROR: Sheet {self.test_name} not found in loaded data")
                raise Exception(f"Sheet '{self.test_name}' not found in loaded data")
        
            sheet_data = self.parent.filtered_sheets[self.test_name]['data'].copy()
            debug_print(f"DEBUG: Found loaded sheet data with shape: {sheet_data.shape}")
    
            # Determine format based on test type
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                columns_per_sample = 8  # [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
                debug_print("DEBUG: Saving in User Test Simulation format")
            else:
                columns_per_sample = 12  # [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM, etc.]
                debug_print("DEBUG: Saving in standard format")
    
            total_data_written = 0
    
            # Save data for each sample
            for sample_idx in range(self.num_samples):
                sample_id = f"Sample {sample_idx + 1}"
                col_offset = sample_idx * columns_per_sample
        
                debug_print(f"DEBUG: Saving data for {sample_id} with column offset {col_offset}")
        
                sample_data_written = 0
                data_start_row = 4  # Data starts at row 5 (0-indexed row 4)
        
                # Clear existing data in this sample's area first
                for clear_row in range(data_start_row, len(sheet_data)):
                    for clear_col in range(col_offset, col_offset + columns_per_sample):
                        if clear_col < len(sheet_data.columns):
                            sheet_data.iloc[clear_row, clear_col] = ""
        
                # Write new data for this sample
                for i in range(len(self.data[sample_id]["puffs"])):
                    data_row_idx = data_start_row + i
                
                    # Ensure we don't exceed the DataFrame bounds
                    if data_row_idx >= len(sheet_data):
                        break
                
                    try:
                        # Prepare data values based on test type
                        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                            # User Test Simulation format: [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
                            values_to_write = [
                                self.data[sample_id]["chronography"][i] if i < len(self.data[sample_id]["chronography"]) else "",
                                self.data[sample_id]["puffs"][i] if i < len(self.data[sample_id]["puffs"]) else "",
                                self.data[sample_id]["before_weight"][i] if i < len(self.data[sample_id]["before_weight"]) else "",
                                self.data[sample_id]["after_weight"][i] if i < len(self.data[sample_id]["after_weight"]) else "",
                                self.data[sample_id]["draw_pressure"][i] if i < len(self.data[sample_id]["draw_pressure"]) else "",
                                self.data[sample_id]["smell"][i] if i < len(self.data[sample_id]["smell"]) else "",  # "Failure" stored in smell field
                                self.data[sample_id]["notes"][i] if i < len(self.data[sample_id]["notes"]) else "",
                                self.data[sample_id]["tpm"][i] if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None else ""
                            ]
                    
                        else:
                            # Standard format: [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM]
                            values_to_write = [
                                self.data[sample_id]["puffs"][i] if i < len(self.data[sample_id]["puffs"]) else "",
                                self.data[sample_id]["before_weight"][i] if i < len(self.data[sample_id]["before_weight"]) else "",
                                self.data[sample_id]["after_weight"][i] if i < len(self.data[sample_id]["after_weight"]) else "",
                                self.data[sample_id]["draw_pressure"][i] if i < len(self.data[sample_id]["draw_pressure"]) else "",
                                self.data[sample_id]["resistance"][i] if i < len(self.data[sample_id]["resistance"]) else "",
                                self.data[sample_id]["smell"][i] if i < len(self.data[sample_id]["smell"]) else "",
                                self.data[sample_id]["clog"][i] if i < len(self.data[sample_id]["clog"]) else "",
                                self.data[sample_id]["notes"][i] if i < len(self.data[sample_id]["notes"]) else "",
                                self.data[sample_id]["tpm"][i] if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None else ""
                            ]
                
                        # Write values to the appropriate columns
                        for col_idx, value in enumerate(values_to_write):
                            target_col = col_offset + col_idx
                            if target_col < len(sheet_data.columns):
                                sheet_data.iloc[data_row_idx, target_col] = value
                
                        sample_data_written += 1
                
                    except Exception as e:
                        debug_print(f"DEBUG: Error writing row {i} for {sample_id}: {e}")
                        continue
        
                total_data_written += sample_data_written
                debug_print(f"DEBUG: Wrote {sample_data_written} data rows for {sample_id}")
    
            # Update the loaded sheet data in memory
            self.parent.filtered_sheets[self.test_name]['data'] = sheet_data
            debug_print(f"DEBUG: Updated loaded sheet data with {total_data_written} total data rows")
    
            # Also update the UI state in all_filtered_sheets
            if hasattr(self.parent, 'all_filtered_sheets'):
                for file_data in self.parent.all_filtered_sheets:
                    if self.test_name in file_data.get('filtered_sheets', {}):
                        file_data['filtered_sheets'][self.test_name]['data'] = sheet_data.copy()
                        debug_print("DEBUG: Updated all_filtered_sheets with new data")
                        break
    
            debug_print("DEBUG: Successfully saved data to loaded sheets")
    
            debug_print("DEBUG: Loaded sheets save completed - main GUI data should now be current")
    
            # Ensure the main GUI's current filtered_sheets reflects the updated data
            if hasattr(self.parent, 'filtered_sheets') and self.test_name in self.parent.filtered_sheets:
                # The data should already be updated, but let's ensure it's current
                debug_print(f"DEBUG: Main GUI filtered_sheets for {self.test_name} is current")

        except Exception as e:
            debug_print(f"ERROR: Failed to save data to loaded sheets: {e}")
            import traceback
            traceback.print_exc()
            raise e

    def _update_application_state(self):
        """Update the main application's state if this is a VAP3 file."""
        # Check if the parent has methods to update state
        if hasattr(self.parent, 'filtered_sheets') and hasattr(self.parent, 'file_path'):
            if self.parent.file_path and self.parent.file_path.endswith('.vap3'):
                debug_print("DEBUG: Updating VAP3 file and application state")
                try:
                    # Import here to avoid circular imports
                    from vap_file_manager import VapFileManager
                    
                    vap_manager = VapFileManager()
                    
                    # Get current application state
                    image_crop_states = getattr(self.parent, 'image_crop_states', {})
                    plot_settings = {
                        'selected_plot_type': getattr(self.parent, 'selected_plot_type', tk.StringVar()).get()
                    }
                    
                    # Save to VAP3 file
                    success = vap_manager.save_to_vap3(
                        self.parent.file_path,
                        self.parent.filtered_sheets,
                        getattr(self.parent, 'sheet_images', {}),
                        getattr(self.parent, 'plot_options', []),
                        image_crop_states,
                        plot_settings
                    )
                    
                    if success:
                        debug_print("DEBUG: VAP3 file updated successfully")
                    else:
                        debug_print("DEBUG: Failed to update VAP3 file")
                        
                except Exception as e:
                    debug_print(f"DEBUG: Error updating VAP3 file: {e}")
    
    def create_sample_tab(self, parent_frame, sample_id, sample_index):
        """Create a tab for a single sample with data entry controls."""
        debug_print(f"DEBUG: Creating sample tab for {sample_id}")
        debug_print(f"DEBUG: Test name: '{self.test_name}'")

        # Create main container that will hold everything
        main_container = ttk.Frame(parent_frame)
        main_container.pack(fill="both", expand=True)
    
        # Configure main container to expand properly
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        # Determine columns based on test type
        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
            debug_print(f"DEBUG: Setting up User Test Simulation columns with chronography")
            columns = ["Chronography", "Puffs", "Before Weight", "After Weight", "Draw Pressure", "Failure", "Notes", "TPM"]
            column_widths = [100, 80, 100, 100, 100, 80, 150, 80]
            self.tree_columns = len(columns)

            # Ensure chronography data exists
            if "chronography" not in self.data[sample_id]:
                debug_print(f"DEBUG: Creating chronography data structure for {sample_id}")
                self.data[sample_id]["chronography"] = []
                existing_length = len(self.data[sample_id]["puffs"])
                for j in range(existing_length):
                    self.data[sample_id]["chronography"].append("")
        
            debug_print(f"DEBUG: User Test Simulation columns: {columns}")
        else:
            columns = ["Puffs", "Before Weight", "After Weight", "Draw Pressure", "Resistance", "Smell", "Clog", "Notes", "TPM"]
            column_widths = [80, 100, 100, 100, 100, 80, 80, 150, 80]
            self.tree_columns = len(columns)
            debug_print(f"DEBUG: Standard test columns: {columns}")

        # Create the Treeview directly in the main container
        tree = ttk.Treeview(main_container, columns=columns, show="headings")
    
        # Configure columns
        for i, (col, width) in enumerate(zip(columns, column_widths)):
            tree.heading(col, text=col)
            tree.column(col, width=width, minwidth=50, anchor='center')
            debug_print(f"DEBUG: Configured column {i}: {col} with width {width}")

        # Configure highlighting tags
        tree.tag_configure('highlight_selected', background='#CCE5FF', foreground='black')
        tree.tag_configure('tpm_calculated', background='#C6EFCE', foreground='black')
        tree.tag_configure('evenrow', background='#F8F8F8')
        tree.tag_configure('oddrow', background='white')

        # Create scrollbars for the treeview
        tree_v_scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=tree.yview)
        tree_h_scrollbar = ttk.Scrollbar(main_container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=tree_v_scrollbar.set, xscrollcommand=tree_h_scrollbar.set)

        # Grid everything to fill the container
        tree.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        tree_v_scrollbar.grid(row=0, column=1, sticky="ns")
        tree_h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Populate treeview with initial data
        self.update_treeview(tree, sample_id)

        # Bind events for editing
        tree.bind("<Button-1>", lambda event: self.on_tree_click(event, tree, sample_id))
        tree.bind("<Double-1>", lambda event: self.on_tree_click(event, tree, sample_id))
        tree.bind("<KeyPress>", lambda event: self.on_tree_key_press(event, tree, sample_id))

        debug_print(f"DEBUG: Sample tab created for {sample_id} with {len(columns)} columns")
        return tree

    def on_tree_key_press(self, event, tree, sample_id):
        """Handle key press events in the treeview."""
        if event.keysym == "Tab":
            return self.handle_tab_key(event, tree, sample_id)
        elif event.keysym in ["Up", "Down", "Left", "Right"]:
            return self.handle_arrow_key(event, tree, sample_id, event.keysym.lower())
        elif event.char.isprintable():
            return self.start_edit_on_typing(event, tree, sample_id)
    
        return None

    def update_treeview(self, tree, sample_id):
        """Update the treeview with current data, highlighting only TPM cells."""
        debug_print(f"DEBUG: Updating treeview for {sample_id}")

        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
    
        # Get data length
        data_length = len(self.data[sample_id]["puffs"])
        debug_print(f"DEBUG: Data length for {sample_id}: {data_length}")
        
        has_real_data = False
        for i in range(data_length):
            if (self.data[sample_id]["puffs"][i] or 
                self.data[sample_id]["before_weight"][i] or 
                self.data[sample_id]["after_weight"][i]):
                has_real_data = True
                break
        debug_print(f"DEBUG: {sample_id} has real data: {has_real_data}")

        # Insert data rows with alternating colors for better visual separation
        for i in range(data_length):
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                # User Test Simulation format: [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
                tpm_value = ""
                has_tpm = False
                if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None:
                    tpm_value = f"{self.data[sample_id]['tpm'][i]:.6f}"
                    has_tpm = True
            
                values = [
                    self.data[sample_id]["chronography"][i] if i < len(self.data[sample_id]["chronography"]) else "",
                    self.data[sample_id]["puffs"][i] if i < len(self.data[sample_id]["puffs"]) else "",
                    self.data[sample_id]["before_weight"][i] if i < len(self.data[sample_id]["before_weight"]) else "",
                    self.data[sample_id]["after_weight"][i] if i < len(self.data[sample_id]["after_weight"]) else "",
                    self.data[sample_id]["draw_pressure"][i] if i < len(self.data[sample_id]["draw_pressure"]) else "",
                    self.data[sample_id]["smell"][i] if i < len(self.data[sample_id]["smell"]) else "",  # Failure column (stored in smell field)
                    self.data[sample_id]["notes"][i] if i < len(self.data[sample_id]["notes"]) else "",
                    tpm_value  # TPM column
                ]
            else:
                # Standard format: [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM]
                tpm_value = ""
                has_tpm = False
                if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None:
                    tpm_value = f"{self.data[sample_id]['tpm'][i]:.6f}"
                    has_tpm = True
            
                values = [
                    self.data[sample_id]["puffs"][i] if i < len(self.data[sample_id]["puffs"]) else "",
                    self.data[sample_id]["before_weight"][i] if i < len(self.data[sample_id]["before_weight"]) else "",
                    self.data[sample_id]["after_weight"][i] if i < len(self.data[sample_id]["after_weight"]) else "",
                    self.data[sample_id]["draw_pressure"][i] if i < len(self.data[sample_id]["draw_pressure"]) else "",
                    self.data[sample_id]["resistance"][i] if i < len(self.data[sample_id]["resistance"]) else "",
                    self.data[sample_id]["smell"][i] if i < len(self.data[sample_id]["smell"]) else "",
                    self.data[sample_id]["clog"][i] if i < len(self.data[sample_id]["clog"]) else "",
                    self.data[sample_id]["notes"][i] if i < len(self.data[sample_id]["notes"]) else "",
                    tpm_value  # TPM column
                ]
          
                if i < 3:
                    debug_print(f"DEBUG: Row {i} values being inserted: {values}")
            # Determine row coloring and TPM highlighting
            tags = []
            if i % 2 == 0:
                tags.append('evenrow')
            else:
                tags.append('oddrow')
        
            # Only add TPM highlighting if this row has calculated TPM
            if has_tpm:
                tags.append('tpm_calculated')
        
            item = tree.insert("", "end", values=values, tags=tags)
    
        debug_print(f"DEBUG: Treeview updated for {sample_id} with {data_length} rows")

    def refresh_main_gui_after_save_old(self):
        """Refresh the main GUI to show updated data after saving."""
        debug_print("DEBUG: Starting main GUI refresh")

        try:
            # Use test_name as the sheet to refresh
            current_sheet_name = self.test_name
            debug_print(f"DEBUG: Target sheet: {current_sheet_name}")
    
            # Check parent and file_manager
            debug_print(f"DEBUG: Has parent: {hasattr(self, 'parent')}")
            debug_print(f"DEBUG: Parent exists: {self.parent is not None}")
            debug_print(f"DEBUG: Has file_manager: {hasattr(self.parent, 'file_manager')}")
            debug_print(f"DEBUG: File manager exists: {self.parent.file_manager is not None}")
    
            # Handle .vap3 files differently - they don't need reloading from disk
            if self.file_path.endswith('.vap3'):
                debug_print("DEBUG: .vap3 file detected - updating UI without reloading from disk")
            
                # For .vap3 files, the data is already updated in memory
                # Just refresh the display without reloading
                if hasattr(self.parent, 'selected_sheet'):
                    if not self.parent.selected_sheet:
                        import tkinter as tk
                        self.parent.selected_sheet = tk.StringVar()
                        debug_print("DEBUG: Created selected_sheet variable")
                    self.parent.selected_sheet.set(current_sheet_name)
                    debug_print(f"DEBUG: Set sheet to {current_sheet_name}")

                # Update the display
                if hasattr(self.parent, 'update_displayed_sheet'):
                    debug_print("DEBUG: Calling update_displayed_sheet")
                    self.parent.update_displayed_sheet(current_sheet_name)
                    debug_print("DEBUG: Display updated")
                else:
                    debug_print("DEBUG: No update_displayed_sheet method available")
        
            else:
                # Regular Excel file handling
                debug_print(f"DEBUG: Force reloading Excel file {self.file_path}")
        
                # Clear any existing cache for this file first
                if hasattr(self.parent, 'file_manager') and self.parent.file_manager:
                    if hasattr(self.parent.file_manager, 'loaded_files_cache'):
                        cache_key = f"{self.file_path}_None"
                        if cache_key in self.parent.file_manager.loaded_files_cache:
                            del self.parent.file_manager.loaded_files_cache[cache_key]
                            debug_print("DEBUG: Cleared file cache")
            
                    self.parent.file_manager.load_excel_file(
                        self.file_path, 
                        skip_database_storage=True,
                        force_reload=True
                    )
                    debug_print("DEBUG: Excel file reloaded")

                    # Update all_filtered_sheets
                    if hasattr(self.parent, 'all_filtered_sheets') and hasattr(self.parent, 'current_file'):
                        debug_print(f"DEBUG: Updating all_filtered_sheets for {self.parent.current_file}")
                        for file_data in self.parent.all_filtered_sheets:
                            if file_data.get("file_name") == self.parent.current_file:
                                file_data["filtered_sheets"] = copy.deepcopy(self.parent.filtered_sheets)
                                debug_print("DEBUG: Updated all_filtered_sheets entry")
                                break

                    # Set the sheet selection
                    if hasattr(self.parent, 'selected_sheet'):
                        if not self.parent.selected_sheet:
                            import tkinter as tk
                            self.parent.selected_sheet = tk.StringVar()
                            debug_print("DEBUG: Created selected_sheet variable")
                        self.parent.selected_sheet.set(current_sheet_name)
                        debug_print(f"DEBUG: Set sheet to {current_sheet_name}")

                    # Update the display
                    if hasattr(self.parent, 'update_displayed_sheet'):
                        debug_print("DEBUG: Calling update_displayed_sheet")
                        self.parent.update_displayed_sheet(current_sheet_name)
                        debug_print("DEBUG: Display updated")
                    else:
                        debug_print("DEBUG: No update_displayed_sheet method available")
                else:
                    debug_print("DEBUG: Cannot reload - file_manager not available")
        
            # Force GUI update
            if hasattr(self.parent, 'root'):
                self.parent.root.update_idletasks()
                debug_print("DEBUG: Forced GUI update")

            debug_print("DEBUG: Refresh complete")

        except Exception as e:
            debug_print(f"ERROR: Refresh failed: {e}")
            import traceback
            traceback.print_exc()

    def refresh_main_gui_after_save(self):
        """Update the main GUI data structures after saving data."""
        try:
            debug_print("DEBUG: Refreshing main GUI after data collection save")
        
            # For .vap3 files or database files, the data is already updated in memory
            # We just need to ensure the main GUI reflects the current state
            if self.file_path.endswith('.vap3') or not os.path.exists(self.file_path):
                self._update_vap3_data_in_main_gui()
            else:
                # For regular Excel files, reload from file
                self._update_excel_data_in_main_gui()
            
            # Mark this file as modified in the staging area
            self._mark_file_as_modified()
        
            debug_print("DEBUG: Main GUI refresh completed")
        
        except Exception as e:
            debug_print(f"ERROR: Failed to refresh main GUI: {e}")
            import traceback
            traceback.print_exc()

    def _update_excel_data_in_main_gui(self):
        """Update main GUI data for Excel files."""
        try:
            # Only do this for actual Excel files that exist
            if not os.path.exists(self.file_path) or not self.file_path.endswith(('.xlsx', '.xls')):
                debug_print("DEBUG: File doesn't exist or isn't Excel, skipping Excel update")
                return
            
            # Reload just this sheet from the Excel file
            import pandas as pd
            new_sheet_data = pd.read_excel(self.file_path, sheet_name=self.test_name)
        
            # Update the main GUI's filtered_sheets
            if hasattr(self.parent, 'filtered_sheets') and self.test_name in self.parent.filtered_sheets:
                self.parent.filtered_sheets[self.test_name]['data'] = new_sheet_data
                self.parent.filtered_sheets[self.test_name]['is_empty'] = new_sheet_data.empty
                debug_print(f"DEBUG: Updated sheet {self.test_name} in main GUI filtered_sheets")
        
            # Update all_filtered_sheets for the current file
            if hasattr(self.parent, 'all_filtered_sheets') and hasattr(self.parent, 'current_file'):
                for file_data in self.parent.all_filtered_sheets:
                    if file_data["file_name"] == self.parent.current_file:
                        if self.test_name in file_data["filtered_sheets"]:
                            file_data["filtered_sheets"][self.test_name]['data'] = new_sheet_data
                            file_data["filtered_sheets"][self.test_name]['is_empty'] = new_sheet_data.empty
                            debug_print(f"DEBUG: Updated sheet {self.test_name} in all_filtered_sheets for file {self.parent.current_file}")
                        break
        
        except Exception as e:
            debug_print(f"ERROR: Failed to update Excel data in main GUI: {e}")
            raise

    def _update_vap3_data_in_main_gui(self):
        """Update main GUI data for .vap3 files - data is already in memory."""
        try:
            debug_print("DEBUG: Updating .vap3 data in main GUI - data should already be current")
        
            # For .vap3 files, the data should have already been updated in the save process
            # The _save_to_loaded_sheets method should have updated the main GUI data structures
        
            # Just ensure the all_filtered_sheets stays synchronized
            if hasattr(self.parent, 'filtered_sheets') and hasattr(self.parent, 'all_filtered_sheets'):
                current_file = getattr(self.parent, 'current_file', None)
                if current_file:
                    for file_data in self.parent.all_filtered_sheets:
                        if file_data["file_name"] == current_file:
                            # Ensure the all_filtered_sheets data matches current filtered_sheets
                            file_data["filtered_sheets"] = copy.deepcopy(self.parent.filtered_sheets)
                            debug_print(f"DEBUG: Synchronized .vap3 data for file {current_file}")
                            break
        
            debug_print("DEBUG: .vap3 data update completed")
        
        except Exception as e:
            debug_print(f"ERROR: Failed to update .vap3 data in main GUI: {e}")
            raise

    def _mark_file_as_modified(self):
        """Mark the current file as modified in the staging area."""
        try:
            if hasattr(self.parent, 'all_filtered_sheets') and hasattr(self.parent, 'current_file'):
                for file_data in self.parent.all_filtered_sheets:
                    if file_data["file_name"] == self.parent.current_file:
                        file_data["is_modified"] = True
                        file_data["last_modified"] = time.time()
                        debug_print(f"DEBUG: Marked file {self.parent.current_file} as modified")
                        break
        
            # Update the main GUI window title to show modified status
            if hasattr(self.parent, 'root') and hasattr(self.parent, 'current_file'):
                current_title = self.parent.root.title()
                if not current_title.endswith(" *"):
                    self.parent.root.title(current_title + " *")
                    debug_print("DEBUG: Updated window title to show modified status")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to mark file as modified: {e}")

    def calculate_tpm(self, sample_id):
        """Calculate TPM (Total Particulate Matter) for all rows with valid data."""
        valid_tpm_values = []
    
        for i in range(len(self.data[sample_id]["puffs"])):
            try:
                # Get weight values
                before_weight_str = self.data[sample_id]["before_weight"][i]
                after_weight_str = self.data[sample_id]["after_weight"][i]
            
                # Skip if either weight is missing
                if not before_weight_str or not after_weight_str:
                    continue
                
                # Convert to float
                before_weight = float(before_weight_str)
                after_weight = float(after_weight_str)
            
                # Validate weights
                if before_weight <= after_weight:
                    continue
                
                # Calculate puffs in this interval
                puff_interval = int(self.data[sample_id]["puffs"][i])
                puffs_in_interval = puff_interval
            
                if i > 0:
                    prev_puff = int(self.data[sample_id]["puffs"][i - 1])
                    puffs_in_interval = puff_interval - prev_puff
                
                # Skip if invalid puff interval
                if puffs_in_interval <= 0:
                    continue
                
                # Calculate TPM (mg/puff)
                weight_consumed = before_weight - after_weight  # in grams
                tpm = (weight_consumed * 1000) / puffs_in_interval  # Convert to mg per puff
            
                # Store result
                self.data[sample_id]["tpm"][i] = round(tpm, 6)
                valid_tpm_values.append(tpm)
                
            except Exception:
                # Ensure tpm list is long enough even for failed calculations
                while len(self.data[sample_id]["tpm"]) <= i:
                    self.data[sample_id]["tpm"].append(None)
    
        # Update average TPM
        self.data[sample_id]["avg_tpm"] = sum(valid_tpm_values) / len(valid_tpm_values) if valid_tpm_values else 0.0
    
        return len(valid_tpm_values) > 0

    def validate_weight_entry(self, sample_id, row_idx, column_name, value):
        """
        Validate weight entries to ensure data consistency.
        Returns True if valid, False otherwise.
        """
        try:
            if not value or not value.strip():
                return True  # Empty values are allowed
                
            weight_value = float(value.strip())
            
            # Check for reasonable weight values (between 0.001g and 100g)
            if weight_value < 0.001 or weight_value > 100:
                debug_print(f"WARNING: Weight value {weight_value}g seems unreasonable for row {row_idx}")
                return False
                
            # If this is an after_weight, check that it's less than before_weight
            if column_name == "after_weight":
                before_weight_str = self.data[sample_id]["before_weight"][row_idx]
                if before_weight_str and before_weight_str.strip():
                    try:
                        before_weight = float(before_weight_str.strip())
                        if weight_value >= before_weight:
                            debug_print(f"WARNING: After weight ({weight_value}g) should be less than before weight ({before_weight}g)")
                            return False
                    except ValueError:
                        pass
                        
            # If this is a before_weight, check that it's greater than after_weight
            elif column_name == "before_weight":
                after_weight_str = self.data[sample_id]["after_weight"][row_idx]
                if after_weight_str and after_weight_str.strip():
                    try:
                        after_weight = float(after_weight_str.strip())
                        if weight_value <= after_weight:
                            debug_print(f"WARNING: Before weight ({weight_value}g) should be greater than after weight ({after_weight}g)")
                            return False
                    except ValueError:
                        pass
                        
            return True
            
        except ValueError:
            debug_print(f"ERROR: Invalid weight value '{value}' - must be a number")
            return False
    
    def load_existing_data_from_loaded_sheets(self):
        """Load existing data from already-loaded sheet data (for .vap3 files)."""
        debug_print(f"DEBUG: Loading existing data from loaded sheets for test: {self.test_name}")
    
        try:
            # Get the loaded sheet data
            if not hasattr(self.parent, 'filtered_sheets') or self.test_name not in self.parent.filtered_sheets:
                debug_print(f"ERROR: Sheet {self.test_name} not found in loaded data")
                return
        
            sheet_data = self.parent.filtered_sheets[self.test_name]['data']
            debug_print(f"DEBUG: Found loaded sheet data with shape: {sheet_data.shape}")
    
            debug_print("DEBUG: First 10 rows and 16 columns of actual DataFrame data:")
            for i in range(min(10, len(sheet_data))):
                row_preview = []
                for j in range(min(16, len(sheet_data.columns))):
                    val = sheet_data.iloc[i, j]
                    if pd.isna(val):
                        row_preview.append("NaN")
                    else:
                        val_str = str(val)[:10]  # Truncate long values
                        row_preview.append(f"'{val_str}'")
                debug_print(f"  Row {i}: {row_preview}")

            # Determine the data format based on test type
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                columns_per_sample = 8
                debug_print("DEBUG: Loading User Test Simulation format")
            else:
                columns_per_sample = 12
                debug_print("DEBUG: Loading standard format")
    
            # Load data for each sample
            loaded_data_count = 0
        
            for sample_idx in range(self.num_samples):
                sample_id = f"Sample {sample_idx + 1}"
                col_offset = sample_idx * columns_per_sample
        
                debug_print(f"DEBUG: Loading data for {sample_id} with column offset {col_offset}")
            
                debug_print(f"DEBUG: Clearing existing template data for {sample_id}")
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    self.data[sample_id]["chronography"] = []
                    self.data[sample_id]["puffs"] = []
                    self.data[sample_id]["before_weight"] = []
                    self.data[sample_id]["after_weight"] = []
                    self.data[sample_id]["draw_pressure"] = []
                    self.data[sample_id]["smell"] = []  
                    self.data[sample_id]["notes"] = []
                    self.data[sample_id]["tpm"] = []
                else:
                    self.data[sample_id]["puffs"] = []
                    self.data[sample_id]["before_weight"] = []
                    self.data[sample_id]["after_weight"] = []
                    self.data[sample_id]["draw_pressure"] = []
                    self.data[sample_id]["resistance"] = []
                    self.data[sample_id]["smell"] = []
                    self.data[sample_id]["clog"] = []
                    self.data[sample_id]["notes"] = []
                    self.data[sample_id]["tpm"] = []
        
                # Load data starting from row 5 (DataFrame index 4)
                sample_had_data = False
                row_count = 0
        
                data_start_row = 4  # Data starts at row 5 (0-indexed row 4)
        
                for data_row_idx in range(data_start_row, min(len(sheet_data), 100)):  # Limit to reasonable number of rows
                    try:
                        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                            # User Test Simulation format: [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
                            chronography_col = 0 + col_offset
                            puffs_col = 1 + col_offset
                            before_weight_col = 2 + col_offset
                            after_weight_col = 3 + col_offset
                            draw_pressure_col = 4 + col_offset
                            failure_col = 5 + col_offset  # Stored in smell field
                            notes_col = 6 + col_offset
                            tpm_col = 7 + col_offset
                    
                            # Extract values if columns exist
                            values = {}
                    
                            # Helper function to clean values and strip quotes
                            def clean_value(raw_val):
                                if pd.isna(raw_val):
                                    return ""
                                val_str = str(raw_val).strip()
                                # Remove surrounding quotes if present
                                if val_str.startswith("'") and val_str.endswith("'"):
                                    val_str = val_str[1:-1]
                                elif val_str.startswith('"') and val_str.endswith('"'):
                                    val_str = val_str[1:-1]
                                return val_str
                        
                            debug_print(f"DEBUG: Processing row {data_row_idx} for {sample_id}")
                        
                            if chronography_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, chronography_col]
                                values['chronography'] = clean_value(raw_val)
                                debug_print(f"DEBUG: chronography raw: {raw_val}, cleaned: {values['chronography']}")
                            else:
                                values['chronography'] = ""
                        
                            if puffs_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, puffs_col]
                                values['puffs'] = clean_value(raw_val)
                                debug_print(f"DEBUG: puffs raw: {raw_val}, cleaned: {values['puffs']}")
                            else:
                                values['puffs'] = ""
                        
                            if before_weight_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, before_weight_col]
                                values['before_weight'] = clean_value(raw_val)
                                debug_print(f"DEBUG: before_weight raw: {raw_val}, cleaned: {values['before_weight']}")
                            else:
                                values['before_weight'] = ""
                        
                            if after_weight_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, after_weight_col]
                                values['after_weight'] = clean_value(raw_val)
                                debug_print(f"DEBUG: after_weight raw: {raw_val}, cleaned: {values['after_weight']}")
                            else:
                                values['after_weight'] = ""
                        
                            if draw_pressure_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, draw_pressure_col]
                                values['draw_pressure'] = clean_value(raw_val)
                            else:
                                values['draw_pressure'] = ""
                        
                            if failure_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, failure_col]
                                values['failure'] = clean_value(raw_val)
                            else:
                                values['failure'] = ""
                        
                            if notes_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, notes_col]
                                values['notes'] = clean_value(raw_val)
                            else:
                                values['notes'] = ""
                        
                            if tpm_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, tpm_col]
                                cleaned_val = clean_value(raw_val)
                                if cleaned_val and cleaned_val.strip():
                                    try:
                                        values['tpm'] = float(cleaned_val)
                                        debug_print(f"DEBUG: tpm raw: {raw_val}, cleaned: {cleaned_val}, float: {values['tpm']}")
                                    except (ValueError, TypeError):
                                        values['tpm'] = None
                                        debug_print(f"DEBUG: tpm conversion failed for: {cleaned_val}")
                                else:
                                    values['tpm'] = None
                            else:
                                values['tpm'] = None
                    
                            # Check if this row has any meaningful data (not all empty)
                            if any([values['chronography'], values['puffs'], values['before_weight'], values['after_weight']]):
                                debug_print(f"DEBUG: Storing data for row {data_row_idx}: puffs={values['puffs']}, before={values['before_weight']}, after={values['after_weight']}")
                                # Append the data (building from scratch)
                                self.data[sample_id]["chronography"].append(values['chronography'])
                                self.data[sample_id]["puffs"].append(values['puffs'])
                                self.data[sample_id]["before_weight"].append(values['before_weight'])
                                self.data[sample_id]["after_weight"].append(values['after_weight'])
                                self.data[sample_id]["draw_pressure"].append(values['draw_pressure'])
                                self.data[sample_id]["smell"].append(values['failure']) 
                                self.data[sample_id]["notes"].append(values['notes'])
                                self.data[sample_id]["tpm"].append(values['tpm'])
                        
                                sample_had_data = True
                                row_count += 1
                            else:
                                debug_print(f"DEBUG: Row {data_row_idx} has no meaningful data, skipping")
                        
                        else:
                            # Standard format: [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM]
                            puffs_col = 0 + col_offset
                            before_weight_col = 1 + col_offset
                            after_weight_col = 2 + col_offset
                            draw_pressure_col = 3 + col_offset
                            resistance_col = 4 + col_offset
                            smell_col = 5 + col_offset
                            clog_col = 6 + col_offset
                            notes_col = 7 + col_offset
                            tpm_col = 8 + col_offset  # TPM is typically in column 9 for standard format
                    
                            # Extract values if columns exist
                            values = {}
                        
                            # Helper function to clean values and strip quotes
                            def clean_value(raw_val):
                                if pd.isna(raw_val):
                                    return ""
                                val_str = str(raw_val).strip()
                                # Remove surrounding quotes if present
                                if val_str.startswith("'") and val_str.endswith("'"):
                                    val_str = val_str[1:-1]
                                elif val_str.startswith('"') and val_str.endswith('"'):
                                    val_str = val_str[1:-1]
                                return val_str
                        
                            debug_print(f"DEBUG: Processing row {data_row_idx} for {sample_id}")
                    
                            for field, col_idx in [
                                ('puffs', puffs_col), ('before_weight', before_weight_col), 
                                ('after_weight', after_weight_col), ('draw_pressure', draw_pressure_col),
                                ('resistance', resistance_col), ('smell', smell_col), 
                                ('clog', clog_col), ('notes', notes_col)
                            ]:
                                if col_idx < len(sheet_data.columns):
                                    raw_val = sheet_data.iloc[data_row_idx, col_idx]
                                    values[field] = clean_value(raw_val)
                                    debug_print(f"DEBUG: {field} raw: {raw_val}, cleaned: {values[field]}")
                                else:
                                    values[field] = ""
                    
                            # Handle TPM separately
                            if tpm_col < len(sheet_data.columns):
                                raw_val = sheet_data.iloc[data_row_idx, tpm_col]
                                cleaned_val = clean_value(raw_val)
                                if cleaned_val and cleaned_val.strip():
                                    try:
                                        values['tpm'] = float(cleaned_val)
                                        debug_print(f"DEBUG: tpm raw: {raw_val}, cleaned: {cleaned_val}, float: {values['tpm']}")
                                    except (ValueError, TypeError):
                                        values['tpm'] = None
                                        debug_print(f"DEBUG: tpm conversion failed for: {cleaned_val}")
                                else:
                                    values['tpm'] = None
                            else:
                                values['tpm'] = None
                    
                            # Check if this row has any meaningful data (not all empty)
                            if any([values['puffs'], values['before_weight'], values['after_weight']]):
                                debug_print(f"DEBUG: Storing data for row {data_row_idx}: puffs={values['puffs']}, before={values['before_weight']}, after={values['after_weight']}")
                                # Append the data (building from scratch)
                                self.data[sample_id]["puffs"].append(values['puffs'])
                                self.data[sample_id]["before_weight"].append(values['before_weight'])
                                self.data[sample_id]["after_weight"].append(values['after_weight'])
                                self.data[sample_id]["draw_pressure"].append(values['draw_pressure'])
                                self.data[sample_id]["resistance"].append(values['resistance'])
                                self.data[sample_id]["smell"].append(values['smell'])
                                self.data[sample_id]["clog"].append(values['clog'])
                                self.data[sample_id]["notes"].append(values['notes'])
                                self.data[sample_id]["tpm"].append(values['tpm'])
                        
                                sample_had_data = True
                                row_count += 1
                            else:
                                debug_print(f"DEBUG: Row {data_row_idx} has no meaningful data, skipping")
                
                    except Exception as e:
                        debug_print(f"DEBUG: Error processing row {data_row_idx} for {sample_id}: {e}")
                        continue
        
                if sample_had_data:
                    loaded_data_count += 1
                    debug_print(f"DEBUG: {sample_id} - Loaded {row_count} rows of existing data (cleared template)")
                else:
                    debug_print(f"DEBUG: {sample_id} - No existing data found")
                    # If no data was found, ensure we still have empty arrays (not template data)
                    if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                        self.data[sample_id]["chronography"] = []
                        self.data[sample_id]["puffs"] = []
                        self.data[sample_id]["before_weight"] = []
                        self.data[sample_id]["after_weight"] = []
                        self.data[sample_id]["draw_pressure"] = []
                        self.data[sample_id]["smell"] = []
                        self.data[sample_id]["notes"] = []
                        self.data[sample_id]["tpm"] = []
                    else:
                        self.data[sample_id]["puffs"] = []
                        self.data[sample_id]["before_weight"] = []
                        self.data[sample_id]["after_weight"] = []
                        self.data[sample_id]["draw_pressure"] = []
                        self.data[sample_id]["resistance"] = []
                        self.data[sample_id]["smell"] = []
                        self.data[sample_id]["clog"] = []
                        self.data[sample_id]["notes"] = []
                        self.data[sample_id]["tpm"] = []
    
            debug_print(f"DEBUG: Successfully loaded existing data for {loaded_data_count} samples from loaded sheets")
    
            # Recalculate TPM values for all samples
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                self.calculate_tpm(sample_id)
                debug_print(f"DEBUG: Calculated TPM for {sample_id}")

            # Update all treeviews to show the loaded data
            for i, sample_tree in enumerate(self.sample_trees):
                sample_id = f"Sample {i + 1}"
                self.update_treeview(sample_tree, sample_id)
                debug_print(f"DEBUG: Updated treeview for {sample_id}")

            # Update the stats panel
            self.update_stats_panel()

            debug_print("DEBUG: Existing data loading from loaded sheets completed successfully")
    
        except Exception as e:
            debug_print(f"ERROR: Failed to load existing data from loaded sheets: {e}")
            import traceback
            traceback.print_exc()

    def load_existing_data_from_file(self):
        """Load existing data from file or loaded sheets depending on file type."""
        debug_print(f"DEBUG: Loading existing data for file: {self.file_path}")
    
        # Check if this is a .vap3 file or if the file doesn't exist (temporary file)
        if self.file_path.endswith('.vap3') or not os.path.exists(self.file_path):
            debug_print("DEBUG: Detected .vap3 file or non-existent file, loading from loaded sheets")
            self.load_existing_data_from_loaded_sheets()
        else:
            debug_print("DEBUG: Detected Excel file, loading from file using openpyxl")
            self.load_existing_data_from_excel_file()

    def load_existing_data_from_excel_file(self):
        """Load existing data from the Excel file into the data collection interface."""
        debug_print(f"DEBUG: Loading existing data from file: {self.file_path}")

        try:
            # Load the workbook and calculate formulas
            wb = openpyxl.load_workbook(self.file_path, data_only=True)  # data_only=True evaluates formulas

            # Check if the test sheet exists
            if self.test_name not in wb.sheetnames:
                debug_print(f"DEBUG: Sheet '{self.test_name}' not found in file")
                return

            ws = wb[self.test_name]
            debug_print(f"DEBUG: Successfully opened sheet '{self.test_name}'")

            # Determine column layout based on test type
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                columns_per_sample = 8  # Including chronography
                debug_print(f"DEBUG: Loading User Test Simulation format with 8 columns per sample")
            else:
                columns_per_sample = 12  # Standard format
                debug_print(f"DEBUG: Loading standard format with 12 columns per sample")

            # Load data for each sample
            loaded_data_count = 0
            for sample_idx in range(self.num_samples):
                sample_id = f"Sample {sample_idx + 1}"
                col_offset = sample_idx * columns_per_sample

                debug_print(f"DEBUG: Loading data for {sample_id} with column offset {col_offset}")

                # CLEAR ALL EXISTING TEMPLATE DATA for this sample (like in load_existing_data_from_loaded_sheets)
                debug_print(f"DEBUG: Clearing existing template data for {sample_id}")
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    self.data[sample_id]["chronography"] = []
                    self.data[sample_id]["puffs"] = []
                    self.data[sample_id]["before_weight"] = []
                    self.data[sample_id]["after_weight"] = []
                    self.data[sample_id]["draw_pressure"] = []
                    self.data[sample_id]["smell"] = []  # Used for failure in user simulation
                    self.data[sample_id]["notes"] = []
                    self.data[sample_id]["tpm"] = []
                else:
                    self.data[sample_id]["puffs"] = []
                    self.data[sample_id]["before_weight"] = []
                    self.data[sample_id]["after_weight"] = []
                    self.data[sample_id]["draw_pressure"] = []
                    self.data[sample_id]["resistance"] = []
                    self.data[sample_id]["smell"] = []
                    self.data[sample_id]["clog"] = []
                    self.data[sample_id]["notes"] = []
                    self.data[sample_id]["tpm"] = []

                # Read data starting from row 5 (Excel row 5 = index 4)
                row_count = 0
                sample_had_data = False

                for excel_row_idx in range(5, min(ws.max_row + 1, 100)):  # Start from row 5, limit to 100 rows max
                    try:
                        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                            # User Test Simulation format: [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
                            chronography_col = 1 + col_offset  # Column A + offset
                            puffs_col = 2 + col_offset          # Column B + offset  
                            before_weight_col = 3 + col_offset  # Column C + offset
                            after_weight_col = 4 + col_offset   # Column D + offset
                            draw_pressure_col = 5 + col_offset  # Column E + offset
                            failure_col = 6 + col_offset        # Column F + offset (stored in smell field)
                            notes_col = 7 + col_offset          # Column G + offset
                            tpm_col = 8 + col_offset            # Column H + offset

                            # Extract values and clean them
                            def clean_excel_value(cell_value):
                                if cell_value is None:
                                    return ""
                                return str(cell_value).strip().strip('"').strip("'")

                            values = {}
                            values['chronography'] = clean_excel_value(ws.cell(row=excel_row_idx, column=chronography_col).value)
                            values['puffs'] = clean_excel_value(ws.cell(row=excel_row_idx, column=puffs_col).value)
                            values['before_weight'] = clean_excel_value(ws.cell(row=excel_row_idx, column=before_weight_col).value)
                            values['after_weight'] = clean_excel_value(ws.cell(row=excel_row_idx, column=after_weight_col).value)
                            values['draw_pressure'] = clean_excel_value(ws.cell(row=excel_row_idx, column=draw_pressure_col).value)
                            values['failure'] = clean_excel_value(ws.cell(row=excel_row_idx, column=failure_col).value)
                            values['notes'] = clean_excel_value(ws.cell(row=excel_row_idx, column=notes_col).value)
                        
                            # Handle TPM calculation
                            tpm_cell_value = ws.cell(row=excel_row_idx, column=tpm_col).value
                            if tpm_cell_value is not None:
                                try:
                                    values['tpm'] = float(tpm_cell_value)
                                except (ValueError, TypeError):
                                    values['tpm'] = None
                            else:
                                values['tpm'] = None

                            # Check if this row has any meaningful data
                            if any([values['puffs'], values['before_weight'], values['after_weight']]):
                                debug_print(f"DEBUG: Appending User Test Simulation data for Excel row {excel_row_idx}: puffs={values['puffs']}, before={values['before_weight']}, after={values['after_weight']}")
                            
                               
                                self.data[sample_id]["chronography"].append(values['chronography'])
                                self.data[sample_id]["puffs"].append(values['puffs'])
                                self.data[sample_id]["before_weight"].append(values['before_weight'])
                                self.data[sample_id]["after_weight"].append(values['after_weight'])
                                self.data[sample_id]["draw_pressure"].append(values['draw_pressure'])
                                self.data[sample_id]["smell"].append(values['failure'])  # failure stored in smell field
                                self.data[sample_id]["notes"].append(values['notes'])
                                self.data[sample_id]["tpm"].append(values['tpm'])

                                sample_had_data = True
                                row_count += 1
                            else:
                                debug_print(f"DEBUG: Excel row {excel_row_idx} has no meaningful data, skipping")

                        else:
                            # Standard format: [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM]
                            puffs_col = 1 + col_offset          # Column A + offset
                            before_weight_col = 2 + col_offset  # Column B + offset
                            after_weight_col = 3 + col_offset   # Column C + offset
                            draw_pressure_col = 4 + col_offset  # Column D + offset
                            resistance_col = 5 + col_offset     # Column E + offset
                            smell_col = 6 + col_offset          # Column F + offset
                            clog_col = 7 + col_offset           # Column G + offset
                            notes_col = 8 + col_offset          # Column H + offset
                            tpm_col = 9 + col_offset            # Column I + offset

                            # Extract values and clean them
                            def clean_excel_value(cell_value):
                                if cell_value is None:
                                    return ""
                                return str(cell_value).strip().strip('"').strip("'")

                            values = {}
                            values['puffs'] = clean_excel_value(ws.cell(row=excel_row_idx, column=puffs_col).value)
                            values['before_weight'] = clean_excel_value(ws.cell(row=excel_row_idx, column=before_weight_col).value)
                            values['after_weight'] = clean_excel_value(ws.cell(row=excel_row_idx, column=after_weight_col).value)
                            values['draw_pressure'] = clean_excel_value(ws.cell(row=excel_row_idx, column=draw_pressure_col).value)
                            values['resistance'] = clean_excel_value(ws.cell(row=excel_row_idx, column=resistance_col).value)
                            values['smell'] = clean_excel_value(ws.cell(row=excel_row_idx, column=smell_col).value)
                            values['clog'] = clean_excel_value(ws.cell(row=excel_row_idx, column=clog_col).value)
                            values['notes'] = clean_excel_value(ws.cell(row=excel_row_idx, column=notes_col).value)

                            # Handle TPM calculation
                            tpm_cell_value = ws.cell(row=excel_row_idx, column=tpm_col).value
                            if tpm_cell_value is not None:
                                try:
                                    values['tpm'] = float(tpm_cell_value)
                                except (ValueError, TypeError):
                                    values['tpm'] = None
                            else:
                                values['tpm'] = None

                            # Check if this row has any meaningful data
                            if any([values['puffs'], values['before_weight'], values['after_weight']]):
                                debug_print(f"DEBUG: Appending standard data for Excel row {excel_row_idx}: puffs={values['puffs']}, before={values['before_weight']}, after={values['after_weight']}")
                            
                             
                                self.data[sample_id]["puffs"].append(values['puffs'])
                                self.data[sample_id]["before_weight"].append(values['before_weight'])
                                self.data[sample_id]["after_weight"].append(values['after_weight'])
                                self.data[sample_id]["draw_pressure"].append(values['draw_pressure'])
                                self.data[sample_id]["resistance"].append(values['resistance'])
                                self.data[sample_id]["smell"].append(values['smell'])
                                self.data[sample_id]["clog"].append(values['clog'])
                                self.data[sample_id]["notes"].append(values['notes'])
                                self.data[sample_id]["tpm"].append(values['tpm'])

                                sample_had_data = True
                                row_count += 1
                            else:
                                debug_print(f"DEBUG: Excel row {excel_row_idx} has no meaningful data, skipping")

                    except Exception as e:
                        debug_print(f"DEBUG: Error processing Excel row {excel_row_idx} for {sample_id}: {e}")
                        continue

                if sample_had_data:
                    loaded_data_count += 1
                    debug_print(f"DEBUG: {sample_id} - Loaded {row_count} rows of existing data from Excel")
                else:
                    debug_print(f"DEBUG: {sample_id} - No existing data found in Excel")

            wb.close()
        
            if loaded_data_count > 0:
                debug_print(f"DEBUG: Successfully loaded existing data from Excel file for {loaded_data_count} samples")
            
                # Recalculate TPM values for all samples FIRST
                for i in range(self.num_samples):
                    sample_id = f"Sample {i + 1}"
                    self.calculate_tpm(sample_id)
                    debug_print(f"DEBUG: Calculated TPM for {sample_id}")

                # Update all treeviews to show the loaded data WITH calculated TPM
                if hasattr(self, 'sample_trees') and self.sample_trees:
                    for i, sample_tree in enumerate(self.sample_trees):
                        if i < self.num_samples:
                            sample_id = f"Sample {i + 1}"
                            self.update_treeview(sample_tree, sample_id)  # FIXED: Pass both tree and sample_id
                            debug_print(f"DEBUG: Updated treeview for {sample_id} with TPM values")
                else:
                    debug_print("DEBUG: sample_trees not available yet, TPM calculation completed")
                
                # Update the stats panel if available
                if hasattr(self, 'update_stats_panel'):
                    self.update_stats_panel()
            else:
                debug_print("DEBUG: No existing data found in Excel file")

        except Exception as e:
            debug_print(f"ERROR: Failed to load existing data from Excel file: {e}")
            import traceback
            traceback.print_exc()
    
    def ensure_initial_tpm_calculation(self):
        """Ensure TPM is calculated and displayed when window opens."""
        debug_print("DEBUG: Ensuring initial TPM calculation and display")
    
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"
        
            # Calculate TPM for this sample
            self.calculate_tpm(sample_id)
        
            # Update the treeview to show calculated TPM with highlighting
            if i < len(self.sample_trees):
                tree = self.sample_trees[i]
                self.update_treeview(tree, sample_id)
    
        # Update stats panel to show current TPM data
        self.update_stats_panel()
    
        debug_print("DEBUG: Initial TPM calculation and display completed")

    def setup_event_handlers(self):
        """Set up event handlers for the window."""
        # Handle window close with auto-save
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        # Set up hotkeys
        self.setup_hotkeys()
    
    def setup_hotkeys(self):
        """Set up keyboard shortcuts for navigation."""
        if not hasattr(self, 'hotkey_bindings'):
            self.hotkey_bindings = {}
    
        # Clear any existing bindings
        for key, binding_id in self.hotkey_bindings.items():
            self.window.unbind(key, binding_id)
    
        self.hotkey_bindings.clear()
    
        # Bind Ctrl+S for quick save
        binding_id = self.window.bind("<Control-s>", lambda e: self.save_data(show_confirmation=False) if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-s>"] = binding_id
    
    
        # Bind Ctrl+Left/Right for sample navigation
        binding_id = self.window.bind("<Control-Left>", lambda e: self.go_to_previous_sample() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-Left>"] = binding_id
    
        binding_id = self.window.bind("<Control-Right>", lambda e: self.go_to_next_sample() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-Right>"] = binding_id
    
    def on_window_close(self):
        """Handle window close event with auto-save."""
        self.log("Window close event triggered", "debug")
 
        # Cancel auto-save timer
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)
    
        # Auto-save if there are unsaved changes
        if self.has_unsaved_changes:
            if messagebox.askyesno("Save Changes", 
                                 "You have unsaved changes. Save before closing?"):
                try:
                    self.save_data(show_confirmation=False)
                    self.result = "load_file"
                except Exception as e:
                    messagebox.showerror("Save Error", f"Failed to save: {e}")
                    return  # Don't close if save failed
            else:
                self.result = "cancel"
        else:
            self.result = "load_file" if self.last_save_time else "cancel"
    
        self.window.destroy()
    
    def save_data(self, exit_after=False, show_confirmation=True, auto_save=False):
        """Unified method for saving data."""
        # End any active editing
        self.end_editing()
    
        # For auto-save, we'll set a flag to avoid certain behaviors
        if auto_save:
            self._auto_save_in_progress = True
    
        # Confirm save if requested and not auto-saving
        if show_confirmation and not auto_save:
            if not messagebox.askyesno("Confirm Save", "Save the collected data to the file?"):
                return False
    
        try:
            # Calculate TPM values for all samples
            for i in range(self.num_samples):
                sample_id = f"Sample {i+1}"
                self.calculate_tpm(sample_id)
        
            # Save to Excel file
            self._save_to_excel()
        
            # Update application state if needed
            if hasattr(self.parent, 'filtered_sheets'):
                self._update_application_state()
        
            # Refresh the main GUI if not auto-saving
            if not auto_save:
                self.refresh_main_gui_after_save()
        
            # Mark as saved
            self.has_unsaved_changes = False
            self.update_save_status(False)
        
            # Show confirmation if requested (not for auto-save)
            if show_confirmation and not auto_save and not exit_after:
                show_success_message("Save Complete", "Data saved successfully.", self.window)
        
            # Clean up auto-save flag
            if auto_save:
                self._auto_save_in_progress = False
        
            # Exit if requested
            if exit_after:
                self.result = "load_file"
                self.window.destroy()
            
            return True
        
        except Exception as e:
            # Clean up auto-save flag
            if auto_save:
                self._auto_save_in_progress = False
            
            error_msg = f"Failed to save data: {e}"
            self.log(error_msg, "error")
        
            if not auto_save:  # Don't show errors for auto-save
                messagebox.showerror("Save Error", error_msg)
            
            return False
    
    def initialize_data(self):
        """Initialize data structures for all samples."""
        self.data = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
    
            # Basic data structure
            self.data[sample_id] = {
                "current_row_index": 0, 
                "avg_tpm": 0.0
            }
    
            # Add standard columns (smell field will be used for both "smell" and "failure" depending on test type)
            columns = ["puffs", "before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog", "notes", "tpm"]
            for column in columns:
                self.data[sample_id][column] = []
    
            # Add special columns for User Test Simulation
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                self.data[sample_id]["chronography"] = []
    
            # Pre-initialize rows
            for j in range(50):
                puff = (j + 1) * self.puff_interval
                self.data[sample_id]["puffs"].append(puff)
        
                # Initialize other columns with empty values
                for column in columns:
                    if column == "puffs":
                        continue  # Already initialized
                    elif column == "tpm":
                        self.data[sample_id][column].append(None)
                    else:
                        self.data[sample_id][column].append("")
        
                # Initialize chronography if needed
                if "chronography" in self.data[sample_id]:
                    self.data[sample_id]["chronography"].append("")
            
            self.log(f"Initialized data for Sample {i+1}", "debug")

    def on_cancel(self):
        """Handle cancel button click or window close."""
        self.on_window_close()
    
    def show(self):
        """
        Show the window and wait for user input.
        
        Returns:
            str: "load_file" if data was saved and file should be loaded for viewing,
                 "cancel" if the user cancelled.
        """
        debug_print("DEBUG: Showing DataCollectionWindow")
        self.window.lift()
        self.window.focus_force()
        self.window.grab_set()  # Ensure it maintains focus
    
        self.window.wait_window()
        
        # Clean up auto-save timer
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)
        
        debug_print(f"DEBUG: DataCollectionWindow closed with result: {self.result}")
        return self.result  

    def create_tpm_stats_panel(self):
        """Create the enhanced TPM statistics panel with plot on the right side."""
        debug_print("DEBUG: Creating enhanced TPM statistics panel")

        # Clear existing widgets
        for widget in self.stats_frame.winfo_children():
            widget.destroy()

        # Create scrollable frame for the stats panel
        canvas = tk.Canvas(self.stats_frame, bg='white')
        scrollbar = ttk.Scrollbar(self.stats_frame, orient="vertical", command=canvas.yview)
        scrollable_stats_frame = ttk.Frame(canvas)

        scrollable_stats_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_stats_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Create a title with minimal padding
        title_label = ttk.Label(scrollable_stats_frame, text="TPM Statistics", 
                               font=("Arial", 12, "bold"), style='TLabel')
        title_label.pack(pady=(2, 5)) 

        # Create frame for current sample statistics - 1/3 of height
        stats_container = ttk.Frame(scrollable_stats_frame, style='TFrame')
        stats_container.pack(fill="x", padx=5, pady=2) 

        # Initialize TPM labels dictionary
        self.tpm_labels = {}

        # Create placeholder for current sample stats (will be updated in update_stats_panel)
        self.current_sample_stats_frame = ttk.Frame(stats_container, style='TFrame')
        self.current_sample_stats_frame.pack(fill="x", pady=2)  

        # Create frame for TPM plot - 2/3 of height with minimal padding
        plot_frame = ttk.LabelFrame(scrollable_stats_frame, text="TPM Over Time", style='TLabelframe')
        plot_frame.pack(fill="both", expand=True, padx=5, pady=(5, 2)) 
    
        # Add resize handling for the plot frame
        plot_frame.bind('<Configure>', self._on_plot_frame_resize)
    
        # Store reference to plot_frame for resize handling
        self.plot_frame = plot_frame

        # Create matplotlib figure for TPM plot with responsive sizing
        plt.style.use('default')  # Ensure we're using default style
    
        # Calculate initial size based on frame - will be updated on resize
        initial_width = 4  # Fallback width in inches
        initial_height = 3  # Fallback height in inches
    
        debug_print(f"DEBUG: Creating TPM figure with initial size: {initial_width}x{initial_height}")
    
        self.tpm_figure = plt.Figure(figsize=(initial_width, initial_height), dpi=80, 
                                    facecolor='white', tight_layout=True)
        self.tpm_ax = self.tpm_figure.add_subplot(111)

        # Create canvas with minimal padding and center it
        canvas_frame = ttk.Frame(plot_frame)
        canvas_frame.pack(fill="both", expand=True)
    
        self.tpm_canvas = FigureCanvasTkAgg(self.tpm_figure, canvas_frame)
        self.tpm_canvas.get_tk_widget().pack(fill="both", expand=True, padx=1, pady=1)  

        # Configure the plot with tight layout
        self.tpm_figure.tight_layout(pad=0.1)  

        # Initialize empty plot
        self.tpm_ax.set_xlabel('Puffs', fontsize=9)
        self.tpm_ax.set_ylabel('TPM', fontsize=9)
        self.tpm_ax.set_title('TPM Over Time', fontsize=10)
        self.tpm_ax.grid(True, alpha=0.3)

        # Apply tight layout and draw
        self.tpm_figure.tight_layout(pad=0.1)
        self.tpm_canvas.draw()

        # Update the statistics for the current sample
        self.update_stats_panel()

        debug_print("DEBUG: Enhanced TPM statistics panel created successfully")

    def _on_plot_frame_resize(self, event):
        """Handle plot frame resize to maintain responsive plot sizing."""
        # Only handle resize events for the plot_frame itself, not its children
        if event.widget != self.plot_frame:
            return
        
        try:
            # Get current frame dimensions
            frame_width = self.plot_frame.winfo_width()
            frame_height = self.plot_frame.winfo_height()
        
            # Skip if dimensions are too small or not ready
            if frame_width <= 1 or frame_height <= 1:
               
                return
            
            # Calculate figure size in inches
            dpi = 80
            padding_pixels = 4  
        
            # Use 2/3 of available height for plot (66.7%)
            available_plot_height = frame_height * 0.95  # Use almost all available space
        
            # Calculate dimensions with minimal padding
            new_width_inches = max(2.0, (frame_width - padding_pixels) / dpi)
            new_height_inches = max(1.5, (available_plot_height - padding_pixels) / dpi)
                
            # Only resize if the change is significant (avoid excessive redraws)
            current_size = self.tpm_figure.get_size_inches()
            width_diff = abs(current_size[0] - new_width_inches)
            height_diff = abs(current_size[1] - new_height_inches)
        
            if width_diff > 0.1 or height_diff > 0.1:            
                # Resize the figure
                self.tpm_figure.set_size_inches(new_width_inches, new_height_inches)
            
                # Apply very tight layout and redraw
                self.tpm_figure.tight_layout(pad=0.05)  # Very minimal padding
                self.tpm_canvas.draw()
                
        except Exception as e:
            debug_print(f"DEBUG: Error during plot resize: {e}")

    def update_plot_size(self):
        """Update plot size to fit container."""
        if hasattr(self, 'tpm_canvas') and hasattr(self, 'tpm_figure'):
            try:
                # Get the container size
                canvas_widget = self.tpm_canvas.get_tk_widget()
                canvas_widget.update_idletasks()
            
                width = canvas_widget.winfo_width()
                height = canvas_widget.winfo_height()
            
                # Only resize if we have valid dimensions
                if width > 50 and height > 50:
                    # Convert pixels to inches (80 DPI)
                    fig_width = max(3, width / 80.0)   
                    fig_height = max(2, height / 80.0) 
                
                    # Update figure size
                    self.tpm_figure.set_size_inches(fig_width, fig_height)
                    self.tpm_figure.tight_layout(pad=0.5)
                    self.tpm_canvas.draw()
                
            except Exception as e:
                debug_print(f"DEBUG: Error updating plot size: {e}")

    def update_stats_panel(self, event=None):
        """Update the TPM statistics panel to show only current sample with enhanced stats."""
        debug_print("DEBUG: Updating enhanced TPM statistics panel")

        # Get currently selected sample
        try:
            current_tab_index = self.notebook.index(self.notebook.select())
            current_sample_id = f"Sample {current_tab_index + 1}"
            debug_print(f"DEBUG: Updating stats for {current_sample_id}")
        except:
            current_sample_id = "Sample 1"  
            current_tab_index = 0

        # Clear current sample stats frame
        for widget in self.current_sample_stats_frame.winfo_children():
            widget.destroy()

        # Get sample name
        sample_name = "Unknown Sample"
        if current_tab_index < len(self.header_data.get('samples', [])):
            sample_name = self.header_data['samples'][current_tab_index].get('id', f"Sample {current_tab_index + 1}")

        # Calculate TPM values if needed
        self.calculate_tpm(current_sample_id)

        # Get TPM values and data for current sample (filtering out None values)
        tpm_values = [v for v in self.data[current_sample_id]["tpm"] if v is not None]

        # Create sample header with minimal padding
        sample_header = ttk.Label(self.current_sample_stats_frame, 
                                 text=f"{current_sample_id}: {sample_name}",
                                 font=("Arial", 11, "bold"), style='TLabel')
        sample_header.pack(anchor="w", pady=(0, 2))  

        ttk.Separator(self.current_sample_stats_frame, orient="horizontal").pack(fill="x", pady=1) 

        # Create statistics grid with minimal padding
        stat_grid = ttk.Frame(self.current_sample_stats_frame, style='TFrame')
        stat_grid.pack(fill="x", pady=2)  

        # Configure grid columns
        stat_grid.columnconfigure(0, weight=1)
        stat_grid.columnconfigure(1, weight=0)

        if tpm_values:
            # Calculate statistics
            avg_tpm = sum(tpm_values) / len(tpm_values)
            latest_tpm = tpm_values[-1]
    
            # Calculate standard deviation of last 5 sessions (or all if < 5)
            recent_tpm_values = tpm_values[-5:] if len(tpm_values) >= 5 else tpm_values
            std_dev = statistics.stdev(recent_tpm_values) if len(recent_tpm_values) > 1 else 0.0
    
            # Find current puff count (furthest down row with after_weight data)
            current_puff_count = 0
            for i in range(len(self.data[current_sample_id]["after_weight"]) - 1, -1, -1):
                if (self.data[current_sample_id]["after_weight"][i] and 
                    str(self.data[current_sample_id]["after_weight"][i]).strip()):
                    current_puff_count = self.data[current_sample_id]["puffs"][i] if i < len(self.data[current_sample_id]["puffs"]) else 0
                    break
    
            self.data[current_sample_id]["avg_tpm"] = avg_tpm
        else:
            avg_tpm = 0.0
            latest_tpm = 0.0
            std_dev = 0.0
            current_puff_count = 0
            recent_tpm_values = []

        # Row 1: Average TPM
        ttk.Label(stat_grid, text="Average TPM:", style='TLabel').grid(row=0, column=0, sticky="w", pady=1)  
        avg_tpm_label = ttk.Label(stat_grid, text=f"{avg_tpm:.6f}" if tpm_values else "N/A", 
                                 font=("Arial", 10, "bold"), style='TLabel')
        avg_tpm_label.grid(row=0, column=1, sticky="e", pady=1)

        # Row 2: Latest TPM
        ttk.Label(stat_grid, text="Latest TPM:", style='TLabel').grid(row=1, column=0, sticky="w", pady=1)
        latest_tpm_label = ttk.Label(stat_grid, text=f"{latest_tpm:.6f}" if tpm_values else "N/A", 
                                    font=("Arial", 10), style='TLabel')
        latest_tpm_label.grid(row=1, column=1, sticky="e", pady=1)

        # Row 3: Standard Deviation (last 5 sessions)
        sessions_text = f"(last {len(recent_tpm_values)} sessions)" if tpm_values else ""
        ttk.Label(stat_grid, text=f"Std Dev {sessions_text}:", style='TLabel').grid(row=2, column=0, sticky="w", pady=1)
        std_dev_label = ttk.Label(stat_grid, text=f"{std_dev:.6f}" if tpm_values else "N/A", 
                                 font=("Arial", 10), style='TLabel')
        std_dev_label.grid(row=2, column=1, sticky="e", pady=1)

        # Row 4: Current Puff Count
        ttk.Label(stat_grid, text="Current Puffs:", style='TLabel').grid(row=3, column=0, sticky="w", pady=1)
        puff_count_label = ttk.Label(stat_grid, text=str(current_puff_count), style='TLabel')
        puff_count_label.grid(row=3, column=1, sticky="e", pady=1)

        # Store references to labels for current sample
        self.tpm_labels[current_sample_id] = {
            "avg_tpm": avg_tpm_label,
            "latest_tpm": latest_tpm_label,
            "std_dev": std_dev_label,
            "puff_count": puff_count_label
        }

        # Update TPM plot for current sample
        self.update_tpm_plot(current_sample_id)

        debug_print(f"DEBUG: Enhanced stats updated for {current_sample_id}")

    def update_tpm_plot(self, sample_id):
        """Update the TPM plot for the specified sample with proper autosizing."""
        debug_print(f"DEBUG: Updating TPM plot for {sample_id}")
    
        # Clear the plot
        self.tpm_ax.clear()
    
        # Get data for the sample
        tpm_values = [v for v in self.data[sample_id]["tpm"] if v is not None]
        puff_values = []
    
        # Get corresponding puff values for non-None TPM values
        for i, tpm in enumerate(self.data[sample_id]["tpm"]):
            if tpm is not None and i < len(self.data[sample_id]["puffs"]):
                puff_values.append(self.data[sample_id]["puffs"][i])
    
        if tpm_values and puff_values and len(tpm_values) == len(puff_values):
            # Plot TPM over puffs
            self.tpm_ax.plot(puff_values, tpm_values, marker='o', linewidth=2, markersize=4, color='blue')
            self.tpm_ax.set_xlabel('Puffs', fontsize=9)
            self.tpm_ax.set_ylabel('TPM', fontsize=9)
            self.tpm_ax.set_title(f'TPM Over Time - {sample_id}', fontsize=10)
            self.tpm_ax.grid(True, alpha=0.3)
        
            # Set reasonable y-axis limits
            if max(tpm_values) <= 9:
                self.tpm_ax.set_ylim(0, 9)
            else:
                self.tpm_ax.set_ylim(0, max(tpm_values) * 1.1)
            
            # Adjust tick label sizes for better fit
            self.tpm_ax.tick_params(axis='both', which='major', labelsize=8)
        
        else:
            # Show empty plot with labels
            self.tpm_ax.set_xlabel('Puffs', fontsize=9)
            self.tpm_ax.set_ylabel('TPM', fontsize=9)
            self.tpm_ax.set_title(f'TPM Over Time - {sample_id}', fontsize=10)
            self.tpm_ax.grid(True, alpha=0.3)
            self.tpm_ax.text(0.5, 0.5, 'No TPM data available', 
                            transform=self.tpm_ax.transAxes, ha='center', va='center', fontsize=9)
    
        self.tpm_figure.tight_layout(pad=0.5)
    
        # Refresh the canvas
        self.tpm_canvas.draw()

        self.window.after(50, self.update_plot_size)
    
        debug_print(f"DEBUG: TPM plot updated for {sample_id} with autosizing")
    
    def go_to_previous_sample(self):
        """Navigate to the previous sample tab."""
        if not self.hotkeys_enabled:
            return
        current_tab = self.notebook.index(self.notebook.select())
        target_tab = (current_tab - 1) % len(self.sample_frames) 
        self.notebook.select(target_tab)

    def go_to_next_sample(self):
        """Navigate to the next sample tab."""
        if not self.hotkeys_enabled:
            return
        current_tab = self.notebook.index(self.notebook.select())
        target_tab = (current_tab + 1) % len(self.sample_frames)  
        self.notebook.select(target_tab)
    
    def get_column_name(self, col_idx):
        """Get the internal column name based on column index."""
        # Column mapping varies by test type
        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
            # User Test Simulation: [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
            column_map = {
                0: "chronography",
                1: "puffs", 
                2: "before_weight", 
                3: "after_weight", 
                4: "draw_pressure", 
                5: "smell",  # "Failure" column maps to "smell" field in data structure
                6: "notes",
                7: "tpm"  # TPM column (read-only)
            }
        else:
            # Standard test: [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM]
            column_map = {
                0: "puffs", 
                1: "before_weight", 
                2: "after_weight", 
                3: "draw_pressure", 
                4: "resistance",
                5: "smell", 
                6: "clog",
                7: "notes",
                8: "tpm"  # TPM column (read-only)
            }
    
        # Return appropriate column name or default to a safe value
        return column_map.get(col_idx, "notes")

    def create_header_section(self, parent_frame):
        """Create the header section with test information and save status."""
        header_frame = ttk.Frame(parent_frame, style='TFrame')
        header_frame.pack(fill="x", pady=(0, 10))
    
        # Left side of header
        header_left = ttk.Frame(header_frame, style='TFrame')
        header_left.pack(side="left", fill="x", expand=True)
    
        # Use styled labels
        ttk.Label(header_left, text=f"Test: {self.test_name}", style='Header.TLabel').pack(side="left")
        ttk.Label(header_left, text=f"Tester: {self.header_data['common']['tester']}", 
                 style='SubHeader.TLabel').pack(side="left", padx=(20, 0))
    
        # Right side of header - save status
        header_right = ttk.Frame(header_frame, style='TFrame')
        header_right.pack(side="right")
    
        self.save_status_label = ttk.Label(header_right, text="●", 
                                         style='SubHeader.TLabel', foreground="red")
        self.save_status_label.pack(side="right")
    
        self.save_status_text = ttk.Label(header_right, text="Unsaved changes", 
                                        style='SubHeader.TLabel')
        self.save_status_text.pack(side="right", padx=(0, 5))
    
        self.log("Header section created", "debug")

    def create_sample_tabs(self):
        """Create tabs for all samples in the notebook."""
        # Initialize lists to store references
        self.sample_frames = []
        self.sample_trees = []
    
        # Create a tab for each sample
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
        
            # Create frame for this sample
            sample_frame = ttk.Frame(self.notebook, padding=10, style='TFrame')
        
            # Get sample name safely
            sample_name = "Unknown Sample"
            if i < len(self.header_data.get('samples', [])):
                sample_name = self.header_data['samples'][i].get('id', f"Sample {i+1}")
        
            # Add tab to notebook
            self.notebook.add(sample_frame, text=f"Sample {i+1} - {sample_name}")
            self.sample_frames.append(sample_frame)
        
            # Create tab content
            tree = self.create_sample_tab(sample_frame, sample_id, i)
            self.sample_trees.append(tree)
    
        self.log(f"Created {self.num_samples} sample tabs", "debug")

    def create_control_buttons(self, parent_frame):
        """Create control buttons at the bottom of the window."""
        button_frame = ttk.Frame(parent_frame, style='TFrame')
        button_frame.pack(fill="x", pady=(10, 0))

        # Left side controls
        left_controls = ttk.Frame(button_frame, style='TFrame')
        left_controls.pack(side="left", fill="x")

        ttk.Label(left_controls, text="Puff Interval:", style='TLabel').pack(side="left")
        self.puff_interval_var = tk.IntVar(value=self.puff_interval)
    
        # Define update function inline
        def update_interval():
            try:
                self.puff_interval = self.puff_interval_var.get()
                debug_print(f"DEBUG: Updated puff interval to {self.puff_interval}")
            except:
                messagebox.showerror("Error", "Invalid puff interval. Please enter a positive number.")
                self.puff_interval_var.set(self.puff_interval)
    
        puff_spinbox = ttk.Spinbox(
            left_controls, 
            from_=1, 
            to=100, 
            textvariable=self.puff_interval_var, 
            width=5,
            command=update_interval 
        )
        puff_spinbox.pack(side="left", padx=5)

        # Sample navigation
        nav_frame = ttk.Frame(button_frame, style='TFrame')
        nav_frame.pack(side="left", padx=20)

        ttk.Button(nav_frame, text="← Prev Sample", command=self.go_to_previous_sample).pack(side="left")
        ttk.Button(nav_frame, text="Next Sample →", command=self.go_to_next_sample).pack(side="left", padx=5)

        # Right side controls - consolidated save methods
        ttk.Button(button_frame, text="Quick Save", 
                  command=lambda: self.save_data(show_confirmation=False)).pack(side="right", padx=(10, 0))
        ttk.Button(button_frame, text="Save & Exit", 
                  command=lambda: self.save_data(exit_after=True)).pack(side="right")
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side="right", padx=(0, 10))

        self.log("Control buttons created", "debug")

    def finish_edit(self, event=None):
        """Finalize editing of a cell, update data and UI, and optionally navigate."""
        if not self.editing or not hasattr(self, 'current_edit'):
            return

        edit = self.current_edit
        value = edit["entry"].get()
        sample_id = edit["sample_id"]
        tree = edit["tree"]
        item = edit["item"]
        column = edit["column"]
        column_name = edit["column_name"]
        row_idx = edit["row_idx"]

        # Attempt to convert value
        value = self.convert_cell_value(value, column_name)

        # Update internal data if changed
        if row_idx < len(self.data[sample_id][column_name]):
            old_value = self.data[sample_id][column_name][row_idx]
            if old_value != value:
                self.data[sample_id][column_name][row_idx] = value
                self.mark_unsaved_changes()

                # Auto-progression features
                if column_name == "after_weight":
                    self.auto_progress_weight(tree, sample_id, row_idx, value)
                elif column_name == "puffs":
                    self.auto_progress_puffs(tree, sample_id, row_idx, value)

        # Update Treeview display
        col_idx = int(column[1:]) - 1
        values = list(tree.item(item, "values"))
        values[col_idx] = str(value) if value != "" else ""
        tree.item(item, values=values)

        # Live TPM calculation and display update
        if column_name in ["before_weight", "after_weight", "puffs"]:
            debug_print(f"DEBUG: Triggering live TPM calculation for {sample_id} due to {column_name} change")
        
            # Calculate TPM for this specific row and potentially affected rows
            self.calculate_tpm_for_row(sample_id, row_idx)
        
            # Update the entire treeview to reflect TPM changes
            self.update_treeview(tree, sample_id)
        
            # Update stats panel
            self.update_stats_panel()
        
            debug_print(f"DEBUG: Live TPM update completed for {sample_id}")

        # Finish edit mode before navigating
        self.end_editing()

        # Optional keyboard navigation
        if event and event.keysym in ["Right", "Left"]:
            direction = "right" if event.keysym == "Right" else "left"
            self.handle_arrow_key(event, tree, sample_id, direction)

    def calculate_tpm_for_row(self, sample_id, changed_row_idx):
        """Calculate TPM for a specific row and related rows when data changes."""
        debug_print(f"DEBUG: Calculating TPM for {sample_id}, row {changed_row_idx}")
    
        # Calculate TPM for the changed row
        self.calculate_single_row_tpm(sample_id, changed_row_idx)
    
        # If puffs changed, recalculate all subsequent rows since intervals may have changed
        if changed_row_idx < len(self.data[sample_id]["puffs"]) - 1:
            for i in range(changed_row_idx + 1, len(self.data[sample_id]["puffs"])):
                self.calculate_single_row_tpm(sample_id, i)
    
        # Update average TPM
        valid_tpm_values = [v for v in self.data[sample_id]["tpm"] if v is not None]
        self.data[sample_id]["avg_tpm"] = sum(valid_tpm_values) / len(valid_tpm_values) if valid_tpm_values else 0.0

    def calculate_single_row_tpm(self, sample_id, row_idx):
        """Calculate TPM for a single row."""
        try:
            # Ensure TPM list is long enough
            while len(self.data[sample_id]["tpm"]) <= row_idx:
                self.data[sample_id]["tpm"].append(None)
        
            # Get weight values for this row
            before_weight_str = self.data[sample_id]["before_weight"][row_idx] if row_idx < len(self.data[sample_id]["before_weight"]) else ""
            after_weight_str = self.data[sample_id]["after_weight"][row_idx] if row_idx < len(self.data[sample_id]["after_weight"]) else ""
        
            # Skip if either weight is missing
            if not before_weight_str or not after_weight_str:
                self.data[sample_id]["tpm"][row_idx] = None
                return
            
            # Convert to float
            before_weight = float(before_weight_str)
            after_weight = float(after_weight_str)
        
            # Validate weights
            if before_weight <= after_weight:
                self.data[sample_id]["tpm"][row_idx] = None
                return
            
            # Calculate puffs in this interval
            current_puff = int(self.data[sample_id]["puffs"][row_idx]) if row_idx < len(self.data[sample_id]["puffs"]) else 0
            puffs_in_interval = current_puff
        
            if row_idx > 0:
                prev_puff = int(self.data[sample_id]["puffs"][row_idx - 1]) if (row_idx - 1) < len(self.data[sample_id]["puffs"]) else 0
                puffs_in_interval = current_puff - prev_puff
        
            # Skip if invalid puff interval
            if puffs_in_interval <= 0:
                self.data[sample_id]["tpm"][row_idx] = None
                return
            
            # Calculate TPM (mg/puff)
            weight_consumed = before_weight - after_weight  # in grams
            tpm = (weight_consumed * 1000) / puffs_in_interval  # Convert to mg per puff
        
            # Store result
            self.data[sample_id]["tpm"][row_idx] = round(tpm, 6)
        
            debug_print(f"DEBUG: Calculated TPM for {sample_id} row {row_idx}: {tpm:.6f} mg/puff")
        
        except (ValueError, TypeError, IndexError) as e:
            debug_print(f"DEBUG: Error calculating TPM for {sample_id} row {row_idx}: {e}")
            self.data[sample_id]["tpm"][row_idx] = None

    def convert_cell_value(self, value, column_name):
        """Convert user-entered value to appropriate type."""
        value = value.strip()
        if not value:
            return "" if column_name != "puffs" else 0

        try:
            if column_name == "puffs":
                return int(float(value))
            elif column_name in ["before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog"]:
                return float(value)
            elif column_name in ["chronography", "notes"]:
                return str(value)
        except ValueError:
            return value  # Keep as-is if conversion fails
        return value

    def auto_progress_weight(self, tree, sample_id, row_idx, value):
        """Auto-fill the next row's before_weight with current after_weight."""
        if value in ["", 0, None]:
            return
        try:
            val = float(value)
            next_row = row_idx + 1
            if next_row < len(self.data[sample_id]["before_weight"]):
                self.data[sample_id]["before_weight"][next_row] = val
                if next_row < len(tree.get_children()):
                    next_item = tree.get_children()[next_row]
                    values = list(tree.item(next_item, "values"))
                    col_idx = 2 if self.test_name in ["User Test Simulation", "User Simulation Test"] else 1
                    values[col_idx] = str(val)
                    tree.item(next_item, values=values)
                    debug_print(f"DEBUG: Auto-set Sample {sample_id} row {next_row} before_weight to {val}")
        except (ValueError, TypeError):
            debug_print("DEBUG: Failed auto weight progression due to invalid value")

    def auto_progress_puffs(self, tree, sample_id, row_idx, value):
        """Auto-fill puffs for subsequent rows unless test is user-driven."""
        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
            debug_print("DEBUG: Skipping auto-puff progression for user simulation test")
            return

        try:
            start_val = int(value)
            for i in range(row_idx + 1, len(self.data[sample_id]["puffs"])):
                increment = (i - row_idx) * self.puff_interval
                puff_val = start_val + increment
                self.data[sample_id]["puffs"][i] = puff_val
                if i < len(tree.get_children()):
                    item = tree.get_children()[i]
                    values = list(tree.item(item, "values"))
                    puff_col = 1 if self.test_name == "User Test Simulation" else 0
                    values[puff_col] = str(puff_val)
                    tree.item(item, values=values)
            debug_print(f"DEBUG: Auto-puff progression applied from row {row_idx}")
        except (ValueError, TypeError):
            debug_print("DEBUG: Error in auto-puff progression")

    def edit_cell(self, tree, item, column, row_idx, sample_id):
        """Single unified method for cell editing."""
        # Get column name based on column index
        col_idx = int(column[1:]) - 1
        column_name = self.get_column_name(col_idx)
    
        # Check if this is the TPM column - make it read-only
        if column_name == "tpm":
            debug_print("DEBUG: TPM column is read-only (calculated automatically)")
            return  # Don't allow editing of TPM column
    
        # Check if the column exists in our data structure
        if column_name not in self.data[sample_id]:
            debug_print(f"DEBUG: Column {column_name} not found in data structure")
            return
    
        # End any existing edit first
        self.end_editing()
    
        # Get cell coordinates and create editing frame
        x, y, width, height = tree.bbox(item, column)
        frame = tk.Frame(tree, borderwidth=0, highlightthickness=1, highlightbackground="black")
        frame.place(x=x, y=y, width=width, height=height)
    
        # Create entry widget with current value
        entry = tk.Entry(frame, borderwidth=0)
        entry.pack(fill="both", expand=True)
    
        # Set current value
        current_value = self.data[sample_id][column_name][row_idx] if row_idx < len(self.data[sample_id][column_name]) else ""
        if current_value:
            entry.insert(0, str(current_value))
        entry.select_range(0, tk.END)
    
        # Store edit state
        self.editing = True
        self.hotkeys_enabled = False
        self.current_edit = {
            "frame": frame, "entry": entry, "tree": tree, "item": item,
            "column": column, "column_name": column_name, "row_idx": row_idx, "sample_id": sample_id
        }
    
        # Set up event bindings
        entry.focus_set()
        entry.bind("<Return>", self.finish_edit)
        entry.bind("<Tab>", self.handle_tab_in_edit)
        entry.bind("<Escape>", self.cancel_edit)
        entry.bind("<FocusOut>", self.finish_edit)
        entry.bind("<Left>", self.handle_arrow_in_edit)
        entry.bind("<Right>", self.handle_arrow_in_edit)
    
        self.log(f"Started editing cell - sample: {sample_id}, row: {row_idx}, column: {column_name}", "debug")
    
    def on_tree_click(self, event, tree, sample_id):
        """Unified method for handling tree clicks."""
        import time
    
        item = tree.identify("item", event.x, event.y)
        column = tree.identify("column", event.x, event.y)
        region = tree.identify("region", event.x, event.y)
    
        # Ignore clicks outside cells
        if not item or not column or region != "cell":
            return
    
        # Store current selection
        self.current_item = item
        self.current_column = column
    
        # Check for double-click (based on time since last click)
        current_time = time.time() * 1000
        is_double_click = (current_time - self.last_click_time < 300 and
                          item == getattr(self, 'last_clicked_item', None) and
                          column == getattr(self, 'last_clicked_column', None))
    
        if is_double_click:
            # Start editing on double-click
            row_idx = tree.index(item)
            self.edit_cell(tree, item, column, row_idx, sample_id)
        else:
            # Single click - just highlight cell
            self.highlight_cell(tree, item, column)
    
        # Update click tracking
        self.last_click_time = current_time
        self.last_clicked_item = item
        self.last_clicked_column = column

    def handle_tab_key(self, event, tree, sample_id):
        """Handle Tab key press in the treeview."""
        if self.editing:
            return "break"  # Already handled in edit mode
        
        # Get the current selection
        item = tree.focus()
        if not item:
            # Select the first item and the first column (puffs - now editable)
            if tree.get_children():
                item = tree.get_children()[0]
                tree.selection_set(item)
                tree.focus(item)
                self.edit_cell(tree, item, "#1", 0, sample_id, 'puffs')  # Start with puffs column
            return "break"
        
        # Get the current column
        column = tree.identify_column(event.x) if event else None
        if not column:
            # Start at the first column (puffs)
            column = "#1"
        
        # Get column index (1-based)
        col_idx = int(column[1:])
    
        # Get next column
        next_col_idx = col_idx + 1
    
        # If at the last column, move to the first column of the next row
        if next_col_idx > 6:  # We have 6 columns
            # Get the next item
            items = tree.get_children()
            idx = items.index(item)
        
            if idx < len(items) - 1:
                # Move to next row
                next_item = items[idx + 1]
                tree.selection_set(next_item)
                tree.focus(next_item)
                tree.see(next_item)
            
                # Edit the first column (puffs - now editable)
                self.edit_cell(tree, next_item, "#1", idx + 1, sample_id, 'puffs')
            else:
                # At the last row, move to next sample
                self.go_to_next_sample()
        else:
            # Edit the next column in the same row
            self.current_column = f"#{next_col_idx}"
            self.highlight_cell(tree, item, self.current_column)
        
        return "break"  # Stop event propagation

    def start_edit_on_typing(self, event, tree, sample_id):
        """Start editing if a debug_printable character is typed while a cell is selected."""
        if not event.char.isprintable():
            return  # Skip control keys

        item = tree.focus()
        column = getattr(self, 'current_column', '#2')

        if not item:
            items = tree.get_children()
            if not items:
                return
            item = items[0]
            tree.selection_set(item)
            tree.focus(item)

        row_idx = tree.index(item)
    
        # Start editing with the key typed
        self.edit_cell(tree, item, column, row_idx, sample_id)
    
        if self.current_edit and self.current_edit["entry"]:
            entry = self.current_edit["entry"]
            entry.delete(0, tk.END)
            entry.insert(0, event.char)
            entry.icursor(1)

    def handle_tab_in_edit(self, event):
        """Handle tab key pressed while editing a cell."""
        self.log("Tab pressed during edit", "debug")
    
        # Store the current edit info before finishing
        current_tree = self.current_edit["tree"]
        current_sample_id = self.current_edit["sample_id"]
    
        # Save the current edit
        self.finish_edit(event)
    
        # Use the handle_tab_navigation method
        return self.handle_tab_navigation(event)

    def handle_tab_navigation(self, event):
        """Handle tab key navigation during editing."""
        if not self.editing or not self.current_edit:
            return
    
        # Store current edit values before finishing
        tree = self.current_edit["tree"]
        sample_id = self.current_edit["sample_id"]
        column = self.current_edit["column"]
        row_idx = self.current_edit["row_idx"]
    
        # Get column index and direction
        col_idx = int(column[1:])
        shift_pressed = event.state & 0x0001  # Check if Shift key is pressed
        direction = "left" if shift_pressed else "right"
    
        # Finish current edit
        self.finish_edit()
    
        # Navigate based on direction
        items = tree.get_children()
    
        if direction == "right":
            # Navigate to next column or row
            next_col_idx = col_idx + 1
        
            if next_col_idx <= self.tree_columns:
                # Go to next column in same row
                next_column = f"#{next_col_idx}"
                self.edit_cell(tree, self.current_item, next_column, row_idx, sample_id)
            elif row_idx < len(items) - 1:
                # Go to first column of next row
                next_item = items[row_idx + 1]
                tree.selection_set(next_item)
                tree.focus(next_item)
                tree.see(next_item)
                self.current_item = next_item
                self.edit_cell(tree, next_item, "#1", row_idx + 1, sample_id)
        else:
            # Navigate to previous column or row
            prev_col_idx = col_idx - 1
        
            if prev_col_idx >= 1:
                # Go to previous column in same row
                prev_column = f"#{prev_col_idx}"
                self.edit_cell(tree, self.current_item, prev_column, row_idx, sample_id)
            elif row_idx > 0:
                # Go to last column of previous row
                prev_item = items[row_idx - 1]
                tree.selection_set(prev_item)
                tree.focus(prev_item)
                tree.see(prev_item)
                self.current_item = prev_item
                self.edit_cell(tree, prev_item, f"#{self.tree_columns}", row_idx - 1, sample_id)
    
        return "break"  # Stop event propagation

    def handle_arrow_key(self, event, tree, sample_id, direction):
        """Handle arrow key press in the treeview."""
        if self.editing:
            return  # Let the entry widget handle arrow keys when editing
        
        # Get the current selection
        item = tree.focus()
        if not item:
            return "break"
        
        # Get current column if stored, otherwise use first editable column
        current_column = getattr(self, 'current_column', '#2')
        col_idx = int(current_column[1:])
    
        if direction in ["left", "right"]:
            items = tree.get_children()
            row_idx = items.index(item)

            if direction == "right":
                if col_idx < 6:
                    col_idx += 1
                elif row_idx < len(items) - 1:
                    row_idx += 1
                    col_idx = 2
                else:
                    return "break"
            elif direction == "left":
                if col_idx > 2:
                    col_idx -= 1
                elif row_idx > 0:
                    row_idx -= 1
                    col_idx = 6
                else:
                    return "break"

            next_item = items[row_idx]
            next_column = f"#{col_idx}"
            self.current_column = next_column
            self.highlight_cell(tree, next_item, next_column)
            debug_print(f"DEBUG: Arrow navigation - moved to row {row_idx}, column {next_column}")
        
        elif direction in ["up", "down"]:
            # Vertical navigation
            items = tree.get_children()
            current_idx = items.index(item)
        
            if direction == "down":
                if current_idx < len(items) - 1:
                    next_item = items[current_idx + 1]
                    tree.selection_set(next_item)
                    tree.focus(next_item)
                    tree.see(next_item)
                    self.current_column = current_column
                    self.highlight_cell(tree, next_item, self.current_column)
                    debug_print(f"DEBUG: Arrow navigation - moved down to row {current_idx + 1}")
            else:  # up
                if current_idx > 0:
                    prev_item = items[current_idx - 1]
                    tree.selection_set(prev_item)
                    tree.focus(prev_item)
                    tree.see(prev_item)
                    self.current_column = current_column
                    self.highlight_cell(tree, prev_item, self.current_column)
                    debug_print(f"DEBUG: Arrow navigation - moved up to row {current_idx - 1}")
    
        return "break"  # Stop event propagation

    def handle_arrow_in_edit(self, event):
        """Handle arrow keys pressed while editing a cell."""
        # Only handle left/right arrows
        if event.keysym not in ["Left", "Right"]:
            return
    
        # Check if at beginning or end of text
        entry = self.current_edit["entry"]
        cursor_pos = entry.index(tk.INSERT)
    
        # If at beginning and pressing left, or at end and pressing right, navigate to next cell
        if (cursor_pos == 0 and event.keysym == "Left") or \
           (cursor_pos == len(entry.get()) and event.keysym == "Right"):
            self.log(f"Arrow key navigation from edit - {event.keysym}", "debug")
            self.finish_edit(event)
            return "break"  # Stop event propagation

    def cancel_edit(self, event=None):
        self.end_editing()
        return "break"

    def end_editing(self):
        if not self.editing:
            return
        if self.current_edit and "frame" in self.current_edit:
            self.current_edit["frame"].destroy()
        self.current_edit = None
        self.editing = False
        self.hotkeys_enabled = True

    def highlight_cell(self, tree, item, column):
        """Highlight a specific cell using enhanced treeview selection."""
        # Clear previous selection
        tree.selection_set()
    
        # Select the item and focus on it
        tree.selection_set(item)
        tree.focus(item)
        tree.see(item)
    
        # Store current selection for keyboard navigation
        self.current_item = item
        self.current_column = column
    
        # Add visual emphasis by applying a special tag
        current_tags = list(tree.item(item, 'tags'))
    
        # Remove any existing highlight tags from other items
        for child in tree.get_children():
            tags = list(tree.item(child, 'tags'))
            tags = [tag for tag in tags if not tag.startswith('highlight_')]
            tree.item(child, tags=tags)
    
        # Add highlight tag to current item
        if 'highlight_selected' not in current_tags:
            current_tags.append('highlight_selected')
            tree.item(item, tags=current_tags)
    
        debug_print(f"DEBUG: Cell highlighted - item: {item}, column: {column}")

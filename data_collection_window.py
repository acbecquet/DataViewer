"""
data_collection_window.py
Developed by Charlie Becquet.
Interface for rapid test data collection with enhanced saving and menu functionality.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tksheet import Sheet
import datetime
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
from utils import FONT, debug_print, show_success_message, load_excel_file_with_formulas
import threading
import subprocess

from data_collection_ui import DataCollectionUI
from data_collection_handlers import DataCollectionHandlers
from data_collection_data import DataCollectionData
from data_collection_file_io import DataCollectionFileIO

def setup_logging():
    """Set up logging with proper error handling and writable directory."""
    try:
        # Use user's home directory for log files (always writable)
        log_dir = os.path.expanduser("~/.dataviewer")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'datacollection.log')

        # Configure logging with file and console handlers
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        debug_print(f"DEBUG: Logging configured successfully. Log file: {log_file}")

    except Exception as e:
        # Fallback to console-only logging if file logging fails
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )
        debug_print(f"WARNING: Could not set up file logging ({e}). Using console only.")

# Call the setup function
setup_logging()

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

    "Big Headspace High T Test": """
SOP - Big Headspace High T Test:

Drain device to 30% remaining and then place in the oven at 40C.
After 1 hour, collect 10 puffs, tracking TPM and draw resistance for each puff. Repeat this 3 times, for 30 total puffs.
Be sure to take detailed notes on bubbling and any failure modes.

Key Points:
- 40C big headspace test
- Monitor for thermal effects
- Document temperature-related observations
    """,

    "Big Headspace Serial Test": """
SOP - Big Headspace Serial Test:

Drain device to 30% remaining and then place in the oven upright at 40C.
After 1 hour, collect 10 puffs, tracking TPM and draw resistance for each puff. Repeat this 3 times,
the second time placing the samples horizontal, airway up, and the third placed horizontal, airway down.
Be sure to take detailed notes on clogging or any other failure modes.

Key Points:
- 40C big headspace test
- Monitor for thermal effects
- Document temperature-related observations
    """,

    "Big Headspace Low T Test": """
SOP - Big Headspace Low T Test:

Drain device to 30% remaining and then place in refrigerator at 4C.
After 1 hour, collect 10 puffs, tracking TPM and draw resistance for each puff. Repeat this 3 times, for 30 total puffs.
Be sure to take detailed notes on viscosity changes and any failure modes.

Key Points:
- 4C big headspace test
- Monitor for cold temperature effects
- Document viscosity and flow changes
    """,

    "Extended Test": """
SOP - Extended Test:

Long-duration testing to assess device lifetime and performance degradation.
Sessions of 10 puffs with a 60mL/3s/30s puffing regime. Rest 15 minutes between sessions.
Monitor for performance and consistency over time. Measure initial and final draw resistance,
and measure TPM every 10 puffs.

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

    "Lifetime Test": """
SOP - Lifetime Test:

Evaluation of Performance over a device lifetime. Observe TPM change over time, draw pressure change, and clogging or oil accumulation.
Monitor device until complete depletion or failure.

Key Points:
- Full device lifetime assessment
- Track performance degradation
- Document failure modes
- Measure until device end-of-life
    """,

    "Device Life Test": """
SOP - Device Life Test:

Comprehensive evaluation of device performance throughout its operational lifetime.
Test until device failure or oil depletion. Record TPM, draw resistance, and any performance changes.
Document all failure modes and performance degradation patterns.

Key Points:
- Complete lifecycle testing
- Performance tracking over time
- Failure mode documentation
- End-of-life characterization
    """,

    "Horizontal Puffing Test": """
SOP - Horizontal Puffing Test:

Assessment of device performance while placed horizontally.
Focus on key performance indicators such as clogging and burn.
Test device in horizontal position for entire duration.

Key Points:
- Horizontal orientation testing
- Monitor for position-related issues
- Track clogging and burn performance
- Compare to vertical performance
    """,

    "Long Puff Test": """
SOP - Long Puff Test:

Extended puff duration testing to evaluate device performance under prolonged draw conditions.
Use extended puff durations (5-10 seconds) at standard volume. Monitor for overheating and performance changes.
Record TPM, draw resistance, and any thermal effects.

Key Points:
- Extended puff duration protocol
- Monitor thermal performance
- Track device response to long draws
- Document overheating issues
    """,

    "Rapid Puff Test": """
SOP - Rapid Puff Test:

High-frequency puffing protocol to stress-test device performance.
Conduct rapid successive puffs with minimal rest intervals (5-10 seconds between puffs).
Monitor for overheating, performance degradation, and device failure.

Key Points:
- High-frequency puffing protocol
- Minimal rest intervals
- Stress testing conditions
- Monitor thermal and performance effects
    """,

    "Viscosity Compatibility": """
SOP - Viscosity Compatibility Test:

Test device performance with oils of varying viscosity levels.
Use oils ranging from low (20-50 cP) to high viscosity (200-500 cP).
Record TPM, draw resistance, and flow characteristics for each viscosity level.

Key Points:
- Multiple viscosity levels
- Flow performance assessment
- Compatibility evaluation
- Document viscosity-related issues
    """,

    "Upside Down Test": """
SOP - Upside Down Test:

Evaluate device performance when inverted (upside down orientation).
Test for leakage, flow disruption, and performance changes.
Monitor for air bubble formation and oil flow issues.

Key Points:
- Inverted orientation testing
- Monitor for leakage
- Assess flow performance
- Document orientation-related effects
    """,

    "Big Headspace Pocket Test": """
SOP - Big Headspace Pocket Test:

Drain device to 30% remaining and simulate pocket storage conditions.
Subject device to body temperature (37C) and mechanical stress.
Test performance after exposure to pocket-like conditions.

Key Points:
- Body temperature exposure
- Mechanical stress simulation
- Post-exposure performance testing
- Document storage-related effects
    """,

    "Low Temperature Stability": """
SOP - Low Temperature Stability Test:

Evaluate device performance and oil stability at low temperatures.
Store device at 4C for extended period, then test performance.
Monitor for crystallization, viscosity changes, and flow issues.

Key Points:
- Low temperature storage
- Oil stability assessment
- Performance after cold exposure
- Document temperature-related changes
    """,

    "Vacuum Test": """
SOP - Vacuum Test:

Test device performance under reduced atmospheric pressure conditions.
Simulate high-altitude or vacuum environments.
Monitor for vapor formation, oil expansion, and performance changes.

Key Points:
- Reduced pressure conditions
- Monitor vapor formation
- Assess performance under vacuum
- Document pressure-related effects
    """,

    "Negative Pressure Test": """
SOP - Negative Pressure Test:

Evaluate device response to negative pressure conditions.
Test for structural integrity and performance under reduced pressure.
Monitor for oil expansion, air bubble formation, and flow disruption.

Key Points:
- Negative pressure conditions
- Structural integrity assessment
- Flow performance evaluation
- Document pressure-related issues
    """,

    "Various Oil Compatibility": """
SOP - Various Oil Compatibility Test:

Test device compatibility with different oil formulations and additives.
Use representative oil types including different carrier oils, viscosity modifiers, and additives.
Evaluate performance, compatibility, and any material interactions.

Key Points:
- Multiple oil formulations
- Compatibility assessment
- Material interaction evaluation
- Performance comparison across oils
    """,

    "Sheet1": """
SOP - General Test Protocol:

Standard testing protocol for miscellaneous or custom test configurations.
Follow appropriate measurement procedures based on test objectives.
Record all relevant data and observations.

Key Points:
- Flexible test protocol
- Standard measurement procedures
- Comprehensive data recording
- Adapt to specific test requirements
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

class DataCollectionWindow(DataCollectionUI, DataCollectionHandlers, DataCollectionData, DataCollectionFileIO):
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
        self.updating_notes = False

        if hasattr(parent, 'root'):
            self.main_window_was_visible = parent.root.winfo_viewable()
            parent.root.withdraw()  # Hide main window
            debug_print("DEBUG: Main GUI window hidden")

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
        self.window.state('zoomed')
        self.window.minsize(1250, 625)

        self.window.lift()  # Bring to top of window stack
        self.window.focus_force()  # Force focus to this window

        # Make it stay on top temporarily
        self.window.attributes('-topmost', True)
        self.window.after(100, lambda: self.window.attributes('-topmost', False))

        self.sample_images = {}
        self.sample_image_crop_states = {}

        # Default puff interval
        self.puff_interval = 10  # Default to 10

        # Initialize tab tracking for notes handling
        self.previous_tab_index = 0
        self.updating_notes = False
        debug_print("DEBUG: Initialized tab change tracking variables")

        # Set up keyboard shortcut flags
        self.hotkeys_enabled = True
        self.hotkey_bindings = {}

        # Create the style for ttk widgets
        self.style = ttk.Style()
        self.setup_styles()

        # Initialize data structures
        self.initialize_data()

        # Create the menu bar first
        self.create_menu_bar()

        # Create the UI
        self.create_widgets()

        self.load_existing_data_from_file()

        self.load_existing_sample_images_from_vap3()

        # Center the window
        self.center_window()

        # Set up event handlers
        self.setup_event_handlers()

        # Start auto-save timer
        self.start_auto_save_timer()

        self.ensure_initial_tpm_calculation()

        self.log(f"DataCollectionWindow initialized for {test_name} with {self.num_samples} samples", "debug")

    def calculate_dynamic_plot_size(self, parent_frame):
        """Calculate plot size that directly scales with window size."""
        debug_print("DEBUG: Starting dynamic plot size calculation")

        # Force geometry update to get current dimensions
        self.window.update_idletasks()

        if hasattr(self, 'stats_frame') and self.stats_frame.winfo_exists():
            parent_frame = self.stats_frame
            parent_frame.update_idletasks()

        # Get the actual dimensions
        available_width = parent_frame.winfo_width()
        available_height = parent_frame.winfo_height()

        debug_print(f"DEBUG: Parent frame: {parent_frame.__class__.__name__}")
        debug_print(f"DEBUG: Available dimensions - Width: {available_width}px, Height: {available_height}px")

        # Simple fallback for initial sizing or very small windows
        if available_width < 200 or available_height < 200:
            debug_print("DEBUG: Using fallback size for small window")
            return (6, 4)

        # Reserve space for stats text and controls
        text_space = 120  # Space for text statistics
        control_space = 50  # Space for labels and padding

        # Calculate available space for the plot
        plot_height_available = available_height - text_space - control_space
        plot_width_available = available_width - 40  # 20px margin on each side

        # Use most of the available space for the plot
        plot_width_pixels = max(plot_width_available, 200)
        plot_height_pixels = max(plot_height_available, 150)

        debug_print(f"DEBUG: Plot space in pixels - Width: {plot_width_pixels}px, Height: {plot_height_pixels}px")

        # Convert to inches for matplotlib (using standard 100 DPI)
        plot_width_inches = plot_width_pixels / 100.0
        plot_height_inches = plot_height_pixels / 100.0

        # Apply minimum and maximum sizes
        plot_width_inches = max(min(plot_width_inches, 12.0), 3.0)
        plot_height_inches = max(min(plot_height_inches, 8.0), 2.0)

        debug_print(f"DEBUG: FINAL plot size - Width: {plot_width_inches:.2f} inches, Height: {plot_height_inches:.2f} inches")

        return (plot_width_inches, plot_height_inches)

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
            if hasattr(self, 'sample_sheets') and current_tab < len(self.sample_sheets):
                sheet = self.sample_sheets[current_tab]
                self.update_tksheet(sheet, sample_id)
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
                if hasattr(self, 'sample_sheets') and i < len(self.sample_sheets):
                    sheet = self.sample_sheets[i]
                    self.update_tksheet(sheet, sample_id)

            self.update_stats_panel()
            self.mark_unsaved_changes()
            debug_print("DEBUG: Cleared all data for all samples")

    def add_row(self):
        """Add a new row to all samples."""
        debug_print("DEBUG: Adding new row to all samples")

        # Get the current sample instead of checking all samples
        try:
            current_tab_index = self.notebook.index(self.notebook.select())
            current_sample_id = f"Sample {current_tab_index + 1}"
            debug_print(f"DEBUG: Current sample for add row: {current_sample_id}")
        except:
            current_sample_id = "Sample 1"
            current_tab_index = 0
            debug_print("DEBUG: Defaulting to Sample 1 for add row")

        # Find the last non-empty puff value in the current sample
        last_puff = 0
        for puff_val in reversed(self.data[current_sample_id]["puffs"]):
            if puff_val and str(puff_val).strip():
                try:
                    last_puff = int(float(puff_val))
                    debug_print(f"DEBUG: Found last puff in {current_sample_id}: {last_puff}")
                    break
                except (ValueError, TypeError):
                    continue

        # Calculate new puff value for current sample
        new_puff = last_puff + self.puff_interval
        debug_print(f"DEBUG: Calculated new_puff: {new_puff} (last puff: {last_puff}, interval: {self.puff_interval})")

        # Add row to all samples, but use appropriate puff values for each
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"

            # Ensure all required fields exist in the data structure (defensive check)
            required_fields = ["puffs", "before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog", "notes", "tpm"]
            for field in required_fields:
                if field not in self.data[sample_id]:
                    self.data[sample_id][field] = []
                    debug_print(f"DEBUG: Added missing field '{field}' to {sample_id}")

            # Add chronography for User Test Simulation if needed
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                if "chronography" not in self.data[sample_id]:
                    self.data[sample_id]["chronography"] = []

            if sample_id == current_sample_id:
                # Use calculated puff for current sample
                self.data[sample_id]["puffs"].append(new_puff)
            else:
                # For other samples, find their last puff and add interval
                last_sample_puff = 0
                for puff_val in reversed(self.data[sample_id]["puffs"]):
                    if puff_val and str(puff_val).strip():
                        try:
                            last_sample_puff = int(float(puff_val))
                            break
                        except (ValueError, TypeError):
                            continue

                other_sample_new_puff = last_sample_puff + self.puff_interval
                self.data[sample_id]["puffs"].append(other_sample_new_puff)
                debug_print(f"DEBUG: Added puff {other_sample_new_puff} to {sample_id}")

            # Add empty values for other columns
            self.data[sample_id]["before_weight"].append("")
            self.data[sample_id]["after_weight"].append("")
            self.data[sample_id]["draw_pressure"].append("")
            self.data[sample_id]["smell"].append("")
            self.data[sample_id]["notes"].append("")
            self.data[sample_id]["tpm"].append(None)

            # Add resistance and clog for standard tests
            if self.test_name not in ["User Test Simulation", "User Simulation Test"]:
                self.data[sample_id]["resistance"].append("")
                self.data[sample_id]["clog"].append("")

            # Add chronography for User Test Simulation
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                self.data[sample_id]["chronography"].append("")

        # Update all tksheets to show the new row
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"
            self.update_tksheet(sample_id)

        # Update TPM calculations and stats
        self.update_all_tpm_calculations()

        # Mark as having unsaved changes
        self.mark_unsaved_changes()

        debug_print(f"DEBUG: Added new row with puff count {new_puff} to {current_sample_id}")

    def add_multiple_rows(self):
        """Add multiple rows to all samples with user input."""
        debug_print("DEBUG: Adding multiple rows to all samples")

        # Create a dialog to get the number of rows
        dialog = tk.Toplevel(self.window)
        dialog.title("Add Multiple Rows")
        dialog.geometry("300x150")
        dialog.transient(self.window)
        dialog.grab_set()
        dialog.resizable(False, False)

        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        # Result variable
        result = {"rows": 0, "confirmed": False}

        # Create the dialog content
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Label
        ttk.Label(main_frame, text="Number of rows to add:", font=("Arial", 10)).pack(pady=(0, 10))

        # Entry with validation
        rows_var = tk.IntVar(value=1)
        entry = ttk.Spinbox(
            main_frame,
            from_=1,
            to=100,
            textvariable=rows_var,
            width=10,
            justify="center"
        )
        entry.pack(pady=(0, 20))
        entry.focus_set()

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")

        def on_ok():
            try:
                num_rows = rows_var.get()
                if num_rows < 1:
                    messagebox.showerror("Invalid Input", "Number of rows must be at least 1.")
                    return
                result["rows"] = num_rows
                result["confirmed"] = True
                dialog.destroy()
            except tk.TclError:
                messagebox.showerror("Invalid Input", "Please enter a valid number.")

        def on_cancel():
            result["confirmed"] = False
            dialog.destroy()

        # Buttons
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="right")
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side="right", padx=(0, 5))

        # Bind Enter key to OK
        dialog.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())

        # Wait for dialog to close
        dialog.wait_window()

        # If user confirmed, add the rows
        if result["confirmed"] and result["rows"] > 0:
            num_rows_to_add = result["rows"]
            debug_print(f"DEBUG: User requested to add {num_rows_to_add} rows")

            # Get the current sample instead of checking all samples
            try:
                current_tab_index = self.notebook.index(self.notebook.select())
                current_sample_id = f"Sample {current_tab_index + 1}"
                debug_print(f"DEBUG: Current sample for add multiple rows: {current_sample_id}")
            except:
                current_sample_id = "Sample 1"
                current_tab_index = 0
                debug_print("DEBUG: Defaulting to Sample 1 for add multiple rows")

            # Find the last non-empty puff value in the current sample
            last_puff = 0
            for puff_val in reversed(self.data[current_sample_id]["puffs"]):
                if puff_val and str(puff_val).strip():
                    try:
                        last_puff = int(float(puff_val))
                        debug_print(f"DEBUG: Found last puff in {current_sample_id}: {last_puff}")
                        break
                    except (ValueError, TypeError):
                        continue

            # Add the specified number of rows
            for row_num in range(num_rows_to_add):
                # Calculate new puff value for current sample
                new_puff = last_puff + ((row_num + 1) * self.puff_interval)
                debug_print(f"DEBUG: Adding row {row_num + 1}/{num_rows_to_add} with puff {new_puff}")

                # Add row to all samples
                for i in range(self.num_samples):
                    sample_id = f"Sample {i + 1}"

                    # Ensure all required fields exist in the data structure (defensive check)
                    required_fields = ["puffs", "before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog", "notes", "tpm"]
                    for field in required_fields:
                        if field not in self.data[sample_id]:
                            self.data[sample_id][field] = []
                            debug_print(f"DEBUG: Added missing field '{field}' to {sample_id}")

                    # Add chronography for User Test Simulation if needed
                    if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                        if "chronography" not in self.data[sample_id]:
                            self.data[sample_id]["chronography"] = []

                    if sample_id == current_sample_id:
                        # Use calculated puff for current sample
                        self.data[sample_id]["puffs"].append(new_puff)
                    else:
                        # For other samples, find their last puff and add interval
                        last_sample_puff = 0
                        for puff_val in reversed(self.data[sample_id]["puffs"]):
                            if puff_val and str(puff_val).strip():
                                try:
                                    last_sample_puff = int(float(puff_val))
                                    break
                                except (ValueError, TypeError):
                                    continue

                        other_sample_new_puff = last_sample_puff + ((row_num + 1) * self.puff_interval)
                        self.data[sample_id]["puffs"].append(other_sample_new_puff)

                    # Add empty values for other columns
                    self.data[sample_id]["before_weight"].append("")
                    self.data[sample_id]["after_weight"].append("")
                    self.data[sample_id]["draw_pressure"].append("")
                    self.data[sample_id]["smell"].append("")
                    self.data[sample_id]["notes"].append("")
                    self.data[sample_id]["tpm"].append(None)

                    # Add resistance and clog for standard tests
                    if self.test_name not in ["User Test Simulation", "User Simulation Test"]:
                        self.data[sample_id]["resistance"].append("")
                        self.data[sample_id]["clog"].append("")

                    # Add chronography for User Test Simulation
                    if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                        self.data[sample_id]["chronography"].append("")

            # Update all tksheets to show the new rows
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                self.update_tksheet(sample_id)

            # Update TPM calculations and stats
            self.update_all_tpm_calculations()

            # Mark as having unsaved changes
            self.mark_unsaved_changes()

            debug_print(f"DEBUG: Successfully added {num_rows_to_add} rows to all samples")

            # Show success message
            self.show_temp_status_message(f"Added {num_rows_to_add} rows to all samples", 3000)
        else:
            debug_print("DEBUG: Add multiple rows cancelled by user")

    def remove_last_row(self):
        """Remove the last row from all samples."""
        if messagebox.askyesno("Confirm Remove", "Remove the last row from all samples?"):
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                for key in self.data[sample_id]:
                    if self.data[sample_id][key]:  # Only remove if list is not empty
                        self.data[sample_id][key].pop()

                if hasattr(self, 'sample_sheets') and i < len(self.sample_sheets):
                    sheet = self.sample_sheets[i]
                    self.update_tksheet(sheet, sample_id)

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



    def get_sample_images_summary(self):
        """Get a summary of all loaded sample images."""
        summary = {}
        total_images = 0

        for sample_id, images in self.sample_images.items():
            summary[sample_id] = len(images)
            total_images += len(images)

        debug_print(f"DEBUG: Sample images summary: {summary} (Total: {total_images})")
        return summary, total_images



    def ensure_initial_tpm_calculation(self):
        """Ensure TPM is calculated and displayed when window opens."""
        debug_print("DEBUG: Ensuring initial TPM calculation and display")

        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"

            # Calculate TPM for this sample
            self.calculate_tpm(sample_id)

        if hasattr(self, 'sample_sheets') and i < len(self.sample_sheets):
            sheet = self.sample_sheets[i]
            self.update_tksheet(sheet, sample_id)

        # Update stats panel to show current TPM data
        self.update_stats_panel()

        debug_print("DEBUG: Initial TPM calculation and display completed")

    def _transfer_sample_notes_to_main_gui(self):
        """Transfer sample notes to main GUI with proper structure for display."""
        try:
            debug_print("DEBUG: Transferring sample notes to main GUI")

            # Ensure we save current notes before transfer
            self._save_current_notes_before_tab_switch()

            # Check if we have sample notes to transfer
            if not hasattr(self, 'data') or not self.data:
                debug_print("DEBUG: No sample data to transfer notes from")
                return

            debug_print(f"DEBUG: Processing notes for {len(self.data)} samples")

            # Collect all sample notes
            sample_notes_data = {}
            notes_found = False

            for sample_id, sample_data in self.data.items():
                sample_notes = sample_data.get("sample_notes", "")
                if sample_notes.strip():  # Only include non-empty notes
                    sample_notes_data[sample_id] = sample_notes
                    notes_found = True
                    debug_print(f"DEBUG: Found notes for {sample_id}: '{sample_notes[:50]}...'")

            if not notes_found:
                debug_print("DEBUG: No sample notes found to transfer")
                return

            # Ensure header_data structure exists and is properly formatted
            if not hasattr(self, 'header_data') or not self.header_data:
                self.header_data = {'samples': []}

            # Ensure header_data has enough sample entries
            while len(self.header_data['samples']) < self.num_samples:
                self.header_data['samples'].append({})

            # Update header_data with current notes
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                if sample_id in sample_notes_data:
                    self.header_data['samples'][i]['sample_notes'] = sample_notes_data[sample_id]
                    debug_print(f"DEBUG: Updated header_data with notes for {sample_id}")

            # Store sample notes in parent for main GUI processing
            if sample_notes_data:
                self.parent.pending_sample_notes = {
                    'test_name': self.test_name,
                    'header_data': self.header_data.copy(),
                    'notes_data': sample_notes_data.copy()
                }
                debug_print(f"DEBUG: Stored sample notes for main GUI - test: {self.test_name}")

                # Store notes metadata for reverse lookup (similar to images)
                if not hasattr(self.parent, 'sample_notes_metadata'):
                    self.parent.sample_notes_metadata = {}
                if self.parent.current_file not in self.parent.sample_notes_metadata:
                    self.parent.sample_notes_metadata[self.parent.current_file] = {}

                self.parent.sample_notes_metadata[self.parent.current_file][self.test_name] = {
                    'header_data': self.header_data.copy(),
                    'notes_data': sample_notes_data.copy(),
                    'test_name': self.test_name
                }
                debug_print(f"DEBUG: Stored sample notes metadata for reverse lookup")

                # If parent has a method to immediately process notes, call it
                if hasattr(self.parent, 'process_pending_sample_notes'):
                    self.parent.process_pending_sample_notes()

        except Exception as e:
            debug_print(f"ERROR: Failed to transfer sample notes to main GUI: {e}")
            import traceback
            traceback.print_exc()

    def _transfer_sample_images_to_main_gui(self):
        """Transfer sample images to main GUI with proper labeling for display."""
        try:
            debug_print("DEBUG: Transferring sample images to main GUI")

            # Check if we have sample images to transfer
            if not hasattr(self, 'sample_images') or not self.sample_images:
                debug_print("DEBUG: No sample images to transfer")
                return

            debug_print(f"DEBUG: Processing {len(self.sample_images)} sample groups for main GUI")

            # Create formatted images for main GUI display
            formatted_images = []

            for sample_id, image_paths in self.sample_images.items():
                try:
                    # Extract sample index from sample_id (e.g., "Sample 1" -> 0)
                    sample_index = int(sample_id.split()[-1]) - 1

                    if sample_index < len(self.header_data['samples']):
                        sample_info = self.header_data['samples'][sample_index]

                        # Create labels for each image with comprehensive information
                        for img_path in image_paths:
                            # Create descriptive label: "Sample 1 - Test Name - Media - Viscosity - Date"
                            label_parts = [
                                sample_info.get('id', sample_id),
                                self.test_name,
                                sample_info.get('media', 'Unknown Media'),
                                f"{sample_info.get('viscosity', 'Unknown')} cP",
                                datetime.datetime.now().strftime("%Y-%m-%d")
                            ]

                            # Filter out empty parts and join
                            formatted_label = " - ".join(filter(lambda x: x and str(x).strip(), label_parts))

                            formatted_images.append({
                                'path': img_path,
                                'label': formatted_label,
                                'sample_id': sample_id,
                                'sample_info': sample_info,
                                'crop_state': getattr(self, 'sample_image_crop_states', {}).get(img_path, False)
                            })

                            debug_print(f"DEBUG: Created formatted image: {formatted_label}")

                except (ValueError, IndexError) as e:
                    debug_print(f"DEBUG: Error processing sample {sample_id}: {e}")
                    continue

            # Store formatted images in parent for main GUI processing
            if formatted_images:
                self.parent.pending_formatted_images = formatted_images
                # Store the target sheet information
                self.parent.pending_images_target_sheet = self.test_name
                debug_print(f"DEBUG: Stored {len(formatted_images)} formatted images for main GUI")
                debug_print(f"DEBUG: Target sheet stored as: {self.test_name}")

                # Store sample-specific image metadata for reverse lookup
                if not hasattr(self.parent, 'sample_image_metadata'):
                    self.parent.sample_image_metadata = {}
                if self.parent.current_file not in self.parent.sample_image_metadata:
                    self.parent.sample_image_metadata[self.parent.current_file] = {}
                if self.test_name not in self.parent.sample_image_metadata[self.parent.current_file]:
                    self.parent.sample_image_metadata[self.parent.current_file][self.test_name] = {}

                # Store the sample-to-image mapping for later retrieval
                self.parent.sample_image_metadata[self.parent.current_file][self.test_name] = {
                    'sample_images': self.sample_images.copy(),
                    'sample_image_crop_states': getattr(self, 'sample_image_crop_states', {}).copy(),
                    'header_data': self.header_data.copy(),
                    'test_name': self.test_name
                }
                debug_print(f"DEBUG: Stored sample metadata for reverse lookup")

                # If parent has a method to immediately process these, call it
                if hasattr(self.parent, 'process_pending_sample_images'):
                    self.parent.process_pending_sample_images()

        except Exception as e:
            debug_print(f"ERROR: Failed to transfer sample images to main GUI: {e}")
            import traceback
            traceback.print_exc()

    def save_sample_images_to_vap3(self):
        """Save sample-specific images to the VAP3 file."""
        try:
            debug_print("DEBUG: Saving sample images to VAP3 file")

            if not self.sample_images:
                debug_print("DEBUG: No sample images to save")
                return

            # We'll modify the VAP file manager to handle this
            # For now, store the sample images in the parent for later saving
            if hasattr(self.parent, 'pending_sample_images'):
                self.parent.pending_sample_images = self.sample_images.copy()
                self.parent.pending_sample_image_crop_states = self.sample_image_crop_states.copy()
                self.parent.pending_sample_header_data = self.header_data.copy()
                debug_print(f"DEBUG: Stored {len(self.sample_images)} sample image groups for later saving")

        except Exception as e:
            debug_print(f"ERROR: Failed to save sample images: {e}")
            import traceback
            traceback.print_exc()

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

            self._save_sample_images()

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
                debug_print("DEBUG: Save and exit requested - calling on_window_close()")
                self.result = "load_file"
                # Call on_window_close() to properly restore main window before destroying
                self.on_window_close()
                return True

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

    def _save_sample_images(self):
        """Save sample images to the appropriate file format."""
        try:
            debug_print("DEBUG: _save_sample_images() starting")

            # Check if we have sample images to save
            if not hasattr(self, 'sample_images') or not self.sample_images:
                debug_print("DEBUG: No sample images to save")
                return

            debug_print(f"DEBUG: Saving sample images for {len(self.sample_images)} samples")

            # For VAP3 files or if the parent supports it, store in memory for later saving
            if self.file_path.endswith('.vap3') or hasattr(self.parent, 'filtered_sheets'):
                debug_print("DEBUG: Storing sample images in parent for VAP3 save")

                # Store sample images in parent for later VAP3 save
                self.parent.pending_sample_images = self.sample_images.copy()
                self.parent.pending_sample_image_crop_states = getattr(self, 'sample_image_crop_states', {}).copy()
                self.parent.pending_sample_header_data = self.header_data.copy()

                debug_print(f"DEBUG: Stored {len(self.sample_images)} sample groups in parent")

                # If this is already a VAP3 file, save it now
                if self.file_path.endswith('.vap3'):
                    self._save_vap3_with_sample_images()

            debug_print("DEBUG: Sample images saved successfully")

        except Exception as e:
            debug_print(f"ERROR: Failed to save sample images: {e}")
            import traceback
            traceback.print_exc()
            # Don't fail the entire save process for image save issues

    def _save_vap3_with_sample_images(self):
        """Save the VAP3 file with sample images included."""
        try:
            debug_print("DEBUG: Saving VAP3 file with sample images")

            # Import the enhanced VAP file manager
            from vap_file_manager import VapFileManager
            vap_manager = VapFileManager()

            # Get current application state from parent
            filtered_sheets = getattr(self.parent, 'filtered_sheets', {})
            sheet_images = getattr(self.parent, 'sheet_images', {})
            plot_options = getattr(self.parent, 'plot_options', [])
            image_crop_states = getattr(self.parent, 'image_crop_states', {})

            # Plot settings
            plot_settings = {}
            if hasattr(self.parent, 'selected_plot_type'):
                plot_settings['selected_plot_type'] = self.parent.selected_plot_type.get()

            # Sample images
            sample_images = getattr(self.parent, 'pending_sample_images', {})
            sample_image_crop_states = getattr(self.parent, 'pending_sample_image_crop_states', {})
            sample_header_data = getattr(self.parent, 'pending_sample_header_data', {})

            # Save to VAP3 file with sample images
            success = vap_manager.save_to_vap3(
                self.file_path,
                filtered_sheets,
                sheet_images,
                plot_options,
                image_crop_states,
                plot_settings,
                sample_images,
                sample_image_crop_states,
                sample_header_data
            )

            if success:
                debug_print("DEBUG: VAP3 file with sample images saved successfully")
            else:
                debug_print("ERROR: Failed to save VAP3 file with sample images")

        except Exception as e:
            debug_print(f"ERROR: Failed to save VAP3 with sample images: {e}")
            import traceback
            traceback.print_exc()
            # Don't fail the entire save for this

    def initialize_data(self):
        """Initialize data structures for all samples."""
        self.data = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"

            # Basic data structure
            self.data[sample_id] = {
                "current_row_index": 0,
                "avg_tpm": 0.0,
                "sample_notes": ""
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


"""
data_collection_ui.py
Developed by Charlie Becquet.
UI modules for data collection window
"""

# pylint: disable=no-member
# This module is part of a multiple inheritance structure where attributes
# are defined in other parent classes (DataCollectionData, DataCollectionHandlers, etc.)

import json
import logging
import os

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tksheet import Sheet

from utils import debug_print, show_success_message

def load_test_sops():
    """
    Load Standard Operating Procedures from JSON file.
    Returns the TEST_SOPS dictionary or default values if file not found.
    """
    debug_print("DEBUG: Loading test SOPs from JSON file")

    # Try multiple possible paths for the JSON file
    possible_paths = [
        "test_sops.json",                           # Same directory as script
        os.path.join(".", "test_sops.json"),       # Current working directory
        os.path.join("resources", "test_sops.json"), # Resources folder
        os.path.join(os.path.dirname(__file__), "test_sops.json"), # Same directory as this script
    ]

    for json_path in possible_paths:
        try:
            if os.path.exists(json_path):
                debug_print(f"DEBUG: Found test_sops.json at: {json_path}")
                with open(json_path, 'r', encoding='utf-8') as f:
                    test_sops = json.load(f)
                debug_print(f"DEBUG: Successfully loaded {len(test_sops)} SOPs from JSON")
                return test_sops
        except (json.JSONDecodeError, IOError) as e:
            debug_print(f"ERROR: Failed to load SOPs from {json_path}: {e}")
            continue

    # Fallback to default SOPs if JSON file not found
    debug_print("WARNING: Could not load test_sops.json, using default fallback SOPs")
    return {
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
        """.strip()
    }

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

class DataCollectionUI:
    """UI creation and widget management"""

    # pylint: disable=too-many-instance-attributes
    # This is a mixin class - attributes are defined in other parent classes

    def __init__(self):
        """Initialize UI-specific attributes."""
        # UI State attributes
        self.sop_visible = False
        self.puff_interval = getattr(self, 'puff_interval',10)
        self.auto_save_interval = getattr(self, 'auto_save_interval', 300000)

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

    def create_menu_bar(self):
        """Create a comprehensive menu bar for the data collection window."""
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Save", accelerator="Ctrl+S",
                            command=lambda: self.save_data(show_confirmation=False))
        file_menu.add_command(label="Save and Exit",
                            command=lambda: self.save_data(exit_after=True))
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
        edit_menu.add_command(label="Add Rows", command=self.add_multiple_rows)
        edit_menu.add_command(label="Remove Last Row", command=self.remove_last_row)
        edit_menu.add_separator()
        edit_menu.add_command(label="Recalculate TPM", command=self.recalculate_all_tpm)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        # Navigate menu
        navigate_menu = tk.Menu(menubar, tearoff=0)
        navigate_menu.add_command(label="Previous Sample", accelerator="Alt+Left",
                                command=self.go_to_previous_sample)
        navigate_menu.add_command(label="Next Sample", accelerator="Alt+Right",
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
        content_frame.add(data_frame, weight=2)  # 50% width

        # Create stats panel (right side)
        self.stats_frame = ttk.Frame(content_frame)
        content_frame.add(self.stats_frame, weight=2)  # 50% width

        # Create notebook with sample tabs
        self.notebook = ttk.Notebook(data_frame)
        self.notebook.pack(fill="both", expand=True, pady=(0, 10))  # Add bottom padding

        # Create sample tabs
        self.create_sample_tabs()

        # Create stats panel
        self.create_tpm_stats_panel()

        # Bind notebook tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

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
            width=6,
            command=self.toggle_sop_visibility
        )
        self.toggle_sop_button.pack(side="left", padx=(0, 5))

        # Add title label after the button
        title_label = ttk.Label(header_frame, text="Standard Operating Procedure (SOP)",
                               font=("Arial", 10, "bold"), style='TLabel')
        title_label.pack(side="left")

        # NEW: Load Images for Sample button
        load_images_button = ttk.Button(
            header_frame,
            text="Load Images for Sample",
            command=self.load_images_for_sample
        )
        load_images_button.pack(side="right", padx=(5, 0))

        test_sops = load_test_sops()

        # Get the SOP text for the current test
        sop_text = test_sops.get(self.test_name, test_sops.get("default", "No SOP available"))

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

        debug_print("DEBUG: SOP section created successfully with Load Images button")

    def create_sample_tabs(self):
        """Create tabs for all samples in the notebook."""
        # This method manages the overall process
        self.sample_frames = []
        self.sample_sheets = []

        # Calls create_sample_tab() for EACH sample
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            sample_frame = ttk.Frame(self.notebook, padding=10, style='TFrame')
            sample_name = self.header_data['samples'][i].get('id', f"Sample {i+1}")

            self.notebook.add(sample_frame, text=f"Sample {i+1} - {sample_name}")
            self.sample_frames.append(sample_frame)

            # Create individual tab content
            sheet = self.create_sample_tab(sample_frame, sample_id, i)
            self.sample_sheets.append(sheet)

        # REMOVED: Duplicate binding that was conflicting
        # self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        debug_print("DEBUG: Sample tabs created without duplicate binding")

    def create_sample_tab(self, parent_frame, sample_id, sample_index):
        """Create a tab for a single sample with tksheet instead of treeview."""
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
            headers = ["Chronography", "Puffs", "Before Weight", "After Weight", "Draw Pressure", "Failure", "Notes", "TPM"]
            self.tree_columns = len(headers)

            # Ensure chronography data exists
            if "chronography" not in self.data[sample_id]:
                debug_print(f"DEBUG: Creating chronography data structure for {sample_id}")
                self.data[sample_id]["chronography"] = []
                existing_length = len(self.data[sample_id]["puffs"])
                for j in range(existing_length):
                    self.data[sample_id]["chronography"].append("")

            debug_print(f"DEBUG: User Test Simulation columns: {headers}")
        else:
            headers = ["Puffs", "Before Weight", "After Weight", "Draw Pressure", "Resistance", "Smell", "Clog", "Notes", "TPM"]
            self.tree_columns = len(headers)
            debug_print(f"DEBUG: Standard test columns: {headers}")

        # Create the tksheet widget
        sheet = Sheet(
            main_container,
            headers=headers,
            height=400,
            width=900,
            show_table=True,
            show_top_left=True,
            show_row_index=True,
            show_header=True,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            empty_horizontal=0,
            empty_vertical=50
        )

        # Configure sheet appearance
        sheet.set_options(
            table_bg="#FFFFFF",
            table_fg="#000000",
            header_bg="#D0D0D0",
            header_fg="#000000",
            index_bg="#F0F0F0",
            index_fg="#000000",
            selected_cells_bg="#CCE5FF",
            selected_cells_fg="#000000",
            table_selected_cells_border_fg="#4472C4",
            table_selected_cells_bg="#CCE5FF",
            table_selected_rows_bg="#E8F4FD",
            table_selected_rows_fg="#000000"
        )

        sheet.enable_bindings([
            "single_select",
            "drag_select",
            "column_select",
            "row_select",
            "arrowkeys",
            "tab",
            "delete",
            "backspace",
            "f2",
            "edit_cell",
            "begin_edit_cell",
            "ctrl_c",
            "ctrl_v",
            "ctrl_x",
            "column_width_resize",
            "row_height_resize",
            "right_click_popup_menu"
        ])

        # Grid the sheet to fill the container
        sheet.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        # Make TPM column read-only and pin it to the right
        tpm_col_idx = len(headers) - 1
        sheet.readonly_columns([tpm_col_idx])

        # CENTER ALIGN all columns
        for i in range(len(headers)):
            sheet.align_columns([i], align="center")

        # Set up enhanced event bindings for Excel-like behavior
        self.setup_enhanced_sheet_bindings(sheet, sample_id, sample_index)

        # Populate sheet with initial data
        self.populate_tksheet_data(sheet, sample_id)

        self.old_sheet_data = sheet.get_sheet_data()

        debug_print(f"DEBUG: Sample tab created for {sample_id} with {len(headers)} columns using tksheet")
        return sheet

    def create_tpm_stats_panel(self):
        """Create the enhanced TPM statistics panel with responsive resizing."""
        debug_print("DEBUG: Creating enhanced TPM statistics panel with resize support")

        # Clear existing widgets
        for widget in self.stats_frame.winfo_children():
            widget.destroy()

        # Create the main stats frame with proper expansion
        main_stats_frame = ttk.LabelFrame(self.stats_frame, text="TPM Statistics", padding=10)
        main_stats_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Store reference for resize handling
        self.plot_container = main_stats_frame

        # Sample selection for stats display (fixed height section)
        control_frame = ttk.Frame(main_stats_frame)
        control_frame.pack(side='top', fill='x', pady=(0, 5))

        # Current sample label
        self.current_sample_label = ttk.Label(control_frame, text="Sample 1: Unknown Sample",
                                             font=('Arial', 12, 'bold'))
        self.current_sample_label.pack(anchor='w')

        # Stats text frame (fixed height section) - now with grid layout for 3 columns
        text_frame = ttk.Frame(main_stats_frame)
        text_frame.pack(side='top', fill='x', pady=(5, 10))
        self.text_frame_ref = text_frame  # Store the reference properly

        # Configure grid weights for equal column sizing: 33% / 33% / 33%
        text_frame.columnconfigure(0, weight=1, uniform="equal_columns")
        text_frame.columnconfigure(1, weight=1, uniform="equal_columns")
        text_frame.columnconfigure(2, weight=1, uniform="equal_columns")

        # Left column - TPM statistics - CHANGED: Use Labels instead of Text widget
        left_column = ttk.LabelFrame(text_frame, text="TPM Statistics", padding=5)
        left_column.grid(row=0, column=0, sticky='nsew', padx=(0, 2))
        left_column.grid_columnconfigure(0, weight=1)

        # Create label for stats display - FIXED: Use Label instead of Text
        self.sample_stats_label = ttk.Label(left_column, text="No TPM data available",
                                           font=('Arial', 10), background='#f8f8f8',
                                           relief='flat', borderwidth=1, anchor='nw', justify='left')
        self.sample_stats_label.pack(fill='both', expand=True, padx=2, pady=2)
        debug_print("DEBUG: Created TPM stats label widget")

        # Middle column - Sample information - CHANGED: Use Labels instead of Text widget
        middle_column = ttk.LabelFrame(text_frame, text="Sample Information", padding=5)
        middle_column.grid(row=0, column=1, sticky='nsew', padx=2)
        middle_column.grid_columnconfigure(0, weight=1)

        # Create label for sample info - FIXED: Use Label instead of Text
        self.sample_info_label = ttk.Label(middle_column, text="No sample data available",
                                          font=('Arial', 10), background='#f8f8f8',
                                          relief='flat', borderwidth=1, anchor='nw', justify='left')
        self.sample_info_label.pack(fill='both', expand=True, padx=2, pady=2)
        debug_print("DEBUG: Created sample info label widget")

        # Right column - Sample test notes - KEPT: Text widget for editing
        right_column = ttk.LabelFrame(text_frame, text="Sample Test Notes", padding=5)
        right_column.grid(row=0, column=2, sticky='nsew', padx=(2, 0))
        right_column.grid_columnconfigure(0, weight=1)

        # Create text widget for notes with scrollbar in right column
        notes_container = ttk.Frame(right_column)
        notes_container.pack(fill='both', expand=True)

        self.sample_notes_text = tk.Text(notes_container, height=6, wrap='word', font=('Arial', 10))
        notes_scrollbar = ttk.Scrollbar(notes_container, orient='vertical', command=self.sample_notes_text.yview)
        self.sample_notes_text.configure(yscrollcommand=notes_scrollbar.set)

        self.sample_notes_text.pack(side='left', fill='both', expand=True)
        notes_scrollbar.pack(side='right', fill='y')

        # Bind notes text changes to save function
        self.sample_notes_text.bind('<KeyRelease>', self.on_notes_changed)
        self.sample_notes_text.bind('<FocusOut>', self.on_notes_changed)
        debug_print("DEBUG: Bound notes change events")

        # Plot container frame (use original expandable approach)

        # Plot container frame (use original expandable approach)
        plot_container = ttk.Frame(main_stats_frame)
        plot_container.pack(side='top', fill='both', expand=True)

        # Store reference to the container for proper sizing
        self.stats_frame_container = plot_container

        # Create the matplotlib plot using your original method
        self.setup_stats_plot_canvas(plot_container)

        # Initialize TPM labels dictionary if not exists
        if not hasattr(self, 'tpm_labels'):
            self.tpm_labels = {}

        # Update the statistics for the current sample (this will handle the plot update)
        self.update_stats_panel()

        debug_print("DEBUG: Enhanced TPM statistics panel created with 3-column layout and original plot sizing")

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

    def recreate_sample_tabs(self):
        """Recreate all sample tabs when sample count changes."""
        debug_print("DEBUG: Recreating sample tabs")

        # Clear existing tabs
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        # Just call the existing method
        self.create_sample_tabs()

        # Recreate the TPM stats panel
        self.create_tpm_stats_panel()

        debug_print(f"DEBUG: Recreated {self.num_samples} sample tabs")

    def initialize_plot_size(self):
        """Initialize plot size after window is fully rendered."""
        debug_print("DEBUG: Initializing plot size after window render")

        # Give window time to fully render
        self.window.update_idletasks()

        # Update plot size if available
        if hasattr(self, 'update_plot_size_for_resize'):
            self.window.after(500, self.update_plot_size_for_resize)

        debug_print("DEBUG: Initial plot size set")

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

    def setup_stats_plot_canvas(self, parent):
        """Create the matplotlib canvas for the TPM plot with dynamic responsive sizing."""
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

        # Calculate dynamic plot size based on available space
        dynamic_width, dynamic_height = self.calculate_dynamic_plot_size(parent)

        # Create figure with calculated responsive sizing
        self.stats_fig, self.stats_ax = plt.subplots(figsize=(dynamic_width, dynamic_height))
        self.stats_fig.patch.set_facecolor('white')
        self.stats_fig.subplots_adjust(left=0.15, right=0.95, top=0.90, bottom=0.15)

        # Create canvas with proper expansion configuration
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill='both', expand=True)

        # Store reference to canvas_frame for resize handling
        self.canvas_frame = canvas_frame

        # Create canvas
        self.stats_canvas = FigureCanvasTkAgg(self.stats_fig, canvas_frame)
        canvas_widget = self.stats_canvas.get_tk_widget()
        canvas_widget.pack(fill='both', expand=True)

        debug_print("DEBUG: Stats plot canvas created with responsive sizing")

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

    def update_header_display(self):
        """Update the header display and sample information across all tabs."""
        debug_print("DEBUG: Updating header display after header edit")

        try:
            # Update the sample information display
            self.update_all_sample_info()

            # If the number of samples changed, recreate the tabs
            current_sample_count = len(self.sample_tabs)
            new_sample_count = self.header_data.get("num_samples", 1)

            if current_sample_count != new_sample_count:
                debug_print(f"DEBUG: Sample count changed from {current_sample_count} to {new_sample_count}, recreating tabs")
                self.recreate_sample_tabs()
            else:
                # Just update the existing tab labels and sample info
                for i, sample_data in enumerate(self.header_data.get("samples", [])):
                    if i < len(self.sample_tabs):
                        sample_id = sample_data.get("id", f"Sample {i+1}") if isinstance(sample_data, dict) else str(sample_data)
                        self.notebook.tab(i, text=f"Sample {i+1} - {sample_id}")

            # Update the window title
            self.update_window_title()

            # Refresh the current tab display
            current_tab = self.notebook.select()
            if current_tab:
                tab_index = self.notebook.index(current_tab)
                self.on_tab_change(tab_index)

            debug_print("DEBUG: Header display update completed")

        except Exception as e:
            debug_print(f"ERROR: Failed to update header display: {e}")
            import traceback
            traceback.print_exc()

    def end_editing(self):
        """End any active editing in tksheets."""
        # For tksheet, we don't need to explicitly end editing like with treeview
        # This method is kept for compatibility with the save process
        debug_print("DEBUG: Ending any active editing (tksheet)")

    def update_tpm_in_sheet(self, sheet, sample_id):
        """Update the TPM column in the sheet with calculated values"""
        # Get the number of headers to find TPM column index
        debug_print(f"DEBUG: attempting to calculate sheet header length")
        if hasattr(sheet, 'columns'):
            tpm_col_idx = len(sheet.columns) - 1  # TPM is always the last column
            debug_print(f"DEBUG: calculated sheet header length")
        else:
            # Fallback if headers attribute not available
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                tpm_col_idx = 7  # TPM column index for user simulation
            else:
                tpm_col_idx = 8  # TPM column index for standard tests

        for row_idx in range(len(self.data[sample_id]["tpm"])):
            tpm_value = self.data[sample_id]["tpm"][row_idx]
            if tpm_value is not None:
                sheet.set_cell_data(row_idx, tpm_col_idx, f"{tpm_value:.6f}")
            else:
                sheet.set_cell_data(row_idx, tpm_col_idx, "")

    def update_tksheet(self, sheet, sample_id):
        """Update the tksheet with current data (replaces update_treeview)"""
        debug_print(f"DEBUG: Updating tksheet for {sample_id}")

        try:
            # Populate the sheet with current data
            self.populate_tksheet_data(sheet, sample_id)

            # Highlight TPM cells that have calculated values
            #self.highlight_tpm_cells(sheet, sample_id)

            debug_print(f"DEBUG: tksheet updated for {sample_id}")
        except Exception as e:
            debug_print(f"DEBUG: Error updating tksheet: {e}")

    def populate_tksheet_data(self, sheet, sample_id):
        """Populate tksheet with data from self.data"""
        debug_print(f"DEBUG: Populating tksheet for {sample_id}")

        data_rows = []
        data_length = len(self.data[sample_id]["puffs"])

        for i in range(data_length):
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                row = [
                    self.data[sample_id]["chronography"][i] if i < len(self.data[sample_id]["chronography"]) else "",
                    self.data[sample_id]["puffs"][i] if i < len(self.data[sample_id]["puffs"]) else "",
                    self.data[sample_id]["before_weight"][i] if i < len(self.data[sample_id]["before_weight"]) else "",
                    self.data[sample_id]["after_weight"][i] if i < len(self.data[sample_id]["after_weight"]) else "",
                    self.data[sample_id]["draw_pressure"][i] if i < len(self.data[sample_id]["draw_pressure"]) else "",
                    self.data[sample_id]["smell"][i] if i < len(self.data[sample_id]["smell"]) else "",  # Failure column
                    self.data[sample_id]["notes"][i] if i < len(self.data[sample_id]["notes"]) else "",
                    f"{self.data[sample_id]['tpm'][i]:.6f}" if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None else ""
                ]
            else:
                row = [
                    self.data[sample_id]["puffs"][i] if i < len(self.data[sample_id]["puffs"]) else "",
                    self.data[sample_id]["before_weight"][i] if i < len(self.data[sample_id]["before_weight"]) else "",
                    self.data[sample_id]["after_weight"][i] if i < len(self.data[sample_id]["after_weight"]) else "",
                    self.data[sample_id]["draw_pressure"][i] if i < len(self.data[sample_id]["draw_pressure"]) else "",
                    self.data[sample_id]["resistance"][i] if i < len(self.data[sample_id]["resistance"]) else "",
                    self.data[sample_id]["smell"][i] if i < len(self.data[sample_id]["smell"]) else "",
                    self.data[sample_id]["clog"][i] if i < len(self.data[sample_id]["clog"]) else "",
                    self.data[sample_id]["notes"][i] if i < len(self.data[sample_id]["notes"]) else "",
                    f"{self.data[sample_id]['tpm'][i]:.6f}" if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None else ""
                ]
            data_rows.append(row)

        # Add empty rows to reach the standard 50 rows
        while len(data_rows) < 50:
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                empty_row = ["", "", "", "", "", "", "", ""]
            else:
                empty_row = ["", "", "", "", "", "", "", "", ""]
            data_rows.append(empty_row)

        sheet.set_sheet_data(data_rows)
        sheet.set_all_cell_sizes_to_text(width=60)
        debug_print(f"DEBUG: Populated tksheet with {len(data_rows)} rows")

    def load_images_for_sample(self):
        """Load images for the currently selected sample."""
        try:
            # Get current sample
            current_tab_index = self.notebook.index(self.notebook.select())
            sample_id = f"Sample {current_tab_index + 1}"

            debug_print(f"DEBUG: Loading images for {sample_id}")

            # Create a temporary frame for the image loader
            temp_window = tk.Toplevel(self.window)
            temp_window.title(f"Load Images for {sample_id}")
            temp_window.geometry("600x400")
            temp_window.transient(self.window)
            temp_window.grab_set()

            # Create frame for image loader
            image_frame = ttk.Frame(temp_window)
            image_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Create ImageLoader instance for this sample
            from image_loader import ImageLoader

            def on_images_selected(image_paths):
                """Callback when images are selected for this sample."""
                debug_print(f"DEBUG: Images selected for {sample_id}: {image_paths}")

                # Store images for this sample
                if sample_id not in self.sample_images:
                    self.sample_images[sample_id] = []

                # Add new images (avoid duplicates)
                for img_path in image_paths:
                    if img_path not in self.sample_images[sample_id]:
                        self.sample_images[sample_id].append(img_path)
                        # Store crop state (default from current settings)
                        self.sample_image_crop_states[img_path] = False  # Default crop state

                debug_print(f"DEBUG: Total images for {sample_id}: {len(self.sample_images[sample_id])}")

                # Update the status label
                status_label.config(text=f"Images loaded for {sample_id}: {len(self.sample_images[sample_id])} files")

            # Create ImageLoader
            image_loader = ImageLoader(
                image_frame,
                is_plotting_sheet=False,  # Not a plotting sheet context
                on_images_selected=on_images_selected,
                main_gui=None  # No main GUI connection yet
            )

            # Load existing images for this sample if any
            if sample_id in self.sample_images:
                image_loader.load_images_from_list(self.sample_images[sample_id])
                # Restore crop states
                for img_path in self.sample_images[sample_id]:
                    if img_path in self.sample_image_crop_states:
                        image_loader.image_crop_states[img_path] = self.sample_image_crop_states[img_path]

            # Add status label
            status_label = ttk.Label(temp_window, text=f"Current images for {sample_id}: {len(self.sample_images.get(sample_id, []))}")
            status_label.pack(pady=5)

            # Add control buttons
            button_frame = ttk.Frame(temp_window)
            button_frame.pack(pady=10)

            ttk.Button(button_frame, text="Add More Images",
                      command=image_loader.add_images).pack(side="left", padx=5)

            ttk.Button(button_frame, text="Close",
                      command=temp_window.destroy).pack(side="left", padx=5)

            # Center the window
            temp_window.transient(self.window)
            temp_window.grab_set()

            # Position relative to parent
            x = self.window.winfo_x() + 50
            y = self.window.winfo_y() + 50
            temp_window.geometry(f"600x400+{x}+{y}")

        except Exception as e:
            debug_print(f"ERROR: Failed to load images for sample: {e}")
            import traceback
            traceback.print_exc()

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

    def exit_without_saving(self):
        """Exit without saving (with confirmation)."""
        if self.has_unsaved_changes:
            if not messagebox.askyesno("Confirm Exit",
                                     "You have unsaved changes. Are you sure you want to exit without saving?"):
                return

        debug_print("DEBUG: Exiting without saving")
        self.result = "cancel"
        self.window.destroy()

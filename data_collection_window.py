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
import openpyxl
from openpyxl.styles import PatternFill
from utils import FONT
import threading
import subprocess

class DataCollectionWindow:
    def __init__(self, parent, file_path, test_name, header_data):
        """
        Initialize the data collection window.
        
        Args:
            parent (tk.Tk): The parent window.
            file_path (str): Path to the Excel file.
            test_name (str): Name of the test being conducted.
            header_data (dict): Dictionary containing header data.
        """
        self.parent = parent
        self.root = parent.root
        self.file_path = file_path
        self.test_name = test_name
        self.header_data = header_data
        self.num_samples = header_data["num_samples"]
        self.result = None
        
        # Auto-save settings
        self.auto_save_interval = 5 * 60 * 1000  # 5 minutes in milliseconds
        self.auto_save_timer = None
        self.has_unsaved_changes = False
        self.last_save_time = None
        
        # Create the window
        self.window = tk.Toplevel(self.root)
        self.window.title(f"Data Collection - {test_name}")
        self.window.geometry("1100x600")  # Wider window to accommodate TPM panel
        self.window.minsize(1000, 500)
        
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

        # Click tracking variables - IMPROVED APPROACH
        self.last_click_time = 0
        self.last_click_xy = (0,0)
        self.double_click_threshold = 500  # milliseconds
        self.cell_border_frame = None # Initialize cell border frame
        self.single_click_after_id = None
        
        # Data storage
        self.data = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            self.data[sample_id] = {
                "puffs": [],           # Will be populated dynamically
                "before_weight": [],   # Will be populated dynamically
                "after_weight": [],    # Will be populated dynamically
                "draw_pressure": [],   # Will be populated dynamically
                "smell": [],           # Will be populated dynamically
                "notes": [],           # Will be populated dynamically
                "tpm": [],             # Will store calculated TPM values
                "current_row_index": 0, # Track the current editable row
                "avg_tpm": 0.0         # Track average TPM
            }
            
            # Pre-initialize 50 rows
            for j in range(50):
                puff = (j + 1) * self.puff_interval
                self.data[sample_id]["puffs"].append(puff)
                self.data[sample_id]["before_weight"].append("")
                self.data[sample_id]["after_weight"].append("")
                self.data[sample_id]["draw_pressure"].append("")
                self.data[sample_id]["smell"].append("")
                self.data[sample_id]["notes"].append("")
                self.data[sample_id]["tpm"].append(None)

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
        
        print(f"DEBUG: DataCollectionWindow initialized for {test_name} with {self.num_samples} samples")
    
    def create_menu_bar(self):
        """Create a comprehensive menu bar for the data collection window."""
        print("DEBUG: Creating menu bar for DataCollectionWindow")
        
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Save", accelerator="Ctrl+S", command=self.quick_save)
        file_menu.add_command(label="Save and Continue", command=self.save_and_continue)
        file_menu.add_command(label="Save and Exit", command=self.save_and_exit)
        file_menu.add_separator()
        file_menu.add_command(label="Open Raw Excel File", command=self.open_raw_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Export CSV", command=self.export_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit without Saving", command=self.exit_without_saving)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Edit menu  
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Edit Headers", command = self.edit_headers)
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
        navigate_menu.add_command(label="Previous Sample", accelerator="Ctrl+Left", command=self.go_to_previous_sample)
        navigate_menu.add_command(label="Next Sample", accelerator="Ctrl+Right", command=self.go_to_next_sample)
        navigate_menu.add_separator()
        navigate_menu.add_command(label="Go to Sample...", command=self.go_to_sample_dialog)
        menubar.add_cascade(label="Navigate", menu=navigate_menu)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Change Puff Interval", command=self.change_puff_interval_dialog)
        tools_menu.add_command(label="Auto-Save Settings", command=self.auto_save_settings_dialog)
        tools_menu.add_separator()
        tools_menu.add_command(label="Switch Test", command=self.switch_test_dialog)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_keyboard_shortcuts)
        help_menu.add_command(label="About Data Collection", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        print("DEBUG: Menu bar created successfully")
    
    def setup_styles(self):
        """Set up styles for ttk widgets to ensure no blue backgrounds."""
        # Configure ttk styles to use system defaults
        self.style.configure('TFrame', background='SystemButtonFace')
        self.style.configure('TLabel', background='SystemButtonFace')
        self.style.configure('TLabelframe', background='SystemButtonFace')
        self.style.configure('TLabelframe.Label', background='SystemButtonFace')
        self.style.configure('TNotebook', background='SystemButtonFace')
        self.style.configure('TNotebook.Tab', background='SystemButtonFace')

        # Create special styles for headers
        self.style.configure('Header.TLabel', 
                             font=("Arial", 14, "bold"), 
                             background='SystemButtonFace')

        self.style.configure('SubHeader.TLabel', 
                             font=("Arial", 12), 
                             background='SystemButtonFace')

        # Style for stats panel
        self.style.configure('Stats.TLabel', 
                             font=("Arial", 14, "bold"), 
                             background='SystemButtonFace')

        # Style for sample info
        self.style.configure('SampleInfo.TLabel',
                             font=("Arial", 11),
                             background='SystemButtonFace')
                     
        # Style for Treeview with enhanced visual separation
        self.style.configure('Treeview', 
                            background='white',
                            fieldbackground='white',
                            borderwidth=1)

        # Add gridlines to the Treeview cells
        self.style.configure('Treeview', rowheight=30)  # Increased for better visibility
    
        # Configure selection colors
        self.style.map('Treeview', 
                      background=[('selected', '#3874CC')],
                      foreground=[('selected', 'white')])

        # Style headers with a distinct look
        self.style.configure('Treeview.Heading',
                            background='#D0D0D0',
                            foreground='black',
                            relief='raised',
                            borderwidth=2,
                            font=('Arial', 10, 'bold'))
    
        # Add hover effect for headers
        self.style.map('Treeview.Heading',
                      background=[('active', '#C0C0C0')])

        print("DEBUG: Styles configured with enhanced visual separation")
    
    def center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """Create the data collection UI."""
        # Set window background to system default
        self.window.configure(background='SystemButtonFace')
        
        # Create a horizontal split layout
        main_frame = ttk.Frame(self.window, padding=10, style='TFrame')
        main_frame.pack(fill="both", expand=True)
        
        # Header with test information and save status
        header_frame = ttk.Frame(main_frame, style='TFrame')
        header_frame.pack(fill="x", pady=(0, 10))
        
        # Left side of header
        header_left = ttk.Frame(header_frame, style='TFrame')
        header_left.pack(side="left", fill="x", expand=True)
        
        # Use styled labels
        ttk.Label(header_left, text=f"Test: {self.test_name}", style='Header.TLabel').pack(side="left")
        ttk.Label(header_left, text=f"Tester: {self.header_data['common']['tester']}", style='SubHeader.TLabel').pack(side="left", padx=(20, 0))
        
        # Right side of header - save status
        header_right = ttk.Frame(header_frame, style='TFrame')
        header_right.pack(side="right")
        
        self.save_status_label = ttk.Label(header_right, text="●", style='SubHeader.TLabel', foreground="red")
        self.save_status_label.pack(side="right")
        
        self.save_status_text = ttk.Label(header_right, text="Unsaved changes", style='SubHeader.TLabel')
        self.save_status_text.pack(side="right", padx=(0, 5))
        
        # Create a horizontal paned window to split the main area
        paned_window = ttk.PanedWindow(main_frame, orient="horizontal")
        paned_window.pack(fill="both", expand=True)
        
        # Left side - Data entry area
        data_frame = ttk.Frame(paned_window, style='TFrame')
        paned_window.add(data_frame, weight=3)  # 75% of the width
        
        # Right side - TPM stats panel
        self.stats_frame = ttk.Frame(paned_window, style='TFrame')
        paned_window.add(self.stats_frame, weight=1)  # 25% of the width
        
        # Setup data entry area with notebook
        self.notebook = ttk.Notebook(data_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Create a tab for each sample
        self.sample_frames = []
        self.sample_trees = []
        
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            sample_frame = ttk.Frame(self.notebook, padding=10, style='TFrame')
            self.notebook.add(sample_frame, text=f"Sample {i+1} - {self.header_data['samples'][i]['id']}")
            self.sample_frames.append(sample_frame)
            
            # Create the sample tab content
            tree = self.create_sample_tab(sample_frame, sample_id, i)
            self.sample_trees.append(tree)
        
        # Create the TPM stats panel
        self.create_tpm_stats_panel()
        
        # Control buttons at the bottom
        button_frame = ttk.Frame(main_frame, style='TFrame')
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Left side controls
        left_controls = ttk.Frame(button_frame, style='TFrame')
        left_controls.pack(side="left", fill="x")
        
        ttk.Label(left_controls, text="Puff Interval:", style='TLabel').pack(side="left")
        self.puff_interval_var = tk.IntVar(value=self.puff_interval)
        puff_spinbox = ttk.Spinbox(
            left_controls, 
            from_=1, 
            to=100, 
            textvariable=self.puff_interval_var, 
            width=5,
            command=self.update_puff_interval
        )
        puff_spinbox.pack(side="left", padx=5)
        
        # Sample navigation
        nav_frame = ttk.Frame(button_frame, style='TFrame')
        nav_frame.pack(side="left", padx=20)
        
        ttk.Button(nav_frame, text="← Prev Sample", command=self.go_to_previous_sample).pack(side="left")
        ttk.Button(nav_frame, text="Next Sample →", command=self.go_to_next_sample).pack(side="left", padx=5)
        
        # Right side controls
        ttk.Button(button_frame, text="Quick Save", command=self.quick_save).pack(side="right", padx=(10, 0))
        ttk.Button(button_frame, text="Save & Exit", command=self.save_and_exit).pack(side="right")
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side="right", padx=(0, 10))
        
        # Bind notebook tab change to update stats panel
        self.notebook.bind("<<NotebookTabChanged>>", self.update_stats_panel)
        
        print("DEBUG: Main widgets created successfully")
    
    # Auto-save functionality
    def start_auto_save_timer(self):
        """Start the auto-save timer."""
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)
        
        self.auto_save_timer = self.window.after(self.auto_save_interval, self.auto_save)
        print(f"DEBUG: Auto-save timer started - will save in {self.auto_save_interval/1000/60:.1f} minutes")
    
    def auto_save(self):
        """Perform automatic save without user confirmation."""
        if self.has_unsaved_changes:
            print("DEBUG: Performing auto-save...")
            try:
                self.save_data_internal(show_confirmation=False, auto_save=True)
                self.update_save_status(False)  # Mark as saved
                print("DEBUG: Auto-save completed successfully")
            except Exception as e:
                print(f"DEBUG: Auto-save failed: {e}")
                # Continue even if auto-save fails
        
        # Restart the timer
        self.start_auto_save_timer()
    
    def mark_unsaved_changes(self):
        """Mark that there are unsaved changes."""
        if not self.has_unsaved_changes:
            self.has_unsaved_changes = True
            self.update_save_status(True)
            print("DEBUG: Marked as having unsaved changes")
    
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
    
    # Menu action methods
    def quick_save(self):
        """Quick save without closing the window."""
        try:
            self.save_data_internal(show_confirmation=False)
            self.update_save_status(False)
            messagebox.showinfo("Save Complete", "Data saved and main view updated.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save data: {e}")
    
    def save_and_continue(self):
        """Save data and continue working."""
        print("DEBUG: Save and continue initiated")
        try:
            self.save_data_internal(show_confirmation=True)
            self.update_save_status(False)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save data: {e}")
    
    def save_and_exit(self):
        """Save data and exit the window."""
        print("DEBUG: Save and exit initiated")
        try:
            self.save_data_internal(show_confirmation=False)
            self.result = "load_file"
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save data: {e}")
    
    def exit_without_saving(self):
        """Exit without saving (with confirmation)."""
        if self.has_unsaved_changes:
            if not messagebox.askyesno("Confirm Exit", 
                                     "You have unsaved changes. Are you sure you want to exit without saving?"):
                return
        
        print("DEBUG: Exiting without saving")
        self.result = "cancel"
        self.window.destroy()
    
    def open_raw_excel(self):
        """Open the raw Excel file for direct editing."""
        print(f"DEBUG: Opening raw Excel file: {self.file_path}")
        try:
            if os.path.exists(self.file_path):
                os.startfile(self.file_path)
            else:
                messagebox.showerror("File Not Found", f"Could not find file: {self.file_path}")
        except Exception as e:
            messagebox.showerror("Open Error", f"Failed to open Excel file: {e}")
    
    def export_csv(self):
        """Export current data to CSV files."""
        print("DEBUG: Exporting data to CSV")
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
                
            messagebox.showinfo("Export Complete", f"Exported {self.num_samples} CSV files to {directory}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV files: {e}")
    
    def edit_headers(self):
        """Open header editing dialog from within data collection."""
        print("DEBUG: Edit Headers requested from data collection window")
        
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
            print("DEBUG: Header data updated successfully")
            # Update our internal header data
            old_header_data = self.header_data.copy()
            self.header_data = new_header_data
            
            # Check if number of samples changed
            old_num_samples = old_header_data.get("num_samples", 0)
            new_num_samples = new_header_data.get("num_samples", 0)
            
            if old_num_samples != new_num_samples:
                print(f"DEBUG: Number of samples changed from {old_num_samples} to {new_num_samples}")
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
            print("DEBUG: Header editing was cancelled")

    def handle_sample_count_change(self, old_count, new_count):
        """Handle changes in the number of samples."""
        print(f"DEBUG: Handling sample count change: {old_count} -> {new_count}")
        
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
            
            print(f"DEBUG: Added {new_count - old_count} new samples")
            
        elif new_count < old_count:
            # Remove excess samples
            for i in range(new_count, old_count):
                sample_id = f"Sample {i+1}"
                if sample_id in self.data:
                    del self.data[sample_id]
            
            print(f"DEBUG: Removed {old_count - new_count} samples")
        
        # Update the number of samples
        self.num_samples = new_count
        
        # Recreate the UI to reflect the new sample count
        self.recreate_sample_tabs()

    def recreate_sample_tabs(self):
        """Recreate all sample tabs when sample count changes."""
        print("DEBUG: Recreating sample tabs")
        
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
        
        print(f"DEBUG: Recreated {self.num_samples} sample tabs")

    def update_header_display(self):
        """Update the header information displayed in the UI."""
        print("DEBUG: Updating header display")
        
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
            
            print("DEBUG: Header display updated")
            
        except Exception as e:
            print(f"DEBUG: Error updating header display: {e}")

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
            print(f"DEBUG: Error updating label: {e}")

    def apply_header_changes_to_file(self):
        """Apply header changes to the Excel file."""
        print("DEBUG: Applying header changes to Excel file")
        try:
            import openpyxl
            wb = openpyxl.load_workbook(self.file_path)
            
            if self.test_name not in wb.sheetnames:
                print(f"DEBUG: Sheet {self.test_name} not found")
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
            print("DEBUG: Header changes applied to Excel file")
            
        except Exception as e:
            print(f"DEBUG: Error applying header changes to file: {e}")

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
        
        if messagebox.askyesno("Confirm Clear", 
                             f"Are you sure you want to clear all data for {sample_id}?"):
            # Clear all data except puffs
            for key in ["before_weight", "after_weight", "draw_pressure", "smell", "notes", "tpm"]:
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
            
            print(f"DEBUG: Cleared data for {sample_id}")
    
    def clear_all_data(self):
        """Clear all data for all samples."""
        if messagebox.askyesno("Confirm Clear All", 
                             "Are you sure you want to clear ALL data for ALL samples?"):
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                for key in ["before_weight", "after_weight", "draw_pressure", "smell", "notes", "tpm"]:
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
            print("DEBUG: Cleared all data for all samples")
    
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
            
            # Update treeview
            tree = self.sample_trees[i]
            self.update_treeview(tree, sample_id)
        
        self.mark_unsaved_changes()
        print(f"DEBUG: Added new row with puff count {new_puff}")
    
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
            print("DEBUG: Removed last row from all samples")
    
    def recalculate_all_tpm(self):
        """Recalculate TPM for all samples."""
        print("DEBUG: Recalculating TPM for all samples")
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"
            self.calculate_tpm(sample_id)
        
        self.update_stats_panel()
        self.mark_unsaved_changes()
        messagebox.showinfo("Recalculation Complete", "TPM values have been recalculated for all samples.")
    
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
            print(f"DEBUG: Changed puff interval to {new_interval}")
    
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
            print(f"DEBUG: Changed auto-save interval to {new_minutes} minutes")
    
    def switch_test_dialog(self):
        """Show dialog to switch to a different test in the same file."""
        # This would require integration with the main application
        # For now, just show an info message
        messagebox.showinfo("Switch Test", 
                          "To switch tests, please save your current work and use the main application menu.")
    
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
• Ctrl+O - Open raw Excel file

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
        messagebox.showinfo("About Data Collection", about_text)
    
    # Enhanced saving functionality
    def save_data_internal(self, show_confirmation=True, auto_save=False):
        """
        Internal save method that handles both Excel and VAP3 files.
    
        Args:
            show_confirmation (bool): Whether to show confirmation dialog
            auto_save (bool): Whether this is an auto-save operation
        """
        # End any active editing
        self.end_editing()
    
        # Confirm save if not auto-save
        if show_confirmation and not auto_save:
            if not messagebox.askyesno("Confirm Save", "Save the collected data to the file?"):
                return
    
        print(f"DEBUG: Starting save operation (auto_save: {auto_save})")
    
        # Ensure TPM values are calculated for all samples
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            self.calculate_tpm(sample_id)
    
        # Save to Excel file
        self._save_to_excel()
    
        # Update application state if this is the main file
        self._update_application_state()
    
        # Refresh the main GUI to show updated data (but not during auto-save to avoid disruption)
        if not auto_save:
            print("DEBUG: About to call refresh_main_gui_after_save()")
            self.refresh_main_gui_after_save()
            print("DEBUG: refresh_main_gui_after_save() call completed")
        else:
            print("DEBUG: Skipping GUI refresh for auto-save")
    
        # Mark as saved
        self.has_unsaved_changes = False
    
        if not auto_save:
            print("DEBUG: Save operation completed successfully")
    
    def _save_to_excel(self):
        """Save data to the Excel file."""
        print(f"DEBUG: _save_to_excel() starting - file: {self.file_path}")
    
        # Load the workbook
        wb = openpyxl.load_workbook(self.file_path)
        print(f"DEBUG: Loaded workbook, sheets: {wb.sheetnames}")
    
        # Get the sheet for this test
        if self.test_name not in wb.sheetnames:
            raise Exception(f"Sheet '{self.test_name}' not found in the file.")
        
        ws = wb[self.test_name]
        print(f"DEBUG: Opened sheet '{self.test_name}'")
    
        # Define green fill for TPM cells
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    
        # Track how much data we're actually writing
        total_data_written = 0
    
        # For each sample, write the data
        for sample_idx in range(self.num_samples):
            sample_id = f"Sample {sample_idx+1}"
        
            # Calculate column offset (12 columns per sample)
            col_offset = sample_idx * 12
        
            print(f"DEBUG: Writing data for {sample_id} at column offset {col_offset}")
        
            sample_data_written = 0
        
            # Write the puffs data starting at row 5
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
            
                if not has_data:
                    continue  # Skip empty rows
                
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
            
                # Smell column (F + offset)
                if self.data[sample_id]["smell"][i]:
                    try:
                        ws.cell(row=row, column=6 + col_offset, value=float(self.data[sample_id]["smell"][i]))
                    except:
                        ws.cell(row=row, column=6 + col_offset, value=self.data[sample_id]["smell"][i])
            
                # Notes column (H + offset)
                if self.data[sample_id]["notes"][i]:
                    ws.cell(row=row, column=8 + col_offset, value=str(self.data[sample_id]["notes"][i]))
            
                # TPM column (I + offset) - if calculated
                if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None:
                    tpm_cell = ws.cell(row=row, column=9 + col_offset, value=float(self.data[sample_id]["tpm"][i]))
                    tpm_cell.fill = green_fill
            
                sample_data_written += 1
            
            total_data_written += sample_data_written
            print(f"DEBUG: Wrote {sample_data_written} data rows for {sample_id}")
    
        # Save the workbook
        print(f"DEBUG: Saving workbook with {total_data_written} total data rows written...")

        print("DEBUG: About to save Excel file - verifying data to be saved:")
        for sample_idx in range(min(2, self.num_samples)):  # Check first 2 samples
            sample_id = f"Sample {sample_idx+1}"
            col_offset = sample_idx * 12
            
            print(f"DEBUG: Sample {sample_idx+1} data preview:")
            for i in range(min(3, len(self.data[sample_id]["puffs"]))):  # First 3 rows
                row = i + 5
                puff_val = ws.cell(row=row, column=1 + col_offset).value
                before_val = ws.cell(row=row, column=2 + col_offset).value
                after_val = ws.cell(row=row, column=3 + col_offset).value
                tpm_val = ws.cell(row=row, column=9 + col_offset).value
                print(f"DEBUG:   Row {i}: Puff={puff_val}, Before={before_val}, After={after_val}, TPM={tpm_val}")


        wb.save(self.file_path)
        print(f"DEBUG: Excel file saved successfully to {self.file_path}")
    
    def _update_application_state(self):
        """Update the main application's state if this is a VAP3 file."""
        # Check if the parent has methods to update state
        if hasattr(self.parent, 'filtered_sheets') and hasattr(self.parent, 'file_path'):
            if self.parent.file_path and self.parent.file_path.endswith('.vap3'):
                print("DEBUG: Updating VAP3 file and application state")
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
                        print("DEBUG: VAP3 file updated successfully")
                    else:
                        print("DEBUG: Failed to update VAP3 file")
                        
                except Exception as e:
                    print(f"DEBUG: Error updating VAP3 file: {e}")
    
    # Keep all existing methods for UI creation and interaction
    # (create_sample_tab, update_treeview, editing methods, etc.)
    # These remain the same as in your current implementation
    
    def create_sample_tab(self, parent_frame, sample_id, sample_index):
        """Create a tab for a single sample with fast data entry."""
        # Sample metadata display
        info_frame = ttk.Frame(parent_frame, style='TFrame')
        info_frame.pack(fill="x", pady=(0, 10))

        # Display sample metadata without background color
        ttk.Label(info_frame, 
                 text=f"Sample ID: {self.header_data['samples'][sample_index]['id']}", 
                 style='SampleInfo.TLabel').pack(side="left", padx=(0, 20))
         
        ttk.Label(info_frame, 
                 text=f"Resistance: {self.header_data['samples'][sample_index]['resistance']} Ω", 
                 style='SampleInfo.TLabel').pack(side="left", padx=(0, 20))
         
        ttk.Label(info_frame, 
                 text=f"Voltage: {self.header_data['common']['voltage']} V", 
                 style='SampleInfo.TLabel').pack(side="left", padx=(0, 20))

        # Create a frame for the data table with border
        table_container = ttk.Frame(parent_frame, style='TFrame', relief='solid', borderwidth=1)
        table_container.pack(fill="both", expand=True)

        # Create inner frame for spacing
        table_frame = ttk.Frame(table_container, style='TFrame')
        table_frame.pack(fill="both", expand=True, padx=1, pady=1)

        # Create the treeview (table)
        columns = ("puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes")
        tree = ttk.Treeview(table_frame, columns=columns, show="tree headings", 
                           selectmode="browse", height=20, style='Treeview')

        # Hide the tree column
        tree.column("#0", width=0, stretch=False)

        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        y_scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=y_scrollbar.set)

        # Define column headings
        tree.heading("puffs", text="Puffs")
        tree.heading("before_weight", text="Before weight/g")
        tree.heading("after_weight", text="After weight/g")
        tree.heading("draw_pressure", text="Draw Pressure (kPa)")
        tree.heading("smell", text="Smell")
        tree.heading("notes", text="Notes")

        # Define column widths
        tree.column("puffs", width=80, anchor="center")
        tree.column("before_weight", width=120, anchor="center")
        tree.column("after_weight", width=120, anchor="center")
        tree.column("draw_pressure", width=120, anchor="center")
        tree.column("smell", width=80, anchor="center")
        tree.column("notes", width=150, anchor="w")

        tree.pack(fill="both", expand=True)

        # Add the initial row
        self.update_treeview(tree, sample_id)

        tree.bind("<Button-1>", lambda e: self.on_tree_click(e, tree, sample_id))
        tree.bind("<Double-Button-1>", lambda e: self.on_tree_double_click(e, tree, sample_id))

        tree.bind("<KeyPress>", lambda e: self.start_edit_on_typing(e, tree, sample_id))

        # Bind keyboard navigation
        tree.bind("<Tab>", lambda e: self.handle_tab_key(e, tree, sample_id))
        tree.bind("<Left>", lambda e: self.handle_arrow_key(e, tree, sample_id, "left"))
        tree.bind("<Right>", lambda e: self.handle_arrow_key(e, tree, sample_id, "right"))
        tree.bind("<Up>", lambda e: self.handle_arrow_key(e, tree, sample_id, "up"))
        tree.bind("<Down>", lambda e: self.handle_arrow_key(e, tree, sample_id, "down"))

        print(f"DEBUG: Bound single click handler for sample {sample_id}")

        # Store reference to the tree
        self.data[sample_id]["tree"] = tree

        return tree
    
    def update_treeview(self, tree, sample_id):
        """Update the treeview with current data."""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)

        # Configure tags for alternating row colors
        tree.tag_configure('oddrow', background='#F5F5F5')
        tree.tag_configure('evenrow', background='white')

        # Add rows from data
        for i in range(len(self.data[sample_id]["puffs"])):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        
            # Convert values to strings for display
            values = (
                str(self.data[sample_id]["puffs"][i]) if self.data[sample_id]["puffs"][i] != "" else "",
                str(self.data[sample_id]["before_weight"][i]) if self.data[sample_id]["before_weight"][i] != "" else "",
                str(self.data[sample_id]["after_weight"][i]) if self.data[sample_id]["after_weight"][i] != "" else "",
                str(self.data[sample_id]["draw_pressure"][i]) if self.data[sample_id]["draw_pressure"][i] != "" else "",
                str(self.data[sample_id]["smell"][i]) if self.data[sample_id]["smell"][i] != "" else "",
                str(self.data[sample_id]["notes"][i]) if self.data[sample_id]["notes"][i] != "" else ""
            )
        
            tree.insert("", "end", values=values, tags=(tag,))
    
    def finish_edit(self, event=None):
        """Save the current edit and move to the next cell if needed."""
        if not self.editing or not hasattr(self, 'current_edit'):
            return

        value = self.current_edit["entry"].get()
        tree = self.current_edit["tree"]
        item = self.current_edit["item"]
        column = self.current_edit["column"]
        column_name = self.current_edit["column_name"]
        row_idx = self.current_edit["row_idx"]
        sample_id = self.current_edit["sample_id"]

        # Convert value to appropriate type based on column
        if column_name in ["puffs", "before_weight", "after_weight", "draw_pressure", "smell"]:
            if value.strip():
                try:
                    if column_name == "puffs":
                        converted_value = int(float(value.strip()))
                    else:
                        converted_value = float(value.strip())
                    value = converted_value
                except ValueError:
                    pass  # Keep as string if conversion fails
            else:
                if column_name in ["before_weight", "after_weight", "draw_pressure", "smell"]:
                    value = "" 
                else:
                    value = 0

        # Update data storage
        if row_idx < len(self.data[sample_id][column_name]):
            old_value = self.data[sample_id][column_name][row_idx]
            self.data[sample_id][column_name][row_idx] = value

            # Only mark as changed if value actually changed
            if old_value != value:
                self.mark_unsaved_changes()

                # AUTO-WEIGHT PROGRESSION: If after_weight was changed, update next row's before_weight
                if column_name == "after_weight" and value != "" and value != 0:
                    try:
                        after_weight_value = float(value)
                        next_row_idx = row_idx + 1
                        if next_row_idx < len(self.data[sample_id]["before_weight"]):
                            self.data[sample_id]["before_weight"][next_row_idx] = after_weight_value
                            # Update the tree display for the next row
                            next_item = tree.get_children()[next_row_idx] if next_row_idx < len(tree.get_children()) else None
                            if next_item:
                                next_values = list(tree.item(next_item, "values"))
                                next_values[1] = str(after_weight_value)
                                tree.item(next_item, values=next_values)
                    except (ValueError, TypeError):
                        pass

                # AUTO-PUFF PROGRESSION: If puffs was changed, update all subsequent rows
                if column_name == "puffs" and value != "" and value != 0:
                    try:
                        new_puff_value = int(value)
                        total_rows = len(self.data[sample_id]["puffs"])
                        for subsequent_row_idx in range(row_idx + 1, total_rows):
                            rows_ahead = subsequent_row_idx - row_idx
                            expected_puff = new_puff_value + (rows_ahead * self.puff_interval)
                            self.data[sample_id]["puffs"][subsequent_row_idx] = expected_puff
                            # Update the tree display
                            if subsequent_row_idx < len(tree.get_children()):
                                subsequent_item = tree.get_children()[subsequent_row_idx]
                                subsequent_values = list(tree.item(subsequent_item, "values"))
                                subsequent_values[0] = str(expected_puff)
                                tree.item(subsequent_item, values=subsequent_values)
                    except (ValueError, TypeError):
                        pass

        # Update the tree display
        col_idx = int(column[1:]) - 1
        values = list(tree.item(item, "values"))
        values[col_idx] = str(value) if value != "" else ""
        tree.item(item, values=values)

        # Calculate TPM if weight OR puffs was changed
        if column_name in ["before_weight", "after_weight", "puffs"]:
            self.calculate_tpm(sample_id)
            self.update_stats_panel()

        # End the current edit BEFORE navigation
        self.end_editing()

        # Handle navigation keys
        if event and event.keysym in ["Right", "Left"]:
            if event.keysym == "Right":
                self.handle_arrow_key(event, tree, sample_id, "right")
            elif event.keysym == "Left":
                self.handle_arrow_key(event, tree, sample_id, "left")
    
    def refresh_main_gui_after_save(self):
        """Refresh the main GUI to show updated data after saving."""
        print("DEBUG: Starting main GUI refresh")

        try:
            # Use test_name as the sheet to refresh
            current_sheet_name = self.test_name
            print(f"DEBUG: Target sheet: {current_sheet_name}")
        
            # Check parent and file_manager
            print(f"DEBUG: Has parent: {hasattr(self, 'parent')}")
            print(f"DEBUG: Parent exists: {self.parent is not None}")
            print(f"DEBUG: Has file_manager: {hasattr(self.parent, 'file_manager')}")
            print(f"DEBUG: File manager exists: {self.parent.file_manager is not None}")
        
            # Force reload the file
            if hasattr(self.parent, 'file_manager') and self.parent.file_manager:
                print(f"DEBUG: Force reloading {self.file_path}")
            
                # Clear any existing cache for this file first
                if hasattr(self.parent.file_manager, 'loaded_files_cache'):
                    cache_key = f"{self.file_path}_None"
                    if cache_key in self.parent.file_manager.loaded_files_cache:
                        del self.parent.file_manager.loaded_files_cache[cache_key]
                        print("DEBUG: Cleared file cache")
            
                self.parent.file_manager.load_excel_file(
                    self.file_path, 
                    skip_database_storage=True,
                    force_reload=True
                )
                print("DEBUG: File reloaded")
            else:
                print("DEBUG: Cannot reload - file_manager not available")

            # Update all_filtered_sheets
            if hasattr(self.parent, 'all_filtered_sheets') and hasattr(self.parent, 'current_file'):
                print(f"DEBUG: Updating all_filtered_sheets for {self.parent.current_file}")
                for file_data in self.parent.all_filtered_sheets:
                    if file_data.get("file_name") == self.parent.current_file:
                        file_data["filtered_sheets"] = copy.deepcopy(self.parent.filtered_sheets)
                        print("DEBUG: Updated all_filtered_sheets entry")
                        break

            # Set the sheet selection
            if hasattr(self.parent, 'selected_sheet'):
                if not self.parent.selected_sheet:
                    import tkinter as tk
                    self.parent.selected_sheet = tk.StringVar()
                    print("DEBUG: Created selected_sheet variable")
                self.parent.selected_sheet.set(current_sheet_name)
                print(f"DEBUG: Set sheet to {current_sheet_name}")

            # Update the display
            if hasattr(self.parent, 'update_displayed_sheet'):
                print("DEBUG: Calling update_displayed_sheet")
                self.parent.update_displayed_sheet(current_sheet_name)
                print("DEBUG: Display updated")
            else:
                print("DEBUG: No update_displayed_sheet method available")
        
            # Force GUI update
            if hasattr(self.parent, 'root'):
                self.parent.root.update_idletasks()
                print("DEBUG: Forced GUI update")
    
            print("DEBUG: Refresh complete")

        except Exception as e:
            print(f"ERROR: Refresh failed: {e}")
            import traceback
            traceback.print_exc()

    def calculate_tpm(self, sample_id):
        """Calculate TPM for all rows with before and after weights."""
        calculated_count = 0
        for i in range(len(self.data[sample_id]["puffs"])):
            try:
                before_weight_str = self.data[sample_id]["before_weight"][i]
                after_weight_str = self.data[sample_id]["after_weight"][i]
        
                # Skip if either weight is missing or empty
                if not before_weight_str or before_weight_str == "":
                    continue
                if not after_weight_str or after_weight_str == "":
                    continue
            
                # Convert to float
                try:
                    before_weight = float(before_weight_str)
                    after_weight = float(after_weight_str)
                except (ValueError, TypeError):
                    continue
        
                # Validate that before_weight > after_weight
                if before_weight <= after_weight:
                    continue
        
                # Get puff interval
                try:
                    puff_interval = int(self.data[sample_id]["puffs"][i]) if self.data[sample_id]["puffs"][i] != "" else 0
                except (ValueError, TypeError):
                    continue
        
                # Calculate puffs in this interval
                if i > 0:
                    try:
                        prev_puff = int(self.data[sample_id]["puffs"][i - 1]) if self.data[sample_id]["puffs"][i - 1] != "" else 0
                        puffs_in_interval = puff_interval - prev_puff
                    except (ValueError, TypeError):
                        continue
                else:
                    puffs_in_interval = puff_interval
        
                # Calculate TPM
                if puffs_in_interval > 0:
                    weight_consumed = before_weight - after_weight  # in grams
                    tpm = (weight_consumed * 1000) / puffs_in_interval  # Convert to mg per puff
            
                    # Ensure tpm list is long enough
                    while len(self.data[sample_id]["tpm"]) <= i:
                        self.data[sample_id]["tpm"].append(None)
                
                    self.data[sample_id]["tpm"][i] = round(tpm, 6)
                    calculated_count += 1
            
            except Exception as e:
                # Ensure tpm list is long enough even for failed calculations
                while len(self.data[sample_id]["tpm"]) <= i:
                    self.data[sample_id]["tpm"].append(None)
        
        # Update average TPM for the sample
        valid_tpm_values = [v for v in self.data[sample_id]["tpm"] if v is not None]
        if valid_tpm_values:
            self.data[sample_id]["avg_tpm"] = sum(valid_tpm_values) / len(valid_tpm_values)
        else:
            self.data[sample_id]["avg_tpm"] = 0.0

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
                print(f"WARNING: Weight value {weight_value}g seems unreasonable for row {row_idx}")
                return False
                
            # If this is an after_weight, check that it's less than before_weight
            if column_name == "after_weight":
                before_weight_str = self.data[sample_id]["before_weight"][row_idx]
                if before_weight_str and before_weight_str.strip():
                    try:
                        before_weight = float(before_weight_str.strip())
                        if weight_value >= before_weight:
                            print(f"WARNING: After weight ({weight_value}g) should be less than before weight ({before_weight}g)")
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
                            print(f"WARNING: Before weight ({weight_value}g) should be greater than after weight ({after_weight}g)")
                            return False
                    except ValueError:
                        pass
                        
            return True
            
        except ValueError:
            print(f"ERROR: Invalid weight value '{value}' - must be a number")
            return False
    
    def load_existing_data_from_file(self):
        """Load existing data from the Excel file into the data collection interface."""
        print(f"DEBUG: Loading existing data from file: {self.file_path}")
    
        try:
            # Load the workbook
            wb = openpyxl.load_workbook(self.file_path)
        
            # Check if the test sheet exists
            if self.test_name not in wb.sheetnames:
                print(f"DEBUG: Sheet '{self.test_name}' not found in file")
                return
            
            ws = wb[self.test_name]
            print(f"DEBUG: Successfully opened sheet '{self.test_name}'")
        
            # Load data for each sample
            loaded_data_count = 0
            for sample_idx in range(self.num_samples):
                sample_id = f"Sample {sample_idx + 1}"
                col_offset = sample_idx * 12  # 12 columns per sample
            
                print(f"DEBUG: Loading data for {sample_id} with column offset {col_offset}")
            
                # Read data starting from row 5 (Excel row 5 = index 4)
                row_count = 0
                sample_had_data = False
            
                for row_idx in range(4, ws.max_row + 1):  # Start from row 5 (index 4)
                    excel_row = row_idx + 1  # Convert to 1-based Excel row
                    data_row_idx = row_idx - 4  # Convert to 0-based data index
                
                    # Make sure our data lists are long enough
                    while len(self.data[sample_id]["puffs"]) <= data_row_idx:
                        # Extend all lists if needed
                        for key in ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes", "tpm"]:
                            self.data[sample_id][key].append("" if key != "tpm" else None)
                
                    # Read each column for this row
                    puffs_cell = ws.cell(row=excel_row, column=1 + col_offset)
                    before_weight_cell = ws.cell(row=excel_row, column=2 + col_offset)
                    after_weight_cell = ws.cell(row=excel_row, column=3 + col_offset)
                    draw_pressure_cell = ws.cell(row=excel_row, column=4 + col_offset)
                    smell_cell = ws.cell(row=excel_row, column=6 + col_offset)  # Column F
                    notes_cell = ws.cell(row=excel_row, column=8 + col_offset)  # Column H
                    tpm_cell = ws.cell(row=excel_row, column=9 + col_offset)    # Column I
                
                    # Check if this row has any data (if puffs is empty, assume row is empty)
                    if puffs_cell.value is None or puffs_cell.value == "":
                        # No more data for this sample
                        print(f"DEBUG: {sample_id} - Found end of data at row {data_row_idx}")
                        break
                
                    # Load the data with proper type conversion
                    try:
                        # Puffs - convert to int
                        puffs_value = int(puffs_cell.value) if puffs_cell.value is not None else ""
                        self.data[sample_id]["puffs"][data_row_idx] = puffs_value
                    
                        # Before weight - convert to float or empty string
                        before_weight_value = float(before_weight_cell.value) if before_weight_cell.value is not None else ""
                        self.data[sample_id]["before_weight"][data_row_idx] = before_weight_value
                    
                        # After weight - convert to float or empty string
                        after_weight_value = float(after_weight_cell.value) if after_weight_cell.value is not None else ""
                        self.data[sample_id]["after_weight"][data_row_idx] = after_weight_value
                    
                        # Draw pressure - convert to float or empty string
                        draw_pressure_value = float(draw_pressure_cell.value) if draw_pressure_cell.value is not None else ""
                        self.data[sample_id]["draw_pressure"][data_row_idx] = draw_pressure_value
                    
                        # Smell - convert to float or empty string
                        smell_value = float(smell_cell.value) if smell_cell.value is not None else ""
                        self.data[sample_id]["smell"][data_row_idx] = smell_value
                    
                        # Notes - keep as string
                        notes_value = str(notes_cell.value) if notes_cell.value is not None else ""
                        self.data[sample_id]["notes"][data_row_idx] = notes_value
                    
                        # TPM - convert to float or None
                        tpm_value = float(tpm_cell.value) if tpm_cell.value is not None else None
                        self.data[sample_id]["tpm"][data_row_idx] = tpm_value
                    
                        sample_had_data = True
                        row_count += 1
                    
                        print(f"DEBUG: {sample_id} Row {data_row_idx}: puffs={puffs_value}, before={before_weight_value}, after={after_weight_value}, tpm={tpm_value}")
                    
                    except (ValueError, TypeError) as e:
                        print(f"DEBUG: Error converting data for {sample_id} row {data_row_idx}: {e}")
                        # Keep default values if conversion fails
                        continue
            
                if sample_had_data:
                    loaded_data_count += 1
                    print(f"DEBUG: {sample_id} - Loaded {row_count} rows of existing data")
                else:
                    print(f"DEBUG: {sample_id} - No existing data found")
        
            print(f"DEBUG: Successfully loaded existing data for {loaded_data_count} samples")
        
            # Update all treeviews to show the loaded data
            for i, sample_tree in enumerate(self.sample_trees):
                sample_id = f"Sample {i + 1}"
                self.update_treeview(sample_tree, sample_id)
                print(f"DEBUG: Updated treeview for {sample_id}")
        
            # Recalculate TPM values for all samples
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                self.calculate_tpm(sample_id)
        
            # Update the stats panel
            self.update_stats_panel()
        
            print("DEBUG: Existing data loading completed successfully")
        
        except Exception as e:
            print(f"ERROR: Failed to load existing data from file: {e}")
            import traceback
            traceback.print_exc()
            # Don't show error to user - just continue with empty data
            # messagebox.showerror("Load Error", f"Failed to load existing data: {e}")
    
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
        binding_id = self.window.bind("<Control-s>", lambda e: self.quick_save() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-s>"] = binding_id
        
        # Bind Ctrl+O for open raw Excel
        binding_id = self.window.bind("<Control-o>", lambda e: self.open_raw_excel() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-o>"] = binding_id
        
        # Bind Ctrl+Left/Right for sample navigation
        binding_id = self.window.bind("<Control-Left>", lambda e: self.go_to_previous_sample() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-Left>"] = binding_id
        
        binding_id = self.window.bind("<Control-Right>", lambda e: self.go_to_next_sample() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Control-Right>"] = binding_id
    
    def on_window_close(self):
        """Handle window close event with auto-save."""
        print("DEBUG: Window close event triggered")
        
        # Cancel auto-save timer
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)
        
        # Auto-save if there are unsaved changes
        if self.has_unsaved_changes:
            if messagebox.askyesno("Save Changes", 
                                 "You have unsaved changes. Save before closing?"):
                try:
                    self.save_data_internal(show_confirmation=False)
                    self.result = "load_file"
                except Exception as e:
                    messagebox.showerror("Save Error", f"Failed to save: {e}")
                    return  # Don't close if save failed
            else:
                self.result = "cancel"
        else:
            self.result = "load_file" if self.last_save_time else "cancel"
        
        self.window.destroy()
    
    def save_data(self):
        """Legacy save method for compatibility."""
        self.save_and_exit()
    
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
        print("DEBUG: Showing DataCollectionWindow")
        self.window.wait_window()
        
        # Clean up auto-save timer
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)
        
        print(f"DEBUG: DataCollectionWindow closed with result: {self.result}")
        return self.result

    # Include all the remaining methods from your original implementation
    # (I'm including the essential ones here, but you should merge all editing, 
    # navigation, and UI methods from your current data_collection_window.py)
    
    def start_edit_on_typing(self, event, tree, sample_id):
        """Start editing if a printable character is typed while a cell is selected."""
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
        col_idx = int(column[1:]) - 1

        # REMOVED: if col_idx == 0: return  # Don't allow editing 'puffs' column
        # Now puffs column (col_idx == 0) is editable

        column_name = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][col_idx]

        # Start editing with the key typed
        self.edit_cell(tree, item, column, row_idx, sample_id, column_name)
        if self.current_edit and self.current_edit["entry"]:
            entry = self.current_edit["entry"]
            entry.delete(0, tk.END)
            entry.insert(0, event.char)
            entry.icursor(1)

    def create_tpm_stats_panel(self):
        """Create the TPM statistics panel on the right side."""
        # Clear existing widgets
        for widget in self.stats_frame.winfo_children():
            widget.destroy()
            
        # Create a title without background color
        title_label = ttk.Label(self.stats_frame, text="TPM Statistics", style='Stats.TLabel')
        title_label.pack(pady=(0, 20))
        
        # Create frame for statistics
        stats_container = ttk.Frame(self.stats_frame, style='TFrame')
        stats_container.pack(fill="both", expand=True)
        
        # Create individual stat frames for each sample
        self.tpm_labels = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            sample_name = self.header_data["samples"][i]["id"]
            
            # Create a frame for this sample - using standard ttk frames instead of LabelFrame
            sample_frame = ttk.Frame(stats_container, style='TFrame')
            sample_frame.pack(fill="x", pady=5, padx=10)
            
            # Add a header label
            ttk.Label(sample_frame, 
                     text=f"Sample {i+1}: {sample_name}",
                     style='SampleInfo.TLabel').pack(anchor="w", pady=(0, 5))
            
            # Add a separator
            ttk.Separator(sample_frame, orient="horizontal").pack(fill="x", pady=2)
            
            # Add TPM statistics
            stat_grid = ttk.Frame(sample_frame, style='TFrame')
            stat_grid.pack(fill="x", pady=5, padx=10)
            
            # Row 1: Average TPM
            ttk.Label(stat_grid, text="Average TPM:", style='TLabel').grid(row=0, column=0, sticky="w", pady=2)
            avg_tpm_label = ttk.Label(stat_grid, text="N/A", font=("Arial", 10, "bold"), style='TLabel')
            avg_tpm_label.grid(row=0, column=1, sticky="e", pady=2)
            
            # Row 2: Latest TPM
            ttk.Label(stat_grid, text="Latest TPM:", style='TLabel').grid(row=1, column=0, sticky="w", pady=2)
            latest_tpm_label = ttk.Label(stat_grid, text="N/A", font=("Arial", 10), style='TLabel')
            latest_tpm_label.grid(row=1, column=1, sticky="e", pady=2)
            
            # Row 3: Puff Count
            ttk.Label(stat_grid, text="Puff Count:", style='TLabel').grid(row=2, column=0, sticky="w", pady=2)
            puff_count_label = ttk.Label(stat_grid, text="0", style='TLabel')
            puff_count_label.grid(row=2, column=1, sticky="e", pady=2)
            
            # Store references to labels
            self.tpm_labels[sample_id] = {
                "avg_tpm": avg_tpm_label,
                "latest_tpm": latest_tpm_label,
                "puff_count": puff_count_label
            }
            
        # Update the statistics for the current sample
        self.update_stats_panel()
    
    def update_stats_panel(self, event=None):
        """Update the TPM statistics panel based on current data."""
        # Update stats for all samples
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            
            # Calculate TPM values if needed
            self.calculate_tpm(sample_id)
            
            # Get TPM values (filtering out None values)
            tpm_values = [v for v in self.data[sample_id]["tpm"] if v is not None]
            
            # Update labels
            if tpm_values:
                avg_tpm = sum(tpm_values) / len(tpm_values)
                self.data[sample_id]["avg_tpm"] = avg_tpm
                self.tpm_labels[sample_id]["avg_tpm"].config(text=f"{avg_tpm:.6f}")
                self.tpm_labels[sample_id]["latest_tpm"].config(text=f"{tpm_values[-1]:.6f}")
                puff_count = self.data[sample_id]["puffs"][-1] if self.data[sample_id]["puffs"] else 0
                self.tpm_labels[sample_id]["puff_count"].config(text=str(puff_count))
            else:
                self.tpm_labels[sample_id]["avg_tpm"].config(text="N/A")
                self.tpm_labels[sample_id]["latest_tpm"].config(text="N/A")
                self.tpm_labels[sample_id]["puff_count"].config(text="0")

    # Add all remaining methods from your current implementation
    # These include all the cell editing, navigation, and interaction methods
    # For brevity, I'm not including them all here, but you should merge them from your current file
    
    def go_to_previous_sample(self):
        """Navigate to the previous sample tab."""
        if not self.hotkeys_enabled:
            return
            
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab > 0:
            self.notebook.select(current_tab - 1)
        else:
            # Wrap around to last tab
            self.notebook.select(len(self.sample_frames) - 1)
    
    def go_to_next_sample(self):
        """Navigate to the next sample tab."""
        if not self.hotkeys_enabled:
            return
            
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab < len(self.sample_frames) - 1:
            self.notebook.select(current_tab + 1)
        else:
            # Wrap around to first tab
            self.notebook.select(0)
    
    def update_puff_interval(self):
        """Update the puff interval for future rows."""
        try:
            self.puff_interval = self.puff_interval_var.get()
            print(f"DEBUG: Updated puff interval to {self.puff_interval}")
        except:
            messagebox.showerror("Error", "Invalid puff interval. Please enter a positive number.")
            self.puff_interval_var.set(self.puff_interval)

    # ============================================================================
    # COMPLETE CELL EDITING AND NAVIGATION METHODS
    # ============================================================================
    
    def edit_cell(self, tree, item, column, row_idx, sample_id, column_name):
        """Create an entry widget for editing a cell."""
        # Cancel any existing edit
        self.end_editing()
        
        # Remove cell border during editing
        if hasattr(self, 'cell_border_frame') and self.cell_border_frame:
            self.cell_border_frame.destroy()
            self.cell_border_frame = None

        # Mark that we're editing
        self.editing = True
        self.hotkeys_enabled = False
        
        # Get the current value
        current_value = ""
        if row_idx < len(self.data[sample_id][column_name]):
            current_value = self.data[sample_id][column_name][row_idx]
        
        # Get the cell coordinates
        x, y, width, height = tree.bbox(item, column)
        
        # Create a frame for the entry
        frame = tk.Frame(tree, borderwidth=0, highlightthickness=1, highlightbackground="black")
        frame.place(x=x, y=y, width=width, height=height)
        
        # Create the entry widget
        entry = tk.Entry(frame, borderwidth=0)
        entry.pack(fill="both", expand=True)
        
        # Set current value
        if current_value:
            entry.insert(0, current_value)
        
        # Select all text
        entry.select_range(0, tk.END)
        
        # Save references
        self.current_edit = {
            "frame": frame,
            "entry": entry,
            "tree": tree,
            "item": item,
            "column": column,
            "column_name": column_name,
            "row_idx": row_idx,
            "sample_id": sample_id
        }
        
        # Focus the entry
        entry.focus_set()
        
        # Bind events for the entry widget
        entry.bind("<Tab>", lambda e: self.move_to_next_cell_during_edit(e, tree, sample_id, direction="right"))
        entry.bind("<Return>", lambda e: self.move_to_next_cell(tree, sample_id, direction="down"))
        entry.bind("<Escape>", self.cancel_edit)
        entry.bind("<FocusOut>", self.finish_edit)
        entry.bind("<Left>", self.handle_arrow_in_edit)
        entry.bind("<Right>", self.handle_arrow_in_edit)
        
        print(f"DEBUG: Started editing cell - sample: {sample_id}, row: {row_idx}, column: {column_name}")
    
    def on_tree_click(self, event, tree, sample_id):
        """Delay single-click action to detect double-clicks."""
        # Cancel any previous scheduled single-click
        if self.single_click_after_id:
            self.window.after_cancel(self.single_click_after_id)
            self.single_click_after_id = None

        # Schedule the single-click action after 300ms
        self.single_click_after_id = self.window.after(
            300,
            lambda: self._handle_single_click(event, tree, sample_id)
        )

    def on_tree_double_click(self, event, tree, sample_id):
        """Handle double-click and cancel pending single-click."""
        if self.single_click_after_id:
            self.window.after_cancel(self.single_click_after_id)
            self.single_click_after_id = None

        item = tree.identify("item", event.x, event.y)
        column = tree.identify("column", event.x, event.y)
        region = tree.identify("region", event.x, event.y)

        if not item or not column or region != "cell":
            return

        col_idx = int(column[1:])
        if col_idx == 1:
            print("DEBUG: Double-click on puffs column — not editable")
            return

        row_idx = tree.index(item)
        column_name = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][col_idx - 1]
        print(f"DEBUG: Double-click - editing item: {item}, column: {column}, row_idx: {row_idx}")
        self.edit_cell(tree, item, column, row_idx, sample_id, column_name)

    def _handle_single_click(self, event, tree, sample_id):
        """Actual logic for handling a single-click."""
        item = tree.identify("item", event.x, event.y)
        column = tree.identify("column", event.x, event.y)
        region = tree.identify("region", event.x, event.y)

        if not item or not column or region != "cell":
            return

        self.end_editing()
        self.current_column = column
        self.current_item = item
        self.highlight_cell(tree, item, column)
        print(f"DEBUG: Single-click - item: {item}, column: {column}")

    def move_to_next_cell_during_edit(self, event, tree, sample_id, direction="right"):
        """Handle Tab navigation during editing."""
        if not self.editing or not self.current_edit:
            return "break"

        current_item = self.current_edit["item"]
        current_column = self.current_edit["column"]
        row_idx = self.current_edit["row_idx"]
        col_idx = int(current_column[1:])

        self.finish_edit()

        items = tree.get_children()

        if direction == "right":
            if col_idx < 6:
                # Go to next column in the same row
                next_column = f"#{col_idx + 1}"
                column_name = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][col_idx]
                self.edit_cell(tree, current_item, next_column, row_idx, sample_id, column_name)
            elif row_idx < len(items) - 1:
                # Wrap to next row, first editable column
                next_item = items[row_idx + 1]
                next_column = "#2"
                column_name = "before_weight"
                self.edit_cell(tree, next_item, next_column, row_idx + 1, sample_id, column_name)

        return "break"

    def move_to_next_cell(self, tree, sample_id, direction="right"):
        """Handle Tab or Enter navigation during editing."""
        self.finish_edit()

        item = self.current_item
        column = self.current_column
        items = tree.get_children()

        if not item or not column:
            return "break"

        col_idx = int(column[1:])
        row_idx = items.index(item)

        if direction == "right":
            next_col_idx = col_idx + 1
            if next_col_idx > 6:
                return "break"
            next_column = f"#{next_col_idx}"
            column_name = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][next_col_idx - 1]
            self.edit_cell(tree, item, next_column, row_idx, sample_id, column_name)
        elif direction == "down":
            if row_idx >= len(items) - 1:
                return "break"
            next_item = items[row_idx + 1]
            tree.selection_set(next_item)
            tree.focus(next_item)
            tree.see(next_item)
            column_name = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][col_idx - 1]
            self.edit_cell(tree, next_item, f"#{col_idx}", row_idx + 1, sample_id, column_name)

        return "break"

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
                self.edit_cell(tree, item, "#1", 0, sample_id, "puffs")  # Start with puffs column
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
                self.edit_cell(tree, next_item, "#1", idx + 1, sample_id, "puffs")
            else:
                # At the last row, move to next sample
                self.go_to_next_sample()
        else:
            # Edit the next column in the same row
            self.current_column = f"#{next_col_idx}"
            self.highlight_cell(tree, item, self.current_column)
        
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
            print(f"DEBUG: Arrow navigation - moved to row {row_idx}, column {next_column}")
        
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
                    print(f"DEBUG: Arrow navigation - moved down to row {current_idx + 1}")
            else:  # up
                if current_idx > 0:
                    prev_item = items[current_idx - 1]
                    tree.selection_set(prev_item)
                    tree.focus(prev_item)
                    tree.see(prev_item)
                    self.current_column = current_column
                    self.highlight_cell(tree, prev_item, self.current_column)
                    print(f"DEBUG: Arrow navigation - moved up to row {current_idx - 1}")
    
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
            print(f"DEBUG: Arrow key navigation from edit - {event.keysym}")
            self.finish_edit(event)
            return "break"  # Stop event propagation

    def cancel_edit(self, event=None):
        """Cancel the current edit without saving."""
        print("DEBUG: Edit cancelled")
        self.end_editing()
        return "break"  # Stop event propagation

    def end_editing(self):
        """Clean up editing widgets and state."""
        if not self.editing:
            return
    
        if hasattr(self, 'current_edit') and self.current_edit:
            if "frame" in self.current_edit:
                self.current_edit["frame"].destroy()
            self.current_edit = None
    
        self.editing = False
        self.hotkeys_enabled = True
    
        # Also remove the cell border frame if it exists
        if hasattr(self, 'cell_border_frame') and self.cell_border_frame is not None:
            self.cell_border_frame.destroy()
            self.cell_border_frame = None
        
        print("DEBUG: Editing ended and cleanup completed")

    def highlight_cell(self, tree, item, column):
        """Highlight a specific cell instead of the entire row."""
        # Remove any existing cell highlights
        self.clear_cell_highlights(tree)
    
        # Create a unique tag for this cell
        cell_tag = f"selected_cell_{item}_{column}"
    
        # Get all existing tags for this item
        existing_tags = list(tree.item(item, 'tags'))
    
        # Add our cell selection tag
        if cell_tag not in existing_tags:
            existing_tags.append(cell_tag)
    
        # Apply the tags
        tree.item(item, tags=existing_tags)
    
        # Configure the cell highlight
        # We'll use a custom approach to highlight individual cells
        col_idx = int(column[1:]) - 1
    
        # Store current selection info
        self.selected_cell = {
            'tree': tree,
            'item': item,
            'column': column,
            'col_idx': col_idx
        }
    
        # Use selection to track the row but we'll add visual indicator for the cell
        tree.selection_set(item)
        tree.focus(item)
    
        # Draw a border around the selected cell
        self.draw_cell_border(tree, item, column)

    def clear_cell_highlights(self, tree):
        """Clear all cell highlight tags."""
        for item in tree.get_children():
            tags = list(tree.item(item, 'tags'))
            # Remove any cell selection tags
            tags = [tag for tag in tags if not tag.startswith('selected_cell_')]
            tree.item(item, tags=tags)
    
        # Remove any existing cell border
        if hasattr(self, 'cell_border_frame') and self.cell_border_frame is not None:
            self.cell_border_frame.destroy()
            self.cell_border_frame = None

    def draw_cell_border(self, tree, item, column):
        """Draw a transparent border around a cell without obscuring its content."""
        # Remove old border
        if self.cell_border_frame:
            self.cell_border_frame.destroy()
            self.cell_border_frame = None

        bbox = tree.bbox(item, column)
        if not bbox:
            return

        x, y, width, height = bbox

        # Draw an invisible frame with just border
        self.cell_border_frame = tk.Frame(
            tree, bg="", highlightbackground="blue",
            highlightthickness=2, bd=0
        )
        self.cell_border_frame.place(x=x - 1, y=y - 1, width=width + 2, height=height + 2)
        self.cell_border_frame.lift()  # Ensure it's above background but doesn't obscure text

    def handle_tab_in_edit(self, event):
        """Handle tab key pressed while editing a cell."""
        print("DEBUG: Tab pressed during edit")
    
        # Store the current edit info before finishing
        current_tree = self.current_edit["tree"]
        current_sample_id = self.current_edit["sample_id"]
    
        # Save the current edit
        self.finish_edit(event)
    
        # Now handle tab navigation
        self.handle_tab_key(event, current_tree, current_sample_id)
    
        return "break"  # Stop event propagation

    # ============================================================================
    # END OF CELL EDITING AND NAVIGATION METHODS
    # ============================================================================
"""
File Management Module for DataViewer Application

This module handles all file-related operations including:
- Loading Excel files (standard and legacy formats)
- Saving and loading .vap3 files (custom format)
- Opening Excel files for editing and monitoring changes
- Managing file selections and UI updates related to files
"""

# Standard library imports
import os
import copy
import shutil
import time
import uuid
import tempfile
import threading
import subprocess
import traceback
from typing import Optional

# Third party imports
import pandas as pd
import psutil
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, Button, ttk, Frame

# Local imports
import processing
from utils import (
    is_valid_excel_file,
    get_resource_path,
    load_excel_file,
    get_save_path,
    is_standard_file,
    FONT,
    APP_BACKGROUND_COLOR
)

class FileManager:
    """File Management Module for DataViewer.
    
    This class is initialized with a reference to the main TestingGUI instance
    so that it can update its state (sheets, selected_sheet, etc.).
    """
    def __init__(self, gui):
        self.gui = gui
        self.root = gui.root
        
        
    def load_excel_file(self, file_path, legacy_mode: str = None) -> None:
        """
        Load the selected Excel file and process its sheets.
        For legacy files (non-standard), convert the file to the proper format,
        move the new formatted file into a folder called 'legacy data', and load that copy.
        """
       
        try:
            # Ensure the file is a valid Excel file.
            if not is_valid_excel_file(os.path.basename(file_path)):
                raise ValueError(f"Invalid Excel file selected: {file_path}")

            if not is_standard_file(file_path):
                # Legacy file: convert to proper format.
                legacy_dir = os.path.join(os.path.abspath("."), "legacy data")
                if not os.path.exists(legacy_dir):
                    os.makedirs(legacy_dir)
                
                legacy_wb = load_workbook(file_path)
                legacy_sheetnames = legacy_wb.sheetnames

                # Load the default template to get its sheet names.
                template_path_default = os.path.join(os.path.abspath("."), "resources", 
                                         "Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
                wb_template = load_workbook(template_path_default)
                template_sheet_names = wb_template.sheetnames

                # If legacy_mode is not explicitly provided, choose based on heuristic.
                if legacy_mode is None:
                    if len(legacy_sheetnames) == 1 and legacy_sheetnames[0] not in template_sheet_names:
                        legacy_mode = "file"
                    else:
                        legacy_mode = "standards"

                final_sheets = {}
                full_sample_data = pd.DataFrame()

                if legacy_mode == "file":

                    converted = processing.convert_legacy_file_using_template(file_path)
                    key = f"Legacy_{os.path.basename(file_path)}"
                    final_sheets = {key: {"data": converted, "is_empty": converted.empty}}
                    full_sample_data = converted

                    self.gui.filtered_sheets = final_sheets
                    self.gui.full_sample_data = full_sample_data
                    default_key = list(final_sheets.keys())[0]
                    self.gui.selected_sheet.set(default_key)

                elif legacy_mode == "standards":

                    print("Legacy Standard File. Processing")
                    
                    converted_dict = processing.convert_legacy_standards_using_template(file_path)
                    
                    self.gui.sheets = converted_dict
                    self.gui.filtered_sheets = {
                        name: {"data": data, "is_empty": data.empty}
                        for name, data in self.gui.sheets.items()
                    }
                    self.gui.full_sample_data = pd.concat([sheet_info["data"] for sheet_info in self.gui.filtered_sheets.values()], axis=1)
                    first_sheet = list(self.gui.filtered_sheets.keys())[0]
                    self.gui.selected_sheet.set(first_sheet)
                    self.gui.update_displayed_sheet(first_sheet)

                                # Debugging: Print the first 20 rows and 15 columns for 'Intense Test' in legacy file
                    if "Intense Test" in self.gui.filtered_sheets:
                        intense_test_df = self.gui.filtered_sheets["Intense Test"]["data"]
                        #print("\n--- Legacy Standard File: 'Intense Test' ---")
                        #print(intense_test_df.iloc[:20, :15])
                else:
                    raise ValueError(f"Unknown legacy mode: {legacy_mode}")
            else:
                # Standard file processing.
                print("Standard File. Processing")
                self.gui.sheets = load_excel_file(file_path)
                self.gui.filtered_sheets = {
                    name: {"data": data, "is_empty": data.empty}
                    for name, data in self.gui.sheets.items()
                }
                self.gui.full_sample_data = pd.concat([sheet_info["data"] for sheet_info in self.gui.filtered_sheets.values()],axis=1)
                first_sheet = list(self.gui.filtered_sheets.keys())[0]
                self.gui.selected_sheet.set(first_sheet)
                self.gui.update_displayed_sheet(first_sheet)
                # Debugging: Print the first 20 rows and 15 columns for 'Intense Test' in normal file
                            # Debugging: Print the first 20 rows and 15 columns for 'Intense Test' in normal file
                if "Intense Test" in self.gui.filtered_sheets:
                    intense_test_df = self.gui.filtered_sheets["Intense Test"]["data"]
                    #print("\n--- Normal Standard File: 'Intense Test' ---")
                    #print(intense_test_df.iloc[:20, :15])


        except Exception as e:
            messagebox.showerror("Error", f"Error occurred while loading file: {e}")

    def load_initial_file(self) -> None:
        """Handle file loading directly on the main thread."""
        file_paths = filedialog.askopenfilenames(
            title="Select Excel File(s)", 
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_paths:
            messagebox.showinfo("Info", "No files were selected")
            return

        self.gui.progress_dialog.show_progress_bar("Loading files...")
        total_files = len(file_paths)
    
        try:
            for index, file_path in enumerate(file_paths, 1):
                # Update progress
                progress = int((index / total_files) * 100)
                self.gui.progress_dialog.update_progress_bar(progress)
            
                # Load and process file
                self.load_excel_file(file_path)
            
                # Store results
                self.gui.all_filtered_sheets.append({
                    "file_name": os.path.basename(file_path),
                    "file_path": file_path,
                    "filtered_sheets": copy.deepcopy(self.gui.filtered_sheets)
                })
            
                # Allow GUI to refresh
                self.root.update_idletasks()

            # Final setup after loading all files
            if self.gui.all_filtered_sheets:
                last_file = self.gui.all_filtered_sheets[-1]
                self.set_active_file(last_file["file_name"])
                self.update_file_dropdown()
                self.update_ui_for_current_file()

        except Exception as e:
            messagebox.showerror("Loading Error", f"Failed to load files: {str(e)}")
        finally:
            self.gui.progress_dialog.hide_progress_bar()
            self.root.update_idletasks()  # Ensure GUI refreshes

    def set_active_file(self, file_name: str) -> None:
        """Set the active file based on the given file name."""
        for file_data in self.gui.all_filtered_sheets:
            if file_data["file_name"] == file_name:
                self.gui.current_file = file_name
                self.gui.file_path = file_data.get("file_path", None)
                self.gui.filtered_sheets = file_data["filtered_sheets"]
                if self.gui.file_path is None:
                    raise ValueError(f"No file path associated with the file '{file_name}'.")
                break
        else:
            raise ValueError(f"File '{file_name}' not found.")

    def update_ui_for_current_file(self) -> None:
        """Update UI components to reflect the currently active file."""
        if not self.gui.current_file:
            return
        self.gui.file_dropdown_var.set(self.gui.current_file)
        self.gui.populate_or_update_sheet_dropdown()
        current_sheet = self.gui.selected_sheet.get()
        if current_sheet not in self.gui.filtered_sheets:
            first_sheet = list(self.gui.filtered_sheets.keys())[0] if self.gui.filtered_sheets else None
            if first_sheet:
                self.gui.selected_sheet.set(first_sheet)
                self.gui.update_displayed_sheet(first_sheet)
        else:
            self.gui.update_displayed_sheet(current_sheet)

    def add_data(self) -> None:
        """Handle adding a new data file directly and update UI accordingly."""
        file_paths = filedialog.askopenfilenames(
            title="Select Excel File(s)", filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_paths:
            messagebox.showinfo("Info", "No files were selected")
            return

        for file_path in file_paths:
            self.load_excel_file(file_path)
            self.gui.all_filtered_sheets.append({
                "file_name": os.path.basename(file_path),
                "file_path": file_path,
                "filtered_sheets": copy.deepcopy(self.gui.filtered_sheets)
            })
        self.update_file_dropdown()
        last_file = self.gui.all_filtered_sheets[-1]
        self.set_active_file(last_file["file_name"])
        self.update_ui_for_current_file()
        messagebox.showinfo("Success", f"Data from {len(file_paths)} file(s) added successfully.")

    def ask_open_file(self) -> str:
        """Prompt the user to select an Excel file."""
        return filedialog.askopenfilename(
            title="Select Standardized Testing File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

    def reload_excel_file(self) -> None:
        """Reload the Excel file into the program, preserving the state of the UI."""
        try:
            self.load_excel_file(self.gui.file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to reload the Excel file: {e}")

    def open_raw_data_in_excel(self, sheet_name=None) -> None:
        """
        Opens an Excel file for a specific sheet, allowing edits, and then
        updates the application when Excel closes.
    
        This improved version forces Excel to open in a new instance and has robust
        file tracking, preventing errors when other Excel files are already open.
    
        Args:
            sheet_name (str, optional): Name of the sheet to open. If None,
                                        uses the currently selected sheet.
        """
        try:
            # Validate the current state
            if not hasattr(self.gui, 'filtered_sheets') or not self.gui.filtered_sheets:
                messagebox.showerror("Error", "No data is currently loaded.")
                return
        
            # If no sheet specified, use the currently selected sheet
            if not sheet_name and hasattr(self.gui, 'selected_sheet'):
                sheet_name = self.gui.selected_sheet.get()
        
            if not sheet_name or sheet_name not in self.gui.filtered_sheets:
                messagebox.showerror("Error", "Please select a valid sheet first.")
                return
        
            # Get the sheet data
            sheet_data = self.gui.filtered_sheets[sheet_name]['data']
        
            # Create a temporary Excel file in a reliable location, with a simple filename
            import tempfile
            import uuid
            from openpyxl import Workbook
        
            # Generate a unique identifier that's short but still unique
            unique_id = str(uuid.uuid4()).split('-')[0]
        
            # Ensure the sheet name only contains valid characters
            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c == ' ')
            safe_sheet_name = safe_sheet_name.replace(' ', '_')[:15]  # Keep it reasonably short
        
            # Create a temporary file in the user's temp directory
            temp_dir = os.path.abspath(tempfile.gettempdir())
            temp_file = os.path.join(temp_dir, f"dataviewer_{safe_sheet_name}_{unique_id}.xlsx")
        
            #print(f"Creating temporary file at: {temp_file}")
        
            # Create a new workbook with just this sheet
            wb = Workbook()
            ws = wb.active
            # Ensure sheet name is valid for Excel (31 chars max, no special chars)
            ws.title = safe_sheet_name[:31]
        
            # Write the headers
            for col_idx, column_name in enumerate(sheet_data.columns, 1):
                ws.cell(row=1, column=col_idx, value=str(column_name))
        
            # Write the data
            for row_idx, row in enumerate(sheet_data.itertuples(index=False), 2):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
                
            # Save the workbook
            wb.save(temp_file)
        
            # Make sure the file exists before proceeding
            if not os.path.exists(temp_file):
                raise FileNotFoundError(f"Failed to create temporary file at {temp_file}")
        
            # Record the file modification time before opening
            original_mod_time = os.path.getmtime(temp_file)
        
            # Create a small status bar notification
            status_text = f"Opening {sheet_name} in Excel. Changes will be imported when Excel closes."
            status_label = ttk.Label(self.gui.root, text=status_text, relief="sunken", anchor="w")
            status_label.pack(side="bottom", fill="x")
            self.gui.root.update_idletasks()
        
            # Use a class to hold shared state that can be modified in the thread
            class FileMonitorState:
                def __init__(self, initial_mod_time):
                    self.last_mod_time = initial_mod_time
                    self.has_changed = False
        
            # Create a state object with the initial modification time
            monitor_state = FileMonitorState(original_mod_time)
        
            # ----- THE KEY DIFFERENCE: FORCE A NEW EXCEL INSTANCE -----
            # Windows Registry by default opens Excel files in the same instance
            # We need to bypass this by using the /x switch to start a new instance
            import subprocess
        
            # The command to force a new Excel instance 
            cmd = f'start /wait "" "excel.exe" /x "{os.path.abspath(temp_file)}"'
        
            # Execute the command which opens Excel in a new instance
            # Note: we use subprocess.call with shell=True to properly handle the start command
            try:
                subprocess.Popen(cmd, shell=True)
                #print(f"Launched Excel with command: {cmd}")
            except Exception as e:
                print(f"Error launching Excel with command: {e}")
                # Fall back to the standard method as a last resort
                os.startfile(os.path.abspath(temp_file))
                #print("Fell back to os.startfile method")
            
            # Wait a moment for Excel to start up
            time.sleep(2.0)
        
            # Start a background thread to monitor the file
            import threading
        
            def monitor_file_lock():
                """Monitor the Excel file until it's no longer locked, then process any changes."""
                file_open = True
            
                #print(f"Starting file lock monitoring for {temp_file}")
            
                while file_open:
                    # Check if the file is still locked by Excel
                    file_locked = False
                    try:
                        # Try to open the file for exclusive access - if it fails, Excel still has it open
                        with open(temp_file, 'r+b') as test_lock:
                            # Successfully opened file - not locked
                            pass
                    except PermissionError:
                        # File is still locked (open in Excel)
                        file_locked = True
                        #print("File is locked - Excel is still using it")
                    except FileNotFoundError:
                        # File was deleted or moved
                        #print(f"File not found during monitoring: {temp_file}")
                        file_locked = False
                        file_open = False
                    except Exception as e:
                        # Some other error - assume file is not locked
                        #print(f"Error checking file lock: {e}")
                        file_locked = False
                
                    # Update file_open status based on lock check
                    file_open = file_locked
                
                    # If file exists, check if it has been modified 
                    if os.path.exists(temp_file):
                        try:
                            current_mod_time = os.path.getmtime(temp_file)
                            if current_mod_time > monitor_state.last_mod_time:
                                #print(f"File modification detected. Old time: {monitor_state.last_mod_time}, New time: {current_mod_time}")
                                monitor_state.has_changed = True
                                monitor_state.last_mod_time = current_mod_time
                        except Exception as e:
                            #print(f"Error checking file modification time: {e}")
                            pass
                    else:
                        #print(f"Warning: File no longer exists at {temp_file}")
                        pass
                    # If the file is no longer open, process any changes
                    if not file_open:
                        #print("File is no longer locked - processing changes")
                        # Use a delay before processing to make sure Excel fully releases the file
                        self.gui.root.after(500, lambda: self._process_excel_changes(temp_file, sheet_name, monitor_state.has_changed, status_label))
                        break
                
                    # Sleep briefly to reduce CPU usage
                    time.sleep(1.0)
        
            # Start the monitoring thread
            monitor_thread = threading.Thread(target=monitor_file_lock, daemon=True)
            self.gui.threads.append(monitor_thread)
            monitor_thread.start()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {e}")
            import traceback
            traceback.print_exc()
            # Clean up any UI changes if there was an error
            for widget in self.gui.root.winfo_children():
                if isinstance(widget, ttk.Label) and "Opening" in widget.cget("text"):
                    widget.destroy()

    def _process_excel_changes(self, temp_file, sheet_name, file_changed, status_label=None):
        """
        Process changes made in Excel after it closes.
    
        Args:
            temp_file (str): Path to the temporary Excel file
            sheet_name (str): The sheet name that was edited
            file_changed (bool): Whether the file was modified
            status_label (ttk.Label, optional): Status label to update and remove
        """
        try:
            # Remove the status label if it exists
            if status_label and status_label.winfo_exists():
                status_label.destroy()
        
            #print(f"Processing Excel changes for {sheet_name}. File changed: {file_changed}")
            #print(f"File path: {temp_file}")
            #print(f"File exists: {os.path.exists(temp_file)}")
        
            if file_changed and os.path.exists(temp_file):
                try:
                    # Read the modified data with a retry mechanism
                    max_retries = 3
                    retry_count = 0
                    read_success = False
                    modified_data = None
                
                    while not read_success and retry_count < max_retries:
                        try:
                            #print(f"Attempting to read Excel file, attempt {retry_count + 1}")
                            modified_data = pd.read_excel(temp_file)
                            read_success = True
                            #print("Successfully read modified Excel data")
                        except Exception as read_error:
                            retry_count += 1
                            #print(f"Error reading Excel file (attempt {retry_count}): {read_error}")
                            time.sleep(1.0)  # Wait before retrying
                
                    if not read_success:
                        raise Exception(f"Failed to read Excel file after {max_retries} attempts")
                
                    # Update the filtered sheets with new data
                    self.gui.filtered_sheets[sheet_name]['data'] = modified_data
                
                    # If we're using a VAP3 file, update it
                    if hasattr(self.gui, 'file_path') and self.gui.file_path.endswith('.vap3'):
                        from vap_file_manager import VapFileManager
                    
                        vap_manager = VapFileManager()
                        #print("Updating VAP3 file with modified data")
                        # Save to the same VAP3 file
                        vap_manager.save_to_vap3(
                            self.gui.file_path,
                            self.gui.filtered_sheets,
                            self.gui.sheet_images,
                            self.gui.plot_options,
                            getattr(self.gui, 'image_crop_states', {})
                        )
                
                    # Refresh the GUI display
                    self.gui.update_displayed_sheet(sheet_name)
                
                    # Show a temporary success message in the status bar
                    success_label = ttk.Label(
                        self.gui.root, 
                        text="Changes from Excel have been imported successfully.",
                        relief="sunken", 
                        anchor="w"
                    )
                    success_label.pack(side="bottom", fill="x")
                
                    # Auto-remove the success message after 3 seconds
                    self.gui.root.after(3000, lambda: success_label.destroy() if success_label.winfo_exists() else None)
                    #print("Excel changes successfully processed and imported")
                
                except Exception as e:
                    #print(f"Error processing Excel changes: {e}")
                    import traceback
                    traceback.print_exc()
                    messagebox.showerror("Error", f"Failed to read modified Excel file: {e}")
            elif file_changed and not os.path.exists(temp_file):
                #print("File was modified but no longer exists")
                messagebox.showinfo("Information", "Excel file was modified but appears to have been moved or renamed. Changes could not be imported.")
            else:
                # File was closed without changes
                #print("Excel file was closed without changes")
                info_label = ttk.Label(
                    self.gui.root, 
                    text="Excel file was closed without changes.",
                    relief="sunken", 
                    anchor="w"
                )
                info_label.pack(side="bottom", fill="x")
                self.gui.root.after(3000, lambda: info_label.destroy() if info_label.winfo_exists() else None)
            
        except Exception as e:
            #print(f"Exception in _process_excel_changes: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to process Excel changes: {e}")
        finally:
            # Always try to clean up the temp file
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                    #print(f"Temporary file removed: {temp_file}")
                except Exception as cleanup_error:
                    #print(f"Warning: Failed to remove temporary file: {cleanup_error}")
                    pass
                    # Don't disrupt the user's workflow with an error message for cleanup issues

    def create_new_template(self, startup_menu: tk.Toplevel) -> None:
        """
        Handle the 'New' button click to create a new template file.
        """
        try:
            startup_menu.destroy()
            template_path = get_resource_path("resources/Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
            if not os.path.exists(template_path):
                messagebox.showerror("Error", "Template file not found. Please check the resources folder.")
                return
            new_file_path = filedialog.asksaveasfilename(
                title="Save New Test File As",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not new_file_path:
                return
            shutil.copy(template_path, new_file_path)
            self.gui.file_path = new_file_path
            self.load_excel_file(new_file_path)
            self.gui.clear_dynamic_frame()
            self.gui.all_filtered_sheets.append({
                "file_name": os.path.basename(new_file_path),
                "file_path": new_file_path,
                "filtered_sheets": copy.deepcopy(self.gui.filtered_sheets)
            })
            self.set_active_file(os.path.basename(new_file_path))
            self.update_ui_for_current_file()
            messagebox.showinfo("Success", "New template created and loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating a new template: {e}")

    def start_file_loading_wrapper(self, startup_menu: tk.Toplevel) -> None:
        """Handle the 'Load' button click in the startup menu."""
        startup_menu.destroy()
        self.load_initial_file()

    def save_to_new_file(self, data, original_file_path) -> None:
        new_file_path = get_save_path
        if not new_file_path:
            return
        with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
            data.to_excel(writer, index=False, sheet_name='Sheet1')
        messagebox.showinfo("Save Successful", f"Data saved to {new_file_path}")

    def update_file_dropdown(self) -> None:
        """Update the file dropdown with loaded file names."""
        file_names = [file_data["file_name"] for file_data in self.gui.all_filtered_sheets]
        #print("DEBUG: file_names = ", file_names)
        self.gui.file_dropdown["values"] = file_names
        if file_names:
            self.gui.file_dropdown_var.set(file_names[-1])
            self.gui.file_dropdown.update_idletasks()

    def add_or_update_file_dropdown(self) -> None:
        """Add a file selection dropdown or update its values if it already exists."""
        if not hasattr(self.gui, 'file_dropdown') or not self.gui.file_dropdown:
            dropdown_frame = ttk.Frame(self.gui.top_frame, width=1400, height=40)
            dropdown_frame.pack(side="left", pady=2, padx=5)
            file_label = ttk.Label(dropdown_frame, text="Select File:", font=FONT, foreground="white", background=APP_BACKGROUND_COLOR)
            file_label.pack(side="left", padx=(0, 0))
            self.gui.file_dropdown_var = tk.StringVar()
            self.gui.file_dropdown = ttk.Combobox(
                dropdown_frame,
                textvariable=self.gui.file_dropdown_var,
                state="readonly",
                font=FONT,
                width=10
            )
            self.gui.file_dropdown.pack(side="left", fill="x", expand=True, padx=(5, 5))
            self.gui.file_dropdown.bind("<<ComboboxSelected>>", self.gui.on_file_selection)
        self.update_file_dropdown()

# ==================== .vap3 File FUNCTIONS ====================

    def save_as_vap3(self, filepath=None) -> None:
        """
        Save the current session to a .vap3 file.
    
        Args:
            filepath: Optional file path. If not provided, a file dialog will be shown.
    
        Returns:
            None
        """
        from vap_file_manager import VapFileManager
    
        if not filepath:
            filepath = filedialog.asksaveasfilename(
                title="Save As VAP3",
                defaultextension=".vap3",
                filetypes=[("VAP3 files", "*.vap3")]
            )
        
        if not filepath:
            return  # User canceled save operation
    
        try:
            self.gui.progress_dialog.show_progress_bar("Saving VAP3 file...")
            self.root.update_idletasks()
        
            vap_manager = VapFileManager()
        
            # Collect plot settings if needed
            plot_settings = {
                'selected_plot_type': self.gui.selected_plot_type.get() if hasattr(self.gui, 'selected_plot_type') else None
            }
        
            # Get image crop states
            image_crop_states = getattr(self.gui, 'image_crop_states', {})
            if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
                # Update with any latest crop states from the image loader
                image_crop_states.update(self.gui.image_loader.image_crop_states)
        
            success = vap_manager.save_to_vap3(
                filepath,
                self.gui.filtered_sheets,
                self.gui.sheet_images,
                self.gui.plot_options,
                image_crop_states,
                plot_settings
            )
        
            self.gui.progress_dialog.update_progress_bar(100)
            self.root.update_idletasks()
        
            if success:
                messagebox.showinfo("Success", f"Data saved successfully to {filepath}")
            else:
                messagebox.showerror("Error", "Failed to save data")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the file: {e}")
        finally:
            self.gui.progress_dialog.hide_progress_bar()

    def load_vap3_file(self, filepath=None) -> bool:
        """
        Load a .vap3 file and update the application state.
    
        Args:
            filepath: Optional file path. If not provided, a file dialog will be shown.
    
        Returns:
            bool: True if successful, False otherwise
        """
        from vap_file_manager import VapFileManager
    
        if not filepath:
            filepath = filedialog.askopenfilename(
                title="Open VAP3 File",
                filetypes=[("VAP3 files", "*.vap3")]
            )
        
        if not filepath:
            return False  # User canceled operation
    
        try:
            self.gui.progress_dialog.show_progress_bar("Loading VAP3 file...")
            self.root.update_idletasks()
        
            vap_manager = VapFileManager()
            result = vap_manager.load_from_vap3(filepath)
        
            # Update application state with loaded data
            self.gui.filtered_sheets = result['filtered_sheets']
            self.gui.sheets = {name: sheet_info['data'] for name, sheet_info in result['filtered_sheets'].items()}
            self.gui.sheet_images = result['sheet_images']
        
            if 'image_crop_states' in result and result['image_crop_states']:
                self.gui.image_crop_states = result['image_crop_states']
        
            if 'plot_options' in result and result['plot_options']:
                self.gui.plot_options = result['plot_options']
        
            if 'plot_settings' in result and result['plot_settings']:
                # Apply plot settings
                if 'selected_plot_type' in result['plot_settings'] and result['plot_settings']['selected_plot_type']:
                    self.gui.selected_plot_type.set(result['plot_settings']['selected_plot_type'])
        
            # Update UI state
            self.gui.current_file = os.path.basename(filepath)
            self.gui.file_path = filepath
        
            # Clear and rebuild all_filtered_sheets
            self.gui.all_filtered_sheets = [{
                "file_name": self.gui.current_file,
                "file_path": filepath,
                "filtered_sheets": self.gui.filtered_sheets
            }]
        
            # Update UI
            self.update_file_dropdown()
            self.update_ui_for_current_file()
        
            self.gui.progress_dialog.update_progress_bar(100)
            self.root.update_idletasks()
        
            messagebox.showinfo("Success", f"VAP3 file loaded successfully: {filepath}")
            return True
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load VAP3 file: {str(e)}")
            return False
        finally:
            self.gui.progress_dialog.hide_progress_bar()
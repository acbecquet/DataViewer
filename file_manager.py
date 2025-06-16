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
import openpyxl
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, Button, ttk, Frame
from test_selection_dialog import TestSelectionDialog
from test_start_menu import TestStartMenu
from header_data_dialog import HeaderDataDialog
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
from database_manager import DatabaseManager

class FileManager:
    """File Management Module for DataViewer.
    
    This class is initialized with a reference to the main TestingGUI instance
    so that it can update its state (sheets, selected_sheet, etc.).
    """
    def __init__(self, gui):
        self.gui = gui
        self.root = gui.root
        

        self.db_manager = DatabaseManager()
        
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

                self._store_file_in_database(file_path)
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
                self._store_file_in_database(file_path)

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

    def _store_file_in_database(self, original_file_path):
        """
        Store the loaded file in the database.

        Steps:
        1. Create a temporary VAP3 file
        2. Store the VAP3 file in the database with proper filename
        3. Store sheet metadata
        4. Store associated images
        """
        try:
            # Show progress dialog
            self.gui.progress_dialog.show_progress_bar("Storing file in database...")
            self.gui.root.update_idletasks()
    
            # Create a temporary VAP3 file
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_vap3_path = temp_file.name
    
            # Save current state as VAP3
            from vap_file_manager import VapFileManager
            vap_manager = VapFileManager()
    
            # Collect plot settings
            plot_settings = {
                'selected_plot_type': self.gui.selected_plot_type.get() if hasattr(self.gui, 'selected_plot_type') else None
            }
    
            # Get image crop states
            image_crop_states = getattr(self.gui, 'image_crop_states', {})
            if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
                image_crop_states.update(self.gui.image_loader.image_crop_states)
    
            # Extract just the filename (without path or extension) from the original file
            original_filename_base = os.path.splitext(os.path.basename(original_file_path))[0]
            display_filename = original_filename_base + '.vap3'
    
            # Save to temporary VAP3 file
            success = vap_manager.save_to_vap3(
                temp_vap3_path,
                self.gui.filtered_sheets,
                self.gui.sheet_images,
                self.gui.plot_options,
                image_crop_states,
                plot_settings
            )
    
            if not success:
                raise Exception("Failed to create temporary VAP3 file")
    
            # Print debug information
            print(f"Original path: {original_file_path}")
            print(f"Extracted display filename: {display_filename}")
    
            # Store meta_data about the original file
            meta_data = {
                'display_filename': display_filename,  # Include the display filename
                'original_filename': os.path.basename(original_file_path),
                'original_path': original_file_path,
                'creation_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                'sheet_count': len(self.gui.filtered_sheets),
                'plot_options': self.gui.plot_options,
                'plot_settings': plot_settings
            }
    
            # Store the VAP3 file in the database with the proper display filename
            file_id = self.db_manager.store_vap3_file(temp_vap3_path, meta_data)
    
            # Store sheet meta_data
            for sheet_name, sheet_info in self.gui.filtered_sheets.items():
                is_plotting = processing.plotting_sheet_test(sheet_name, sheet_info["data"])
                is_empty = sheet_info.get("is_empty", False)
        
                self.db_manager.store_sheet_info(
                    file_id, 
                    sheet_name, 
                    is_plotting, 
                    is_empty
                )
    
            # Store associated images
            if hasattr(self.gui, 'sheet_images') and self.gui.current_file in self.gui.sheet_images:
                for sheet_name, images in self.gui.sheet_images[self.gui.current_file].items():
                    for img_path in images:
                        if os.path.exists(img_path):
                            crop_enabled = image_crop_states.get(img_path, False)
                            self.db_manager.store_image(file_id, img_path, sheet_name, crop_enabled)
    
            # Clean up the temporary file
            try:
                os.unlink(temp_vap3_path)
            except:
                pass
    
            # Update progress
            self.gui.progress_dialog.update_progress_bar(100)
            self.gui.root.update_idletasks()
    
            print(f"File successfully stored in database with ID: {file_id} and name: {display_filename}")
    
        except Exception as e:
            print(f"Error storing file in database: {e}")
            traceback.print_exc()
        finally:
            # Hide progress dialog
            self.gui.progress_dialog.hide_progress_bar()
    
    def load_from_database(self, file_id=None):
        """
        Load a file from the database.
        
        Args:
            file_id: The ID of the file to load. If None, shows a dialog to select from available files.
        """
        try:
            # Show progress dialog
            self.gui.progress_dialog.show_progress_bar("Loading from database...")
            self.gui.root.update_idletasks()
            
            if file_id is None:
                # Show a dialog to select from available files
                file_list = self.db_manager.list_files()
                if not file_list:
                    messagebox.showinfo("Info", "No files found in the database.")
                    return
                
                # Create a simple dialog to choose a file
                dialog = Toplevel(self.gui.root)
                dialog.title("Select File from Database")
                dialog.geometry("600x400")
                dialog.transient(self.gui.root)
                dialog.grab_set()
                
                # Create a frame for the listbox and scrollbar
                frame = Frame(dialog)
                frame.pack(fill="both", expand=True, padx=10, pady=10)
                
                # Create a Listbox to display the files
                listbox = tk.Listbox(frame, font=FONT)
                listbox.pack(side="left", fill="both", expand=True)
                
                # Add a scrollbar
                scrollbar = tk.Scrollbar(frame, orient="vertical", command=listbox.yview)
                scrollbar.pack(side="right", fill="y")
                listbox.config(yscrollcommand=scrollbar.set)
                
                # Populate the listbox
                for file in file_list:
                    created_at = file["created_at"].strftime("%Y-%m-%d %H:%M:%S")
                    listbox.insert(tk.END, f"{file['id']}: {file['filename']} - {created_at}")
                
                # Add buttons
                button_frame = Frame(dialog)
                button_frame.pack(fill="x", padx=10, pady=10)
                
                selected_file_id = [None]  # Use a list to store the selected file ID
                
                def on_select():
                    selection = listbox.curselection()
                    if selection:
                        index = selection[0]
                        file_id = file_list[index]["id"]
                        selected_file_id[0] = file_id
                        dialog.destroy()
                
                def on_cancel():
                    dialog.destroy()
                
                select_button = Button(button_frame, text="Select", command=on_select)
                select_button.pack(side="right", padx=5)
                
                cancel_button = Button(button_frame, text="Cancel", command=on_cancel)
                cancel_button.pack(side="right", padx=5)
                
                # Wait for the dialog to close
                self.gui.root.wait_window(dialog)
                
                # Check if a file was selected
                file_id = selected_file_id[0]
                if file_id is None:
                    return
            
            # Load the file from the database
            file_data = self.db_manager.get_file_by_id(file_id)
            if not file_data:
                messagebox.showerror("Error", f"File with ID {file_id} not found in the database.")
                return
            
            # Save the VAP3 file to a temporary location
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_vap3_path = temp_file.name
                temp_file.write(file_data['file_content'])
            
            # Load the VAP3 file using the existing method
            from vap_file_manager import VapFileManager
            vap_manager = VapFileManager()
            
            # Load the VAP3 file
            self.load_vap3_file(temp_vap3_path)
            
            # Clean up the temporary file
            try:
                os.unlink(temp_vap3_path)
            except:
                pass
            
            # Update the UI to indicate the file was loaded from the database
            file_name = file_data['filename']
            self.gui.root.title(f"DataViewer - {file_name} (from Database)")
            
            # Update progress
            self.gui.progress_dialog.update_progress_bar(100)
            self.gui.root.update_idletasks()
            
            messagebox.showinfo("Success", f"File '{file_name}' loaded from database.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file from database: {e}")
            traceback.print_exc()
        finally:
            # Hide progress dialog
            self.gui.progress_dialog.hide_progress_bar()

    def show_database_browser(self):
        """
        Show a dialog to browse files stored in the database.
        """
        # Create a dialog to browse the database
        dialog = Toplevel(self.gui.root)
        dialog.title("Database Browser")
        dialog.geometry("800x600")
        dialog.transient(self.gui.root)
    
        # Create frames for UI elements
        top_frame = Frame(dialog)
        top_frame.pack(fill="x", padx=10, pady=10)
    
        list_frame = Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
        bottom_frame = Frame(dialog)
        bottom_frame.pack(fill="x", padx=10, pady=10)
    
        # Add a refresh button
        def refresh_list():
            # Clear the current list
            file_listbox.delete(0, tk.END)
        
            # Get the latest file list
            file_list = self.db_manager.list_files()
        
            # Populate the listbox
            for file in file_list:
                created_at = file["created_at"].strftime("%Y-%m-%d %H:%M:%S")
                file_listbox.insert(tk.END, f"{file['id']}: {file['filename']} - {created_at}")
        
            # Update the status label
            status_label.config(text=f"Total files: {len(file_list)}")
    
        refresh_button = Button(top_frame, text="Refresh", command=refresh_list)
        refresh_button.pack(side="left", padx=5)
    
        # Add a status label
        status_label = Label(top_frame, text="Total files: 0")
        status_label.pack(side="right", padx=5)
    
        # Create a listbox with scrollbar for files
        file_listbox = tk.Listbox(list_frame, font=FONT)
        file_listbox.pack(side="left", fill="both", expand=True)
    
        file_scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=file_listbox.yview)
        file_scrollbar.pack(side="right", fill="y")
        file_listbox.config(yscrollcommand=file_scrollbar.set)
    
        # Add buttons for actions
        def load_selected():
            selection = file_listbox.curselection()
            if not selection:
                messagebox.showinfo("Info", "Please select a file to load.")
                return
        
            index = selection[0]
            selected_text = file_listbox.get(index)
            file_id = int(selected_text.split(":")[0])
        
            # Close the dialog
            dialog.destroy()
        
            # Load the file from the database
            self.load_from_database(file_id)
    
        def delete_selected():
            selection = file_listbox.curselection()
            if not selection:
                messagebox.showinfo("Info", "Please select a file to delete.")
                return
        
            index = selection[0]
            selected_text = file_listbox.get(index)
            file_id = int(selected_text.split(":")[0])
            filename = selected_text.split(":")[1].split("-")[0].strip()
        
            # Confirm deletion
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {filename}?"):
                # Delete the file
                success = self.db_manager.delete_file(file_id)
                if success:
                    messagebox.showinfo("Success", f"File {filename} deleted successfully.")
                    # Refresh the list
                    refresh_list()
                else:
                    messagebox.showerror("Error", f"Failed to delete file {filename}.")
    
        load_button = Button(bottom_frame, text="Load Selected", command=load_selected)
        load_button.pack(side="left", padx=5)
    
        delete_button = Button(bottom_frame, text="Delete Selected", command=delete_selected)
        delete_button.pack(side="left", padx=5)
    
        close_button = Button(bottom_frame, text="Close", command=dialog.destroy)
        close_button.pack(side="right", padx=5)
    
        # Initial population of the list
        refresh_list()
    
        # Center the dialog
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() - width) // 2
        y = (dialog.winfo_screenheight() - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")
    
        # Make the dialog modal
        dialog.grab_set()
        self.gui.root.wait_window(dialog)

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

    def create_new_template_old(self, startup_menu: tk.Toplevel) -> None:
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

    def create_new_template(self, parent_window=None):
        """Create a new file from the template."""
        return self.create_new_file_with_tests(parent_window)


    def create_new_file_with_tests(self, parent_window=None):
        """
        Create a new file with only the selected tests.
    
        Args:
            parent_window (tk.Toplevel, optional): Parent window to close. Defaults to None.
    
        Returns:
            str: Path to the created file or None if creation was cancelled.
        """

        # first, get the save path before showing test selection dialog
        file_path = get_save_path(".xlsx")
        if not file_path:
            return None  # User cancelled

        # Get template path
        template_path = get_resource_path("resources/Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
    
        # Load template to get available tests
        wb = openpyxl.load_workbook(template_path)
        available_tests = [sheet for sheet in wb.sheetnames if sheet != "Sheet1"]
    
        # Show test selection dialog
        test_dialog = TestSelectionDialog(self.gui.root, available_tests)
        ok_clicked, selected_tests = test_dialog.show()
    
        if not ok_clicked or not selected_tests:
            return None  # User cancelled or no tests selected
    
        # Create a new file with selected tests
        try:
            # Copy template file
            shutil.copy(template_path, file_path)
        
            # Open the copied file
            new_wb = openpyxl.load_workbook(file_path)
        
            # Remove unselected tests
            for sheet_name in list(new_wb.sheetnames):
                if sheet_name not in selected_tests and sheet_name != "Sheet1":
                    del new_wb[sheet_name]
        
            # Save the modified workbook
            new_wb.save(file_path)
        
            # Close parent window if provided
            if parent_window and parent_window.winfo_exists():
                parent_window.destroy()
        
            # Show test start menu
            self.show_test_start_menu(file_path)
        
            return file_path
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the file: {str(e)}")
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)  # Clean up partial file
                except:
                    pass
            return None

    def show_test_start_menu(self, file_path):
        """
        Show the test start menu for a file.
    
        Args:
            file_path (str): Path to the file.
        """
        print(f"DEBUG: Showing test start menu for file: {file_path}")
        
        # Show the test start menu
        start_menu = TestStartMenu(self.gui.root, file_path)
        result, selected_test = start_menu.show()
    
        if result == "start_test" and selected_test:
            print(f"DEBUG: Starting test: {selected_test}")
            # Show header data dialog
            self.show_header_data_dialog(file_path, selected_test)
        elif result == "view_raw_file":
            print("DEBUG: Viewing raw file requested")
            # Load the file for viewing if not already loaded
            if not hasattr(self.gui, 'file_path') or self.gui.file_path != file_path:
                self.load_file(file_path)
        else:
            print("DEBUG: Test start menu was cancelled or closed")

    def show_header_data_dialog(self, file_path, selected_test):
        """
        Show the header data dialog for a selected test.
    
        Args:
            file_path (str): Path to the file.
            selected_test (str): Name of the selected test.
        """
        print(f"DEBUG: Showing header data dialog for {selected_test}")
        
        # Show the header data dialog
        header_dialog = HeaderDataDialog(self.gui.root, file_path, selected_test)
        result, header_data = header_dialog.show()
    
        if result:
            print("DEBUG: Header data dialog completed successfully")
            # Apply header data to the file
            self.apply_header_data_to_file(file_path, header_data)
        
            print("DEBUG: Proceeding to data collection window")
            # Show the data collection window
            from data_collection_window import DataCollectionWindow
            data_collection = DataCollectionWindow(self.gui.root, file_path, selected_test, header_data)
            data_result = data_collection.show()
        
            if data_result == "load_file":
                # Load the file for viewing
                self.load_file(file_path)
            elif data_result == "cancel":
                # User cancelled data collection, just load the file
                self.load_file(file_path)
        else:
            print("DEBUG: Header data dialog was cancelled")

    def apply_header_data_to_file(self, file_path, header_data):
        """
        Apply the header data to the Excel file.
    
        Args:
            file_path (str): Path to the file.
            header_data (dict): Dictionary containing header data.
        """
        try:
            # Load the workbook
            wb = openpyxl.load_workbook(file_path)
        
            # Get the sheet for the selected test
            if header_data["test"] in wb.sheetnames:
                ws = wb[header_data["test"]]
            
                # Common data mappings (row, column)
                common_mappings = {
                    "media": (2, 2),      # Cell B2
                    "viscosity": (3, 2),  # Cell B3
                    "voltage": (2, 6),    # Cell F2
                    "oil_mass": (2, 8)    # Cell H2
                }
            
                # Apply common data
                for field, (row, col) in common_mappings.items():
                    if header_data["common"][field]:
                        ws.cell(row=row, column=col, value=header_data["common"][field])
            
                # Apply sample-specific data
                num_samples = header_data["num_samples"]
                for i in range(num_samples):
                    # Calculate column offset (12 columns per sample)
                    col_offset = i * 12
                
                    # Sample ID (row 1, column F + offset)
                    sample_id_col = 6 + col_offset  # Column F is 6
                    ws.cell(row=1, column=sample_id_col, value=header_data["samples"][i]["id"])
                
                    # Resistance (row 3, column D + offset)
                    resistance_col = 4 + col_offset  # Column D is 4
                    ws.cell(row=3, column=resistance_col, value=header_data["samples"][i]["resistance"])
            
                # Add tester name to sheet A1
                ws.cell(row=1, column=1, value=f"Tester: {header_data['common']['tester']}")
            
                # Save the workbook
                wb.save(file_path)
            
                print(f"Successfully applied header data to {file_path}")
            else:
                messagebox.showerror("Error", f"Sheet '{header_data['test']}' not found in the file.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while applying header data: {str(e)}")



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



    def load_file(self, file_path):
            """Load a file without showing dialogs - used for data collection flow."""
            print(f"DEBUG: Loading file for data collection: {file_path}")
            try:
                # Load the file using existing logic
                self.load_excel_file(file_path)
            
                # Update the GUI state
                self.gui.all_filtered_sheets.append({
                    "file_name": os.path.basename(file_path),
                    "file_path": file_path,
                    "filtered_sheets": copy.deepcopy(self.gui.filtered_sheets)
                })
            
                self.update_file_dropdown()
                self.set_active_file(os.path.basename(file_path))
                self.update_ui_for_current_file()
            
                print(f"DEBUG: File loaded successfully: {file_path}")
                return True
            
            except Exception as e:
                print(f"DEBUG: Error loading file: {e}")
                messagebox.showerror("Error", f"Failed to load file: {e}")
                return False
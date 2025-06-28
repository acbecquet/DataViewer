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
from resource_utils import get_resource_path
from utils import (
    is_valid_excel_file,
    load_excel_file,
    get_save_path,
    is_standard_file,
    FONT,
    APP_BACKGROUND_COLOR,
    load_excel_file_with_formulas,
    debug_print
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
        
        # Add cache to prevent redundant operations
        self.loaded_files_cache = {}  # Cache for loaded file data
        self.stored_files_cache = set()  # Track files already stored in database
        
    def load_excel_file(self, file_path, legacy_mode: str = None, skip_database_storage: bool = False, force_reload: bool = False) -> None:
        """
        Load the selected Excel file and process its sheets.
        Enhanced with caching and optional database storage skip.

        Args:
            file_path (str): Path to the Excel file
            legacy_mode (str, optional): Legacy processing mode
            skip_database_storage (bool): If True, skip storing in database
            force_reload (bool): If True, bypass cache and reload from file
        """
        debug_print(f"DEBUG: load_excel_file called for {file_path}, skip_db_storage={skip_database_storage}, force_reload={force_reload}")

        # Check cache first (unless force_reload is True)
        cache_key = f"{file_path}_{legacy_mode}"
        if not force_reload and cache_key in self.loaded_files_cache:
            debug_print("DEBUG: Using cached file data instead of reprocessing")
            cached_data = self.loaded_files_cache[cache_key]
            self.gui.filtered_sheets = cached_data['filtered_sheets']
            self.gui.sheets = cached_data.get('sheets', {})
            # Fix: Make sure full_sample_data exists
            if hasattr(self.gui, 'full_sample_data'):
                self.gui.full_sample_data = cached_data.get('full_sample_data', pd.DataFrame())
    
            # Set the selected sheet without triggering full UI update
            if cached_data['filtered_sheets']:
                first_sheet = list(cached_data['filtered_sheets'].keys())[0]
                self.gui.selected_sheet.set(first_sheet)
            return

        # Clear cache entry if force_reload is True
        if force_reload and cache_key in self.loaded_files_cache:
            debug_print(f"DEBUG: Force reload requested - clearing cache entry for {file_path}")
            del self.loaded_files_cache[cache_key]

        try:
            # Ensure the file is a valid Excel file.
            if not is_valid_excel_file(os.path.basename(file_path)):
                raise ValueError(f"Invalid Excel file selected: {file_path}")

            debug_print(f"DEBUG: {'Force reloading' if force_reload else 'Loading'} file from disk: {file_path}")
            debug_print(f"DEBUG: Checking if file is standard format: {file_path}")
        
            if not is_standard_file(file_path):
                debug_print("DEBUG: File is legacy format, processing accordingly")
                # Legacy file processing
                legacy_dir = os.path.join(os.path.abspath("."), "legacy data")
                if not os.path.exists(legacy_dir):
                    os.makedirs(legacy_dir)
                    debug_print(f"DEBUG: Created legacy data directory: {legacy_dir}")
            
                legacy_wb = load_workbook(file_path)
                legacy_sheetnames = legacy_wb.sheetnames
                debug_print(f"DEBUG: Legacy file sheets: {legacy_sheetnames}")

                template_path_default = os.path.join(os.path.abspath("."), "resources", 
                                         "Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
                                     
                if not os.path.exists(template_path_default):
                    raise FileNotFoundError(f"Template file not found: {template_path_default}")
                
                wb_template = load_workbook(template_path_default)
                template_sheet_names = wb_template.sheetnames
                debug_print(f"DEBUG: Template sheets: {template_sheet_names}")

                if legacy_mode is None:
                    if len(legacy_sheetnames) == 1 and legacy_sheetnames[0] not in template_sheet_names:
                        legacy_mode = "file"
                        debug_print("DEBUG: Auto-detected legacy mode: file")
                    else:
                        legacy_mode = "standards"
                        debug_print("DEBUG: Auto-detected legacy mode: standards")

                final_sheets = {}
                full_sample_data = pd.DataFrame()

                if legacy_mode == "file":
                    debug_print("DEBUG: Processing as legacy file")
                    converted = processing.convert_legacy_file_using_template(file_path)
                    key = f"Legacy_{os.path.basename(file_path)}"
                    final_sheets = {key: {"data": converted, "is_empty": converted.empty}}
                    full_sample_data = converted

                    self.gui.filtered_sheets = final_sheets
                    self.gui.full_sample_data = full_sample_data
                    default_key = list(final_sheets.keys())[0]
                    self.gui.selected_sheet.set(default_key)
                    debug_print(f"DEBUG: Legacy file processed, key: {default_key}")

                elif legacy_mode == "standards":
                    debug_print("DEBUG: Processing as legacy standards file")
                    converted_dict = processing.convert_legacy_standards_using_template(file_path)
                
                    self.gui.sheets = converted_dict
                    self.gui.filtered_sheets = {
                        name: {"data": data, "is_empty": data.empty}
                        for name, data in self.gui.sheets.items()
                    }
                    # Fix: Safely concatenate sheets
                    sheet_data_list = []
                    for sheet_info in self.gui.filtered_sheets.values():
                        if not sheet_info["data"].empty:
                            sheet_data_list.append(sheet_info["data"])
                
                    if sheet_data_list:
                        self.gui.full_sample_data = pd.concat(sheet_data_list, axis=1)
                    else:
                        self.gui.full_sample_data = pd.DataFrame()
                    
                    first_sheet = list(self.gui.filtered_sheets.keys())[0]
                    self.gui.selected_sheet.set(first_sheet)
                    debug_print(f"DEBUG: Legacy standards processed, first sheet: {first_sheet}")
                else:
                    raise ValueError(f"Unknown legacy mode: {legacy_mode}")

                # Store in database only if not skipping and not already stored
                if not skip_database_storage and file_path not in self.stored_files_cache:
                    debug_print("DEBUG: Storing legacy file in database")
                    self._store_file_in_database(file_path)
                    self.stored_files_cache.add(file_path)
            else:
                # Standard file processing
                debug_print("DEBUG: Processing as standard file")
                self.gui.sheets = load_excel_file(file_path)
                self.gui.filtered_sheets = {
                    name: {"data": data, "is_empty": data.empty}
                    for name, data in self.gui.sheets.items()
                }
            
                # Fix: Safely concatenate sheets and handle empty sheets
                sheet_data_list = []
                for sheet_info in self.gui.filtered_sheets.values():
                    if not sheet_info["data"].empty:
                        sheet_data_list.append(sheet_info["data"])
            
                if sheet_data_list:
                    self.gui.full_sample_data = pd.concat(sheet_data_list, axis=1)
                else:
                    self.gui.full_sample_data = pd.DataFrame()
                
                first_sheet = list(self.gui.filtered_sheets.keys())[0]
                self.gui.selected_sheet.set(first_sheet)
                debug_print(f"DEBUG: Standard file processed, first sheet: {first_sheet}")
            
                # Store in database only if not skipping and not already stored
                if not skip_database_storage and file_path not in self.stored_files_cache:
                    debug_print("DEBUG: Storing standard file in database")
                    self._store_file_in_database(file_path)
                    self.stored_files_cache.add(file_path)

            # Cache the processed data
            cache_data = {
                'filtered_sheets': copy.deepcopy(self.gui.filtered_sheets),
                'sheets': copy.deepcopy(getattr(self.gui, 'sheets', {})),
                'full_sample_data': self.gui.full_sample_data.copy() if hasattr(self.gui, 'full_sample_data') and not self.gui.full_sample_data.empty else pd.DataFrame()
            }
            self.loaded_files_cache[cache_key] = cache_data
            debug_print(f"DEBUG: Cached processed data for {file_path}")

        except Exception as e:
            error_msg = f"Error occurred while loading file: {e}"
            print(f"ERROR: {error_msg}")
            traceback.print_exc()
            messagebox.showerror("Error", error_msg)

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
            self.root.update_idletasks()

    def _store_file_in_database(self, original_file_path):
        """
        Store the loaded file in the database.
        Enhanced with duplicate checking and better error handling.
        """
        debug_print(f"DEBUG: Checking if file {original_file_path} needs database storage")
        
        # Check if already stored
        if original_file_path in self.stored_files_cache:
            debug_print("DEBUG: File already stored in database, skipping")
            return
            
        try:
            debug_print("DEBUG: Storing new file in database...")
            # Show progress dialog
            self.gui.progress_dialog.show_progress_bar("Storing file in database...")
            self.gui.root.update_idletasks()
    
            # Create a temporary VAP3 file
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_vap3_path = temp_file.name
                debug_print(f"DEBUG: Created temporary VAP3 file: {temp_vap3_path}")
    
            # Save current state as VAP3
            from vap_file_manager import VapFileManager
            vap_manager = VapFileManager()
    
            # Collect plot settings safely
            plot_settings = {}
            if hasattr(self.gui, 'selected_plot_type'):
                plot_settings['selected_plot_type'] = self.gui.selected_plot_type.get()
            
            debug_print(f"DEBUG: Plot settings: {plot_settings}")
    
            # Get image crop states safely
            image_crop_states = getattr(self.gui, 'image_crop_states', {})
            if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
                image_crop_states.update(getattr(self.gui.image_loader, 'image_crop_states', {}))
            
            debug_print(f"DEBUG: Image crop states: {len(image_crop_states)} items")
    
            # Extract just the filename (without path or extension) from the original file
            original_filename_base = os.path.splitext(os.path.basename(original_file_path))[0]
            display_filename = original_filename_base + '.vap3'
            
            debug_print(f"DEBUG: Display filename: {display_filename}")
    
            # Save to temporary VAP3 file
            success = vap_manager.save_to_vap3(
                temp_vap3_path,
                self.gui.filtered_sheets,
                getattr(self.gui, 'sheet_images', {}),
                getattr(self.gui, 'plot_options', []),
                image_crop_states,
                plot_settings
            )
    
            if not success:
                raise Exception("Failed to create temporary VAP3 file")
            
            debug_print("DEBUG: VAP3 file created successfully")
    
            # Store metadata about the original file
            meta_data = {
                'display_filename': display_filename,
                'original_filename': os.path.basename(original_file_path),
                'original_path': original_file_path,
                'creation_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                'sheet_count': len(self.gui.filtered_sheets),
                'plot_options': getattr(self.gui, 'plot_options', []),
                'plot_settings': plot_settings
            }
            
            debug_print(f"DEBUG: Metadata: {meta_data}")
    
            # Store the VAP3 file in the database with the proper display filename
            file_id = self.db_manager.store_vap3_file(temp_vap3_path, meta_data)
            debug_print(f"DEBUG: File stored with ID: {file_id}")
    
            # Store sheet metadata
            for sheet_name, sheet_info in self.gui.filtered_sheets.items():
                is_plotting = processing.plotting_sheet_test(sheet_name, sheet_info["data"])
                is_empty = sheet_info.get("is_empty", False)
        
                sheet_id = self.db_manager.store_sheet_info(
                    file_id, 
                    sheet_name, 
                    is_plotting, 
                    is_empty
                )
                debug_print(f"DEBUG: Stored sheet '{sheet_name}' with ID: {sheet_id}")
    
            # Store associated images
            if hasattr(self.gui, 'sheet_images') and hasattr(self.gui, 'current_file') and self.gui.current_file in self.gui.sheet_images:
                for sheet_name, images in self.gui.sheet_images[self.gui.current_file].items():
                    for img_path in images:
                        if os.path.exists(img_path):
                            crop_enabled = image_crop_states.get(img_path, False)
                            img_id = self.db_manager.store_image(file_id, img_path, sheet_name, crop_enabled)
                            debug_print(f"DEBUG: Stored image '{img_path}' with ID: {img_id}")
    
            # Clean up the temporary file
            try:
                os.unlink(temp_vap3_path)
                debug_print(f"DEBUG: Cleaned up temporary file: {temp_vap3_path}")
            except Exception as cleanup_error:
                print(f"WARNING: Failed to clean up temporary file: {cleanup_error}")
    
            # Update progress
            self.gui.progress_dialog.update_progress_bar(100)
            self.gui.root.update_idletasks()
    
            debug_print(f"SUCCESS: File successfully stored in database with ID: {file_id} and name: {display_filename}")
            
            # Mark as stored in cache
            self.stored_files_cache.add(original_file_path)
    
        except Exception as e:
            error_msg = f"Error storing file in database: {e}"
            print(f"ERROR: {error_msg}")
            traceback.print_exc()
        finally:
            # Hide progress dialog
            self.gui.progress_dialog.hide_progress_bar()
    
    def load_from_database(self, file_id=None):
        """Load a file from the database."""
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
        """Show a dialog to browse files stored in the database."""
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
        """Opens an Excel file for a specific sheet, allowing edits, and then updates the application when Excel closes."""
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
        
            # Create a temporary Excel file
            import tempfile
            import uuid
            from openpyxl import Workbook
        
            # Generate a unique identifier
            unique_id = str(uuid.uuid4()).split('-')[0]
        
            # Ensure the sheet name only contains valid characters
            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c == ' ')
            safe_sheet_name = safe_sheet_name.replace(' ', '_')[:15]
        
            # Create a temporary file
            temp_dir = os.path.abspath(tempfile.gettempdir())
            temp_file = os.path.join(temp_dir, f"dataviewer_{safe_sheet_name}_{unique_id}.xlsx")
        
            # Create a new workbook with just this sheet
            wb = Workbook()
            ws = wb.active
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
        
            if not os.path.exists(temp_file):
                raise FileNotFoundError(f"Failed to create temporary file at {temp_file}")
        
            # Record the file modification time before opening
            original_mod_time = os.path.getmtime(temp_file)
        
            # Create status notification
            status_text = f"Opening {sheet_name} in Excel. Changes will be imported when Excel closes."
            status_label = ttk.Label(self.gui.root, text=status_text, relief="sunken", anchor="w")
            status_label.pack(side="bottom", fill="x")
            self.gui.root.update_idletasks()
        
            # File monitor state
            class FileMonitorState:
                def __init__(self, initial_mod_time):
                    self.last_mod_time = initial_mod_time
                    self.has_changed = False
        
            monitor_state = FileMonitorState(original_mod_time)
        
            # Force a new Excel instance
            import subprocess
            cmd = f'start /wait "" "excel.exe" /x "{os.path.abspath(temp_file)}"'
        
            try:
                subprocess.Popen(cmd, shell=True)
            except Exception as e:
                print(f"Error launching Excel with command: {e}")
                os.startfile(os.path.abspath(temp_file))
            
            time.sleep(2.0)
        
            # Monitor file in background thread
            def monitor_file_lock():
                file_open = True
                while file_open:
                    file_locked = False
                    try:
                        with open(temp_file, 'r+b') as test_lock:
                            pass
                    except PermissionError:
                        file_locked = True
                    except FileNotFoundError:
                        file_locked = False
                        file_open = False
                    except Exception:
                        file_locked = False
                
                    file_open = file_locked
                
                    if os.path.exists(temp_file):
                        try:
                            current_mod_time = os.path.getmtime(temp_file)
                            if current_mod_time > monitor_state.last_mod_time:
                                monitor_state.has_changed = True
                                monitor_state.last_mod_time = current_mod_time
                        except Exception:
                            pass
                            
                    if not file_open:
                        self.gui.root.after(500, lambda: self._process_excel_changes(temp_file, sheet_name, monitor_state.has_changed, status_label))
                        break
                
                    time.sleep(1.0)
        
            # Start monitoring thread
            monitor_thread = threading.Thread(target=monitor_file_lock, daemon=True)
            self.gui.threads.append(monitor_thread)
            monitor_thread.start()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {e}")
            traceback.print_exc()

    def _process_excel_changes(self, temp_file, sheet_name, file_changed, status_label=None):
        """Process changes made in Excel after it closes."""
        try:
            if status_label and status_label.winfo_exists():
                status_label.destroy()
        
            if file_changed and os.path.exists(temp_file):
                try:
                    max_retries = 3
                    retry_count = 0
                    read_success = False
                    modified_data = None
                
                    while not read_success and retry_count < max_retries:
                        try:
                            modified_data = pd.read_excel(temp_file)
                            read_success = True
                        except Exception:
                            retry_count += 1
                            time.sleep(1.0)
                
                    if not read_success:
                        raise Exception(f"Failed to read Excel file after {max_retries} attempts")
                
                    # Update the filtered sheets with new data
                    self.gui.filtered_sheets[sheet_name]['data'] = modified_data
                
                    # If using VAP3 file, update it
                    if hasattr(self.gui, 'file_path') and self.gui.file_path.endswith('.vap3'):
                        from vap_file_manager import VapFileManager
                        vap_manager = VapFileManager()
                        vap_manager.save_to_vap3(
                            self.gui.file_path,
                            self.gui.filtered_sheets,
                            self.gui.sheet_images,
                            self.gui.plot_options,
                            getattr(self.gui, 'image_crop_states', {})
                        )
                
                    # Refresh the GUI display
                    self.gui.update_displayed_sheet(sheet_name)
                
                    # Show success message
                    success_label = ttk.Label(
                        self.gui.root, 
                        text="Changes from Excel have been imported successfully.",
                        relief="sunken", 
                        anchor="w"
                    )
                    success_label.pack(side="bottom", fill="x")
                    self.gui.root.after(3000, lambda: success_label.destroy() if success_label.winfo_exists() else None)
                
                except Exception as e:
                    traceback.print_exc()
                    messagebox.showerror("Error", f"Failed to read modified Excel file: {e}")
            elif file_changed and not os.path.exists(temp_file):
                messagebox.showinfo("Information", "Excel file was modified but appears to have been moved or renamed. Changes could not be imported.")
            else:
                info_label = ttk.Label(
                    self.gui.root, 
                    text="Excel file was closed without changes.",
                    relief="sunken", 
                    anchor="w"
                )
                info_label.pack(side="bottom", fill="x")
                self.gui.root.after(3000, lambda: info_label.destroy() if info_label.winfo_exists() else None)
            
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to process Excel changes: {e}")
        finally:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception:
                    pass

    def create_new_template(self, parent_window=None):
        """Create a new file from the template."""
        return self.create_new_file_with_tests(parent_window)

    def create_new_file_with_tests(self, parent_window=None):
        """Create a new file with only the selected tests."""
        # Get save path before showing test selection dialog
        file_path = get_save_path(".xlsx")
        if not file_path:
            return None

        # Get template path
        template_path = get_resource_path("resources/Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
    
        # Load template to get available tests
        wb = openpyxl.load_workbook(template_path)
        available_tests = [sheet for sheet in wb.sheetnames if sheet != "Sheet1"]
    
        # Show test selection dialog
        test_dialog = TestSelectionDialog(self.gui.root, available_tests)
        ok_clicked, selected_tests = test_dialog.show()
    
        if not ok_clicked or not selected_tests:
            return None
    
        # Create a new file with selected tests
        try:
            shutil.copy(template_path, file_path)
            new_wb = openpyxl.load_workbook(file_path)
        
            # Remove unselected tests
            for sheet_name in list(new_wb.sheetnames):
                if sheet_name not in selected_tests and sheet_name != "Sheet1":
                    del new_wb[sheet_name]
        
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
                    os.remove(file_path)
                except:
                    pass
            return None

    def show_test_start_menu(self, file_path):
        """
        Show the test start menu for a file.
        If file is already loaded, try to extract existing header data and skip the dialog.
        """
        debug_print(f"DEBUG: Showing test start menu for file: {file_path}")
        
        # Show the test start menu
        start_menu = TestStartMenu(self.gui.root, file_path)
        result, selected_test = start_menu.show()
    
        if result == "start_test" and selected_test:
            debug_print(f"DEBUG: Starting test: {selected_test}")
            
            # Check if this file is already loaded and has header data
            if hasattr(self.gui, 'file_path') and self.gui.file_path == file_path:
                debug_print("DEBUG: File is already loaded - attempting to extract existing header data")
                
                # Try to extract existing header data from the loaded file
                existing_header_data = self.extract_existing_header_data(file_path, selected_test)
                
                if existing_header_data and self.validate_header_data(existing_header_data):
                    debug_print("DEBUG: Found complete header data - skipping header dialog")
                    # Go directly to data collection with existing header data
                    self.start_data_collection_with_header_data(file_path, selected_test, existing_header_data)
                else:
                    debug_print("DEBUG: Header data incomplete or missing - showing header dialog")
                    # Show header data dialog for missing/incomplete data
                    self.show_header_data_dialog(file_path, selected_test)
            else:
                debug_print("DEBUG: File not loaded yet - showing header dialog")
                # Show header data dialog for new files
                self.show_header_data_dialog(file_path, selected_test)
                
        elif result == "view_raw_file":
            debug_print("DEBUG: Viewing raw file requested")
            # Load the file for viewing if not already loaded
            if not hasattr(self.gui, 'file_path') or self.gui.file_path != file_path:
                self.load_file(file_path)
        else:
            debug_print("DEBUG: Test start menu was cancelled or closed")

    def extract_existing_header_data(self, file_path, selected_test):
        """Extract existing header data from the Excel file with corrected positioning."""
        debug_print(f"DEBUG: Extracting header data from {file_path} for test {selected_test}")
        
        try:
            wb = openpyxl.load_workbook(file_path)
            
            if selected_test not in wb.sheetnames:
                debug_print(f"DEBUG: Sheet {selected_test} not found in file. Available sheets: {wb.sheetnames}")
                return None
                
            ws = wb[selected_test]
            debug_print(f"DEBUG: Successfully opened sheet '{selected_test}'")
            
            # Extract tester name from corrected position: row 3, column 4
            tester_cell = ws.cell(row=3, column=4)
            tester = ""
            if tester_cell.value:
                tester_value = str(tester_cell.value)
                # Remove any label prefix if present
                if "tester:" in tester_value.lower():
                    tester = tester_value.split(":", 1)[1].strip() if ":" in tester_value else ""
                else:
                    tester = tester_value.strip()
            
            debug_print(f"DEBUG: Extracted tester: '{tester}' from cell D3: '{tester_cell.value}'")
            
            # Extract common data from corrected positions
            common_data = {
                'tester': tester,
                'media': str(ws.cell(row=2, column=2).value or ""),          # Row 2, Col B
                'viscosity': str(ws.cell(row=3, column=2).value or ""),      # Row 3, Col B
                'voltage': str(ws.cell(row=3, column=6).value or ""),        # Row 3, Col F (corrected)
                'oil_mass': str(ws.cell(row=3, column=8).value or "")        # Row 3, Col H (corrected)
            }
            
            debug_print(f"DEBUG: Extracted common data with corrected positions: {common_data}")
            
            # Extract sample data (check up to 10 samples with 12-column blocks, or 8-column if user test simulation)
            # Determine columns per sample based on test type

            if selected_test in ["User Test Simulation", "User Simulation Test"]:
                columns_per_sample = 8
                debug_print(f"DEBUG: Using 8-column layout for {selected_test}")
            else:
                columns_per_sample = 12
                debug_print(f"DEBUG: Using 12-column layout for {selected_test}")

            # Extract sample data (check up to 10 samples)
            samples = []
            sample_count = 0

            for i in range(10):
                col_offset = i * columns_per_sample
                sample_id_col = 6 + col_offset  # Column F + offset (relative to sample block)
                resistance_col = 4 + col_offset  # Column D + offset (relative to sample block)
    
                sample_id_cell = ws.cell(row=1, column=sample_id_col)
                resistance_cell = ws.cell(row=2, column=resistance_col)
    
                sample_id = sample_id_cell.value
                resistance = resistance_cell.value
    
                debug_print(f"DEBUG: Sample {i+1} check - ID cell (row 1, col {sample_id_col}): '{sample_id}', Resistance cell (row 2, col {resistance_col}): '{resistance}'")
    
                # If we find a sample ID, this is a valid sample
                if sample_id and str(sample_id).strip():
                    samples.append({
                        'id': str(sample_id).strip(),
                        'resistance': str(resistance).strip() if resistance else ""
                    })
                    sample_count += 1
                    debug_print(f"DEBUG: Found valid sample {sample_count}: ID='{sample_id}', Resistance='{resistance}'")
                else:
                    # If no sample ID, we've reached the end of samples
                    debug_print(f"DEBUG: No more samples found after checking {i+1} positions")
                    break
            
            if sample_count == 0:
                debug_print("DEBUG: No samples found, using default single sample")
                samples = [{'id': 'Sample 1', 'resistance': ''}]
                sample_count = 1
            
            header_data = {
                'common': common_data,
                'samples': samples,
                'test': selected_test,
                'num_samples': sample_count
            }
            
            debug_print(f"DEBUG: Final extracted header data with corrected positions: {sample_count} samples")
            debug_print(f"DEBUG: Samples: {samples}")
            debug_print(f"DEBUG: Common: {common_data}")
            
            return header_data
            
        except Exception as e:
            print(f"ERROR: Exception extracting header data: {e}")
            traceback.print_exc()
            return None

    def validate_header_data(self, header_data):
        """Validate that header data is complete enough for data collection."""
        debug_print("DEBUG: Validating extracted header data")
        
        if not header_data:
            debug_print("DEBUG: Header data is None")
            return False
            
        # Check for required common data
        common_data = header_data.get('common', {})
        tester = common_data.get('tester', '').strip()
        if not tester:
            debug_print("DEBUG: Tester name is missing or empty")
            return False
        
        debug_print(f"DEBUG: Found tester: '{tester}'")
            
        # Check for samples
        samples = header_data.get('samples', [])
        if not samples:
            debug_print("DEBUG: No samples found")
            return False
            
        # Check that at least one sample has an ID
        has_valid_sample = False
        for sample in samples:
            if sample.get('id', '').strip():
                has_valid_sample = True
                break
                
        if not has_valid_sample:
            debug_print("DEBUG: No samples with valid IDs found")
            return False
            
        debug_print("DEBUG: Header data validation passed")
        return True

    def start_data_collection_with_header_data(self, file_path, selected_test, header_data):
        """Start data collection directly with existing header data."""
        debug_print("DEBUG: Starting data collection with existing header data")
        
        try:
            # Show the data collection window directly
            from data_collection_window import DataCollectionWindow
            data_collection = DataCollectionWindow(self.gui, file_path, selected_test, header_data)
            data_result = data_collection.show()
        
            if data_result == "load_file":
                debug_print("DEBUG: Data collection completed - file should already be loaded, updating UI only")
                self.ensure_file_is_loaded_in_ui(file_path)
            elif data_result == "cancel":
                debug_print("DEBUG: Data collection was cancelled - file should already be loaded, updating UI only")
                self.ensure_file_is_loaded_in_ui(file_path)
                
        except Exception as e:
            debug_print(f"DEBUG: Error starting data collection: {e}")
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to start data collection: {e}")

    def ensure_file_is_loaded_in_ui(self, file_path):
        """Ensure the file is properly loaded in the UI without redundant processing."""
        debug_print(f"DEBUG: Ensuring file {file_path} is loaded in UI")
        
        # Validate file path
        if not file_path or not os.path.exists(file_path):
            print(f"ERROR: Invalid file path: {file_path}")
            return False
        
        # Check if file is already in the UI state
        file_name = os.path.basename(file_path)
        
        # Look for existing entry in all_filtered_sheets
        existing_entry = None
        for file_data in self.gui.all_filtered_sheets:
            if file_data["file_path"] == file_path:
                existing_entry = file_data
                break
        
        if existing_entry:
            debug_print("DEBUG: File already in UI state, just updating active file")
            try:
                self.set_active_file(existing_entry["file_name"])
                self.update_ui_for_current_file()
                return True
            except Exception as e:
                print(f"ERROR: Failed to set active file: {e}")
                return False
        else:
            debug_print("DEBUG: File not in UI state, adding it")
            try:
                # Load file without database storage (skip_database_storage=True)
                self.load_excel_file(file_path, skip_database_storage=True)
                
                # Add to all_filtered_sheets
                self.gui.all_filtered_sheets.append({
                    "file_name": file_name,
                    "file_path": file_path,
                    "filtered_sheets": copy.deepcopy(self.gui.filtered_sheets)
                })
                
                self.update_file_dropdown()
                self.set_active_file(file_name)
                self.update_ui_for_current_file()
                debug_print(f"DEBUG: Successfully added file {file_name} to UI")
                return True
                
            except Exception as e:
                print(f"ERROR: Failed to load file into UI: {e}")
                traceback.print_exc()
                return False

    def show_header_data_dialog(self, file_path, selected_test):
        """Show the header data dialog for a selected test."""
        debug_print(f"DEBUG: Showing header data dialog for {selected_test}")
        
        # Show the header data dialog
        header_dialog = HeaderDataDialog(self.gui.root, file_path, selected_test)
        result, header_data = header_dialog.show()
    
        if result:
            debug_print("DEBUG: Header data dialog completed successfully")
            # Apply header data to the file
            self.apply_header_data_to_file(file_path, header_data)
        
            debug_print("DEBUG: Proceeding to data collection window")
            # Show the data collection window
            from data_collection_window import DataCollectionWindow
            data_collection = DataCollectionWindow(self.gui, file_path, selected_test, header_data)
            data_result = data_collection.show()
        
            if data_result == "load_file":
                # Load the file for viewing
                self.load_file(file_path)
            elif data_result == "cancel":
                # User cancelled data collection, just load the file
                self.load_file(file_path)
        else:
            debug_print("DEBUG: Header data dialog was cancelled")

    def apply_header_data_to_file(self, file_path, header_data):
        """
        Apply the header data to the Excel file.
        Enhanced to correctly apply headers to all sample blocks with proper column mapping.
        """
        try:
            debug_print(f"DEBUG: Applying header data to {file_path} for {header_data['num_samples']} samples")
        
            # Load the workbook
            wb = openpyxl.load_workbook(file_path)

            # Get the sheet for the selected test
            if header_data["test"] in wb.sheetnames:
                ws = wb[header_data["test"]]
            
                debug_print(f"DEBUG: Successfully opened sheet '{header_data['test']}'")

                # Set the test name at row 1, column 1 (this should be done once)
                ws.cell(row=1, column=1, value=header_data["test"])
                debug_print(f"DEBUG: Set test name '{header_data['test']}' at row 1, column 1")

                # Get common data once
                common_data = header_data["common"]

                # Apply sample-specific data for each sample block
                num_samples = header_data["num_samples"]
                for i in range(num_samples):
                    # Calculate column offset (12 columns per sample)
                    # Sample blocks start at column 1, so offsets are 0, 12, 24, 36, etc.
                    col_offset = i * 12
                
                    debug_print(f"DEBUG: Processing sample {i+1} with column offset {col_offset}")
            
                    # Get sample data
                    sample_data = header_data["samples"][i]
                    sample_id = sample_data["id"]
                    sample_resistance = sample_data["resistance"]
                
                    debug_print(f"DEBUG: Sample {i+1}: ID='{sample_id}', Resistance='{sample_resistance}'")
                
                    # Apply sample-specific headers according to template structure:
                    # Row 1: Sample ID goes in column F (6) + offset
                    sample_id_col = 6 + col_offset
                    ws.cell(row=1, column=sample_id_col, value=sample_id)
                    debug_print(f"DEBUG: Set sample ID '{sample_id}' at row 1, column {sample_id_col}")
                
                    # Row 2: Resistance goes in column D (4) + offset  
                    resistance_col = 4 + col_offset
                    if sample_resistance:  # Only set if not empty
                        try:
                            # Try to convert to float for numeric storage
                            resistance_value = float(sample_resistance)
                            ws.cell(row=2, column=resistance_col, value=resistance_value)
                        except ValueError:
                            # If not numeric, store as string
                            ws.cell(row=2, column=resistance_col, value=sample_resistance)
                    debug_print(f"DEBUG: Set resistance '{sample_resistance}' at row 2, column {resistance_col}")

                    # Set tester for each sample at row 3, column D (4) + offset (just name, no "Tester:" prefix)
                    tester_col = 4 + col_offset
                    tester_name = header_data['common']['tester']  # Just the name
                    ws.cell(row=3, column=tester_col, value=tester_name)
                    debug_print(f"DEBUG: Set tester '{tester_name}' at row 3, column {tester_col} for sample {i+1}")

                    # Apply common data to EACH sample block (not just the first one)
                    # Row 2, Column B (2) + offset: Media
                    if common_data["media"]:
                        media_col = 2 + col_offset
                        ws.cell(row=2, column=media_col, value=common_data["media"])
                        debug_print(f"DEBUG: Set media '{common_data['media']}' at row 2, column {media_col} for sample {i+1}")
                
                    # Row 3, Column B (2) + offset: Viscosity 
                    if common_data["viscosity"]:
                        viscosity_col = 2 + col_offset
                        try:
                            viscosity_value = float(common_data["viscosity"])
                            ws.cell(row=3, column=viscosity_col, value=viscosity_value)
                        except ValueError:
                            ws.cell(row=3, column=viscosity_col, value=common_data["viscosity"])
                        debug_print(f"DEBUG: Set viscosity '{common_data['viscosity']}' at row 3, column {viscosity_col} for sample {i+1}")
                
                    # Row 3, Column F (6) + offset: Voltage
                    if common_data["voltage"]:
                        voltage_col = 6 + col_offset
                        try:
                            voltage_value = float(common_data["voltage"])
                            ws.cell(row=3, column=voltage_col, value=voltage_value)
                        except ValueError:
                            ws.cell(row=3, column=voltage_col, value=common_data["voltage"])
                        debug_print(f"DEBUG: Set voltage '{common_data['voltage']}' at row 3, column {voltage_col} for sample {i+1}")
                
                    # Row 3, Column H (8) + offset: Oil mass
                    if common_data["oil_mass"]:
                        oil_mass_col = 8 + col_offset
                        try:
                            oil_mass_value = float(common_data["oil_mass"])
                            ws.cell(row=3, column=oil_mass_col, value=oil_mass_value)
                        except ValueError:
                            ws.cell(row=3, column=oil_mass_col, value=common_data["oil_mass"])
                        debug_print(f"DEBUG: Set oil mass '{common_data['oil_mass']}' at row 3, column {oil_mass_col} for sample {i+1}")
            
                    # Row 2, Column G (7) + offset: Puffing regime
                    puffing_regime = header_data['common'].get('puffing_regime', "Standard - 60mL/3s/30s")
                    puffing_regime_col = 7 + col_offset
                    ws.cell(row=2, column=puffing_regime_col, value=puffing_regime)
                    debug_print(f"DEBUG: Set puffing regime '{puffing_regime}' at row 2, column {puffing_regime_col} for sample {i+1}")

                # NEW: Delete extra columns after the last sample to clean up data visualization
                last_sample_column = num_samples * 12  # 12 columns per sample
                max_column = ws.max_column
            
                if max_column > last_sample_column:
                    debug_print(f"DEBUG: Deleting extra columns {last_sample_column + 1} to {max_column}")
                    # Delete columns from (last_sample_column + 1) to max_column
                    for col in range(max_column, last_sample_column, -1):  # Delete from right to left
                        ws.delete_cols(col)
                    debug_print(f"DEBUG: Deleted {max_column - last_sample_column} extra columns")
                else:
                    debug_print(f"DEBUG: No extra columns to delete. Last sample column: {last_sample_column}, Max column: {max_column}")
        
                # Save the workbook
                wb.save(file_path)
                debug_print(f"DEBUG: Successfully saved workbook to {file_path}")
        
                debug_print(f"SUCCESS: Applied header data for {num_samples} samples to {file_path}")
            
                # Verify the data was written correctly by reading it back
                debug_print("DEBUG: Verifying written data...")
                verification_wb = openpyxl.load_workbook(file_path)
                verification_ws = verification_wb[header_data["test"]]
            
                # Verify test name at row 1, column 1
                test_name_cell = verification_ws.cell(row=1, column=1)
                debug_print(f"DEBUG: Verification - Test name at row 1, col 1: '{test_name_cell.value}'")
            
                # Verify column count
                final_max_column = verification_ws.max_column
                debug_print(f"DEBUG: Verification - Final max column: {final_max_column} (expected: {last_sample_column})")
            
                for i in range(num_samples):
                    col_offset = i * 12
                    sample_id_cell = verification_ws.cell(row=1, column=6 + col_offset)
                    resistance_cell = verification_ws.cell(row=2, column=4 + col_offset)
                    tester_cell = verification_ws.cell(row=3, column=4 + col_offset)
                    media_cell = verification_ws.cell(row=2, column=2 + col_offset)
                    viscosity_cell = verification_ws.cell(row=3, column=2 + col_offset)
                    voltage_cell = verification_ws.cell(row=3, column=6 + col_offset)
                    oil_mass_cell = verification_ws.cell(row=3, column=8 + col_offset)
                    puffing_regime_cell = verification_ws.cell(row=2, column=7 + col_offset)
                
                    debug_print(f"DEBUG: Verification - Sample {i+1}:")
                    debug_print(f"  ID='{sample_id_cell.value}', Resistance='{resistance_cell.value}', Tester='{tester_cell.value}'")
                    debug_print(f"  Media='{media_cell.value}', Viscosity='{viscosity_cell.value}'")
                    debug_print(f"  Voltage='{voltage_cell.value}', Oil mass='{oil_mass_cell.value}'")
                    debug_print(f"  Puffing regime='{puffing_regime_cell.value}'")

                verification_wb.close()
                debug_print("DEBUG: Verification complete")
            
            else:
                error_msg = f"Sheet '{header_data['test']}' not found in the file. Available sheets: {wb.sheetnames}"
                print(f"ERROR: {error_msg}")
                messagebox.showerror("Error", error_msg)
            
        except Exception as e:
            error_msg = f"Error applying header data: {str(e)}"
            print(f"ERROR: {error_msg}")
            debug_print("DEBUG: Full traceback:")
            traceback.print_exc()
            messagebox.showerror("Error", error_msg)

    def start_file_loading_wrapper(self, startup_menu: tk.Toplevel) -> None:
        """Handle the 'Load' button click in the startup menu."""
        startup_menu.destroy()
        self.load_initial_file()

    def update_file_dropdown(self) -> None:
        """Update the file dropdown with loaded file names."""
        file_names = [file_data["file_name"] for file_data in self.gui.all_filtered_sheets]
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
        """Save the current session to a .vap3 file."""
        from vap_file_manager import VapFileManager
    
        if not filepath:
            filepath = filedialog.asksaveasfilename(
                title="Save As VAP3",
                defaultextension=".vap3",
                filetypes=[("VAP3 files", "*.vap3")]
            )
        
        if not filepath:
            return
    
        try:
            self.gui.progress_dialog.show_progress_bar("Saving VAP3 file...")
            self.root.update_idletasks()
        
            vap_manager = VapFileManager()
        
            # Collect plot settings
            plot_settings = {
                'selected_plot_type': self.gui.selected_plot_type.get() if hasattr(self.gui, 'selected_plot_type') else None
            }
        
            # Get image crop states
            image_crop_states = getattr(self.gui, 'image_crop_states', {})
            if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
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
        """Load a .vap3 file and update the application state."""
        from vap_file_manager import VapFileManager
    
        if not filepath:
            filepath = filedialog.askopenfilename(
                title="Open VAP3 File",
                filetypes=[("VAP3 files", "*.vap3")]
            )
        
        if not filepath:
            return False
    
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
        debug_print(f"DEBUG: Loading file for data collection: {file_path}")
        try:
            # Use the optimized ensure method instead of full reload
            self.ensure_file_is_loaded_in_ui(file_path)
            debug_print(f"DEBUG: File loaded successfully: {file_path}")
            return True
            
        except Exception as e:
            debug_print(f"DEBUG: Error loading file: {e}")
            messagebox.showerror("Error", f"Failed to load file: {e}")
            return False
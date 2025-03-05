# file_manager.py
import os
import copy
import shutil
import pandas as pd
import tkinter as tk
from openpyxl import load_workbook
from tkinter import filedialog, messagebox, Toplevel, Label, Button, ttk, Frame
import processing
import time
import psutil
from utils import is_valid_excel_file, get_resource_path, load_excel_file, get_save_path, is_standard_file, FONT, APP_BACKGROUND_COLOR

class FileManager:
    """File Management Module for DataViewer.
    
    This class is initialized with a reference to the main TestingGUI instance
    so that it can update its state (sheets, selected_sheet, etc.).
    """
    def __init__(self, gui):
        self.gui = gui
        self.root = gui.root
        # If you need your own lock, you can create one here:
        

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
                        print("\n--- Legacy Standard File: 'Intense Test' ---")
                        print(intense_test_df.iloc[:20, :15])



                    
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
                    print("\n--- Normal Standard File: 'Intense Test' ---")
                    print(intense_test_df.iloc[:20, :15])


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

    def open_raw_data_in_excel_old(self, sheet_name=None) -> None:
        """Opens an Excel file using the default system application (Excel)."""
        try:
            file_path = self.gui.file_path
            if not isinstance(file_path, str):
                raise ValueError(f"Invalid file path. Expected string, got {type(file_path).__name__}")
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"The file {file_path} does not exist.")
            temp_file = None
            if sheet_name:
                try:
                    # Create a temporary file with the specified sheet active
                    import tempfile
                    from openpyxl import load_workbook
                
                    # Create a temporary file
                    fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
                    os.close(fd)
                    temp_file = temp_path
                
                    # Make a copy of the original file
                    import shutil
                    shutil.copy2(file_path, temp_file)
                
                    # Modify the copy to set the active sheet
                    wb = load_workbook(temp_file)
                    if sheet_name in wb.sheetnames:
                        wb.active = wb[sheet_name]
                        wb.save(temp_file)
                        # Open the modified temporary file
                        os.startfile(temp_file)
                    else:
                        # If sheet doesn't exist, fall back to opening the original
                        os.startfile(file_path)
                except Exception as e:
                    print(f"Error preparing sheet-specific Excel file: {e}")
                    # Fall back to just opening the file
                    os.startfile(file_path)
            else:
                # No sheet specified, just open the file
                os.startfile(file_path)
            #wait for excel to open
            excel_opened = False
            for _ in range(10):
                if any(proc.name().lower() == "excel.exe" for proc in psutil.process_iter()):
                    excel_opened = True
                    break
                time.sleep(0.5)
            if not excel_opened:
                raise TimeoutError("Excel did not open in the expected time.")

            excel_processes = [proc for proc in psutil.process_iter() if proc.name().lower() == "excel.exe"]
            
            dialog = Toplevel(self.root)
            dialog.title("Excel Opened")
            dialog.geometry("300x150")
            dialog.grab_set()

            if sheet_name:
                msg = f"Excel has opened with the '{sheet_name}' sheet active."
            else:
                msg = "Excel has opened the file."
            msg += " This dialog will close automatically when Excel is closed."
        
            label = Label(dialog, text=msg, wraplength=280, justify="center")
            label.pack(pady=20)

            def close_excel_and_popup():
                for proc in excel_processes:
                    if proc.is_running():
                        try:
                            proc.terminate()
                            proc.wait()
                        except:
                            pass
                dialog.destroy()
                # clean up temp file if it exists
                if tep_file and os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass

            dialog.protocol("WM_DELETE_WINDOW", lambda: close_excel_and_popup()) 

            def monitor_excel():
                nonlocal excel_processes
                excel_processes = [proc for proc in excel_processes if proc.is_running()]
                if not any(proc.is_running() for proc in excel_processes):
                    dialog.destroy()
                    # Clean up the temporary file
                    if temp_file and os.path.exists(temp_file):
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                else:
                    dialog.after(100, monitor_excel)

            dialog.after(100, monitor_excel)
            close_button = Button(dialog, text="Close", command=close_excel_and_popup)
            close_button.pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {e}")

    def open_raw_data_in_excel(self, sheet_name=None) -> None:
        """
        Opens an Excel file for a specific sheet, allowing edits, and then
        updates the VAP3 file and GUI when Excel closes.
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
        
            # Create a temporary Excel file
            import tempfile
            from openpyxl import Workbook
        
            # Use a consistent naming scheme so we can identify our temp files
            temp_dir = tempfile.gettempdir()
            temp_file = os.path.join(temp_dir, f"dataviewer_edit_{sheet_name}_{os.getpid()}.xlsx")
        
            # Create a new workbook with just this sheet
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name
        
            # Write the headers
            for col_idx, column_name in enumerate(sheet_data.columns, 1):
                ws.cell(row=1, column=col_idx, value=str(column_name))
            
            # Write the data
            for row_idx, row in enumerate(sheet_data.itertuples(index=False), 2):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
                
            # Save the workbook
            wb.save(temp_file)
        
            # Record the file modification time before opening
            original_mod_time = os.path.getmtime(temp_file)
        
            # Open the file in Excel
            os.startfile(temp_file)
        
            # Wait for Excel to open
            excel_opened = False
            for _ in range(10):
                if any(proc.name().lower() == "excel.exe" for proc in psutil.process_iter()):
                    excel_opened = True
                    break
                time.sleep(0.5)
            
            if not excel_opened:
                raise TimeoutError("Excel did not open in the expected time.")
            
            # Get Excel processes for monitoring
            excel_processes = [proc for proc in psutil.process_iter() if proc.name().lower() == "excel.exe"]
        
            # Create dialog
            from tkinter import Toplevel, Label, Button, Frame
            dialog = Toplevel(self.root)
            dialog.title("Excel Edit Mode")
            dialog.geometry("400x200")
            dialog.grab_set()
        
            main_frame = Frame(dialog)
            main_frame.pack(pady=10, padx=15, fill="both", expand=True)
        
            # Information label
            Label(main_frame, text=f"Excel has opened with the '{sheet_name}' sheet.", 
                  font=("Arial", 10, "bold")).pack(pady=(0,5))
        
            # Instructions
            instruction_text = (
                "Make your changes in Excel and save the file.\n"
                "When you close Excel, your changes will be automatically imported."
            )
            Label(main_frame, text=instruction_text, justify="center", wraplength=360).pack(pady=5)
        
            # Warning/note
            note_text = (
                "Note: Only changes to data values will be preserved. "
                "Formatting, formulas, and structural changes may be lost."
            )
            Label(main_frame, text=note_text, font=("Arial", 8), fg="gray", 
                  wraplength=360, justify="center").pack(pady=5)
        
            def process_excel_changes():
                """Process changes made in Excel after it closes"""
                try:
                    # Check if the file was modified
                    if os.path.exists(temp_file):
                        new_mod_time = os.path.getmtime(temp_file)
                    
                        if new_mod_time > original_mod_time:
                            # File was modified, read it back
                            
                        
                            modified_data = pd.read_excel(temp_file)
                        
                            # Update the filtered sheets with new data
                            self.gui.filtered_sheets[sheet_name]['data'] = modified_data
                        
                            # If we're using a VAP3 file, update it
                            if hasattr(self.gui, 'file_path') and self.gui.file_path.endswith('.vap3'):
                                from vap_file_manager import VapFileManager
                            
                                vap_manager = VapFileManager()
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
                            messagebox.showinfo("Success", "Changes from Excel have been imported successfully.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to process Excel changes: {e}")
                finally:
                    # Always try to clean up the temp file
                    if os.path.exists(temp_file):
                        try:
                            os.remove(temp_file)
                        except:
                            pass
        
            def close_excel_and_process():
                """Close Excel and process any changes"""
                for proc in excel_processes:
                    if proc.is_running():
                        try:
                            proc.terminate()
                            proc.wait()
                        except:
                            pass
                process_excel_changes()
                dialog.destroy()
        
            def monitor_excel():
                """Monitor Excel and process changes when it closes"""
                nonlocal excel_processes
                excel_processes = [proc for proc in excel_processes if proc.is_running()]
                if not any(proc.is_running() for proc in excel_processes):
                    process_excel_changes()
                    dialog.destroy()
                else:
                    dialog.after(100, monitor_excel)
        
            dialog.protocol("WM_DELETE_WINDOW", close_excel_and_process)
            dialog.after(100, monitor_excel)
        
            # Close button
            Button(dialog, text="Close Excel & Import Changes", 
                   command=close_excel_and_process).pack(pady=10)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {e}")

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
        print("DEBUG: file_names = ", file_names)
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
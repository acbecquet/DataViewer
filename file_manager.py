# file_manager.py
import os
import copy
import shutil
import pandas as pd
import tkinter as tk
from openpyxl import load_workbook
from tkinter import filedialog, messagebox, Toplevel, Label, Button, ttk
import processing
import time
import psutil
from utils import plotting_sheet_test, get_resource_path, get_save_path, is_standard_file, clean_columns, FONT, APP_BACKGROUND_COLOR, BUTTON_COLOR, PLOT_CHECKBOX_TITLE

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
            if not processing.is_valid_excel_file(os.path.basename(file_path)):
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
                self.gui.sheets = processing.load_excel_file(file_path)
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

    def open_raw_data_in_excel(self, sheet_name=None) -> None:
        """Opens an Excel file using the default system application (Excel)."""
        try:
            file_path = self.gui.file_path
            if not isinstance(file_path, str):
                raise ValueError(f"Invalid file path. Expected string, got {type(file_path).__name__}")
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"The file {file_path} does not exist.")
            os.startfile(file_path)
            for _ in range(10):
                if any(proc.name().lower() == "excel.exe" for proc in psutil.process_iter()):
                    break
                time.sleep(0.5)
            else:
                raise TimeoutError("Excel did not open in the expected time.")

            excel_processes = [proc for proc in psutil.process_iter() if proc.name().lower() == "excel.exe"]
            from tkinter import Toplevel, Label, Button  # import here if not already imported
            dialog = Toplevel(self.root)
            dialog.title("Excel Opened")
            dialog.geometry("300x150")
            dialog.grab_set()
            dialog.protocol("WM_DELETE_WINDOW", lambda: close_excel_and_popup())
            label = Label(dialog, text=f"The file has been opened in Excel. Please navigate to the sheet '{sheet_name}' if necessary. This dialog will close automatically when Excel is closed.", wraplength=280, justify="center")
            label.pack(pady=20)

            def close_excel_and_popup():
                for proc in excel_processes:
                    if proc.is_running():
                        proc.terminate()
                        proc.wait()
                dialog.destroy()

            def monitor_excel():
                nonlocal excel_processes
                excel_processes = [proc for proc in excel_processes if proc.is_running()]
                if not any(proc.is_running() for proc in excel_processes):
                    dialog.destroy()
                else:
                    dialog.after(100, monitor_excel)
            dialog.after(100, monitor_excel)
            close_button = Button(dialog, text="Close", command=close_excel_and_popup)
            close_button.pack(pady=10)
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
            self.gui.clear_display_frame(is_plotting_sheet=False, is_empty_sheet=False)
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
            dropdown_frame = ttk.Frame(self.gui.top_frame, width=400, height=50)
            dropdown_frame.pack(side="left", pady=2, padx=5)
            file_label = ttk.Label(dropdown_frame, text="Select File:", font=FONT, foreground="white", background=APP_BACKGROUND_COLOR)
            file_label.pack(side="left", padx=(0, 5))
            self.gui.file_dropdown_var = tk.StringVar()
            self.gui.file_dropdown = ttk.Combobox(
                dropdown_frame,
                textvariable=self.gui.file_dropdown_var,
                state="readonly",
                font=FONT,
                width=10
            )
            self.gui.file_dropdown.pack(side="left", fill="x", expand=True, padx=(5, 10))
            self.gui.file_dropdown.bind("<<ComboboxSelected>>", self.gui.on_file_selection)
        self.update_file_dropdown()


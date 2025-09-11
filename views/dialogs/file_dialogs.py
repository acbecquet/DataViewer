# views/dialogs/file_dialogs.py
"""
views/dialogs/file_dialogs.py
File dialog views for file operations.
This contains file dialog UI functionality from file_manager.py and file_selection_dialog.py.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Frame, Label, Button, Listbox, Scrollbar, Checkbutton
from typing import Optional, Dict, Any, List, Callable, Tuple
import os


class FileDialogs:
    """File dialog utilities for opening, saving, and selecting files."""
    
    def __init__(self, parent: tk.Widget):
        """Initialize file dialogs."""
        self.parent = parent
        
        # Configuration
        self.app_background_color = '#EFEFEF'
        self.font = ("Arial", 10)
        
        print("DEBUG: FileDialogs initialized")
    
    def open_excel_file(self, title: str = "Select Excel File") -> Optional[str]:
        """Show dialog to open a single Excel file."""
        try:
            file_path = filedialog.askopenfilename(
                parent=self.parent,
                title=title,
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            
            if file_path:
                print(f"DEBUG: FileDialogs - Excel file selected: {os.path.basename(file_path)}")
            else:
                print("DEBUG: FileDialogs - no Excel file selected")
            
            return file_path if file_path else None
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error opening Excel file dialog: {e}")
            return None
    
    def open_excel_files(self, title: str = "Select Excel Files") -> List[str]:
        """Show dialog to open multiple Excel files."""
        try:
            file_paths = filedialog.askopenfilenames(
                parent=self.parent,
                title=title,
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            
            if file_paths:
                print(f"DEBUG: FileDialogs - {len(file_paths)} Excel files selected")
            else:
                print("DEBUG: FileDialogs - no Excel files selected")
            
            return list(file_paths) if file_paths else []
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error opening Excel files dialog: {e}")
            return []
    
    def open_vap3_file(self, title: str = "Select VAP3 File") -> Optional[str]:
        """Show dialog to open a VAP3 file."""
        try:
            file_path = filedialog.askopenfilename(
                parent=self.parent,
                title=title,
                filetypes=[
                    ("VAP3 files", "*.vap3"),
                    ("All files", "*.*")
                ]
            )
            
            if file_path:
                print(f"DEBUG: FileDialogs - VAP3 file selected: {os.path.basename(file_path)}")
            else:
                print("DEBUG: FileDialogs - no VAP3 file selected")
            
            return file_path if file_path else None
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error opening VAP3 file dialog: {e}")
            return None
    
    def save_vap3_file(self, title: str = "Save As VAP3", default_name: str = "") -> Optional[str]:
        """Show dialog to save a VAP3 file."""
        try:
            file_path = filedialog.asksaveasfilename(
                parent=self.parent,
                title=title,
                defaultextension=".vap3",
                initialvalue=default_name,
                filetypes=[
                    ("VAP3 files", "*.vap3"),
                    ("All files", "*.*")
                ]
            )
            
            if file_path:
                print(f"DEBUG: FileDialogs - VAP3 save path: {file_path}")
            else:
                print("DEBUG: FileDialogs - VAP3 save cancelled")
            
            return file_path if file_path else None
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error saving VAP3 file: {e}")
            return None
    
    def save_excel_file(self, title: str = "Save Excel File", default_name: str = "") -> Optional[str]:
        """Show dialog to save an Excel file."""
        try:
            file_path = filedialog.asksaveasfilename(
                parent=self.parent,
                title=title,
                defaultextension=".xlsx",
                initialvalue=default_name,
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("Excel 97-2003", "*.xls"),
                    ("All files", "*.*")
                ]
            )
            
            if file_path:
                print(f"DEBUG: FileDialogs - Excel save path: {file_path}")
            else:
                print("DEBUG: FileDialogs - Excel save cancelled")
            
            return file_path if file_path else None
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error saving Excel file: {e}")
            return None
    
    def select_directory(self, title: str = "Select Directory") -> Optional[str]:
        """Show dialog to select a directory."""
        try:
            directory = filedialog.askdirectory(
                parent=self.parent,
                title=title
            )
            
            if directory:
                print(f"DEBUG: FileDialogs - directory selected: {directory}")
            else:
                print("DEBUG: FileDialogs - no directory selected")
            
            return directory if directory else None
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error selecting directory: {e}")
            return None
    
    def show_file_selection_dialog(self, available_files: List[Dict[str, Any]], 
                                  title: str = "Select Files", 
                                  min_selection: int = 1) -> Tuple[bool, List[Dict[str, Any]]]:
        """Show dialog to select multiple files from a list."""
        try:
            dialog = FileSelectionDialog(self.parent, available_files, title, min_selection)
            return dialog.show()
        except Exception as e:
            print(f"ERROR: FileDialogs - error showing file selection dialog: {e}")
            return False, []
    
    def show_database_browser(self, files: List[Dict[str, Any]], 
                            callback: Optional[Callable] = None) -> None:
        """Show database browser dialog."""
        try:
            print("DEBUG: FileDialogs - showing database browser")
            
            # Create database browser dialog
            dialog = tk.Toplevel(self.parent)
            dialog.title("Database Browser")
            dialog.geometry("800x600")
            dialog.transient(self.parent)
            dialog.grab_set()
            
            # Center dialog
            self._center_dialog(dialog)
            
            # Create main frame
            main_frame = ttk.Frame(dialog)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Header
            header_frame = ttk.Frame(main_frame)
            header_frame.pack(fill="x", pady=(0, 10))
            
            title_label = ttk.Label(header_frame, text="Database Files", font=("Arial", 14, "bold"))
            title_label.pack(side="left")
            
            count_label = ttk.Label(header_frame, text=f"({len(files)} files)")
            count_label.pack(side="left", padx=(5, 0))
            
            # Create listbox with scrollbar
            list_frame = ttk.Frame(main_frame)
            list_frame.pack(fill="both", expand=True, pady=(0, 10))
            
            # Configure grid weights
            list_frame.grid_rowconfigure(0, weight=1)
            list_frame.grid_columnconfigure(0, weight=1)
            
            scrollbar = ttk.Scrollbar(list_frame)
            scrollbar.grid(row=0, column=1, sticky="ns")
            
            listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=self.font)
            listbox.grid(row=0, column=0, sticky="nsew")
            scrollbar.config(command=listbox.yview)
            
            # Populate listbox
            for file_info in files:
                filename = file_info.get('filename', 'Unknown')
                created_at = file_info.get('created_at', 'Unknown date')
                if hasattr(created_at, 'strftime'):
                    created_at = created_at.strftime('%Y-%m-%d %H:%M')
                display_text = f"{filename} - {created_at}"
                listbox.insert(tk.END, display_text)
            
            # Info frame
            info_frame = ttk.LabelFrame(main_frame, text="File Information", padding=5)
            info_frame.pack(fill="x", pady=(0, 10))
            
            info_text = tk.Text(info_frame, height=4, wrap=tk.WORD, state=tk.DISABLED)
            info_text.pack(fill="x")
            
            def on_selection_change(event):
                """Handle listbox selection change."""
                try:
                    selection = listbox.curselection()
                    if selection:
                        file_info = files[selection[0]]
                        info_text.config(state=tk.NORMAL)
                        info_text.delete(1.0, tk.END)
                        
                        # Display file information
                        info_text.insert(tk.END, f"Filename: {file_info.get('filename', 'N/A')}\n")
                        info_text.insert(tk.END, f"Created: {file_info.get('created_at', 'N/A')}\n")
                        info_text.insert(tk.END, f"Size: {file_info.get('file_size', 'N/A')} bytes\n")
                        info_text.insert(tk.END, f"Sheets: {file_info.get('sheet_count', 'N/A')}")
                        
                        info_text.config(state=tk.DISABLED)
                except Exception as e:
                    print(f"DEBUG: FileDialogs - error updating info: {e}")
            
            listbox.bind('<<ListboxSelect>>', on_selection_change)
            
            # Button frame
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            def on_load():
                """Handle load button click."""
                try:
                    selection = listbox.curselection()
                    if selection and callback:
                        selected_file = files[selection[0]]
                        callback(selected_file)
                        dialog.destroy()
                    elif not selection:
                        messagebox.showwarning("No Selection", "Please select a file to load.")
                except Exception as e:
                    print(f"ERROR: FileDialogs - error loading file: {e}")
            
            def on_cancel():
                """Handle cancel button click."""
                dialog.destroy()
            
            ttk.Button(button_frame, text="Load Selected", command=on_load).pack(side="left", padx=(0, 5))
            ttk.Button(button_frame, text="Refresh", command=lambda: self._refresh_database_list()).pack(side="left", padx=(0, 5))
            ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="right")
            
            print(f"DEBUG: FileDialogs - database browser shown with {len(files)} files")
            
        except Exception as e:
            print(f"ERROR: FileDialogs - error showing database browser: {e}")
    
    def _center_dialog(self, dialog: tk.Toplevel):
        """Center a dialog on the parent window."""
        try:
            dialog.update_idletasks()
            
            # Get parent position and size
            parent_x = self.parent.winfo_x()
            parent_y = self.parent.winfo_y()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            # Get dialog size
            dialog_width = dialog.winfo_width()
            dialog_height = dialog.winfo_height()
            
            # Calculate center position
            x = parent_x + (parent_width - dialog_width) // 2
            y = parent_y + (parent_height - dialog_height) // 2
            
            # Ensure dialog stays on screen
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            
            x = max(0, min(x, screen_width - dialog_width))
            y = max(0, min(y, screen_height - dialog_height))
            
            dialog.geometry(f"+{x}+{y}")
            
        except tk.TclError:
            pass
        except Exception as e:
            print(f"ERROR: FileDialogs - error centering dialog: {e}")
    
    def _refresh_database_list(self):
        """Refresh database list (placeholder)."""
        print("DEBUG: FileDialogs - refresh database list requested")


class FileSelectionDialog:
    """Dialog for selecting multiple files from a list."""
    
    def __init__(self, parent: tk.Widget, available_files: List[Dict[str, Any]], 
                 title: str = "Select Files", min_selection: int = 1):
        """Initialize the file selection dialog."""
        self.parent = parent
        self.available_files = available_files
        self.title = title
        self.min_selection = min_selection
        
        # Dialog components
        self.dialog: Optional[tk.Toplevel] = None
        self.file_vars: Dict[int, tk.BooleanVar] = {}
        self.selected_files: List[Dict[str, Any]] = []
        self.result = False
        
        # UI components
        self.canvas: Optional[tk.Canvas] = None
        self.scrollbar: Optional[Scrollbar] = None
        self.checkbox_frame: Optional[Frame] = None
        
        print("DEBUG: FileSelectionDialog initialized")
    
    def show(self) -> Tuple[bool, List[Dict[str, Any]]]:
        """Show the dialog and return results."""
        try:
            # Create dialog
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title(self.title)
            self.dialog.geometry("600x500")
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            
            # Center dialog
            self._center_dialog()
            
            # Create layout
            self._create_layout()
            
            # Setup cleanup
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
            
            # Wait for dialog to close
            self.dialog.wait_window()
            
            print(f"DEBUG: FileSelectionDialog closed with result: {self.result}")
            return self.result, self.selected_files
            
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error showing dialog: {e}")
            return False, []
    
    def _create_layout(self):
        """Create the dialog layout."""
        try:
            if not self.dialog:
                return
            
            # Configure grid weights
            self.dialog.grid_columnconfigure(0, weight=1)
            self.dialog.grid_columnconfigure(1, weight=0)
            self.dialog.grid_rowconfigure(0, weight=0)
            self.dialog.grid_rowconfigure(1, weight=1)
            self.dialog.grid_rowconfigure(2, weight=0)
            
            # Header
            header_frame = ttk.Frame(self.dialog)
            header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
            
            header_label = ttk.Label(header_frame, 
                                   text="Select files to include in your comparison analysis:",
                                   font=("Arial", 12))
            header_label.pack(anchor="w")
            
            # File list frame
            file_frame = ttk.Frame(self.dialog)
            file_frame.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=5)
            file_frame.grid_columnconfigure(0, weight=1)
            file_frame.grid_rowconfigure(0, weight=1)
            
            # Create scrollable area
            self.canvas = tk.Canvas(file_frame)
            self.scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=self.canvas.yview)
            
            self.checkbox_frame = ttk.Frame(self.canvas)
            self.checkbox_frame.bind("<Configure>", 
                                   lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
            
            # Create window in canvas
            canvas_window = self.canvas.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
            
            def on_canvas_configure(event):
                self.canvas.itemconfig(canvas_window, width=event.width)
            
            self.canvas.bind("<Configure>", on_canvas_configure)
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            
            # Pack scrollable components
            self.canvas.grid(row=0, column=0, sticky="nsew")
            self.scrollbar.grid(row=0, column=1, sticky="ns")
            
            # Populate checkboxes
            self._create_file_checkboxes()
            
            # Control buttons frame
            control_frame = ttk.Frame(self.dialog)
            control_frame.grid(row=1, column=1, sticky="ns", padx=(5, 10), pady=5)
            
            ttk.Button(control_frame, text="Select All", command=self._select_all).pack(fill="x", pady=(0, 5))
            ttk.Button(control_frame, text="Deselect All", command=self._deselect_all).pack(fill="x", pady=(0, 5))
            
            # Bottom buttons
            button_frame = ttk.Frame(self.dialog)
            button_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
            
            ttk.Button(button_frame, text="OK", command=self._on_ok).pack(side="right", padx=(5, 0))
            ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="right")
            
            # Selection info
            self.info_label = ttk.Label(button_frame, text="No files selected")
            self.info_label.pack(side="left")
            
            self._update_selection_info()
            
            print("DEBUG: FileSelectionDialog layout created")
            
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error creating layout: {e}")
    
    def _create_file_checkboxes(self):
        """Create checkboxes for each file."""
        try:
            for i, file_data in enumerate(self.available_files):
                # Create variable for checkbox
                var = tk.BooleanVar()
                self.file_vars[i] = var
                
                # Get display name
                display_name = file_data.get('display_filename', file_data.get('file_name', f'File_{i}'))
                created_at = file_data.get('created_at', '')
                if hasattr(created_at, 'strftime'):
                    created_at = created_at.strftime('%Y-%m-%d %H:%M')
                
                # Create checkbox with file info
                checkbox_text = f"{display_name}"
                if created_at:
                    checkbox_text += f" ({created_at})"
                
                checkbox = ttk.Checkbutton(self.checkbox_frame, text=checkbox_text, variable=var,
                                         command=self._update_selection_info)
                checkbox.grid(row=i, column=0, sticky="w", padx=5, pady=2)
                
                # Bind mousewheel to canvas
                def bind_mousewheel(event):
                    if self.canvas:
                        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
                
                def unbind_mousewheel(event):
                    if self.canvas:
                        self.canvas.unbind_all("<MouseWheel>")
                
                checkbox.bind("<Enter>", bind_mousewheel)
                checkbox.bind("<Leave>", unbind_mousewheel)
            
            print(f"DEBUG: FileSelectionDialog - created {len(self.available_files)} checkboxes")
            
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error creating checkboxes: {e}")
    
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling."""
        try:
            if self.canvas:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception as e:
            print(f"DEBUG: FileSelectionDialog - mousewheel error: {e}")
    
    def _select_all(self):
        """Select all files."""
        try:
            for var in self.file_vars.values():
                var.set(True)
            self._update_selection_info()
            print("DEBUG: FileSelectionDialog - all files selected")
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error selecting all: {e}")
    
    def _deselect_all(self):
        """Deselect all files."""
        try:
            for var in self.file_vars.values():
                var.set(False)
            self._update_selection_info()
            print("DEBUG: FileSelectionDialog - all files deselected")
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error deselecting all: {e}")
    
    def _update_selection_info(self):
        """Update selection information display."""
        try:
            selected_count = sum(1 for var in self.file_vars.values() if var.get())
            total_count = len(self.file_vars)
            
            if hasattr(self, 'info_label') and self.info_label:
                self.info_label.config(text=f"{selected_count} of {total_count} files selected")
            
            print(f"DEBUG: FileSelectionDialog - {selected_count}/{total_count} files selected")
            
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error updating selection info: {e}")
    
    def _on_ok(self):
        """Handle OK button click."""
        try:
            # Get selected files
            selected_indices = [i for i, var in self.file_vars.items() if var.get()]
            
            # Check minimum selection requirement
            if len(selected_indices) < self.min_selection:
                messagebox.showwarning("Insufficient Selection", 
                                     f"Please select at least {self.min_selection} file(s).")
                return
            
            # Get selected file data
            self.selected_files = [self.available_files[i] for i in selected_indices]
            self.result = True
            
            print(f"DEBUG: FileSelectionDialog - {len(self.selected_files)} files selected for processing")
            
            if self.dialog:
                self.dialog.destroy()
                
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error in OK handler: {e}")
    
    def _on_cancel(self):
        """Handle Cancel button click."""
        try:
            self.result = False
            self.selected_files = []
            
            if self.dialog:
                self.dialog.destroy()
            
            print("DEBUG: FileSelectionDialog - cancelled")
            
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error in cancel handler: {e}")
    
    def _center_dialog(self):
        """Center the dialog on parent."""
        try:
            if not self.dialog:
                return
            
            self.dialog.update_idletasks()
            
            # Get parent info
            parent_x = self.parent.winfo_x()
            parent_y = self.parent.winfo_y()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            # Get dialog size
            dialog_width = self.dialog.winfo_width()
            dialog_height = self.dialog.winfo_height()
            
            # Calculate position
            x = parent_x + (parent_width - dialog_width) // 2
            y = parent_y + (parent_height - dialog_height) // 2
            
            self.dialog.geometry(f"+{x}+{y}")
            
        except Exception as e:
            print(f"ERROR: FileSelectionDialog - error centering: {e}")
# views/dialogs/file_dialogs.py
"""
views/dialogs/file_dialogs.py
File selection and management dialogs.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List, Callable, Dict, Any


class FileDialogs:
    """Collection of file-related dialogs."""
    
    def __init__(self, parent: tk.Tk):
        """Initialize file dialogs."""
        self.parent = parent
        
        print("DEBUG: FileDialogs initialized")
    
    def select_excel_file(self) -> Optional[str]:
        """Show dialog to select an Excel file."""
        print("DEBUG: FileDialogs - selecting Excel file")
        
        file_path = filedialog.askopenfilename(
            parent=self.parent,
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            print(f"DEBUG: FileDialogs - selected: {file_path}")
        else:
            print("DEBUG: FileDialogs - no file selected")
        
        return file_path if file_path else None
    
    def select_vap3_file(self) -> Optional[str]:
        """Show dialog to select a VAP3 file."""
        print("DEBUG: FileDialogs - selecting VAP3 file")
        
        file_path = filedialog.askopenfilename(
            parent=self.parent,
            title="Select VAP3 File",
            filetypes=[
                ("VAP3 files", "*.vap3"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            print(f"DEBUG: FileDialogs - selected: {file_path}")
        else:
            print("DEBUG: FileDialogs - no file selected")
        
        return file_path if file_path else None
    
    def save_vap3_file(self) -> Optional[str]:
        """Show dialog to save a VAP3 file."""
        print("DEBUG: FileDialogs - saving VAP3 file")
        
        file_path = filedialog.asksaveasfilename(
            parent=self.parent,
            title="Save VAP3 File",
            defaultextension=".vap3",
            filetypes=[
                ("VAP3 files", "*.vap3"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            print(f"DEBUG: FileDialogs - save path: {file_path}")
        else:
            print("DEBUG: FileDialogs - save cancelled")
        
        return file_path if file_path else None
    
    def select_directory(self, title: str = "Select Directory") -> Optional[str]:
        """Show dialog to select a directory."""
        print(f"DEBUG: FileDialogs - selecting directory: {title}")
        
        directory = filedialog.askdirectory(
            parent=self.parent,
            title=title
        )
        
        if directory:
            print(f"DEBUG: FileDialogs - selected directory: {directory}")
        else:
            print("DEBUG: FileDialogs - no directory selected")
        
        return directory if directory else None
    
    def show_database_browser(self, files: List[Dict[str, Any]], 
                            callback: Optional[Callable] = None) -> None:
        """Show database browser dialog."""
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
        
        # Create listbox with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate listbox
        for file_info in files:
            display_text = f"{file_info.get('filename', 'Unknown')} - {file_info.get('created_at', 'Unknown date')}"
            listbox.insert(tk.END, display_text)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        def on_load():
            selection = listbox.curselection()
            if selection and callback:
                selected_file = files[selection[0]]
                callback(selected_file)
                dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Load", command=on_load).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left")
        
        print(f"DEBUG: FileDialogs - database browser shown with {len(files)} files")
    
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
            
            dialog.geometry(f"+{x}+{y}")
            
        except tk.TclError:
            pass
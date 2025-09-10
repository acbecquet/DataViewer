# views/dialogs/startup_dialog.py
"""
views/dialogs/startup_dialog.py
Startup and new template dialogs.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable


class StartupDialog:
    """Startup dialog for new file creation."""
    
    def __init__(self, parent: tk.Tk):
        """Initialize the startup dialog."""
        self.parent = parent
        self.dialog: Optional[tk.Toplevel] = None
        self.result: Optional[str] = None
        
        print("DEBUG: StartupDialog initialized")
    
    def show_new_template_dialog(self) -> Optional[str]:
        """Show dialog for creating new templates."""
        print("DEBUG: StartupDialog - showing new template dialog")
        
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Create New Template")
        self.dialog.geometry("400x200")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center dialog
        self._center_dialog()
        
        # Create layout
        self._create_template_layout()
        
        # Wait for dialog to close
        self.dialog.wait_window()
        
        return self.result
    
    def _create_template_layout(self):
        """Create the template dialog layout."""
        if not self.dialog:
            return
        
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_frame, text="Create New Template", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Template options
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Button(
            options_frame,
            text="Standard Testing Template",
            command=lambda: self._on_template_selected("standard"),
            width=25
        ).pack(pady=(0, 10))
        
        ttk.Button(
            options_frame,
            text="Custom Template",
            command=lambda: self._on_template_selected("custom"),
            width=25
        ).pack(pady=(0, 10))
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="right")
    
    def _on_template_selected(self, template_type: str):
        """Handle template selection."""
        print(f"DEBUG: StartupDialog - template selected: {template_type}")
        self.result = template_type
        if self.dialog:
            self.dialog.destroy()
    
    def _on_cancel(self):
        """Handle cancel button."""
        print("DEBUG: StartupDialog - cancelled")
        self.result = None
        if self.dialog:
            self.dialog.destroy()
    
    def show_startup_menu(self, callback: Optional[Callable] = None):
        """Show the main startup menu."""
        print("DEBUG: StartupDialog - showing startup menu")
        
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("DataViewer - Welcome")
        self.dialog.geometry("300x150")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center dialog
        self._center_dialog()
        
        # Create layout
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Welcome message
        welcome_label = ttk.Label(main_frame, text="Welcome to DataViewer", font=("Arial", 12, "bold"))
        welcome_label.pack(pady=(0, 20))
        
        # Action buttons
        ttk.Button(
            main_frame,
            text="Create New Template",
            command=lambda: self._startup_action("new_template", callback),
            width=20
        ).pack(pady=(0, 5))
        
        ttk.Button(
            main_frame,
            text="Load Existing File",
            command=lambda: self._startup_action("load_file", callback),
            width=20
        ).pack(pady=(0, 5))
        
        ttk.Button(
            main_frame,
            text="Browse Database",
            command=lambda: self._startup_action("browse_database", callback),
            width=20
        ).pack()
    
    def _startup_action(self, action: str, callback: Optional[Callable]):
        """Handle startup action selection."""
        print(f"DEBUG: StartupDialog - startup action: {action}")
        
        if callback:
            callback(action)
        
        if self.dialog:
            self.dialog.destroy()
            self.dialog = None
    
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        if not self.dialog or not self.parent:
            return
        
        try:
            self.dialog.update_idletasks()
            
            # Get parent position and size
            parent_x = self.parent.winfo_x()
            parent_y = self.parent.winfo_y()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            # Get dialog size
            dialog_width = self.dialog.winfo_width()
            dialog_height = self.dialog.winfo_height()
            
            # Calculate center position
            x = parent_x + (parent_width - dialog_width) // 2
            y = parent_y + (parent_height - dialog_height) // 2
            
            self.dialog.geometry(f"+{x}+{y}")
            
        except tk.TclError:
            pass
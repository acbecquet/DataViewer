# views/dialogs/progress_dialog.py
"""
views/dialogs/progress_dialog.py
Progress dialog view.
This will contain the UI from progress_dialog.py.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional


class ProgressDialog:
    """Progress dialog for showing operation progress."""
    
    def __init__(self, parent: tk.Tk):
        """Initialize the progress dialog."""
        self.parent = parent
        self.dialog: Optional[tk.Toplevel] = None
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar()
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.status_label: Optional[ttk.Label] = None
        
        print("DEBUG: ProgressDialog initialized")
    
    def show_progress_bar(self, title: str = "Processing", message: str = "Please wait..."):
        """Show the progress dialog."""
        if self.dialog:
            return  # Already showing
        
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title(title)
        self.dialog.geometry("400x120")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self._center_dialog()
        
        # Create widgets
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Status message
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.pack(pady=(0, 10))
        self.status_var.set(message)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=300,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 10))
        
        # Progress percentage label
        self.progress_label = ttk.Label(main_frame, text="0%")
        self.progress_label.pack()
        
        # Update display
        self.dialog.update()
        
        print(f"DEBUG: ProgressDialog shown - {title}")
    
    def update_progress_bar(self, progress: float, message: str = ""):
        """Update the progress bar value and message."""
        if not self.dialog:
            return
        
        try:
            # Update progress value
            self.progress_var.set(progress)
            
            # Update progress percentage
            if hasattr(self, 'progress_label'):
                self.progress_label.config(text=f"{progress:.0f}%")
            
            # Update status message if provided
            if message:
                self.status_var.set(message)
            
            # Force update
            self.dialog.update_idletasks()
            self.dialog.update()
            
            print(f"DEBUG: ProgressDialog updated - {progress:.1f}%")
            
        except tk.TclError:
            # Dialog may have been closed
            pass
    
    def hide_progress_bar(self):
        """Hide the progress dialog."""
        if self.dialog:
            try:
                self.dialog.grab_release()
                self.dialog.destroy()
                self.dialog = None
                print("DEBUG: ProgressDialog hidden")
            except tk.TclError:
                pass
    
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        if not self.dialog or not self.parent:
            return
        
        try:
            self.dialog.update_idletasks()
            
            # Get parent window position and size
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
    
    def is_showing(self) -> bool:
        """Check if the progress dialog is currently showing."""
        return self.dialog is not None
# views/dialogs/progress_dialog.py
"""
views/dialogs/progress_dialog.py
Progress dialog view.
This contains the UI from progress_dialog.py refactor file.
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
        self.progress_label: Optional[ttk.Label] = None
        
        # Configuration
        self.app_background_color = '#0504AA'
        self.font = ("Arial", 12)
        
        print("DEBUG: ProgressDialog initialized")
    
    def show_progress_bar(self, title: str = "Processing", message: str = "Please wait..."):
        """Show the progress dialog."""
        try:
            if self.dialog and self.dialog.winfo_exists():
                return  # Already showing
            
            # Create dialog window
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title(title)
            self.dialog.geometry("400x120")
            self.dialog.resizable(False, False)
            self.dialog.configure(bg=self.app_background_color)
            
            # Make modal
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            
            # Center the dialog
            self._center_dialog()
            
            # Create status label with white text
            self.status_label = tk.Label(
                self.dialog,
                textvariable=self.status_var,
                fg="white",
                bg=self.app_background_color,
                font=self.font
            )
            self.status_label.pack(pady=(10, 5))
            self.status_var.set(message)
            
            # Create progress bar
            self.progress_bar = ttk.Progressbar(
                self.dialog,
                variable=self.progress_var,
                orient='horizontal',
                mode='determinate',
                length=300,
                maximum=100
            )
            self.progress_bar.pack(pady=(5, 5))
            self.progress_bar['value'] = 0
            
            # Create percentage label
            self.progress_label = tk.Label(
                self.dialog,
                text="0%",
                fg="white",
                bg=self.app_background_color,
                font=self.font
            )
            self.progress_label.pack(pady=(0, 10))
            
            # Update display
            self.dialog.update_idletasks()
            self.dialog.update()
            
            print(f"DEBUG: ProgressDialog shown - {title}")
            
        except Exception as e:
            print(f"ERROR: ProgressDialog - error showing dialog: {e}")
    
    def update_progress_bar(self, progress: float, message: str = ""):
        """Update the progress bar value and message."""
        try:
            if not self.dialog or not self.dialog.winfo_exists():
                return
            
            # Update progress value (ensure it's within bounds)
            progress = max(0, min(100, progress))
            self.progress_var.set(progress)
            
            # Update progress bar directly
            if self.progress_bar and self.progress_bar.winfo_exists():
                self.progress_bar['value'] = progress
            
            # Update percentage label
            if self.progress_label and self.progress_label.winfo_exists():
                self.progress_label.config(text=f"{progress:.0f}%")
            
            # Update status message if provided
            if message and self.status_var:
                self.status_var.set(message)
            
            # Force update
            self.dialog.update_idletasks()
            self.dialog.update()
            
            print(f"DEBUG: ProgressDialog updated - {progress:.1f}%")
            
        except tk.TclError:
            # Dialog may have been closed
            print("DEBUG: ProgressDialog - dialog was closed during update")
            pass
        except Exception as e:
            print(f"ERROR: ProgressDialog - error updating: {e}")
    
    def update_progress(self, progress: float, message: str = ""):
        """Alias for update_progress_bar for backwards compatibility."""
        self.update_progress_bar(progress, message)
    
    def hide_progress_bar(self):
        """Hide and destroy the progress dialog."""
        try:
            if self.dialog and self.dialog.winfo_exists():
                self.dialog.grab_release()
                self.dialog.destroy()
                self.dialog = None
                print("DEBUG: ProgressDialog hidden")
        except tk.TclError:
            print("DEBUG: ProgressDialog - already destroyed")
            pass
        except Exception as e:
            print(f"ERROR: ProgressDialog - error hiding: {e}")
    
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        try:
            if not self.dialog or not self.parent:
                return
            
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
            
            # Ensure dialog stays on screen
            screen_width = self.dialog.winfo_screenwidth()
            screen_height = self.dialog.winfo_screenheight()
            
            x = max(0, min(x, screen_width - dialog_width))
            y = max(0, min(y, screen_height - dialog_height))
            
            self.dialog.geometry(f"+{x}+{y}")
            
            print(f"DEBUG: ProgressDialog centered at {x},{y}")
            
        except tk.TclError:
            pass
        except Exception as e:
            print(f"ERROR: ProgressDialog - error centering: {e}")
    
    def is_showing(self) -> bool:
        """Check if the progress dialog is currently showing."""
        try:
            return self.dialog is not None and self.dialog.winfo_exists()
        except tk.TclError:
            return False
    
    def set_title(self, title: str):
        """Set the dialog title."""
        try:
            if self.dialog and self.dialog.winfo_exists():
                self.dialog.title(title)
                print(f"DEBUG: ProgressDialog title set to: {title}")
        except tk.TclError:
            pass
        except Exception as e:
            print(f"ERROR: ProgressDialog - error setting title: {e}")
    
    def set_message(self, message: str):
        """Set the status message."""
        try:
            if self.status_var:
                self.status_var.set(message)
                if self.dialog and self.dialog.winfo_exists():
                    self.dialog.update_idletasks()
                print(f"DEBUG: ProgressDialog message set to: {message}")
        except Exception as e:
            print(f"ERROR: ProgressDialog - error setting message: {e}")
    
    def reset(self):
        """Reset progress to 0."""
        try:
            self.update_progress_bar(0, "Starting...")
            print("DEBUG: ProgressDialog reset")
        except Exception as e:
            print(f"ERROR: ProgressDialog - error resetting: {e}")
    
    def __del__(self):
        """Cleanup when dialog is destroyed."""
        try:
            self.hide_progress_bar()
        except:
            pass
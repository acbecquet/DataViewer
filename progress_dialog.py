# progress_dialog.py
import tkinter as tk
from tkinter import ttk, Toplevel, Label, Button
from utils import get_resource_path

class ProgressDialog:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.progress_window = None

    def _center_window(self, window: tk.Toplevel, width: int = None, height: int = None) -> None:
        """Center the given window on the screen."""
        window.update_idletasks()
        w = width or window.winfo_width()
        h = height or window.winfo_height()
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        window.geometry(f"{w}x{h}+{x}+{y}")

    def show_progress_bar(self, message: str) -> None:
        """Display a progress bar in a separate window with a white font."""
        if self.progress_window is not None and self.progress_window.winfo_exists():
            return  # Prevent multiple instances

        # Create a new top-level window for the progress bar
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Progress")
        self.progress_window.geometry("400x100")
        self.progress_window.resizable(False, False)
        self.progress_window.configure(bg="#0504AA")

        # Center the progress window relative to the main application
        self._center_window(self.progress_window, 400, 100)

        # Add a label with white font for the message
        self.progress_label = tk.Label(
            self.progress_window,
            text=message,
            fg="white",
            bg="#0504AA",
            font=("Arial", 12)
        )
        self.progress_label.pack(pady=10)

        # Add the progress bar
        self.progress_bar = ttk.Progressbar(
            self.progress_window,
            orient='horizontal',
            mode='determinate',
            length=300
        )
        self.progress_bar.pack(pady=10)
        self.progress_bar['value'] = 0  # Initialize progress bar

        # Disable interactions with the main window while the progress window is active
        self.progress_window.transient(self.root)  # Make it a child window
        self.progress_window.grab_set()  # Prevent interactions with the main app

    def update_progress_bar(self, value: int) -> None:
        """Update the progress bar value."""
        if self.progress_window is not None and self.progress_bar.winfo_exists():
            self.progress_bar['value'] = value  # Update progress value
            self.progress_window.update_idletasks()  # Refresh the progress window

    def hide_progress_bar(self) -> None:
        """Destroy the progress bar window after completion."""
        if self.progress_window is not None and self.progress_window.winfo_exists():
            self.progress_window.grab_release()  # Release grab from the main app
            self.progress_window.destroy()  # Close the progress window
            self.progress_window = None  # Remove the reference to prevent stale references

# views/dialogs/startup_dialog.py
"""
views/dialogs/startup_dialog.py
Startup and new template dialogs.
This contains startup menu UI functionality from main_gui.py.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable, Dict, Any, Tuple


class StartupDialog:
    """Startup dialog for new file creation and welcome screen."""
    
    def __init__(self, parent: tk.Tk):
        """Initialize the startup dialog."""
        self.parent = parent
        self.dialog: Optional[tk.Toplevel] = None
        self.result: Optional[str] = None
        self.selected_option: Optional[str] = None
        
        # Configuration
        self.app_background_color = '#EFEFEF'
        self.button_color = '#E0E0E0'
        self.font = ("Arial", 12)
        
        # Callbacks
        self.on_new_template: Optional[Callable] = None
        self.on_load_file: Optional[Callable] = None
        self.on_load_vap3: Optional[Callable] = None
        
        print("DEBUG: StartupDialog initialized")
    
    def show_welcome_dialog(self) -> Tuple[str, str]:
        """Show the main welcome/startup dialog."""
        try:
            # Create dialog
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("Welcome")
            self.dialog.geometry("350x200")
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            self.dialog.configure(bg=self.app_background_color)
            
            # Center dialog
            self._center_dialog()
            
            # Create layout
            self._create_welcome_layout()
            
            # Setup cleanup
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
            
            # Wait for dialog
            self.dialog.wait_window()
            
            result = self.result or "cancel"
            option = self.selected_option or ""
            
            print(f"DEBUG: StartupDialog - welcome dialog result: {result}")
            return result, option
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error showing welcome dialog: {e}")
            return "cancel", ""
    
    def show_new_template_dialog(self) -> Tuple[str, str]:
        """Show dialog for creating new templates."""
        try:
            # Create dialog
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("Create New Template")
            self.dialog.geometry("400x250")
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            self.dialog.configure(bg=self.app_background_color)
            
            # Center dialog
            self._center_dialog()
            
            # Create layout
            self._create_template_layout()
            
            # Setup cleanup
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
            
            # Wait for dialog
            self.dialog.wait_window()
            
            result = self.result or "cancel"
            option = self.selected_option or ""
            
            print(f"DEBUG: StartupDialog - template dialog result: {result}")
            return result, option
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error showing template dialog: {e}")
            return "cancel", ""
    
    def show_test_selection_dialog(self, available_tests: list) -> Tuple[str, list]:
        """Show dialog for selecting tests for new template."""
        try:
            test_dialog = TestSelectionDialog(self.parent, available_tests)
            return test_dialog.show()
        except Exception as e:
            print(f"ERROR: StartupDialog - error showing test selection: {e}")
            return "cancel", []
    
    def _create_welcome_layout(self):
        """Create the welcome dialog layout."""
        try:
            if not self.dialog:
                return
            
            # Main frame
            main_frame = ttk.Frame(self.dialog)
            main_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            # Welcome label
            welcome_label = tk.Label(main_frame, 
                                   text="Welcome to DataViewer by SDR!", 
                                   font=("Arial", 14, "bold"),
                                   bg=self.app_background_color,
                                   fg="black")
            welcome_label.pack(pady=(0, 20))
            
            # Description
            desc_label = tk.Label(main_frame,
                                text="Choose an option to get started:",
                                font=self.font,
                                bg=self.app_background_color,
                                fg="black")
            desc_label.pack(pady=(0, 15))
            
            # Button frame
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            # New button
            new_button = ttk.Button(button_frame, text="New Template", 
                                  command=self._on_new_template, width=15)
            new_button.pack(pady=5)
            
            # Load Excel button
            load_button = ttk.Button(button_frame, text="Load Excel File", 
                                   command=self._on_load_excel, width=15)
            load_button.pack(pady=5)
            
            # Load VAP3 button
            vap3_button = ttk.Button(button_frame, text="Load VAP3 File", 
                                   command=self._on_load_vap3, width=15)
            vap3_button.pack(pady=5)
            
            print("DEBUG: StartupDialog - welcome layout created")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error creating welcome layout: {e}")
    
    def _create_template_layout(self):
        """Create the template creation dialog layout."""
        try:
            if not self.dialog:
                return
            
            # Main frame
            main_frame = ttk.Frame(self.dialog)
            main_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            # Title
            title_label = tk.Label(main_frame, 
                                 text="Create New Template", 
                                 font=("Arial", 16, "bold"),
                                 bg=self.app_background_color)
            title_label.pack(pady=(0, 15))
            
            # Description
            desc_text = (
                "Select the type of template you want to create.\n"
                "Templates provide a structured format for data collection."
            )
            desc_label = tk.Label(main_frame, text=desc_text, 
                                font=self.font, justify="center",
                                bg=self.app_background_color)
            desc_label.pack(pady=(0, 20))
            
            # Template options frame
            options_frame = ttk.LabelFrame(main_frame, text="Template Options", padding=10)
            options_frame.pack(fill="x", pady=(0, 20))
            
            # Standard template button
            standard_btn = ttk.Button(options_frame, 
                                    text="Standard Testing Template",
                                    command=lambda: self._on_template_selected("standard"),
                                    width=25)
            standard_btn.pack(pady=(0, 10))
            
            standard_desc = tk.Label(options_frame,
                                   text="Pre-configured template with common test types",
                                   font=("Arial", 9), fg="gray",
                                   bg=self.app_background_color)
            standard_desc.pack(pady=(0, 10))
            
            # Custom template button
            custom_btn = ttk.Button(options_frame, 
                                  text="Custom Template",
                                  command=lambda: self._on_template_selected("custom"),
                                  width=25)
            custom_btn.pack(pady=(0, 10))
            
            custom_desc = tk.Label(options_frame,
                                 text="Select specific tests to include in your template",
                                 font=("Arial", 9), fg="gray",
                                 bg=self.app_background_color)
            custom_desc.pack()
            
            # Button frame
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="right")
            
            print("DEBUG: StartupDialog - template layout created")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error creating template layout: {e}")
    
    # Event handlers
    def _on_new_template(self):
        """Handle new template button click."""
        try:
            self.result = "new_template"
            self.selected_option = "new"
            
            if self.dialog:
                self.dialog.destroy()
            
            print("DEBUG: StartupDialog - new template selected")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error in new template handler: {e}")
    
    def _on_load_excel(self):
        """Handle load Excel button click."""
        try:
            self.result = "load_excel"
            self.selected_option = "excel"
            
            if self.dialog:
                self.dialog.destroy()
            
            print("DEBUG: StartupDialog - load Excel selected")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error in load Excel handler: {e}")
    
    def _on_load_vap3(self):
        """Handle load VAP3 button click."""
        try:
            self.result = "load_vap3"
            self.selected_option = "vap3"
            
            if self.dialog:
                self.dialog.destroy()
            
            print("DEBUG: StartupDialog - load VAP3 selected")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error in load VAP3 handler: {e}")
    
    def _on_template_selected(self, template_type: str):
        """Handle template type selection."""
        try:
            self.result = "template_selected"
            self.selected_option = template_type
            
            if self.dialog:
                self.dialog.destroy()
            
            print(f"DEBUG: StartupDialog - template selected: {template_type}")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error in template selection: {e}")
    
    def _on_cancel(self):
        """Handle cancel button or window close."""
        try:
            self.result = "cancel"
            self.selected_option = ""
            
            if self.dialog:
                self.dialog.destroy()
            
            print("DEBUG: StartupDialog - cancelled")
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error in cancel handler: {e}")
    
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        try:
            if not self.dialog:
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
            
        except Exception as e:
            print(f"ERROR: StartupDialog - error centering dialog: {e}")


class TestSelectionDialog:
    """Dialog for selecting tests when creating custom templates."""
    
    def __init__(self, parent: tk.Widget, available_tests: list):
        """Initialize test selection dialog."""
        self.parent = parent
        self.available_tests = available_tests
        self.dialog: Optional[tk.Toplevel] = None
        
        # Variables
        self.test_vars: Dict[str, tk.BooleanVar] = {}
        self.selected_tests: list = []
        self.result = "cancel"
        
        # UI components
        self.canvas: Optional[tk.Canvas] = None
        self.scrollbar: Optional[tk.Scrollbar] = None
        
        print("DEBUG: TestSelectionDialog initialized")
    
    def show(self) -> Tuple[str, list]:
        """Show the test selection dialog."""
        try:
            # Create dialog
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("Select Tests")
            self.dialog.geometry("500x400")
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            
            # Center dialog
            self._center_dialog()
            
            # Create layout
            self._create_layout()
            
            # Setup cleanup
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
            
            # Wait for dialog
            self.dialog.wait_window()
            
            result = self.result
            selected = self.selected_tests.copy()
            
            print(f"DEBUG: TestSelectionDialog - result: {result}, selected: {len(selected)} tests")
            return result, selected
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error showing dialog: {e}")
            return "cancel", []
    
    def _create_layout(self):
        """Create the dialog layout."""
        try:
            if not self.dialog:
                return
            
            # Configure grid
            self.dialog.grid_columnconfigure(0, weight=1)
            self.dialog.grid_columnconfigure(1, weight=0)
            self.dialog.grid_rowconfigure(0, weight=0)
            self.dialog.grid_rowconfigure(1, weight=1)
            self.dialog.grid_rowconfigure(2, weight=0)
            
            # Header
            header_frame = ttk.Frame(self.dialog)
            header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
            
            header_label = ttk.Label(header_frame, 
                                   text="Select tests to include in your new template:",
                                   font=("Arial", 12, "bold"))
            header_label.pack(anchor="w")
            
            # Tests frame
            tests_frame = ttk.Frame(self.dialog)
            tests_frame.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=5)
            tests_frame.grid_columnconfigure(0, weight=1)
            tests_frame.grid_rowconfigure(0, weight=1)
            
            # Scrollable area
            self.canvas = tk.Canvas(tests_frame)
            self.scrollbar = ttk.Scrollbar(tests_frame, orient="vertical", command=self.canvas.yview)
            
            checkbox_frame = ttk.Frame(self.canvas)
            checkbox_frame.bind("<Configure>", 
                              lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
            
            # Create window in canvas
            canvas_window = self.canvas.create_window((0, 0), window=checkbox_frame, anchor="nw")
            
            def on_canvas_configure(event):
                self.canvas.itemconfig(canvas_window, width=event.width)
            
            self.canvas.bind("<Configure>", on_canvas_configure)
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            
            # Pack canvas and scrollbar
            self.canvas.grid(row=0, column=0, sticky="nsew")
            self.scrollbar.grid(row=0, column=1, sticky="ns")
            
            # Create test checkboxes
            self._create_test_checkboxes(checkbox_frame)
            
            # Control buttons
            control_frame = ttk.Frame(self.dialog)
            control_frame.grid(row=1, column=1, sticky="ns", padx=(5, 10), pady=5)
            
            ttk.Button(control_frame, text="Select All", command=self._select_all).pack(fill="x", pady=(0, 5))
            ttk.Button(control_frame, text="Deselect All", command=self._deselect_all).pack(fill="x", pady=(0, 5))
            
            # Bottom buttons
            button_frame = ttk.Frame(self.dialog)
            button_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
            
            ttk.Button(button_frame, text="Create Template", command=self._on_create).pack(side="right", padx=(5, 0))
            ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="right")
            
            # Selection info
            self.info_label = ttk.Label(button_frame, text="No tests selected")
            self.info_label.pack(side="left")
            
            self._update_selection_info()
            
            print("DEBUG: TestSelectionDialog - layout created")
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error creating layout: {e}")
    
    def _create_test_checkboxes(self, parent: ttk.Frame):
        """Create checkboxes for each test."""
        try:
            for i, test_name in enumerate(self.available_tests):
                # Create variable
                var = tk.BooleanVar()
                self.test_vars[test_name] = var
                
                # Create checkbox
                checkbox = ttk.Checkbutton(parent, text=test_name, variable=var,
                                         command=self._update_selection_info)
                checkbox.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            
            print(f"DEBUG: TestSelectionDialog - created {len(self.available_tests)} test checkboxes")
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error creating test checkboxes: {e}")
    
    def _select_all(self):
        """Select all tests."""
        try:
            for var in self.test_vars.values():
                var.set(True)
            self._update_selection_info()
            print("DEBUG: TestSelectionDialog - all tests selected")
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error selecting all: {e}")
    
    def _deselect_all(self):
        """Deselect all tests."""
        try:
            for var in self.test_vars.values():
                var.set(False)
            self._update_selection_info()
            print("DEBUG: TestSelectionDialog - all tests deselected")
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error deselecting all: {e}")
    
    def _update_selection_info(self):
        """Update selection information."""
        try:
            selected_count = sum(1 for var in self.test_vars.values() if var.get())
            total_count = len(self.test_vars)
            
            if hasattr(self, 'info_label') and self.info_label:
                self.info_label.config(text=f"{selected_count} of {total_count} tests selected")
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error updating selection info: {e}")
    
    def _on_create(self):
        """Handle create template button."""
        try:
            # Get selected tests
            self.selected_tests = [test for test, var in self.test_vars.items() if var.get()]
            
            if not self.selected_tests:
                messagebox.showwarning("No Selection", "Please select at least one test.")
                return
            
            self.result = "create"
            
            if self.dialog:
                self.dialog.destroy()
            
            print(f"DEBUG: TestSelectionDialog - creating template with {len(self.selected_tests)} tests")
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error in create handler: {e}")
    
    def _on_cancel(self):
        """Handle cancel button."""
        try:
            self.result = "cancel"
            self.selected_tests = []
            
            if self.dialog:
                self.dialog.destroy()
            
            print("DEBUG: TestSelectionDialog - cancelled")
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error in cancel handler: {e}")
    
    def _center_dialog(self):
        """Center the dialog on parent."""
        try:
            if not self.dialog:
                return
            
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
            
        except Exception as e:
            print(f"ERROR: TestSelectionDialog - error centering: {e}")
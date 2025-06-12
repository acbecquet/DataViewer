"""
header_data_dialog.py
Developed by Charlie Becquet.
Dialog for entering test header data.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from utils import APP_BACKGROUND_COLOR, FONT

class HeaderDataDialog:
    def __init__(self, parent, file_path, selected_test):
        """
        Initialize the header data dialog.
        
        Args:
            parent (tk.Tk): The parent window.
            file_path (str): Path to the created file.
            selected_test (str): The selected test name.
        """
        self.parent = parent
        self.file_path = file_path
        self.selected_test = selected_test
        self.result = None
        self.header_data = {}
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Header Data - {selected_test}")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=APP_BACKGROUND_COLOR)
        self.dialog.geometry("600x500")  # Initial size
        
        self.create_widgets()
        self.center_window()
    
    def center_window(self):
        """Center the dialog on the screen."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """Create the dialog widgets."""
        # Main container with scrolling capability
        main_container = ttk.Frame(self.dialog)
        main_container.pack(fill="both", expand=True)
        
        # Create a canvas with scrollbar
        canvas = tk.Canvas(main_container, bg=APP_BACKGROUND_COLOR)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Frame to hold the form
        self.form_frame = ttk.Frame(canvas, padding=20)
        
        # Add the form frame to the canvas
        canvas_frame = canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        
        # Configure canvas scrolling
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_frame, width=event.width)
        
        self.form_frame.bind("<Configure>", configure_canvas)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_frame, width=e.width))
        
        # Add mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Header
        header_label = ttk.Label(
            self.form_frame,
            text=f"Enter header data for {self.selected_test}",
            font=("Arial", 14),
            foreground="white",
            background=APP_BACKGROUND_COLOR
        )
        header_label.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 20))
        
        # Tester
        ttk.Label(
            self.form_frame, 
            text="Tester:", 
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).grid(row=1, column=0, sticky="w", pady=5)
        
        self.tester_var = tk.StringVar()
        ttk.Entry(self.form_frame, textvariable=self.tester_var, width=30).grid(
            row=1, column=1, sticky="ew", padx=(10, 0), pady=5
        )
        
        # Number of Samples
        ttk.Label(
            self.form_frame, 
            text="Number of Samples:", 
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).grid(row=2, column=0, sticky="w", pady=5)
        
        self.num_samples_var = tk.IntVar(value=1)
        samples_spinbox = ttk.Spinbox(
            self.form_frame, 
            from_=1, 
            to=10, 
            textvariable=self.num_samples_var, 
            width=5,
            command=self.update_sample_fields
        )
        samples_spinbox.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=5)
        
        # Create a frame for sample-specific data
        self.samples_frame = ttk.Frame(self.form_frame)
        self.samples_frame.grid(row=3, column=0, columnspan=4, sticky="ew", pady=10)
        
        # Media
        ttk.Label(
            self.form_frame, 
            text="Media:", 
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).grid(row=4, column=0, sticky="w", pady=5)
        
        self.media_var = tk.StringVar()
        ttk.Entry(self.form_frame, textvariable=self.media_var, width=30).grid(
            row=4, column=1, sticky="ew", padx=(10, 0), pady=5
        )
        
        # Viscosity
        ttk.Label(
            self.form_frame, 
            text="Viscosity (cP):", 
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).grid(row=5, column=0, sticky="w", pady=5)
        
        self.viscosity_var = tk.StringVar()
        ttk.Entry(self.form_frame, textvariable=self.viscosity_var, width=30).grid(
            row=5, column=1, sticky="ew", padx=(10, 0), pady=5
        )
        
        # Voltage
        ttk.Label(
            self.form_frame, 
            text="Voltage (V):", 
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).grid(row=6, column=0, sticky="w", pady=5)
        
        self.voltage_var = tk.StringVar()
        ttk.Entry(self.form_frame, textvariable=self.voltage_var, width=30).grid(
            row=6, column=1, sticky="ew", padx=(10, 0), pady=5
        )
        
        # Initial Oil Mass
        ttk.Label(
            self.form_frame, 
            text="Initial Oil Mass (g):", 
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).grid(row=7, column=0, sticky="w", pady=5)
        
        self.oil_mass_var = tk.StringVar()
        ttk.Entry(self.form_frame, textvariable=self.oil_mass_var, width=30).grid(
            row=7, column=1, sticky="ew", padx=(10, 0), pady=5
        )
        
        # Configure grid
        self.form_frame.columnconfigure(1, weight=1)
        self.form_frame.columnconfigure(3, weight=1)
        
        # Initialize sample fields
        self.sample_id_vars = []
        self.resistance_vars = []
        self.update_sample_fields()
        
        # Buttons
        button_frame = ttk.Frame(self.form_frame)
        button_frame.grid(row=8, column=0, columnspan=4, sticky="ew", pady=(20, 0))
        
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side="right")
        
        ttk.Button(button_frame, text="Continue", command=self.on_continue).pack(side="right", padx=(0, 5))
    
    def update_sample_fields(self):
        """Update the sample-specific fields based on the number of samples."""
        # Clear existing widgets
        for widget in self.samples_frame.winfo_children():
            widget.destroy()
        
        # Get the number of samples
        num_samples = self.num_samples_var.get()
        
        # Ensure we have enough variables
        while len(self.sample_id_vars) < num_samples:
            self.sample_id_vars.append(tk.StringVar())
            self.resistance_vars.append(tk.StringVar())
        
        # Header for samples section
        sample_header = ttk.Label(
            self.samples_frame,
            text="Sample-specific information:",
            font=("Arial", 11, "bold"),
            foreground="white",
            background=APP_BACKGROUND_COLOR
        )
        sample_header.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        
        # Create fields for each sample
        for i in range(num_samples):
            # Sample ID
            ttk.Label(
                self.samples_frame, 
                text=f"Sample {i+1} ID:", 
                foreground="white",
                background=APP_BACKGROUND_COLOR
            ).grid(row=i+1, column=0, sticky="w", pady=5)
            
            ttk.Entry(self.samples_frame, textvariable=self.sample_id_vars[i], width=30).grid(
                row=i+1, column=1, sticky="ew", padx=(10, 0), pady=5
            )
            
            # Resistance
            ttk.Label(
                self.samples_frame, 
                text=f"Resistance {i+1} (Ω):", 
                foreground="white",
                background=APP_BACKGROUND_COLOR
            ).grid(row=i+1, column=2, sticky="w", pady=5, padx=(20, 0))
            
            ttk.Entry(self.samples_frame, textvariable=self.resistance_vars[i], width=10).grid(
                row=i+1, column=3, sticky="w", padx=(10, 0), pady=5
            )
        
        # Configure grid
        self.samples_frame.columnconfigure(1, weight=1)
        self.samples_frame.columnconfigure(3, weight=1)
    
    def validate_data(self):
        """Validate the header data."""
        # Check for empty required fields
        if not self.tester_var.get().strip():
            messagebox.showerror("Validation Error", "Tester name is required.")
            return False
            
        # Validate numeric fields
        try:
            if self.viscosity_var.get().strip():
                float(self.viscosity_var.get().strip())
            if self.voltage_var.get().strip():
                float(self.voltage_var.get().strip())
            if self.oil_mass_var.get().strip():
                float(self.oil_mass_var.get().strip())
        except ValueError:
            messagebox.showerror("Validation Error", "Viscosity, Voltage, and Oil Mass must be numeric values.")
            return False
            
        # Check sample-specific fields
        num_samples = self.num_samples_var.get()
        for i in range(num_samples):
            if not self.sample_id_vars[i].get().strip():
                messagebox.showerror("Validation Error", f"Sample {i+1} ID is required.")
                return False
                
            # Validate resistance
            try:
                if self.resistance_vars[i].get().strip():
                    float(self.resistance_vars[i].get().strip())
            except ValueError:
                messagebox.showerror("Validation Error", f"Resistance for Sample {i+1} must be a numeric value.")
                return False
                
        return True
    
    def collect_header_data(self):
        """Collect header data from form fields."""
        num_samples = self.num_samples_var.get()
        
        # Common data
        common_data = {
            "tester": self.tester_var.get().strip(),
            "media": self.media_var.get().strip(),
            "viscosity": self.viscosity_var.get().strip(),
            "voltage": self.voltage_var.get().strip(),
            "oil_mass": self.oil_mass_var.get().strip()
        }
        
        # Sample-specific data
        samples = []
        for i in range(num_samples):
            sample = {
                "id": self.sample_id_vars[i].get().strip(),
                "resistance": self.resistance_vars[i].get().strip()
            }
            samples.append(sample)
        
        return {
            "common": common_data,
            "samples": samples,
            "test": self.selected_test,
            "num_samples": num_samples
        }
    
    def on_continue(self):
        """Handle Continue button click."""
        if self.validate_data():
            self.header_data = self.collect_header_data()
            self.result = True
            self.dialog.destroy()
    
    def on_cancel(self):
        """Handle Cancel button click."""
        self.result = False
        self.dialog.destroy()
    
    def show(self):
        """
        Show the dialog and wait for user input.
        
        Returns:
            tuple: (result, header_data) where result is True if Continue was clicked,
                   and header_data is a dictionary of header data.
        """
        self.dialog.wait_window()
        return self.result, self.header_data
"""
viscosity_gui.py
A standalone GUI for viscosity calculations and analysis.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import json
from pathlib import Path

# Import the viscosity calculator functionality
from viscosity_calculator import ViscosityCalculator

# Define constants (copied from utils.py to avoid dependency)
FONT = ('Arial', 12)
APP_BACKGROUND_COLOR = '#0504AA'
BUTTON_COLOR = '#4169E1'

class ViscosityGUI:
    """A standalone GUI for viscosity calculations and analysis."""
    
    def __init__(self):
        """Initialize the Viscosity GUI with its own root window."""
        # Create the main window
        self.root = tk.Tk()
        self.root.title("Viscosity Calculator")
        self.root.geometry("550x550")
        self.root.configure(bg=APP_BACKGROUND_COLOR)
        
        # Set application icon if available
        try:
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                application_path = os.path.dirname(sys.executable)
            else:
                # Running as script
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(application_path, 'resources', 'ccell_icon.png')
            if os.path.exists(icon_path):
                self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
        except Exception as e:
            print(f"Could not load icon: {e}")
        
        # Create an adapter class to provide the same interface expected by ViscosityCalculator
        self.gui_adapter = self.GUIAdapter(self.root)
        
        # Initialize the viscosity calculator with our adapter
        self.calculator = ViscosityCalculator(self.gui_adapter)
        
        # Create the main application menu
        self.create_menu()
        
        # Show the calculator window embedded in our main window
        self.show_calculator()
        
        # Ensure directories exist
        self.ensure_directories()
        
        # Center the window on screen
        self.center_window(self.root)
    
    class GUIAdapter:
        """Adapter class to provide the interface expected by ViscosityCalculator."""
        
        def __init__(self, root):
            self.root = root
            # The original code uses self.gui.root, so we need this adapter
        
        def center_window(self, window, width=None, height=None):
            """Center a window on screen."""
            window.update_idletasks()
            w = width or window.winfo_width()
            h = height or window.winfo_height()
            x = (window.winfo_screenwidth() - w) // 2
            y = (window.winfo_screenheight() - h) // 2
            window.geometry(f"{w}x{h}+{x}+{y}")
    
    def ensure_directories(self):
        """Ensure necessary directories exist."""
        # Create data directory for CSV files
        Path("data").mkdir(exist_ok=True)
        # Create models directory for saving trained models
        Path("models").mkdir(exist_ok=True)
        # Create plots directory for generated plots
        Path("plots").mkdir(exist_ok=True)
    
    def create_menu(self):
        """Create the application menu."""
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Load Data", command=self.calculator.upload_training_data)
        file_menu.add_command(label="Save", command=self.save_database)
        file_menu.add_separator()
        # In the ViscosityGUI.create_menu method, add this line after "Save" but before the separator
        file_menu.add_command(label="Manage Data", command=self.calculator.view_formulation_data)
        file_menu.add_command(label= "Export Terpene Profiles", command = self.calculator.export_terpene_profiles)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Models menu
        models_menu = tk.Menu(menubar, tearoff=0)
        models_menu.add_command(label="Train Models (Unified)", command=self.calculator.train_unified_models)
        models_menu.add_command(label="Analyze Models", command=self.calculator.analyze_models)
        models_menu.add_command(label="Analyze Chemical Properties Impact", command=self.calculator.analyze_chemical_importance)
        models_menu.add_command(label="Arrhenius Analysis", 
                               command=self.calculator.filter_and_analyze_specific_combinations)
        models_menu.add_separator()
        models_menu.add_command(label="Diagnose Issues", command = self.calculator.diagnose_models)
        models_menu.add_command(label="Create Potency Demo Model", 
                               command=self.calculator.create_potency_demo_model)
        models_menu.add_command(label="Analyze Model Response", 
                               command=self.calculator.analyze_model_feature_response)
        menubar.add_cascade(label="Models", menu=models_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Help", command=self.show_help)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def save_database(self):
        """Save the terpene formulation database."""
        try:
            self.calculator.save_formulation_database()
            messagebox.showinfo("Success", "Database saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save database: {str(e)}")
    
    def show_calculator(self):
        """Show the calculator within our main window."""
        # Create a frame to hold the calculator
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # The calculator window would normally be a toplevel window
        # Instead, we'll embed it in our frame
        notebook = ttk.Notebook(frame)
        notebook.pack(fill='both', expand=True)
        
        # Create the calculator tabs by calling the calculator's methods
        self.calculator.create_calculator_tab(notebook)
        self.calculator.create_advanced_tab(notebook)
        self.calculator.create_measure_tab(notebook)
        
        # Bind tab change event to update advanced tab
        notebook.bind("<<NotebookTabChanged>>", lambda e: self.calculator.update_advanced_tab_fields())
    
    def show_about(self):
        """Show the about dialog."""
        messagebox.showinfo("About", "Viscosity Calculator\nVersion 1.0\nDeveloped by Charlie Becquet")
    
    def show_help(self):
        """Show the help dialog."""
        help_text = (
            "Viscosity Calculator Help\n\n"
            "This application helps you calculate terpene percentages based on viscosity.\n\n"
            "Calculator Tab: Calculate terpene percentage using mathematical models.\n"
            "Iterative Method Tab: Use a two-step iterative process for finding optimal terpene percentage.\n"
            "Measure Tab: Record viscosity measurements at different temperatures.\n\n"
            "Use the Models menu to train machine learning models on your data."
        )
        messagebox.showinfo("Help", help_text)
    
    def center_window(self, window, width=None, height=None):
        """Center a window on the screen."""
        window.update_idletasks()
        w = width or window.winfo_width()
        h = height or window.winfo_height()
        x = (window.winfo_screenwidth() - w) // 2
        y = (window.winfo_screenheight() - h) // 2
        window.geometry(f"{w}x{h}+{x}+{y}")
    
    def run(self):
        """Run the application."""
        self.root.mainloop()

def main():
    """Main entry point for the application."""
    app = ViscosityGUI()
    app.run()

if __name__ == "__main__":
    main()
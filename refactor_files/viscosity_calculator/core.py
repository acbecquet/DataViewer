"""
viscosity_calculator/core.py
Core module containing the main ViscosityCalculator class definition.
"""
import os
import json
import tkinter as tk
from tkinter import StringVar, ttk, DoubleVar, IntVar, messagebox

FONT = ('Arial', 12)
APP_BACKGROUND_COLOR = '#0504AA'
BUTTON_COLOR = '#4169E1'

class ViscosityCalculator:
    def __init__(self, gui):
        """
        Initialize the ViscosityCalculator with a reference to the main GUI.
        
        Args:
            gui (TestingGUI): The main application GUI instance
        """
        self.gui = gui
        self.root = gui.root
        
        # Initialize variables for input fields
        self.media_var = StringVar()
        self.media_brand_var = StringVar()
        self.terpene_var = StringVar()
        self.terpene_brand_var = StringVar()
        self.mass_of_oil_var = DoubleVar(value=0.0)
        self.target_viscosity_var = DoubleVar(value=0.0)
        
        # Store these internally but don't show in UI by default
        self._total_potency_var = DoubleVar(value=0.0)
        self._d9_thc_var = DoubleVar(value=0.0)
        self._d8_thc_var = DoubleVar(value=0.0)

        # Initialize variables for result fields
        self.exact_percent_var = StringVar(value="0.0%")
        self.exact_mass_var = StringVar(value="0.0g")
        self.start_percent_var = StringVar(value="0.0%")
        self.start_mass_var = StringVar(value="0.0g")
        
        # Initialize variables for advanced tab
        self.step1_amount_var = StringVar(value="0.0g")
        self.step1_viscosity_var = StringVar(value="")
        self.step2_amount_var = StringVar(value="0.0g")
        self.step2_viscosity_var = StringVar(value="")
        self.expected_viscosity_var = StringVar(value="0.0")

        self.initialize_default_terpene_profiles()
        
        # Define media options
        self.media_options = ["D8", "D9", "Resin", "Rosin", "Liquid Diamonds", "Other"]
        
        # Load or initialize formulation database
        self.formulation_db_path = "terpene_formulations.json"
        self._formulation_db = None  # Lazy load this
        
        # Store notebook reference
        self.notebook = None

        # Lazy load models
        self._consolidated = None
    
    # ---- Database access properties with lazy loading ----
    
    @property
    def formulation_db(self):
        """Lazy load the formulation database"""
        if self._formulation_db is None:
            self._formulation_db = self.load_formulation_database()
        return self._formulation_db

    @property
    def consolidated_models(self):
        """Lazy load the consolidated models"""
        if not hasattr(self, '_consolidated_models') or self._consolidated_models is None:
            self._consolidated_models = {}
            self.load_consolidated_models()
        return self._consolidated_models

    @consolidated_models.setter
    def consolidated_models(self, value):
        """Setter for consolidated_models"""
        self._consolidated_models = value

    @formulation_db.setter
    def formulation_db(self, value):
        self._formulation_db = value
        
    def show_calculator(self):
        """
        Show the viscosity calculator window with tabbed interface.
        
        Returns:
            tk.Toplevel: The calculator window
        """
        # Create a new top-level window
        calculator_window = tk.Toplevel(self.root)
        calculator_window.title("Calculate Terpene % for Viscosity")
        calculator_window.geometry("550x500")
        calculator_window.resizable(False, False)
        calculator_window.configure(bg='#0504AA')
        
        # Make the window modal
        calculator_window.transient(self.root)
        calculator_window.grab_set()
        
        # Center the window on the screen
        self.gui.center_window(calculator_window, 550, 500)
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(calculator_window)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create Tab 1: Calculator
        self.create_calculator_tab(self.notebook)
        
        # Create Tab 2: Iterative Method
        self.create_advanced_tab(self.notebook)

        # Create Tab 3: Measure
        self.create_measure_tab(self.notebook)
        
        # Bind tab change event to update advanced tab
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.update_advanced_tab_fields())
        
        return calculator_window
    
    def load_formulation_database(self):
        """
        Load the terpene formulation database from a JSON file.
        If the file doesn't exist, return an empty database.
        """
        if os.path.exists(self.formulation_db_path):
            try:
                with open(self.formulation_db_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading formulation database: {e}")
                return {}
        else:
            return {}
    
    def save_formulation_database(self):
        """
        Save the terpene formulation database to a JSON file.
        """
        try:
            with open(self.formulation_db_path, 'w') as f:
                json.dump(self.formulation_db, f, indent=4)
        except Exception as e:
            print(f"Error saving formulation database: {e}")
    
    def get_raw_oil_viscosity(self, media_type):
        """
        Get the estimated raw oil viscosity based on media type.
        These are placeholder values - replace with actual data.
        """
        
        viscosity_map = {
            "D8": 54346667,
            "D9": 20000000,
            "Liquid Diamonds": 20000000,
            "Other": 2000000 
        }
        
        return viscosity_map.get(media_type, 18000000)
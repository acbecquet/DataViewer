"""
controllers/calculation_controller.py
Calculation operations controller that coordinates calculation operations.
This handles viscosity calculations, model training, and analysis operations.
"""

import os
import json
import tkinter as tk
from tkinter import StringVar, ttk, DoubleVar, IntVar, messagebox, filedialog, Frame, Label, Entry, Button
from typing import Optional, Dict, Any, List, Tuple
import pandas as pd
import numpy as np
from models.data_model import DataModel

# Font and color constants
FONT = ('Arial', 12)
APP_BACKGROUND_COLOR = '#0504AA'
BUTTON_COLOR = '#4169E1'


class CalculationController:
    """Controller for calculation operations like viscosity calculations."""
    
    def __init__(self, data_model: DataModel):
        """Initialize the calculation controller."""
        self.data_model = data_model
        
        # Viscosity calculator variables
        self.media_var = StringVar()
        self.media_brand_var = StringVar()
        self.terpene_var = StringVar()
        self.terpene_brand_var = StringVar()
        self.mass_of_oil_var = DoubleVar(value=0.0)
        self.target_viscosity_var = DoubleVar(value=0.0)
        
        # Internal variables (not shown in UI by default)
        self._total_potency_var = DoubleVar(value=0.0)
        self._d9_thc_var = DoubleVar(value=0.0)
        self._d8_thc_var = DoubleVar(value=0.0)

        # Result variables
        self.exact_percent_var = StringVar(value="0.0%")
        self.exact_mass_var = StringVar(value="0.0g")
        self.start_percent_var = StringVar(value="0.0%")
        self.start_mass_var = StringVar(value="0.0g")
        
        # Advanced tab variables
        self.step1_amount_var = StringVar(value="0.0g")
        self.step1_viscosity_var = StringVar(value="")
        self.step2_amount_var = StringVar(value="0.0g")
        self.step2_viscosity_var = StringVar(value="")
        self.expected_viscosity_var = StringVar(value="0.0")

        # Media options
        self.media_options = ["D8", "D9", "Resin", "Rosin", "Liquid Diamonds", "Other"]
        
        # Database paths
        self.formulation_db_path = "terpene_formulations.json"
        self._formulation_db = None
        
        # Model storage
        self._consolidated_models = None
        self.notebook = None
        
        # Temperature measurement variables
        self.temperature_blocks = []
        self.speed_vars = []
        self.torque_vars = [[], [], []]  # Three runs
        self.viscosity_vars = [[], [], []]  # Three runs
        self.avg_torque_vars = []
        self.avg_visc_vars = []
        
        # Terpene profiles
        self.terpene_profiles = {}
        
        print("DEBUG: CalculationController initialized")
        print(f"DEBUG: Connected to DataModel")
        
        # Initialize default profiles and databases
        self.initialize_default_terpene_profiles()
        self.load_formulation_database()
    
    def initialize_default_terpene_profiles(self):
        """Initialize default terpene profiles."""
        print("DEBUG: CalculationController initializing default terpene profiles")
        
        self.terpene_profiles = {
            "Limonene": {
                "primary": "Limonene",
                "percentage": 100.0,
                "notes": "Pure limonene profile"
            },
            "Myrcene": {
                "primary": "Myrcene", 
                "percentage": 100.0,
                "notes": "Pure myrcene profile"
            },
            "Pinene": {
                "primary": "Pinene",
                "percentage": 100.0,
                "notes": "Pure pinene profile"
            },
            "Linalool": {
                "primary": "Linalool",
                "percentage": 100.0,
                "notes": "Pure linalool profile"
            },
            "Citrus Blend": {
                "primary": "Limonene",
                "secondary": "Pinene",
                "primary_percentage": 70.0,
                "secondary_percentage": 30.0,
                "notes": "Citrus-focused blend"
            },
            "Relaxing Blend": {
                "primary": "Myrcene",
                "secondary": "Linalool", 
                "primary_percentage": 60.0,
                "secondary_percentage": 40.0,
                "notes": "Relaxation-focused blend"
            }
        }
        
        print(f"DEBUG: Initialized {len(self.terpene_profiles)} default terpene profiles")
    
    def load_formulation_database(self):
        """Load or create the formulation database."""
        print("DEBUG: CalculationController loading formulation database")
        
        try:
            if os.path.exists(self.formulation_db_path):
                with open(self.formulation_db_path, 'r') as f:
                    self._formulation_db = json.load(f)
                print(f"DEBUG: Loaded existing formulation database with {len(self._formulation_db)} entries")
            else:
                self._formulation_db = {}
                print("DEBUG: Created new formulation database")
                
        except Exception as e:
            print(f"ERROR: Failed to load formulation database: {e}")
            self._formulation_db = {}
    
    def save_formulation_database(self):
        """Save the formulation database to file."""
        print("DEBUG: CalculationController saving formulation database")
        
        try:
            with open(self.formulation_db_path, 'w') as f:
                json.dump(self._formulation_db, f, indent=2)
            print("DEBUG: Formulation database saved successfully")
            
        except Exception as e:
            print(f"ERROR: Failed to save formulation database: {e}")
    
    def open_viscosity_calculator_window(self, parent_window=None) -> bool:
        """Open the viscosity calculator window."""
        print("DEBUG: CalculationController opening viscosity calculator window")
        
        try:
            # Create new window
            calculator_window = tk.Toplevel(parent_window) if parent_window else tk.Tk()
            calculator_window.title("Viscosity Calculator")
            calculator_window.geometry("550x500")
            calculator_window.resizable(True, True)
            calculator_window.configure(bg=APP_BACKGROUND_COLOR)
            
            if parent_window:
                calculator_window.transient(parent_window)
                calculator_window.grab_set()
            
            # Create tabbed interface
            self.notebook = ttk.Notebook(calculator_window)
            self.notebook.pack(fill='both', expand=True, padx=8, pady=8)
            
            # Create tabs
            self.create_basic_calculator_tab()
            self.create_advanced_calculator_tab()
            self.create_temperature_measurement_tab()
            self.create_terpene_profile_tab()
            
            print("DEBUG: CalculationController successfully opened viscosity calculator")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to open viscosity calculator: {e}")
            return False
    
    def create_basic_calculator_tab(self):
        """Create the basic viscosity calculator tab."""
        print("DEBUG: Creating basic calculator tab")
        
        # Create tab frame
        basic_frame = ttk.Frame(self.notebook)
        self.notebook.add(basic_frame, text="Basic Calculator")
        
        # Create main container with background
        main_container = Frame(basic_frame, bg=APP_BACKGROUND_COLOR)
        main_container.pack(fill='both', expand=True)
        
        # Create form frame
        form_frame = Frame(main_container, bg=APP_BACKGROUND_COLOR)
        form_frame.pack(pady=20, padx=20)
        
        # Configure grid weights
        form_frame.grid_columnconfigure(1, weight=1)
        
        # Media dropdown
        Label(form_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=0, column=0, sticky="w", pady=5)
        
        media_combo = ttk.Combobox(form_frame, textvariable=self.media_var, 
                                  values=self.media_options, state="readonly", width=15)
        media_combo.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        # Media brand
        Label(form_frame, text="Media Brand:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=1, column=0, sticky="w", pady=5)
        
        brand_entry = Entry(form_frame, textvariable=self.media_brand_var, width=15)
        brand_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Terpene dropdown
        Label(form_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=2, column=0, sticky="w", pady=5)
        
        terpene_options = list(self.terpene_profiles.keys())
        terpene_combo = ttk.Combobox(form_frame, textvariable=self.terpene_var,
                                    values=terpene_options, state="readonly", width=15)
        terpene_combo.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Mass of oil
        Label(form_frame, text="Mass of Oil (g):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=3, column=0, sticky="w", pady=5)
        
        mass_entry = Entry(form_frame, textvariable=self.mass_of_oil_var, width=15)
        mass_entry.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        # Target viscosity
        Label(form_frame, text="Target Viscosity:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=4, column=0, sticky="w", pady=5)
        
        viscosity_entry = Entry(form_frame, textvariable=self.target_viscosity_var, width=15)
        viscosity_entry.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        # Calculate button
        calc_button = Button(form_frame, text="Calculate", command=self.calculate_viscosity,
                            bg=BUTTON_COLOR, fg="white", font=FONT)
        calc_button.grid(row=5, column=0, columnspan=2, pady=20)
        
        # Results section
        ttk.Separator(form_frame, orient='horizontal').grid(row=6, column=0, columnspan=4, sticky="ew", pady=15)
        
        Label(form_frame, text="Results:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=(FONT[0], FONT[1], "bold"), anchor="w").grid(row=7, column=0, sticky="w", pady=5)
        
        # Exact results
        Label(form_frame, text="Exact %:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=8, column=0, sticky="w", pady=3)
        
        Label(form_frame, textvariable=self.exact_percent_var, 
              bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT).grid(row=8, column=1, sticky="w", pady=3)
        
        Label(form_frame, text="Exact Mass:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=9, column=0, sticky="w", pady=3)
        
        Label(form_frame, textvariable=self.exact_mass_var, 
              bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT).grid(row=9, column=1, sticky="w", pady=3)
    
    def create_advanced_calculator_tab(self):
        """Create the advanced calculator tab."""
        print("DEBUG: Creating advanced calculator tab")
        
        # Create tab frame
        advanced_frame = ttk.Frame(self.notebook)
        self.notebook.add(advanced_frame, text="Advanced")
        
        # Create main container
        main_container = Frame(advanced_frame, bg=APP_BACKGROUND_COLOR)
        main_container.pack(fill='both', expand=True)
        
        # Title
        Label(main_container, text="Step-by-Step Viscosity Adjustment", 
              bg=APP_BACKGROUND_COLOR, fg="white", font=(FONT[0], FONT[1]+2, "bold")).pack(pady=10)
        
        # Step 1 frame
        step1_frame = Frame(main_container, bg=APP_BACKGROUND_COLOR, relief="raised", bd=1)
        step1_frame.pack(fill="x", padx=20, pady=10)
        
        Label(step1_frame, text="Step 1: Initial Addition", bg=APP_BACKGROUND_COLOR, fg="white",
              font=(FONT[0], FONT[1], "bold")).pack(pady=5)
        
        step1_form = Frame(step1_frame, bg=APP_BACKGROUND_COLOR)
        step1_form.pack(pady=5)
        
        Label(step1_form, text="Amount to add:", bg=APP_BACKGROUND_COLOR, fg="white").grid(row=0, column=0, padx=5)
        Label(step1_form, textvariable=self.step1_amount_var, bg=APP_BACKGROUND_COLOR, fg="#00b539").grid(row=0, column=1, padx=5)
        
        Label(step1_form, text="Measured viscosity:", bg=APP_BACKGROUND_COLOR, fg="white").grid(row=1, column=0, padx=5, pady=5)
        Entry(step1_form, textvariable=self.step1_viscosity_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        
        # Step 2 frame
        step2_frame = Frame(main_container, bg=APP_BACKGROUND_COLOR, relief="raised", bd=1)
        step2_frame.pack(fill="x", padx=20, pady=10)
        
        Label(step2_frame, text="Step 2: Fine Adjustment", bg=APP_BACKGROUND_COLOR, fg="white",
              font=(FONT[0], FONT[1], "bold")).pack(pady=5)
        
        step2_form = Frame(step2_frame, bg=APP_BACKGROUND_COLOR)
        step2_form.pack(pady=5)
        
        Label(step2_form, text="Additional amount:", bg=APP_BACKGROUND_COLOR, fg="white").grid(row=0, column=0, padx=5)
        Label(step2_form, textvariable=self.step2_amount_var, bg=APP_BACKGROUND_COLOR, fg="#00b539").grid(row=0, column=1, padx=5)
        
        Label(step2_form, text="Expected final viscosity:", bg=APP_BACKGROUND_COLOR, fg="white").grid(row=1, column=0, padx=5, pady=5)
        Label(step2_form, textvariable=self.expected_viscosity_var, bg=APP_BACKGROUND_COLOR, fg="#00b539").grid(row=1, column=1, padx=5, pady=5)
        
        # Calculate button
        Button(main_container, text="Calculate Steps", command=self.calculate_advanced_steps,
               bg=BUTTON_COLOR, fg="white", font=FONT).pack(pady=20)
    
    def create_temperature_measurement_tab(self):
        """Create the temperature measurement tab."""
        print("DEBUG: Creating temperature measurement tab")
        
        # Create tab frame
        temp_frame = ttk.Frame(self.notebook)
        self.notebook.add(temp_frame, text="Temperature Blocks")
        
        # Create main container
        main_container = Frame(temp_frame, bg=APP_BACKGROUND_COLOR)
        main_container.pack(fill='both', expand=True)
        
        # Title
        Label(main_container, text="Temperature Block Measurements", 
              bg=APP_BACKGROUND_COLOR, fg="white", font=(FONT[0], FONT[1]+2, "bold")).pack(pady=10)
        
        # Controls frame
        controls_frame = Frame(main_container, bg=APP_BACKGROUND_COLOR)
        controls_frame.pack(pady=10)
        
        Button(controls_frame, text="Add Temperature Block", command=self.add_temperature_block,
               bg=BUTTON_COLOR, fg="white", font=FONT).pack(side="left", padx=5)
        
        Button(controls_frame, text="Remove Last Block", command=self.remove_temperature_block,
               bg="#ff6b6b", fg="white", font=FONT).pack(side="left", padx=5)
        
        Button(controls_frame, text="Export Data", command=self.export_temperature_data,
               bg="#4ecdc4", fg="white", font=FONT).pack(side="left", padx=5)
        
        # Scrollable frame for temperature blocks
        self.temp_canvas = tk.Canvas(main_container, bg=APP_BACKGROUND_COLOR)
        self.temp_scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.temp_canvas.yview)
        self.temp_scrollable_frame = Frame(self.temp_canvas, bg=APP_BACKGROUND_COLOR)
        
        self.temp_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.temp_canvas.configure(scrollregion=self.temp_canvas.bbox("all"))
        )
        
        self.temp_canvas.create_window((0, 0), window=self.temp_scrollable_frame, anchor="nw")
        self.temp_canvas.configure(yscrollcommand=self.temp_scrollbar.set)
        
        self.temp_canvas.pack(side="left", fill="both", expand=True, padx=(20, 0))
        self.temp_scrollbar.pack(side="right", fill="y", padx=(0, 20))
    
    def create_terpene_profile_tab(self):
        """Create the terpene profile management tab."""
        print("DEBUG: Creating terpene profile tab")
        
        # Create tab frame
        profile_frame = ttk.Frame(self.notebook)
        self.notebook.add(profile_frame, text="Terpene Profiles")
        
        # Create main container
        main_container = Frame(profile_frame, bg=APP_BACKGROUND_COLOR)
        main_container.pack(fill='both', expand=True)
        
        # Title
        Label(main_container, text="Terpene Profile Management", 
              bg=APP_BACKGROUND_COLOR, fg="white", font=(FONT[0], FONT[1]+2, "bold")).pack(pady=10)
        
        # Profile list
        profiles_frame = Frame(main_container, bg=APP_BACKGROUND_COLOR)
        profiles_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # List current profiles
        Label(profiles_frame, text="Current Profiles:", bg=APP_BACKGROUND_COLOR, fg="white",
              font=(FONT[0], FONT[1], "bold")).pack(anchor="w")
        
        for profile_name, profile_data in self.terpene_profiles.items():
            profile_frame = Frame(profiles_frame, bg=APP_BACKGROUND_COLOR, relief="raised", bd=1)
            profile_frame.pack(fill="x", pady=2)
            
            Label(profile_frame, text=f"{profile_name}: {profile_data.get('notes', 'No description')}", 
                  bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=10, pady=5)
        
        # Add new profile button
        Button(profiles_frame, text="Add New Profile", command=self.add_terpene_profile,
               bg=BUTTON_COLOR, fg="white", font=FONT).pack(pady=10)
    
    def calculate_viscosity(self):
        """Calculate viscosity based on current inputs."""
        print("DEBUG: CalculationController calculating viscosity")
        
        try:
            media = self.media_var.get()
            terpene = self.terpene_var.get()
            mass = self.mass_of_oil_var.get()
            target = self.target_viscosity_var.get()
            
            if not all([media, terpene, mass > 0, target > 0]):
                messagebox.showwarning("Missing Data", "Please fill in all required fields")
                return
            
            # Look up formulation in database
            formulation_key = f"{media}_{terpene}"
            
            if formulation_key in self._formulation_db:
                formulation = self._formulation_db[formulation_key]
                base_ratio = formulation.get('ratio', 0.1)  # Default 10%
            else:
                # Use default calculation
                base_ratio = self.calculate_default_ratio(media, terpene)
            
            # Calculate results
            target_ratio = base_ratio * (target / 1000.0)  # Adjust for target viscosity
            exact_percent = target_ratio * 100
            exact_mass = mass * target_ratio
            
            # Calculate starting amounts (conservative approach)
            start_percent = exact_percent * 0.8  # Start with 80% of calculated
            start_mass = exact_mass * 0.8
            
            # Update result variables
            self.exact_percent_var.set(f"{exact_percent:.2f}%")
            self.exact_mass_var.set(f"{exact_mass:.3f}g")
            self.start_percent_var.set(f"{start_percent:.2f}%")
            self.start_mass_var.set(f"{start_mass:.3f}g")
            
            # Save this calculation to database
            self.save_calculation_to_database(media, terpene, mass, target, exact_percent, exact_mass)
            
            print(f"DEBUG: Calculated viscosity - Exact: {exact_percent:.2f}%, {exact_mass:.3f}g")
            
        except Exception as e:
            print(f"ERROR: Viscosity calculation failed: {e}")
            messagebox.showerror("Calculation Error", f"Failed to calculate viscosity: {e}")
    
    def calculate_default_ratio(self, media: str, terpene: str) -> float:
        """Calculate default ratio based on media and terpene types."""
        # Base ratios by media type
        media_ratios = {
            "D8": 0.08,
            "D9": 0.10,
            "Resin": 0.12,
            "Rosin": 0.15,
            "Liquid Diamonds": 0.06,
            "Other": 0.10
        }
        
        # Terpene adjustment factors
        terpene_factors = {
            "Limonene": 1.0,
            "Myrcene": 0.9,
            "Pinene": 1.1,
            "Linalool": 0.8,
            "Citrus Blend": 1.05,
            "Relaxing Blend": 0.85
        }
        
        base_ratio = media_ratios.get(media, 0.10)
        terpene_factor = terpene_factors.get(terpene, 1.0)
        
        return base_ratio * terpene_factor
    
    def calculate_advanced_steps(self):
        """Calculate step-by-step viscosity adjustment."""
        print("DEBUG: Calculating advanced steps")
        
        try:
            # Get basic calculation first
            self.calculate_viscosity()
            
            exact_mass = float(self.exact_mass_var.get().replace('g', ''))
            
            # Step 1: Add 70% of calculated amount
            step1_amount = exact_mass * 0.7
            self.step1_amount_var.set(f"{step1_amount:.3f}g")
            
            # Step 2: Remaining amount based on Step 1 measurement
            step1_viscosity = self.step1_viscosity_var.get()
            
            if step1_viscosity:
                try:
                    measured_visc = float(step1_viscosity)
                    target_visc = self.target_viscosity_var.get()
                    
                    if measured_visc > 0 and target_visc > 0:
                        # Calculate remaining amount based on measured viscosity
                        remaining_ratio = (target_visc - measured_visc) / target_visc
                        step2_amount = exact_mass * remaining_ratio * 0.3
                        
                        self.step2_amount_var.set(f"{step2_amount:.3f}g")
                        
                        # Expected final viscosity
                        expected_final = measured_visc + (step2_amount / exact_mass) * target_visc
                        self.expected_viscosity_var.set(f"{expected_final:.1f}")
                    
                except ValueError:
                    print("WARNING: Invalid viscosity measurement entered")
            
        except Exception as e:
            print(f"ERROR: Advanced calculation failed: {e}")
    
    def add_temperature_block(self):
        """Add a new temperature measurement block."""
        print("DEBUG: Adding temperature block")
        
        try:
            # Get temperature from user
            temp_dialog = tk.Toplevel()
            temp_dialog.title("Add Temperature Block")
            temp_dialog.geometry("300x150")
            temp_dialog.configure(bg=APP_BACKGROUND_COLOR)
            
            Label(temp_dialog, text="Temperature (°C):", bg=APP_BACKGROUND_COLOR, fg="white").pack(pady=10)
            
            temp_var = StringVar()
            temp_entry = Entry(temp_dialog, textvariable=temp_var)
            temp_entry.pack(pady=5)
            
            def add_temp():
                try:
                    temperature = float(temp_var.get())
                    self.create_temperature_block_ui(temperature)
                    temp_dialog.destroy()
                except ValueError:
                    messagebox.showerror("Invalid Temperature", "Please enter a valid temperature")
            
            Button(temp_dialog, text="Add", command=add_temp, bg=BUTTON_COLOR, fg="white").pack(pady=10)
            
        except Exception as e:
            print(f"ERROR: Failed to add temperature block: {e}")
    
    def create_temperature_block_ui(self, temperature: float):
        """Create UI elements for a temperature block."""
        print(f"DEBUG: Creating temperature block UI for {temperature}°C")
        
        # Add to temperature blocks list
        self.temperature_blocks.append((temperature, StringVar()))
        
        # Create block frame
        block_frame = Frame(self.temp_scrollable_frame, bg=APP_BACKGROUND_COLOR, relief="raised", bd=2)
        block_frame.pack(fill="x", padx=10, pady=5)
        
        # Temperature header
        Label(block_frame, text=f"Temperature: {temperature}°C", 
              bg=APP_BACKGROUND_COLOR, fg="white", font=(FONT[0], FONT[1], "bold")).pack(pady=5)
        
        # Speed input
        speed_frame = Frame(block_frame, bg=APP_BACKGROUND_COLOR)
        speed_frame.pack(pady=2)
        
        Label(speed_frame, text="Speed (RPM):", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
        speed_var = StringVar()
        Entry(speed_frame, textvariable=speed_var, width=10).pack(side="left", padx=5)
        self.speed_vars.append((temperature, speed_var))
        
        # Measurement table header
        measurements_frame = Frame(block_frame, bg=APP_BACKGROUND_COLOR)
        measurements_frame.pack(pady=5)
        
        # Headers
        Label(measurements_frame, text="Run", bg=APP_BACKGROUND_COLOR, fg="white", width=5).grid(row=0, column=0, padx=2, pady=2)
        Label(measurements_frame, text="Torque", bg=APP_BACKGROUND_COLOR, fg="white", width=10).grid(row=0, column=1, padx=2, pady=2)
        Label(measurements_frame, text="Viscosity", bg=APP_BACKGROUND_COLOR, fg="white", width=10).grid(row=0, column=2, padx=2, pady=2)
        
        # Three measurement rows
        for run in range(3):
            Label(measurements_frame, text=f"Run {run+1}", bg=APP_BACKGROUND_COLOR, fg="white").grid(row=run+1, column=0, padx=2, pady=1)
            
            # Torque entry
            torque_var = StringVar()
            torque_entry = Entry(measurements_frame, textvariable=torque_var, width=10)
            torque_entry.grid(row=run+1, column=1, padx=2, pady=1)
            torque_entry.bind('<KeyRelease>', lambda e, t=temperature: self.check_auto_calculate(t))
            self.torque_vars[run].append((temperature, torque_var))
            
            # Viscosity entry
            visc_var = StringVar()
            visc_entry = Entry(measurements_frame, textvariable=visc_var, width=10)
            visc_entry.grid(row=run+1, column=2, padx=2, pady=1)
            visc_entry.bind('<KeyRelease>', lambda e, t=temperature: self.check_auto_calculate(t))
            self.viscosity_vars[run].append((temperature, visc_var))
        
        # Average row
        Label(measurements_frame, text="Average", bg=APP_BACKGROUND_COLOR, fg="white", font=(FONT[0], FONT[1], "bold")).grid(row=4, column=0, padx=2, pady=2)
        
        avg_torque_var = StringVar()
        Label(measurements_frame, textvariable=avg_torque_var, bg=APP_BACKGROUND_COLOR, fg="#00b539").grid(row=4, column=1, padx=2, pady=2)
        self.avg_torque_vars.append((temperature, avg_torque_var))
        
        avg_visc_var = StringVar()
        Label(measurements_frame, textvariable=avg_visc_var, bg=APP_BACKGROUND_COLOR, fg="#00b539").grid(row=4, column=2, padx=2, pady=2)
        self.avg_visc_vars.append((temperature, avg_visc_var))
    
    def check_auto_calculate(self, temperature: float):
        """Check if we should automatically calculate averages for a temperature."""
        print(f"DEBUG: Checking auto-calculate for {temperature}°C")
        
        try:
            # Find all viscosity and torque values for this temperature
            visc_values = []
            torque_values = []
            
            for run in range(3):
                # Get viscosity values
                for temp, visc_var in self.viscosity_vars[run]:
                    if temp == temperature:
                        value = visc_var.get().strip()
                        if value:
                            try:
                                visc_float = float(value.replace(',', ''))
                                visc_values.append(visc_float)
                            except ValueError:
                                pass
                
                # Get torque values
                for temp, torque_var in self.torque_vars[run]:
                    if temp == temperature:
                        value = torque_var.get().strip()
                        if value:
                            try:
                                torque_float = float(value.replace(',', ''))
                                torque_values.append(torque_float)
                            except ValueError:
                                pass
            
            # Calculate averages if we have 3 values
            if len(visc_values) == 3:
                avg_visc = sum(visc_values) / len(visc_values)
                for temp, avg_var in self.avg_visc_vars:
                    if temp == temperature:
                        if avg_visc >= 1000:
                            avg_var.set(f"{avg_visc:,.1f}")
                        else:
                            avg_var.set(f"{avg_visc:.1f}")
                        break
            
            if len(torque_values) == 3:
                avg_torque = sum(torque_values) / len(torque_values)
                for temp, avg_var in self.avg_torque_vars:
                    if temp == temperature:
                        if avg_torque >= 1000:
                            avg_var.set(f"{avg_torque:,.1f}")
                        else:
                            avg_var.set(f"{avg_torque:.1f}")
                        break
            
        except Exception as e:
            print(f"ERROR: Auto-calculate failed: {e}")
    
    def remove_temperature_block(self):
        """Remove the last temperature block."""
        print("DEBUG: Removing last temperature block")
        
        try:
            if self.temperature_blocks:
                # Remove from data structures
                removed_temp = self.temperature_blocks.pop()[0]
                
                # Remove from variable lists
                self.speed_vars = [(t, v) for t, v in self.speed_vars if t != removed_temp]
                self.avg_torque_vars = [(t, v) for t, v in self.avg_torque_vars if t != removed_temp]
                self.avg_visc_vars = [(t, v) for t, v in self.avg_visc_vars if t != removed_temp]
                
                for run in range(3):
                    self.torque_vars[run] = [(t, v) for t, v in self.torque_vars[run] if t != removed_temp]
                    self.viscosity_vars[run] = [(t, v) for t, v in self.viscosity_vars[run] if t != removed_temp]
                
                # Recreate UI
                for widget in self.temp_scrollable_frame.winfo_children():
                    widget.destroy()
                
                # Recreate all remaining blocks
                for temp, _ in self.temperature_blocks:
                    self.create_temperature_block_ui(temp)
                
                print(f"DEBUG: Removed temperature block for {removed_temp}°C")
            
        except Exception as e:
            print(f"ERROR: Failed to remove temperature block: {e}")
    
    def export_temperature_data(self):
        """Export temperature measurement data."""
        print("DEBUG: Exporting temperature data")
        
        try:
            if not self.temperature_blocks:
                messagebox.showwarning("No Data", "Please add measurement blocks first.")
                return
            
            # Collect all data
            export_data = []
            
            for temp, _ in self.temperature_blocks:
                # Find speed
                speed = ""
                for t, speed_var in self.speed_vars:
                    if t == temp:
                        speed = speed_var.get()
                        break
                
                block_data = {
                    "temperature": temp,
                    "speed": speed,
                    "runs": []
                }
                
                # Collect run data
                for run in range(3):
                    run_data = {"torque": "", "viscosity": ""}
                    
                    for t, torque_var in self.torque_vars[run]:
                        if t == temp:
                            run_data["torque"] = torque_var.get()
                            break
                    
                    for t, visc_var in self.viscosity_vars[run]:
                        if t == temp:
                            run_data["viscosity"] = visc_var.get()
                            break
                    
                    block_data["runs"].append(run_data)
                
                # Get averages
                for t, avg_var in self.avg_torque_vars:
                    if t == temp:
                        block_data["average_torque"] = avg_var.get()
                        break
                
                for t, avg_var in self.avg_visc_vars:
                    if t == temp:
                        block_data["average_viscosity"] = avg_var.get()
                        break
                
                export_data.append(block_data)
            
            # Save to file
            file_path = filedialog.asksaveasfilename(
                title="Export Temperature Data",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w') as f:
                    json.dump(export_data, f, indent=2)
                
                messagebox.showinfo("Export Complete", f"Temperature data exported to {file_path}")
                print(f"DEBUG: Exported temperature data to {file_path}")
            
        except Exception as e:
            print(f"ERROR: Failed to export temperature data: {e}")
            messagebox.showerror("Export Error", f"Failed to export data: {e}")
    
    def add_terpene_profile(self):
        """Add a new terpene profile."""
        print("DEBUG: Adding new terpene profile")
        
        try:
            # Create dialog for new profile
            profile_dialog = tk.Toplevel()
            profile_dialog.title("Add Terpene Profile")
            profile_dialog.geometry("400x300")
            profile_dialog.configure(bg=APP_BACKGROUND_COLOR)
            
            # Profile name
            Label(profile_dialog, text="Profile Name:", bg=APP_BACKGROUND_COLOR, fg="white").pack(pady=5)
            name_var = StringVar()
            Entry(profile_dialog, textvariable=name_var, width=20).pack(pady=5)
            
            # Primary terpene
            Label(profile_dialog, text="Primary Terpene:", bg=APP_BACKGROUND_COLOR, fg="white").pack(pady=5)
            primary_var = StringVar()
            Entry(profile_dialog, textvariable=primary_var, width=20).pack(pady=5)
            
            # Notes
            Label(profile_dialog, text="Notes:", bg=APP_BACKGROUND_COLOR, fg="white").pack(pady=5)
            notes_var = StringVar()
            Entry(profile_dialog, textvariable=notes_var, width=30).pack(pady=5)
            
            def save_profile():
                name = name_var.get().strip()
                primary = primary_var.get().strip()
                notes = notes_var.get().strip()
                
                if name and primary:
                    self.terpene_profiles[name] = {
                        "primary": primary,
                        "percentage": 100.0,
                        "notes": notes or f"Custom {primary} profile"
                    }
                    messagebox.showinfo("Profile Added", f"Terpene profile '{name}' added successfully")
                    profile_dialog.destroy()
                else:
                    messagebox.showwarning("Missing Information", "Please provide name and primary terpene")
            
            Button(profile_dialog, text="Save Profile", command=save_profile,
                   bg=BUTTON_COLOR, fg="white").pack(pady=20)
            
        except Exception as e:
            print(f"ERROR: Failed to add terpene profile: {e}")
    
    def save_calculation_to_database(self, media: str, terpene: str, mass: float, target: float, percent: float, mass_result: float):
        """Save calculation results to the formulation database."""
        print("DEBUG: Saving calculation to database")
        
        try:
            key = f"{media}_{terpene}"
            
            if key not in self._formulation_db:
                self._formulation_db[key] = {
                    'media': media,
                    'terpene': terpene,
                    'calculations': []
                }
            
            calculation = {
                'mass_oil': mass,
                'target_viscosity': target,
                'result_percent': percent,
                'result_mass': mass_result,
                'ratio': percent / 100.0,
                'timestamp': self._get_current_timestamp()
            }
            
            self._formulation_db[key]['calculations'].append(calculation)
            
            # Keep only last 10 calculations per formulation
            if len(self._formulation_db[key]['calculations']) > 10:
                self._formulation_db[key]['calculations'] = self._formulation_db[key]['calculations'][-10:]
            
            # Update average ratio
            ratios = [calc['ratio'] for calc in self._formulation_db[key]['calculations']]
            self._formulation_db[key]['ratio'] = sum(ratios) / len(ratios)
            
            self.save_formulation_database()
            print("DEBUG: Calculation saved to database successfully")
            
        except Exception as e:
            print(f"ERROR: Failed to save calculation to database: {e}")
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp."""
        from datetime import datetime
        return datetime.now().isoformat()
    
    def train_models_from_data(self) -> bool:
        """Train standard viscosity models from data."""
        print("DEBUG: CalculationController training models from data")
        
        try:
            # This would implement model training logic
            # For now, just show a placeholder message
            messagebox.showinfo("Model Training", "Model training functionality will be implemented here")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to train models: {e}")
            return False
    
    def train_models_with_chemistry(self) -> bool:
        """Train enhanced viscosity models with chemistry data."""
        print("DEBUG: CalculationController training models with chemistry")
        
        try:
            # This would implement enhanced model training logic
            # For now, just show a placeholder message
            messagebox.showinfo("Enhanced Model Training", "Enhanced model training functionality will be implemented here")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to train models with chemistry: {e}")
            return False
    
    def analyze_models(self) -> bool:
        """Analyze viscosity models."""
        print("DEBUG: CalculationController analyzing models")
        
        try:
            # This would implement model analysis logic
            # For now, just show current database stats
            total_formulations = len(self._formulation_db)
            total_calculations = sum(len(form.get('calculations', [])) for form in self._formulation_db.values())
            
            analysis_text = f"Model Analysis Results:\n\n"
            analysis_text += f"Total Formulations: {total_formulations}\n"
            analysis_text += f"Total Calculations: {total_calculations}\n\n"
            
            if self._formulation_db:
                analysis_text += "Formulation Summary:\n"
                for key, data in self._formulation_db.items():
                    media = data.get('media', 'Unknown')
                    terpene = data.get('terpene', 'Unknown')
                    ratio = data.get('ratio', 0)
                    calc_count = len(data.get('calculations', []))
                    analysis_text += f"  {media} + {terpene}: {ratio:.3f} avg ratio ({calc_count} calculations)\n"
            
            # Show in dialog
            analysis_dialog = tk.Toplevel()
            analysis_dialog.title("Model Analysis")
            analysis_dialog.geometry("500x400")
            analysis_dialog.configure(bg=APP_BACKGROUND_COLOR)
            
            text_widget = tk.Text(analysis_dialog, bg="white", fg="black", font=FONT)
            text_widget.pack(fill="both", expand=True, padx=10, pady=10)
            text_widget.insert("1.0", analysis_text)
            text_widget.config(state="disabled")
            
            print("DEBUG: CalculationController successfully analyzed models")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to analyze models: {e}")
            return False
    
    def upload_training_data(self) -> bool:
        """Upload training data for models."""
        print("DEBUG: CalculationController uploading training data")
        
        try:
            file_path = filedialog.askopenfilename(
                title="Select Training Data File",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if not file_path:
                return False
            
            # Load and process training data
            if file_path.endswith(('.xlsx', '.xls')):
                data = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                data = pd.read_csv(file_path)
            else:
                messagebox.showerror("Invalid File", "Please select an Excel or CSV file")
                return False
            
            # Process the training data
            # This would implement the actual training data processing
            messagebox.showinfo("Training Data", f"Loaded training data with {len(data)} rows from {file_path}")
            
            print(f"DEBUG: CalculationController successfully uploaded training data from {file_path}")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to upload training data: {e}")
            messagebox.showerror("Upload Error", f"Failed to upload training data: {e}")
            return False
    
    def embed_in_frame(self, parent_frame):
        """Embed the calculator interface in a parent frame."""
        print("DEBUG: CalculationController embedding in frame")
        
        try:
            # Create notebook in the parent frame
            self.notebook = ttk.Notebook(parent_frame)
            self.notebook.pack(fill='both', expand=True)
            
            # Create all tabs
            self.create_basic_calculator_tab()
            self.create_advanced_calculator_tab()
            self.create_temperature_measurement_tab()
            self.create_terpene_profile_tab()
            
            print("DEBUG: CalculationController successfully embedded in frame")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to embed in frame: {e}")
            return False
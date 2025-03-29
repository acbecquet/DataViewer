"""
viscosity_calculator.py
Module for calculating terpene percentages based on viscosity and vice versa.
"""
import os
import json
import tkinter as tk
from tkinter import ttk, StringVar, DoubleVar, IntVar, Toplevel, Frame, Label, Entry, Button, messagebox, OptionMenu

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
        
        # Define media options
        self.media_options = ["D8", "D9", "Resin", "Rosin", "Liquid Diamonds", "Other"]
        
        # Load or initialize formulation database
        self.formulation_db_path = "terpene_formulations.json"
        self._formulation_db = None  # Lazy load this
        
        # Store notebook reference
        self.notebook = None

        # Lazy load models
        self._viscosity_models = None
        
    # ---- Database access properties with lazy loading ----
    
    @property
    def formulation_db(self):
        """Lazy load the formulation database"""
        if self._formulation_db is None:
            self._formulation_db = self.load_formulation_database()
        return self._formulation_db
    
    @property
    def viscosity_models(self):
        """Lazy load the standard viscosity models"""
        if self._viscosity_models is None:
            self._viscosity_models = self.load_standard_models()
        return self._viscosity_models

    @property
    def enhanced_viscosity_models(self):
        """Lazy load the enhanced viscosity models with chemistry"""
        if not hasattr(self, '_enhanced_viscosity_models') or self._enhanced_viscosity_models is None:
            self._enhanced_viscosity_models = self.load_enhanced_models()
        return self._enhanced_viscosity_models

    @property 
    def combined_viscosity_models(self):
        """Lazy load the combined viscosity models"""
        if not hasattr(self, '_combined_viscosity_models') or self._combined_viscosity_models is None:
            self._enhanced_viscosity_models = self.load_combined_models()
        return self._combined_viscosity_models

    @enhanced_viscosity_models.setter
    def enhanced_viscosity_models(self, value):
        """Setter for enhanced viscosity models"""
        self._enhanced_viscosity_models = value

    @viscosity_models.setter
    def viscosity_models(self, value):
        self._viscosity_models = value

    @combined_viscosity_models.setter
    def combined_viscosity_models(self,value):
        self._combined_viscosity_models = value

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
        calculator_window = Toplevel(self.root)
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
    
    # ---- UI Creation Methods ----
    
    def create_calculator_tab(self, notebook):
        """
        Create the main calculator tab.
        
        Args:
            notebook (ttk.Notebook): The notebook widget to add the tab to
        """
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = '#0504AA'
        
        tab1 = Frame(notebook, bg=APP_BACKGROUND_COLOR)
        notebook.add(tab1, text="Calculator")
        
        # Create a frame for the form with some padding
        form_frame = Frame(tab1, bg=APP_BACKGROUND_COLOR, padx=20, pady=10)
        form_frame.pack(fill='both', expand=True)
        
        # Row 0: Title and explanation
        explanation = Label(form_frame, 
                            text="Calculate terpene % based on viscosity formula.\nIf no formula is known, use the 'Iterative Method' tab.",
                            bg=APP_BACKGROUND_COLOR, fg="white", 
                            font=FONT, justify="center")
        explanation.grid(row=0, column=0, columnspan=4, pady=0)
        
        # Row 1: Media and Terpene
        Label(form_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=1, column=0, sticky="w", pady=5)
        
        media_dropdown = ttk.Combobox(
            form_frame, 
            textvariable=self.media_var,
            values=self.media_options,
            state="readonly",
            width=15
        )
        media_dropdown.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        media_dropdown.current(0)  # Set default value
        
        Label(form_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=1, column=2, sticky="w", pady=5)
        
        terpene_entry = Entry(form_frame, textvariable=self.terpene_var, width=15)
        terpene_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        
        # Row 2: Media Brand and Terpene Brand
        Label(form_frame, text="Media Brand:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=2, column=0, sticky="w", pady=5)
        
        media_brand_entry = Entry(form_frame, textvariable=self.media_brand_var, width=15)
        media_brand_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        Label(form_frame, text="Terpene Brand:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=2, column=2, sticky="w", pady=5)
        
        terpene_brand_entry = Entry(form_frame, textvariable=self.terpene_brand_var, width=15)
        terpene_brand_entry.grid(row=2, column=3, sticky="w", padx=5, pady=5)
        
        # Row 3: Mass of Oil
        Label(form_frame, text="Mass of Oil (g):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=3, column=0, sticky="w", pady=5)
        
        mass_entry = Entry(form_frame, textvariable=self.mass_of_oil_var, width=15)
        mass_entry.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        # Row 4: Target Viscosity
        Label(form_frame, text="Target Viscosity:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=4, column=0, sticky="w", pady=5)
        
        viscosity_entry = Entry(form_frame, textvariable=self.target_viscosity_var, width=15)
        viscosity_entry.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        # Separator for results
        ttk.Separator(form_frame, orient='horizontal').grid(row=5, column=0, columnspan=4, sticky="ew", pady=15)

        # Results section header
        Label(form_frame, text="Results:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=(FONT[0], FONT[1], "bold"), anchor="w").grid(row=6, column=0, sticky="w", pady=5)
        
        
        # Row 7: Exact % and Exact Mass
        Label(form_frame, text="Exact %:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=7, column=0, sticky="w", pady=3)
        
        exact_percent_label = Label(form_frame, textvariable=self.exact_percent_var, 
                              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        exact_percent_label.grid(row=7, column=1, sticky="w", pady=3)
        
        Label(form_frame, text="Exact Mass:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=7, column=2, sticky="w", pady=3)
        
        exact_mass_label = Label(form_frame, textvariable=self.exact_mass_var, 
                           bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        exact_mass_label.grid(row=7, column=3, sticky="w", pady=3)
        
        # Row 8: Start % and Start Mass
        Label(form_frame, text="Start %:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=8, column=0, sticky="w", pady=3)
        
        start_percent_label = Label(form_frame, textvariable=self.start_percent_var, 
                              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        start_percent_label.grid(row=8, column=1, sticky="w", pady=3)
        
        Label(form_frame, text="Start Mass:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=8, column=2, sticky="w", pady=3)
        
        start_mass_label = Label(form_frame, textvariable=self.start_mass_var, 
                           bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        start_mass_label.grid(row=8, column=3, sticky="w", pady=3)
        
        # Create button frame for organized rows of buttons
        button_frame = Frame(form_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.grid(row=9, column=0, columnspan=4, pady=10)

        # Create first row of buttons
        button_row1 = Frame(button_frame, bg=APP_BACKGROUND_COLOR)
        button_row1.pack(fill="x", pady=(0, 5))  # Add some space between rows

        # Calculate button
        calculate_btn = ttk.Button(
            button_row1,
            text="Calculate",
            command=self.calculate_terpene_percentage
        )
        calculate_btn.pack(padx=5,pady=5)

    def create_advanced_tab(self, notebook):
        """
        Create the advanced tab for iterative viscosity calculation with balanced spacing.
        """
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = '#0504AA'
        
        tab2 = Frame(notebook, bg=APP_BACKGROUND_COLOR)
        notebook.add(tab2, text="Iterative Method")

        # Create a form frame with minimal padding
        form_frame = Frame(tab2, bg=APP_BACKGROUND_COLOR, padx=2, pady=2)
        form_frame.pack(fill='both', expand=True)

        # Configure column weights for proper distribution
        for i in range(4):
            form_frame.columnconfigure(i, weight=1, minsize=80)

        # Row 0: Title and Explanation
        explanation = Label(form_frame, 
                           text="Use this method when no formula is available.\nFollow the steps and measure viscosity at each stage.",
                           bg=APP_BACKGROUND_COLOR, fg="white", 
                           font=FONT, justify="center")
        explanation.grid(row=0, column=0, columnspan=4, pady=2)

        # Row 1: Media and Media Brand
        Label(form_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=1, column=0, sticky="w", pady=1)
        Label(form_frame, textvariable=self.media_var, 
              bg=APP_BACKGROUND_COLOR, fg="white", font=FONT).grid(row=1, column=1, sticky="w", pady=1)

        Label(form_frame, text="Media Brand:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=1, column=2, sticky="w", pady=1)
        Label(form_frame, textvariable=self.media_brand_var, 
              bg=APP_BACKGROUND_COLOR, fg="white", font=FONT).grid(row=1, column=3, sticky="w", pady=1)

        # Row 2: Terpene and Terpene Brand
        Label(form_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=2, column=0, sticky="w", pady=1)
        Label(form_frame, textvariable=self.terpene_var, 
              bg=APP_BACKGROUND_COLOR, fg="white", font=FONT).grid(row=2, column=1, sticky="w", pady=1)

        Label(form_frame, text="Terpene Brand:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=2, column=2, sticky="w", pady=1)
        Label(form_frame, textvariable=self.terpene_brand_var, 
              bg=APP_BACKGROUND_COLOR, fg="white", font=FONT).grid(row=2, column=3, sticky="w", pady=1)

        # Row 3: Target Viscosity and Oil Mass
        Label(form_frame, text="Target Viscosity:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=3, column=0, sticky="w", pady=1)
        Label(form_frame, textvariable=self.target_viscosity_var, 
              bg=APP_BACKGROUND_COLOR, fg="white", font=FONT).grid(row=3, column=1, sticky="w", pady=1)

        Label(form_frame, text="Oil Mass (g):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=3, column=2, sticky="w", pady=1)
        Label(form_frame, textvariable=self.mass_of_oil_var, 
              bg=APP_BACKGROUND_COLOR, fg="white", font=FONT).grid(row=3, column=3, sticky="w", pady=1)

        # Separator
        ttk.Separator(form_frame, orient='horizontal').grid(row=4, column=0, columnspan=4, sticky="ew", pady=5)

        # Create a frame for the steps section with better spacing
        steps_frame = Frame(form_frame, bg=APP_BACKGROUND_COLOR)
        steps_frame.grid(row=5, column=0, columnspan=4, sticky="nsew", pady=5)
    
        # Configure the steps frame columns with balanced widths
        steps_frame.columnconfigure(0, weight=1)  # "Step X: Add" label
        steps_frame.columnconfigure(1, weight=1)  # Amount value
        steps_frame.columnconfigure(2, weight=1)  # "terpenes" label
    
        # Step 1 row - Improve the flow with consistent spacing
        Label(steps_frame, text="Step 1: Add", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).grid(row=0, column=0, sticky="e", padx=(5, 0), pady=3)
        Label(steps_frame, textvariable=self.step1_amount_var, 
              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT).grid(row=0, column=1, pady=3)
        Label(steps_frame, text="of terpenes", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).grid(row=0, column=2, sticky="w", padx=(0, 5), pady=3)
    
        # Viscosity input row 1
        visc_frame1 = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        visc_frame1.grid(row=1, column=0, columnspan=3, sticky="ew", pady=3)
        visc_frame1.columnconfigure(0, weight=1)
        visc_frame1.columnconfigure(1, weight=1)
    
        Label(visc_frame1, text="Viscosity @ 25C:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).grid(row=0, column=0, sticky="e", padx=(0, 5))
        Entry(visc_frame1, textvariable=self.step1_viscosity_var, width=10).grid(row=0, column=1, sticky="w")
    
        # Step 2 row - same consistent layout as Step 1
        Label(steps_frame, text="Step 2: Add", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).grid(row=2, column=0, sticky="e", padx=(5, 0), pady=3)
        Label(steps_frame, textvariable=self.step2_amount_var, 
              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT).grid(row=2, column=1, pady=3)
        Label(steps_frame, text="of terpenes", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).grid(row=2, column=2, sticky="w", padx=(0, 5), pady=3)
    
        # Viscosity input row 2
        visc_frame2 = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        visc_frame2.grid(row=3, column=0, columnspan=3, sticky="ew", pady=3)
        visc_frame2.columnconfigure(0, weight=1)
        visc_frame2.columnconfigure(1, weight=1)
    
        Label(visc_frame2, text="Viscosity @ 25C:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).grid(row=0, column=0, sticky="e", padx=(0, 5))
        Entry(visc_frame2, textvariable=self.step2_viscosity_var, width=10).grid(row=0, column=1, sticky="w")
    
        # Expected Viscosity with proper centering
        expected_frame = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        expected_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=8)
        expected_frame.columnconfigure(0, weight=1)
    
        # Create a container to hold both the label and value horizontally centered
        expected_container = Frame(expected_frame, bg=APP_BACKGROUND_COLOR)
        expected_container.grid(row=0, column=0)
    
        Label(expected_container, text="Expected Viscosity: ", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=(FONT[0], FONT[1], "bold")).pack(side="left")
        Label(expected_container, textvariable=self.expected_viscosity_var, 
              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=(FONT[0], FONT[1], "bold")).pack(side="left")
    
        # Buttons with better spacing
        button_frame = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=5)
    
        # Use pack instead of grid for better button spacing
        calculate_btn1 = ttk.Button(
            button_frame,
            text="Calculate Step 1",
            command=self.calculate_step1,
            width=16
        )
        calculate_btn1.pack(side="left", padx=5, expand=True)
    
        calculate_btn2 = ttk.Button(
            button_frame,
            text="Calculate Step 2",
            command=self.calculate_step2,
            width=16
        )
        calculate_btn2.pack(side="left", padx=5, expand=True)
    
        save_btn = ttk.Button(
            button_frame,
            text="Save",
            command=self.save_formulation,
            width=8
        )
        save_btn.pack(side="left", padx=5, expand=True)

    def create_measure_tab(self, notebook):
        """
        Create a measurement tab with temperature blocks layout
        """
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = '#0504AA'
        
        tab3 = Frame(notebook, bg=APP_BACKGROUND_COLOR)
        notebook.add(tab3, text="Measure")
        
        # Main container with padding
        main_frame = Frame(tab3, bg=APP_BACKGROUND_COLOR)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title section
        title_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        title_frame.pack(fill='x', pady=(0, 10))
        
        Label(title_frame, text="Raw Viscosity Measurement", 
            bg=APP_BACKGROUND_COLOR, fg="white", font=(FONT[0], FONT[1]+2, "bold")).pack(pady=(0, 2))
        
        # Create a frame for input fields using grid layout
        input_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        input_frame.pack(fill='x', pady=5)
        
        # Configure grid columns to have appropriate weights
        for i in range(6):
            input_frame.columnconfigure(i, weight=1)
        
        # Media section - now using grid consistently
        Label(input_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white", 
            font=FONT).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        
        media_dropdown = ttk.Combobox(
            input_frame, 
            textvariable=self.media_var,
            values=self.media_options,
            state="readonly",
            width=10
        )
        media_dropdown.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        # Terpene section
        Label(input_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, fg="white", 
            font=FONT).grid(row=0, column=2, sticky="w", padx=5, pady=2)
        
        # Initialize terpene variable if not already done
        if not hasattr(self, 'measure_terpene_var'):
            self.measure_terpene_var = StringVar(value="Raw")  # Default to "Raw"
        
        terpene_entry = Entry(input_frame, textvariable=self.measure_terpene_var, width=12)
        terpene_entry.grid(row=0, column=3, sticky="w", padx=5, pady=2)
        
        # Terpene percentage
        Label(input_frame, text="Terpene %:", bg=APP_BACKGROUND_COLOR, fg="white", 
            font=FONT).grid(row=0, column=4, sticky="w", padx=5, pady=2)
        
        # Initialize terpene percentage variable if not already done
        if not hasattr(self, 'measure_terpene_pct_var'):
            self.measure_terpene_pct_var = DoubleVar(value=0.0)  # Default to 0%
        
        pct_entry = Entry(input_frame, textvariable=self.measure_terpene_pct_var, width=5)
        pct_entry.grid(row=0, column=5, sticky="w", padx=5, pady=2)
        
        # Create a frame with scrolling capability for temperature blocks
        canvas_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        canvas_frame.pack(fill='both', expand=True, pady=10)
    
        canvas = tk.Canvas(canvas_frame, bg=APP_BACKGROUND_COLOR, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
    
        # Container for all temperature blocks
        blocks_frame = Frame(canvas, bg=APP_BACKGROUND_COLOR)
    
        # Configure scrolling
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
        # Create window in canvas for the blocks
        canvas_window = canvas.create_window((0, 0), window=blocks_frame, anchor="nw")
    
        # Update scrollregion when the blocks_frame changes size
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Also update the width of the canvas window
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
    
        blocks_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=canvas.winfo_width()))
    
        # Initialize variables for tracking the blocks and measurements
        self.temperature_blocks = []
        self.speed_vars = []
        self.torque_vars = [[] for _ in range(3)]  # 3 runs for each temperature
        self.viscosity_vars = [[] for _ in range(3)]
    
        # Default temperatures
        default_temps = [25, 30, 40, 50]
    
        # Create a block for each temperature
        for temp in default_temps:
            self.create_temperature_block(blocks_frame, temp)
    
        # Buttons for managing and calculating
        button_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.pack(fill='x', pady=10)
    
        ttk.Button(button_frame, text="Add Block", command=lambda: self.add_temperature_block(blocks_frame)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Calculate Averages", command=self.calculate_viscosity_block_stats).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Save Measurements", command=self.save_block_measurements).pack(side="right", padx=5)

    # ---- Utility Methods ----
    
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
        These are placeholder values - replace with actual data. <--- need to do this
        """
        # These are just example values - replace with actual values based on your data
        viscosity_map = {
            "D8": 18000000,
            "D9": 20000000,
            "Liquid Diamonds": 20000000,
            "Other": 2000000 
        }
        
        return viscosity_map.get(media_type, 18000000)
    
    def update_advanced_tab_fields(self):
        """Update the fields in the advanced tab based on the calculator tab inputs."""
        try:
            # Calculate initial terpene amount for step 1
            try:
                mass_of_oil = float(self.mass_of_oil_var.get())
                target_viscosity = float(self.target_viscosity_var.get())
            
                # For step 1, determine amount to add based on guesstimates
                # The calculation described: if 2x < y0, add 1% terpene else 0.1%
                # We're using target_viscosity as x, and assuming y0 (raw oil viscosity)
            
                # Assume raw oil viscosity based on media type - these are placeholders
                raw_oil_viscosity = self.get_raw_oil_viscosity(self.media_var.get())
            
                if 2 * target_viscosity < raw_oil_viscosity:
                    # Add 1% terpene
                    percent_to_add = 1.0
                else:
                    # Add 0.1% terpene
                    percent_to_add = 0.1
            
                step1_amount = (percent_to_add / 100) * mass_of_oil
                self.step1_amount_var.set(f"{step1_amount:.2f}g")
            
            except (ValueError, tk.TclError):
                # If inputs can't be converted to numbers, don't update
                pass
            
        except Exception as e:
            print(f"Error updating advanced tab fields: {e}")

    # ---- Interactive Methods ----
    
    def calculate_step1(self):
        """Calculate the first step amount of terpenes to add."""
        try:
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
            
            # Get estimated raw oil viscosity based on media type
            raw_oil_viscosity = self.get_raw_oil_viscosity(self.media_var.get())
            
            # Calculate amount to add for step 1
            if 2 * target_viscosity < raw_oil_viscosity:
                # Add 1% terpene if target is much lower than raw oil
                percent_to_add = 1.0
            else:
                # Add 0.1% terpene if target is closer to raw oil
                percent_to_add = 0.1
            
            step1_amount = (percent_to_add / 100) * mass_of_oil
            self.step1_amount_var.set(f"{step1_amount:.2f}g")
            
            messagebox.showinfo("Step 1", 
                               f"Add {step1_amount:.2f}g of {self.terpene_var.get()} terpenes.\n"
                               f"Mix thoroughly and then measure the viscosity at 25C.\n"
                               f"Enter the measured viscosity in the field.")
            
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")
    
    def calculate_step2(self):
        """
        Calculate the second step amount of terpenes to add using an improved model
        that accounts for the non-linear relationship between terpene percentage and viscosity.
        """
        try:
            # Get the inputs
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
            step1_viscosity_text = self.step1_viscosity_var.get()
        
            if not step1_viscosity_text:
                messagebox.showerror("Input Error", 
                                    "Please enter the measured viscosity from Step 1.")
                return
        
            step1_viscosity = float(step1_viscosity_text)
            step1_amount = float(self.step1_amount_var.get().replace('g', ''))
        
            # Get raw oil viscosity based on media type
            raw_oil_viscosity = self.get_raw_oil_viscosity(self.media_var.get())
        
            # Convert to percentages for calculations
            step1_percent = (step1_amount / mass_of_oil) * 100
        
            # First, try to use existing models if available
            have_model = False
        
            if hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models:
                have_model = True
            elif hasattr(self, 'enhanced_viscosity_models') and self.enhanced_viscosity_models:
                have_model = True
            elif hasattr(self, 'viscosity_models') and self.viscosity_models:
                have_model = True
            
            

            if have_model:
                # Use model-based prediction
                try:
                    media = self.media_var.get()
                    
                
                    # Estimate potency (if needed)
                    potency = None
                    if hasattr(self, '_total_potency_var'):
                        potency = self._total_potency_var.get()
                        if potency <= 0:
                            # If no potency is explicitly provided, estimate it from terpene
                            # Assuming potency + terpene % ≈ 100%
                            potency = 100 - step1_percent
                
                    # Only look for general models
                    general_key = f"{media}_general"
                    model_info = None
                    model_source = None
                
                    # Search in combined models first
                    if hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models:
                        if general_key in self.combined_viscosity_models:
                            model_info = self.combined_viscosity_models[general_key]
                            model_source = "combined"
                
                    # Fall back to enhanced models
                    if model_info is None and hasattr(self, 'enhanced_viscosity_models') and self.enhanced_viscosity_models:
                        if general_key in self.enhanced_viscosity_models:
                            model_info = self.enhanced_viscosity_models[general_key]
                            model_source = "enhanced"
                
                    # Fall back to standard models
                    if model_info is None and hasattr(self, 'viscosity_models') and self.viscosity_models:
                        if general_key in self.viscosity_models:
                            model_info = self.viscosity_models[general_key]
                            model_source = "standard"
                
                    if model_info is not None:
                        # We found a suitable model, now use a sampling approach to find 
                        # the terpene percentage that gives a viscosity closest to the target
                    
                        # Define a function to predict viscosity based on terpene percentage
                        def predict_viscosity_at_percent(terpene_pct):
                            if model_source in ["combined", "enhanced"] and potency is not None:
                                return self.predict_model_viscosity(model_info, terpene_pct, 25.0, potency)
                            else:
                                return self.predict_model_viscosity(model_info, terpene_pct, 25.0)
                    
                        # Define a function to calculate error from target
                        def viscosity_error(terpene_pct):
                            visc = predict_viscosity_at_percent(terpene_pct)
                            return abs(visc - target_viscosity) / target_viscosity  # Relative error
                    
                        # Sample a range of percentages and find the one with the lowest error
                        best_percent = step1_percent
                        best_error = viscosity_error(step1_percent)
                    
                        # Sample more densely near the current percentage and more sparsely farther away
                        test_percentages = [step1_percent + 0.1 * i for i in range(1, 10)] + \
                                          [step1_percent + 0.5 * i for i in range(2, 20)] + \
                                          [step1_percent + i for i in range(1, 11)]
                    
                        # Limit to reasonable range (0-15%)
                        test_percentages = [p for p in test_percentages if 0 <= p <= 15]
                    
                        for test_percent in test_percentages:
                            error = viscosity_error(test_percent)
                            if error < best_error:
                                best_percent = test_percent
                                best_error = error
                    
                        # Refine search around the best percentage
                        for offset in [-0.2, -0.1, 0.1, 0.2]:
                            test_percent = best_percent + offset
                            if 0 <= test_percent <= 15:
                                error = viscosity_error(test_percent)
                                if error < best_error:
                                    best_percent = test_percent
                                    best_error = error
                    
                        # Calculate additional percentage needed
                        total_percent_needed = best_percent
                        percent_needed = max(0, total_percent_needed - step1_percent)
                    
                        # Calculate amount for step 2
                        step2_amount = (percent_needed / 100) * mass_of_oil
                    
                        # Update the UI
                        self.step2_amount_var.set(f"{step2_amount:.2f}g")
                    
                        # Predict final viscosity
                        expected_viscosity = predict_viscosity_at_percent(total_percent_needed)
                        self.expected_viscosity_var.set(f"{expected_viscosity:.2f}")
                    
                        # Show information about the model used
                        model_type = {
                            "combined": "combined (terpene+potency)",
                            "enhanced": "potency-based",
                            "standard": "terpene-based"
                        }.get(model_source, model_source)
                    
                        messagebox.showinfo("Step 2 (Model-based)", 
                                           f"Add an additional {step2_amount:.2f}g of {self.terpene_var.get()} terpenes.\n"
                                           f"Mix thoroughly and then measure the final viscosity at 25C.\n"
                                           f"The expected final viscosity is {expected_viscosity:.2f}.\n\n"
                                           f"(Using {model_type} model)")
                    
                        return
                except Exception as e:
                    import traceback
                    print(f"Error using model for prediction: {str(e)}")
                    print(traceback.format_exc())
                    # Fall back to exponential calculation below
        
            # Fallback: Use exponential model based on the two measurements
            import math
        
            # Solve for the decay constant in the exponential model
            # viscosity = raw_oil_viscosity * exp(-k * terpene_percentage)
            if step1_viscosity <= 0 or raw_oil_viscosity <= 0:
                raise ValueError("Viscosity values must be positive")
            
            # Calculate decay constant k
            k = -math.log(step1_viscosity / raw_oil_viscosity) / step1_percent
        
            # Solve for the total percentage needed to reach target viscosity
            if target_viscosity <= 0:
                raise ValueError("Target viscosity must be positive")
            
            total_percent_needed = -math.log(target_viscosity / raw_oil_viscosity) / k
        
            # Calculate additional percentage needed
            percent_needed = max(0, total_percent_needed - step1_percent)
        
            # Calculate amount for step 2
            step2_amount = (percent_needed / 100) * mass_of_oil
        
            # Update the UI
            self.step2_amount_var.set(f"{step2_amount:.2f}g")
        
            # Calculate expected final viscosity using the model
            expected_viscosity = raw_oil_viscosity * math.exp(-k * (step1_percent + percent_needed))
            self.expected_viscosity_var.set(f"{expected_viscosity:.2f}")
        
            messagebox.showinfo("Step 2 (Exponential Model)", 
                               f"Add an additional {step2_amount:.2f}g of {self.terpene_var.get()} terpenes.\n"
                               f"Mix thoroughly and then measure the final viscosity at 25C.\n"
                               f"The expected final viscosity is {expected_viscosity:.2f}.\n\n"
                               f"(Using exponential model based on your measurements)")
        
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")

    def save_formulation(self):
        """
        Save the terpene formulation to the database file and append to the master CSV.
        Creates two separate rows: one for step 1 and one for step 2, each with their specific terpene percentages.
        """
        try:
            # Verify we have all required data
            step1_viscosity_text = self.step1_viscosity_var.get()
            step2_viscosity_text = self.step2_viscosity_var.get()
        
            if not step1_viscosity_text:
                messagebox.showinfo("Input Needed", 
                                  "Please enter the measured viscosity from Step 1 before saving.")
                return
            
            if not step2_viscosity_text:
                messagebox.showinfo("Input Needed", 
                                  "Please enter the final measured viscosity from Step 2 before saving.")
                return
        
            # Get all formulation data
            media = self.media_var.get()
            media_brand = self.media_brand_var.get()
            terpene = self.terpene_var.get()
            terpene_brand = self.terpene_brand_var.get()
            target_viscosity = float(self.target_viscosity_var.get())
        
            step1_amount = float(self.step1_amount_var.get().replace('g', ''))
            step1_viscosity = float(step1_viscosity_text)
            step1_terpene_pct = (step1_amount / float(self.mass_of_oil_var.get())) * 100
        
            # For step 2, we need the total terpene amount and percentage
            step2_amount = float(self.step2_amount_var.get().replace('g', ''))
            step2_viscosity = float(step2_viscosity_text)
            total_oil_mass = float(self.mass_of_oil_var.get())
            total_terpene_mass = step1_amount + step2_amount
            total_terpene_pct = (total_terpene_mass / total_oil_mass) * 100
        
            # Calculate potency as 1 - terpene percent (in decimal)
            step1_potency = 1.0 - (step1_terpene_pct / 100.0)
            final_potency = 1.0 - (total_terpene_pct / 100.0)
        
            # Get cannabinoid content if available
            d9_thc = getattr(self, '_d9_thc_var', None)
            d9_thc_value = d9_thc.get() / 100.0 if d9_thc else None
        
            d8_thc = getattr(self, '_d8_thc_var', None)
            d8_thc_value = d8_thc.get() / 100.0 if d8_thc else None
        
            # Current timestamp
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
            # Create the database entry (for JSON formulation database)
            formulation = {
                "media": media,
                "media_brand": media_brand,
                "terpene": terpene,
                "terpene_brand": terpene_brand,
                "target_viscosity": target_viscosity,
                "step1_amount": step1_amount,
                "step1_viscosity": step1_viscosity,
                "step1_terpene_pct": step1_terpene_pct,
                "step2_amount": step2_amount,
                "step2_viscosity": step2_viscosity,
                "expected_viscosity": float(self.expected_viscosity_var.get()),
                "total_oil_mass": total_oil_mass,
                "total_terpene_mass": total_terpene_mass,
                "total_terpene_pct": total_terpene_pct,
                "step1_potency": step1_potency * 100,  # Store as percentage
                "final_potency": final_potency * 100,  # Store as percentage
                "timestamp": timestamp
            }
        
            # Add to database
            key = f"{formulation['media']}_{formulation['media_brand']}_{formulation['terpene']}_{formulation['terpene_brand']}"
        
            if key not in self.formulation_db:
                self.formulation_db[key] = []
        
            self.formulation_db[key].append(formulation)
        
            # Save database to file
            self.save_formulation_database()
        
            # Now create entries for the master CSV file
            import pandas as pd
            import os
        
            # Create two separate rows - one for step 1, one for step 2 (final)
            csv_rows = []
        
            # Common fields for both rows
            common_fields = {
                'media': media,
                'media_brand': media_brand,
                'terpene': terpene,
                'terpene_brand': terpene_brand,
                'temperature': 25.0,  # Standard measurement temperature
                'timestamp': timestamp
            }
        
            # Add cannabinoid fields if available
            if d9_thc_value is not None:
                common_fields['d9_thc'] = d9_thc_value
        
            if d8_thc_value is not None:
                common_fields['d8_thc'] = d8_thc_value
        
            # Step 1 row
            step1_row = common_fields.copy()
            step1_row.update({
                'terpene_pct': step1_terpene_pct / 100.0,  # Store as decimal in CSV
                'total_potency': step1_potency,  # Store as decimal
                'viscosity': step1_viscosity,
                'measurement_stage': 'step1'
            })
            csv_rows.append(step1_row)
        
            # Step 2 row (final formulation)
            step2_row = common_fields.copy()
            step2_row.update({
                'terpene_pct': total_terpene_pct / 100.0,  # Store as decimal in CSV
                'total_potency': final_potency,  # Store as decimal
                'viscosity': step2_viscosity,
                'measurement_stage': 'step2'
            })
            csv_rows.append(step2_row)
        
            # Append to master CSV file
            master_file = './data/Master_Viscosity_Data_processed.csv'
        
            try:
                # Create DataFrame from new rows
                new_rows_df = pd.DataFrame(csv_rows)
            
                # Check if file exists
                if os.path.exists(master_file):
                    # Load existing data
                    master_df = pd.read_csv(master_file)
                
                    # Ensure all columns are present in both dataframes
                    for col in new_rows_df.columns:
                        if col not in master_df.columns:
                            master_df[col] = None
                
                    for col in master_df.columns:
                        if col not in new_rows_df.columns:
                            new_rows_df[col] = None
                
                    # Append new rows
                    master_df = pd.concat([master_df, new_rows_df], ignore_index=True)
                else:
                    # Create directory if needed
                    os.makedirs(os.path.dirname(master_file), exist_ok=True)
                    # First time creating the file
                    master_df = new_rows_df
            
                # Save to CSV
                master_df.to_csv(master_file, index=False)
                print(f"Added {len(csv_rows)} new measurements to master CSV file.")
            except Exception as e:
                import traceback
                print(f"Error updating master CSV: {str(e)}")
                print(traceback.format_exc())
                # Don't stop execution - we've already saved to the database
        
            # Show success message
            messagebox.showinfo("Success", 
                              f"Formulation saved successfully!\n\n"
                              f"Total terpene percentage: {total_terpene_pct:.2f}%\n"
                              f"Total terpene mass: {total_terpene_mass:.2f}g\n"
                              f"Added step 1 ({step1_terpene_pct:.2f}%) and step 2 ({total_terpene_pct:.2f}%) data points.")
        
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            import traceback
            traceback_str = traceback.format_exc()
            print(f"Error in save_formulation: {str(e)}")
            print(traceback_str)
            messagebox.showerror("Save Error", f"An error occurred: {str(e)}")

    def view_formulation_data(self):
        """
        Display a window for viewing and managing data from the Master_Viscosity_Data_processed.csv file.
        Allows viewing, filtering, and deleting entries.
        """
        import tkinter as tk
        from tkinter import ttk, messagebox
        import pandas as pd
        import os
        import datetime
    
        # CSV file path
        csv_file = './data/Master_Viscosity_Data_processed.csv'
    
        # Check if the file exists
        if not os.path.exists(csv_file):
            messagebox.showerror("File Not Found", 
                               f"The master CSV file was not found at:\n{csv_file}")
            return
    
        # Create the window
        data_window = tk.Toplevel(self.root)
        data_window.title("Formulation Data Manager")
        data_window.geometry("1000x600")
        data_window.minsize(900, 500)
        data_window.configure(bg=APP_BACKGROUND_COLOR)
    
        # Center the window
        self.gui.center_window(data_window)
    
        # Load the CSV data
        try:
            df = pd.read_csv(csv_file)
            original_df = df.copy()  # Keep a copy for comparison when saving changes
        
            # Convert datetime columns if present
            if 'timestamp' in df.columns:
                df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
        
            # Ensure terpene_pct is in decimal format
            if 'terpene_pct' in df.columns:
                # If any values are >1, assume they're in percentage and convert to decimal
                if (df['terpene_pct'] > 1).any():
                    df.loc[df['terpene_pct'] > 1, 'terpene_pct'] = df.loc[df['terpene_pct'] > 1, 'terpene_pct'] / 100
        
            # Same for total_potency
            if 'total_potency' in df.columns:
                if (df['total_potency'] > 1).any():
                    df.loc[df['total_potency'] > 1, 'total_potency'] = df.loc[df['total_potency'] > 1, 'total_potency'] / 100
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
            return
    
        # Create main frames
        top_frame = tk.Frame(data_window, bg=APP_BACKGROUND_COLOR, padx=10, pady=10)
        top_frame.pack(fill="x")
    
        # Add title
        title_label = tk.Label(top_frame, text="Master Viscosity Data Manager", 
                             font=("Arial", 14, "bold"), bg=APP_BACKGROUND_COLOR, fg="white")
        title_label.pack(pady=5)
    
        # Filter frame
        filter_frame = tk.Frame(top_frame, bg=APP_BACKGROUND_COLOR, padx=10, pady=5)
        filter_frame.pack(fill="x")
    
        # Get unique values for dropdowns
        media_options = ['All'] + sorted([str(x) for x in df['media'].unique()]) if 'media' in df.columns else ['All']
        terpene_options = ['All'] + sorted([str(x) for x in df['terpene'].unique()]) if 'terpene' in df.columns else ['All']
    
        # Add step filter options if available
        step_options = ['All']
        if 'measurement_stage' in df.columns:
            step_options += sorted([str(x) for x in df['measurement_stage'].unique()])
    
        # Filter controls
        tk.Label(filter_frame, text="Filter by:", bg=APP_BACKGROUND_COLOR, fg="white", 
               font=FONT).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    
        # Media filter
        tk.Label(filter_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white").grid(
            row=0, column=1, sticky="w", padx=5, pady=5)
    
        media_var = tk.StringVar(value="All")
        media_combo = ttk.Combobox(filter_frame, textvariable=media_var, values=media_options, width=10)
        media_combo.grid(row=0, column=2, sticky="w", padx=5, pady=5)
    
        # Terpene filter
        tk.Label(filter_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, fg="white").grid(
            row=0, column=3, sticky="w", padx=5, pady=5)
    
        terpene_var = tk.StringVar(value="All")
        terpene_combo = ttk.Combobox(filter_frame, textvariable=terpene_var, values=terpene_options, width=15)
        terpene_combo.grid(row=0, column=4, sticky="w", padx=5, pady=5)
    
        # Step filter
        tk.Label(filter_frame, text="Stage:", bg=APP_BACKGROUND_COLOR, fg="white").grid(
            row=0, column=5, sticky="w", padx=5, pady=5)
    
        step_var = tk.StringVar(value="All")
        step_combo = ttk.Combobox(filter_frame, textvariable=step_var, values=step_options, width=10)
        step_combo.grid(row=0, column=6, sticky="w", padx=5, pady=5)
    
        # Date range filters
        date_frame = tk.Frame(filter_frame, bg=APP_BACKGROUND_COLOR)
        date_frame.grid(row=1, column=0, columnspan=7, sticky="w", pady=5)
    
        tk.Label(date_frame, text="Date Range:", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
    
        today = datetime.datetime.now()
        # Default start date is 30 days ago
        start_date = today - datetime.timedelta(days=10000)
    
        # Format dates for display
        start_date_str = start_date.strftime("%Y-%m-%d")
        end_date_str = today.strftime("%Y-%m-%d")
    
        start_date_var = tk.StringVar(value=start_date_str)
        end_date_var = tk.StringVar(value=end_date_str)
    
        tk.Label(date_frame, text="From:", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
        start_date_entry = ttk.Entry(date_frame, textvariable=start_date_var, width=12)
        start_date_entry.pack(side="left", padx=2)
    
        tk.Label(date_frame, text="To:", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
        end_date_entry = ttk.Entry(date_frame, textvariable=end_date_var, width=12)
        end_date_entry.pack(side="left", padx=2)
    
        # Info about date format
        tk.Label(date_frame, text="(YYYY-MM-DD)", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
    
        # Create main treeview for data display
        tree_frame = tk.Frame(data_window)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
        # Add scrollbars
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        y_scroll.pack(side="right", fill="y")
    
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        x_scroll.pack(side="bottom", fill="x")
    
        # Define columns based on CSV headers
        columns = df.columns.tolist()
    
        # Create treeview
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                          yscrollcommand=y_scroll.set,
                          xscrollcommand=x_scroll.set)
    
        # Configure columns
        for col in columns:
            # Format column headings nicely
            display_name = col.replace('_', ' ').title()
            tree.heading(col, text=display_name)
        
            # Set column widths based on data type
            if col == 'timestamp':
                tree.column(col, width=150)
            elif 'media' in col or 'terpene' in col:
                tree.column(col, width=100)
            elif 'pct' in col or 'potency' in col:
                tree.column(col, width=80)
            else:
                tree.column(col, width=120)
    
        # Link scrollbars
        y_scroll.config(command=tree.yview)
        x_scroll.config(command=tree.xview)
    
        # Pack the treeview
        tree.pack(fill="both", expand=True)
    
        # Button frame
        button_frame = tk.Frame(data_window, bg=APP_BACKGROUND_COLOR, padx=10, pady=10)
        button_frame.pack(fill="x")
    
        # Status variable
        status_var = tk.StringVar(value=f"Loaded {len(df)} records")
        status_label = tk.Label(button_frame, textvariable=status_var, 
                              bg=APP_BACKGROUND_COLOR, fg="white")
        status_label.pack(side="left", padx=5)
    
        # Apply filter function
        def apply_filter():
            # Start with the full dataframe
            filtered_df = df.copy()
            print(f"Starting with {len(filtered_df)} records")  # Debug print
        
            # Apply media filter
            if media_var.get() != "All" and 'media' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['media'] == media_var.get()]
        
            # Apply terpene filter
            if terpene_var.get() != "All" and 'terpene' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['terpene'] == terpene_var.get()]
        
            # Apply step filter
            if step_var.get() != "All" and 'measurement_stage' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['measurement_stage'] == step_var.get()]
        
            # Apply date filter if timestamp column exists
            if 'timestamp' in filtered_df.columns:
                try:
                    start_date = pd.to_datetime(start_date_var.get())
                    end_date = pd.to_datetime(end_date_var.get())
        
                    # Add one day to end_date to include the entire end day
                    end_date = end_date + pd.Timedelta(days=1)
        
                    # Check how many records have valid timestamps versus NaN
                    valid_timestamps = filtered_df['timestamp'].notna().sum()
                    print(f"Records with valid timestamps: {valid_timestamps} out of {len(filtered_df)}")
        
                    # Only apply date filter to records with timestamps
                    # For records without timestamps, keep them regardless of date range
                    mask = (
                        ((filtered_df['timestamp'] >= start_date) & (filtered_df['timestamp'] < end_date)) | 
                        filtered_df['timestamp'].isna()
                    )
                    filtered_df = filtered_df[mask]
                    print(f"After date filter: {len(filtered_df)} records")
        
                except Exception as e:
                    print(f"Date parsing error: {e}")
                    # If date parsing fails, don't apply date filter
        
            # Clear existing data
            for i in tree.get_children():
                tree.delete(i)

            print(f"Final filtered records: {len(filtered_df)}")
        
            # Populate treeview with filtered data
            for idx, row in filtered_df.iterrows():
                values = []
            
                for col in columns:
                    # Format based on column type
                    if col == 'timestamp' and pd.notna(row[col]):
                        if isinstance(row[col], pd.Timestamp):
                            values.append(row[col].strftime('%Y-%m-%d %H:%M'))
                        else:
                            values.append(str(row[col]))
                    elif col in ['terpene_pct', 'total_potency']:
                        # Display as percentage
                        if pd.notna(row[col]):
                            values.append(f"{row[col]*100:.2f}%")
                        else:
                            values.append("")
                    else:
                        values.append(str(row[col]) if pd.notna(row[col]) else "")
            
                tree.insert('', 'end', values=values, tags=('row',))
        
            # Update status
            status_var.set(f"Displaying {len(filtered_df)} of {len(df)} records")
    
        # Function to delete selected records
        def delete_selected():
            nonlocal df
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("No Selection", "Please select records to delete.")
                return

            # Confirm deletion
            if not messagebox.askyesno("Confirm Delete", 
                                        f"Are you sure you want to delete {len(selected)} selected records?"):
                return

            # Create a list to store indices to delete
            to_delete = []
    
            # For debugging
            print(f"Attempting to delete {len(selected)} records")
    
            # Save original dataframe for potential rollback
            original_df_copy = df.copy()
    
            # Process each selected item
            for item_id in selected:
                # Get values from the treeview
                item_values = tree.item(item_id)['values']
        
                # Find matching rows using more flexible matching criteria
                if len(item_values) >= 3:  # Ensure we have enough columns to match
                    # Use key identifiable columns for matching
                    media_value = str(item_values[columns.index('media')] if 'media' in columns else '')
                    terpene_value = str(item_values[columns.index('terpene')] if 'terpene' in columns else '')
                    temp_value = item_values[columns.index('temperature')] if 'temperature' in columns else None
                    visc_value = item_values[columns.index('viscosity')] if 'viscosity' in columns else None
            
                    # Print debug info
                    print(f"Looking for: media={media_value}, terpene={terpene_value}, temp={temp_value}, visc={visc_value}")
            
                    # Create a flexible matching criteria
                    for idx, row in df.iterrows():
                        # Check if critical columns match
                        match = True
                
                        if media_value and 'media' in row and str(row['media']) != media_value:
                            match = False
                    
                        if terpene_value and 'terpene' in row and str(row['terpene']) != terpene_value:
                            match = False
                
                        # For numeric values, use approximate matching
                        if temp_value is not None and 'temperature' in row:
                            try:
                                if abs(float(row['temperature']) - float(temp_value)) > 0.01:
                                    match = False
                            except (ValueError, TypeError):
                                match = False
                
                        if visc_value is not None and 'viscosity' in row:
                            try:
                                # Handle percentage string or numeric value
                                tree_visc = visc_value
                                if isinstance(visc_value, str) and '%' in visc_value:
                                    tree_visc = float(visc_value.replace('%', ''))
                        
                                # Allow for some margin of error in floating point comparisons
                                if abs(float(row['viscosity']) - float(tree_visc)) > 0.1:
                                    match = False
                            except (ValueError, TypeError):
                                match = False
                
                        if match:
                            to_delete.append(idx)
                            print(f"Found match at index {idx}")
                            break

            # Remove duplicates
            to_delete = list(set(to_delete))
    
            if to_delete:
                print(f"Will delete {len(to_delete)} records")
        
                try:
                    # Step 1: Delete from dataframe without leaving gaps
                    # Create a copy of the dataframe with the filtered rows
                    df_updated = df.drop(to_delete).reset_index(drop=True)
            
                    # Step 2: Try to save to the master file
                    master_file = csv_file  # Use the same file path loaded earlier
            
                    # Create backup before modifying anything
                    backup_file = f"{master_file}.bak.{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    original_df.to_csv(backup_file, index=False)
                    print(f"Backup saved to {backup_file}")
            
                    # Try to save the updated dataframe
                    df_updated.to_csv(master_file, index=False)
                    print(f"Updated dataset saved to {master_file}")
            
                    # If we got here, the file save was successful
                    # Update the global dataframe reference
                    
                    df = df_updated
            
                    # Refresh the view
                    apply_filter()
            
                    # Update status
                    status_var.set(f"Deleted {len(to_delete)} records. {len(df)} records remaining. File updated.")
            
                except Exception as e:
                    # If any error occurs during the save, roll back to the original dataframe
                    error_msg = f"Error updating file: {str(e)}\nChanges have been reverted to maintain consistency."
                    print(error_msg)
            
                    # Restore the original dataframe
                    
                    df = original_df_copy
            
                    # Show error message
                    messagebox.showerror("Save Error", error_msg)
            
                    # Refresh the view with original data
                    apply_filter()
            
            else:
                print("No matching records found for deletion")
                messagebox.showwarning("No Matches", "Could not identify the selected records in the dataset.")
    
        # Function to save changes to the CSV file
        def save_changes():
            nonlocal original_df
            if len(df) == len(original_df) and all(df.eq(original_df).all()):
                messagebox.showinfo("No Changes", "No changes have been made to the data.")
                return
        
            # Confirm save
            if not messagebox.askyesno("Confirm Save", 
                                        "Save changes to the master CSV file?\nThis will overwrite the existing file."):
                return
        
            try:
                # Create backup
                backup_path = f"{csv_file}.bak"
                original_df.to_csv(backup_path, index=False)
            
                # Save changes
                df.to_csv(csv_file, index=False)
            
                messagebox.showinfo("Save Complete", 
                                    f"Changes saved successfully to {csv_file}\nBackup created at {backup_path}")
            
                # Update original_df reference
                
                original_df = df.copy()
            except Exception as e:
                messagebox.showerror("Save Error", f"Error saving changes: {str(e)}")
    
        # Function to export selected records
        def export_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("No Selection", "Please select records to export.")
                return
        
            # Ask for file location
            from tkinter import filedialog
            export_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Selected Records"
            )
        
            if not export_path:
                return  # User cancelled
        
            try:
                # Get selected indices
                selected_indices = []
                for item_id in selected:
                    item_values = tree.item(item_id)['values']
                
                    # Match with dataframe rows (same logic as delete function)
                    typed_values = []
                    for i, col in enumerate(columns):
                        val = item_values[i]
                    
                        if col in ['terpene_pct', 'total_potency'] and isinstance(val, str) and '%' in val:
                            val = float(val.replace('%', '')) / 100
                    
                        if col == 'timestamp' and isinstance(val, str):
                            try:
                                val = pd.to_datetime(val)
                            except:
                                pass
                    
                        typed_values.append(val)
                
                    for idx, row in df.iterrows():
                        match = True
                        for i, col in enumerate(columns):
                            row_val = row[col]
                            tree_val = typed_values[i]
                        
                            if col == 'timestamp':
                                if isinstance(row_val, pd.Timestamp) and isinstance(tree_val, str):
                                    if row_val.strftime('%Y-%m-%d %H:%M') != tree_val:
                                        match = False
                                        break
                                elif row_val != tree_val:
                                    match = False
                                    break
                            elif col in ['terpene_pct', 'total_potency']:
                                row_float = float(row_val) if pd.notna(row_val) else 0
                                tree_float = float(tree_val.replace('%', '')) / 100 if isinstance(tree_val, str) and '%' in tree_val else float(tree_val) if pd.notna(tree_val) else 0
                                if abs(row_float - tree_float) > 0.0001:
                                    match = False
                                    break
                            elif str(row_val) != str(tree_val):
                                match = False
                                break
                    
                        if match:
                            selected_indices.append(idx)
                            break
            
                # Remove duplicates
                selected_indices = list(set(selected_indices))
            
                if selected_indices:
                    # Create export dataframe
                    export_df = df.loc[selected_indices].copy()
                
                    # Save to CSV
                    export_df.to_csv(export_path, index=False)
                
                    messagebox.showinfo("Export Complete", 
                                        f"Successfully exported {len(export_df)} records to {export_path}")
                else:
                    messagebox.showwarning("No Matches", "Could not identify the selected records for export.")
            except Exception as e:
                messagebox.showerror("Export Error", f"Error exporting data: {str(e)}")
    
        # Add buttons for filter apply and refresh
        filter_button = ttk.Button(filter_frame, text="Apply Filter", command=apply_filter)
        filter_button.grid(row=0, column=7, padx=5, pady=5)
    
        # Add action buttons
        delete_button = ttk.Button(button_frame, text="Delete Selected", command=delete_selected)
        delete_button.pack(side="right", padx=5)
    
        save_button = ttk.Button(button_frame, text="Save Changes", command=save_changes)
        save_button.pack(side="right", padx=5)
    
        export_button = ttk.Button(button_frame, text="Export Selected", command=export_selected)
        export_button.pack(side="right", padx=5)
    
        # Apply initial filter
        apply_filter()
    
        # Add alternating row colors
        tree.tag_configure('row', background='#f0f0ff')
    
        # Set multiple selection mode
        tree.config(selectmode='extended')
    
        # Make the window modal
        data_window.transient(self.root)
        data_window.grab_set()
        data_window.focus_set()   

    # ---- Temperature Block Methods (Measure Tab) ----

    def create_temperature_block(self, parent, temperature):
        """
        Create a block for a temperature with a table for 3 runs
        
        Args:
            parent: Parent frame to add the block to
            temperature: Temperature value for this block
        """
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = '#0504AA'
        
        # Create a frame for this temperature block with a border
        block_frame = Frame(parent, bg=APP_BACKGROUND_COLOR, bd=1, relief="solid")
        block_frame.pack(fill="x", expand=True, pady=5, padx=5)

        # Track this block
        self.temperature_blocks.append((temperature, block_frame))

        # Temperature header row
        temp_header = Frame(block_frame, bg=APP_BACKGROUND_COLOR)
        temp_header.pack(fill="x", padx=2, pady=2)

        temp_label = Label(temp_header, text=f"{temperature}C", 
                        bg="#000080", fg="white", font=FONT, width=10)
        temp_label.pack(side="left", fill="x", expand=True)

        # Add a remove button for this block
        remove_btn = ttk.Button(
            temp_header,
            text="x",
            width=3,
            command=lambda t=temperature, bf=block_frame: self.remove_temperature_block(t, bf)
        )
        remove_btn.pack(side="right", padx=2)

        # Speed input row
        speed_frame = Frame(block_frame, bg=APP_BACKGROUND_COLOR)
        speed_frame.pack(fill="x", padx=2, pady=2)

        Label(speed_frame, text="Speed:", bg=APP_BACKGROUND_COLOR, fg="white", 
            font=FONT).pack(side="left", padx=5)

        speed_var = StringVar(value="")
        speed_entry = Entry(speed_frame, textvariable=speed_var, width=15)
        speed_entry.pack(side="left", padx=5)

        Label(speed_frame, text="(manual input)", bg=APP_BACKGROUND_COLOR, fg="white", 
            font=FONT).pack(side="left", padx=5)

        self.speed_vars.append((temperature, speed_var))

        # Create the table for runs
        table_frame = Frame(block_frame, bg=APP_BACKGROUND_COLOR)
        table_frame.pack(fill="x", padx=5, pady=5)

        # Table headers
        headers = ["", "Torque", "Viscosity"]
        col_widths = [10, 15, 15]

        # Create header row
        for col, header in enumerate(headers):
            Label(table_frame, text=header, bg="#000080", fg="white", 
                font=FONT, width=col_widths[col], relief="raised").grid(
                row=0, column=col, sticky="nsew", padx=1, pady=1)

        # Create rows for each run
        for run in range(3):
            # Row label (Run 1, Run 2, Run 3)
            Label(table_frame, text=f"Run {run+1}", bg=APP_BACKGROUND_COLOR, fg="white", 
                font=FONT, width=col_widths[0]).grid(
                row=run+1, column=0, sticky="nsew", padx=1, pady=1)
        
            # Create and track variables for this run
            torque_var = StringVar(value="")
            visc_var = StringVar(value="")
        
            # Make sure we have lists for this temperature
            while len(self.torque_vars) <= run:
                self.torque_vars.append([])
            while len(self.viscosity_vars) <= run:
                self.viscosity_vars.append([])
        
            # Store the variables with temperature for tracking
            self.torque_vars[run].append((temperature, torque_var))
            self.viscosity_vars[run].append((temperature, visc_var))
        
            # Create entry for torque
            torque_entry = Entry(table_frame, textvariable=torque_var, width=col_widths[1])
            torque_entry.grid(row=run+1, column=1, sticky="nsew", padx=1, pady=1)
        
            # Create entry for viscosity with event binding
            visc_entry = Entry(table_frame, textvariable=visc_var, width=col_widths[2])
            visc_entry.grid(row=run+1, column=2, sticky="nsew", padx=1, pady=1)
            
            # Add event binding for value changes in viscosity field
            visc_var.trace_add("write", lambda name, index, mode, temp=temperature: 
                            self.check_run_completion(temp))

        # Add average row
        Label(table_frame, text="Average", bg="#000080", fg="white", 
            font=FONT, width=col_widths[0]).grid(
            row=4, column=0, sticky="nsew", padx=1, pady=1)

        # Create variables for averages
        avg_torque_var = StringVar(value="")
        avg_visc_var = StringVar(value="")

        # Store them for tracking
        if not hasattr(self, 'avg_torque_vars'):
            self.avg_torque_vars = []
        if not hasattr(self, 'avg_visc_vars'):
            self.avg_visc_vars = []

        self.avg_torque_vars.append((temperature, avg_torque_var))
        self.avg_visc_vars.append((temperature, avg_visc_var))

        # Create labels for averages
        Label(table_frame, textvariable=avg_torque_var, bg="#90EE90", 
            width=col_widths[1]).grid(row=4, column=1, sticky="nsew", padx=1, pady=1)
        Label(table_frame, textvariable=avg_visc_var, bg="#90EE90", 
            width=col_widths[2]).grid(row=4, column=2, sticky="nsew", padx=1, pady=1)

    def check_run_completion(self, temperature):
        """
        Check if all three runs have viscosity values for a given temperature.
        If so, automatically calculate the averages.
        
        Args:
            temperature: The temperature block to check
        """
        # Find all viscosity variables for this temperature
        visc_values = []
        torque_values = []
        
        # Check all runs for this temperature
        for run in range(3):
            # Get viscosity values
            for temp, visc_var in self.viscosity_vars[run]:
                if temp == temperature:
                    try:
                        value = visc_var.get().strip()
                        if value:  # Check if not empty
                            # Convert to float, handling commas if present
                            visc_float = float(value.replace(',', ''))
                            visc_values.append(visc_float)
                    except (ValueError, AttributeError) as e:
                        print(f"Error converting viscosity value: {e}")
            
            # Get torque values
            for temp, torque_var in self.torque_vars[run]:
                if temp == temperature:
                    try:
                        torque_value = torque_var.get().strip()
                        if torque_value:  # Check if not empty
                            # Convert to float, handling commas if present
                            torque_float = float(torque_value.replace(',', ''))
                            torque_values.append(torque_float)
                    except (ValueError, AttributeError) as e:
                        print(f"Error converting torque value: {e}")
        
        # If we have three viscosity values, calculate average
        if len(visc_values) == 3:
            try:
                # Calculate viscosity average
                avg_visc = sum(visc_values) / len(visc_values)
                
                # Update the average viscosity variable
                for temp, avg_var in self.avg_visc_vars:
                    if temp == temperature:
                        # Format with commas for larger numbers
                        if avg_visc >= 1000:
                            avg_var.set(f"{avg_visc:,.1f}")
                        else:
                            avg_var.set(f"{avg_visc:.1f}")
                        break
                
                # Calculate torque average if we have values
                if torque_values:
                    avg_torque = sum(torque_values) / len(torque_values)
                    
                    # Update the average torque variable
                    for temp, avg_var in self.avg_torque_vars:
                        if temp == temperature:
                            avg_var.set(f"{avg_torque:.1f}")
                            break
                            
                # Force update of the UI
                self.root.update_idletasks()
                
            except Exception as e:
                print(f"Error calculating averages: {e}")
    
    def add_temperature_block(self, parent):
        """Add a new temperature block with a user-specified temperature"""
        # Ask user for the temperature
        from tkinter import simpledialog
        new_temp = simpledialog.askinteger("Temperature", "Enter temperature (C):",
                                         initialvalue=25, minvalue=0, maxvalue=100)
        if new_temp is not None:
            # Check if this temperature already exists
            temp_values = [t for t, _ in self.temperature_blocks]
            if new_temp in temp_values:
                messagebox.showwarning("Warning", f"A block for {new_temp}C already exists.")
                return
        
            # Create a new block for this temperature
            self.create_temperature_block(parent, new_temp)

    def remove_temperature_block(self, temperature, block_frame):
        """Remove a temperature block"""
        # Remove from tracking lists
        self.temperature_blocks = [(t, bf) for t, bf in self.temperature_blocks if t != temperature]
        self.speed_vars = [(t, v) for t, v in self.speed_vars if t != temperature]
    
        for run in range(len(self.torque_vars)):
            self.torque_vars[run] = [(t, v) for t, v in self.torque_vars[run] if t != temperature]
            self.viscosity_vars[run] = [(t, v) for t, v in self.viscosity_vars[run] if t != temperature]
    
        self.avg_torque_vars = [(t, v) for t, v in self.avg_torque_vars if t != temperature]
        self.avg_visc_vars = [(t, v) for t, v in self.avg_visc_vars if t != temperature]
    
        # Destroy the block frame
        block_frame.destroy()
    
    def calculate_viscosity_block_stats(self):
        """Calculate averages for each temperature block"""
        
        calculations_performed = False
        
        for temp, _ in self.temperature_blocks:
            # Get the torque and viscosity values for this temperature
            torque_values = []
            visc_values = []
        
            for run in range(3):
                # Collect viscosity values for this temperature
                for t, visc_var in self.viscosity_vars[run]:
                    if t == temp:
                        try:
                            visc_value = visc_var.get().strip()
                            if visc_value:  # Check if not empty
                                visc = float(visc_value.replace(',', ''))
                                visc_values.append(visc)
                        except ValueError:
                            pass
                
                # Collect torque values for this temperature
                for t, torque_var in self.torque_vars[run]:
                    if t == temp:
                        try:
                            torque_value = torque_var.get().strip()
                            if torque_value:  # Check if not empty
                                torque = float(torque_value.replace(',', ''))
                                torque_values.append(torque)
                        except ValueError:
                            pass
        
            # Calculate averages if we have values
            if torque_values:
                avg_torque = sum(torque_values) / len(torque_values)
                # Find the average torque variable for this temperature
                for t, avg_var in self.avg_torque_vars:
                    if t == temp:
                        avg_var.set(f"{avg_torque:.1f}")
                        calculations_performed = True
                        break
        
            if visc_values:
                avg_visc = sum(visc_values) / len(visc_values)
                # Find the average viscosity variable for this temperature
                for t, avg_var in self.avg_visc_vars:
                    if t == temp:
                        # Format with commas for larger numbers
                        if avg_visc >= 1000:
                            avg_var.set(f"{avg_visc:,.1f}")
                        else:
                            avg_var.set(f"{avg_visc:.1f}")
                        calculations_performed = True
                        break

        # Show a message to confirm calculation
        if calculations_performed:
            messagebox.showinfo("Calculation Complete", "Averages have been calculated successfully.")

    def save_block_measurements(self):
        """Save the block-based viscosity measurements to the database with better error handling"""
        import datetime
        import os
        import traceback
        
        try:
            # Get the terpene value, defaulting to "Raw" if empty
            terpene_value = self.measure_terpene_var.get().strip()
            if not terpene_value:
                terpene_value = "Raw"
            
            # Get the terpene percentage, defaulting to 0.0 if empty
            try:
                terpene_pct = float(self.measure_terpene_pct_var.get())
            except (ValueError, tk.TclError):
                terpene_pct = 0.0
            
            # Create a data structure to save
            measurements = {
                "media": self.media_var.get(),
                "terpene": terpene_value,
                "terpene_pct": terpene_pct,
                "timestamp": datetime.datetime.now().isoformat(),
                "temperature_data": []
            }
            
            # Validate we have at least one temperature block with data
            if not self.temperature_blocks:
                messagebox.showwarning("Missing Data", "No temperature blocks found. Please add measurement blocks first.")
                return
            
            # Collect data from each temperature block
            for temp, _ in self.temperature_blocks:
                # Find the speed for this temperature
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
                
                # Collect data for each run
                for run in range(3):
                    run_data = {"torque": "", "viscosity": ""}
                    
                    # Find the torque and viscosity for this run at this temperature
                    for t, torque_var in self.torque_vars[run]:
                        if t == temp:
                            run_data["torque"] = torque_var.get()
                            break
                    
                    for t, visc_var in self.viscosity_vars[run]:
                        if t == temp:
                            run_data["viscosity"] = visc_var.get()
                            break
                    
                    block_data["runs"].append(run_data)
                
                # Find the averages
                avg_torque = ""
                avg_visc = ""
                
                for t, avg_var in self.avg_torque_vars:
                    if t == temp:
                        avg_torque = avg_var.get()
                        break
                
                for t, avg_var in self.avg_visc_vars:
                    if t == temp:
                        avg_visc = avg_var.get()
                        break
                
                block_data["average_torque"] = avg_torque
                block_data["average_viscosity"] = avg_visc
                
                measurements["temperature_data"].append(block_data)
            
            # Check if measurements has any valid viscosity data
            has_valid_data = False
            for block in measurements["temperature_data"]:
                if block["average_viscosity"] and block["average_viscosity"].strip():
                    has_valid_data = True
                    break
                for run in block["runs"]:
                    if run["viscosity"] and run["viscosity"].strip():
                        has_valid_data = True
                        break
            
            if not has_valid_data:
                messagebox.showwarning("Missing Data", 
                                    "No valid viscosity measurements found. Please enter at least one viscosity value.")
                return
            
            # Add measurements to the database
            key = f"{measurements['media']}_{measurements['terpene']}_{measurements['terpene_pct']}"
            
            if key not in self.formulation_db:
                self.formulation_db[key] = []
            
            self.formulation_db[key].append(measurements)
            
            # Save to file
            self.save_formulation_database()
            
            # Also save as CSV for machine learning
            self.save_as_csv(measurements)
            
            messagebox.showinfo("Success", "Viscosity measurements saved successfully!")
            
        except Exception as e:
            traceback_str = traceback.format_exc()
            error_msg = f"Error processing measurements: {str(e)}"
            print(error_msg)
            print(traceback_str)
            messagebox.showerror("Save Error", error_msg)

    def save_as_csv(self, measurements):
        """
        Save the measurements to the master CSV file.
    
        Args:
            measurements (dict): The measurements data structure
        """
        # Lazy import pandas only when needed
        import pandas as pd
        import os
        import datetime
    
        try:
            # Load existing master file if it exists
            master_file = './data_scraper/Master_Viscosity_Data_processed.csv'
            if os.path.exists(master_file):
                master_df = pd.read_csv(master_file)
            else:
                # Create directory if it doesn't exist
                os.makedirs('./data_scraper', exist_ok=True)
                master_df = pd.DataFrame()
        
            # Create rows for the CSV
            rows = []
        
            media = measurements['media']
            terpene = measurements['terpene']
            terpene_pct = measurements['terpene_pct']
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
            # Process each temperature block
            for temp_block in measurements['temperature_data']:
                temperature = temp_block['temperature']
                speed = temp_block.get('speed', '')
            
                # Process each individual run
                for run in temp_block['runs']:
                    torque = run.get('torque', '')
                    viscosity = run.get('viscosity', '')
                
                    # Only add rows with valid viscosity values
                    if viscosity and viscosity.strip():
                        try:
                            viscosity_float = float(viscosity.replace(',', ''))
                            row = {
                                'media': media,
                                'media_brand': measurements.get('media_brand', ''),
                                'terpene': terpene,
                                'terpene_brand': measurements.get('terpene_brand', ''),
                                'terpene_pct': terpene_pct,
                                'temperature': temperature,
                                'speed': speed,
                                'torque': torque,
                                'viscosity': viscosity_float,
                                'timestamp': timestamp
                            }
                            rows.append(row)
                        except ValueError as e:
                            print(f"Warning: Could not convert viscosity value '{viscosity}' to float: {e}")
            
                # Add the average if available
                avg_viscosity = temp_block.get('average_viscosity', '')
                if avg_viscosity and avg_viscosity.strip():
                    try:
                        avg_viscosity_float = float(avg_viscosity.replace(',', ''))
                        row = {
                            'media': media,
                            'media_brand': measurements.get('media_brand', ''),
                            'terpene': terpene,
                            'terpene_brand': measurements.get('terpene_brand', ''),
                            'terpene_pct': terpene_pct,
                            'temperature': temperature,
                            'speed': speed,
                            'torque': temp_block.get('average_torque', ''),
                            'viscosity': avg_viscosity_float,
                            'is_average': True,
                            'timestamp': timestamp
                        }
                        rows.append(row)
                    except ValueError as e:
                        print(f"Warning: Could not convert average viscosity value '{avg_viscosity}' to float: {e}")
        
            # Create a DataFrame and append to master file
            if rows:
                new_df = pd.DataFrame(rows)
                if master_df.empty:
                    master_df = new_df
                else:
                    # Ensure all columns exist in both dataframes
                    for col in new_df.columns:
                        if col not in master_df.columns:
                            master_df[col] = None
                    for col in master_df.columns:
                        if col not in new_df.columns:
                            new_df[col] = None
                
                    master_df = pd.concat([master_df, new_df], ignore_index=True)
            
                # Save back to master file
                master_df.to_csv(master_file, index=False)
            
                messagebox.showinfo("Success", 
                                   f"Added {len(rows)} new measurements to master data file.")
                return master_file
            else:
                print("No valid data rows to save")
                return None
            
        except Exception as e:
            import traceback
            print(f"Error saving to master file: {e}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Failed to save measurements: {str(e)}")
            return None

    # ---- Advanced Data Analysis Methods (Lazily Loaded) ----

    def load_models(self):
        """
        Load trained viscosity prediction models from disk.
    
        Returns:
            dict: Dictionary of trained models for different media/terpene combinations
        """
        import pickle
        import os
    
        model_path = 'models/viscosity_models.pkl'
        if os.path.exists(model_path):
            try:
                with open(model_path, 'rb') as f:
                    return pickle.load(f)
            except Exception as e:
                print(f"Error loading models: {e}")
        return {}

    def upload_training_data(self):
        """
        Allow the user to upload the master CSV data file for training viscosity models.
    
        Only accepts the Master_Viscosity_Data_processed.csv file
        """
        from tkinter import filedialog
        import pandas as pd
        import os
        import shutil

        # Prompt user to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select Master Viscosity Data CSV",
            filetypes=[("CSV files", "*.csv")]
        )

        if not file_path:
            return None  # User canceled

        # Check if this is the correct file
        file_name = os.path.basename(file_path)
        if not file_name.startswith("Master_Viscosity_Data"):
            messagebox.showerror("Invalid File", 
                               "Only Master_Viscosity_Data_processed.csv is accepted.\n"
                               "Please select the correct file.")
            return None

        try:
            # Load the data
            data = pd.read_csv(file_path)
    
            # Validate the data has the required columns
            required_cols = ['media', 'terpene', 'terpene_pct', 'temperature', 'viscosity']
            missing_cols = [col for col in required_cols if col not in data.columns]
    
            if missing_cols:
                messagebox.showerror("Error", 
                                   f"CSV missing required columns: {', '.join(missing_cols)}")
                return None
    
            # Copy the file to the data_scraper directory
            os.makedirs('./data_scraper', exist_ok=True)
            dest_path = './data_scraper/Master_Viscosity_Data_processed.csv'
            shutil.copy2(file_path, dest_path)
    
            messagebox.showinfo("Success", 
                              f"Loaded {len(data)} data points from {os.path.basename(file_path)}.\n"
                              f"File copied to {dest_path}")
    
            return data
    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
            return None

    def get_actual_model(self,model_obj):
        """Extract the actual model from a dictionary if needed"""
        if isinstance(model_obj, dict) and 'model' in model_obj:
            return model_obj['model']
        return model_obj

    def analyze_models(self, show_dialog=True):
        """Analyze trained two-step viscosity models with a focus on residual performance."""
        import pandas as pd
        import numpy as np
        import matplotlib.pyplot as plt
        from tkinter import Toplevel, Text, Scrollbar, Label, Frame, Button
        import os
        import pickle
        import traceback
        from sklearn.metrics import mean_squared_error, r2_score
        from sklearn.impute import SimpleImputer  # Add this import for NaN handling

        # Check if models exist
        standard_models_exist = hasattr(self, 'viscosity_models') and self.viscosity_models
        enhanced_models_exist = hasattr(self, 'enhanced_viscosity_models') and self.enhanced_viscosity_models
        combined_models_exist = hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models

        if not standard_models_exist and not enhanced_models_exist and not combined_models_exist:
            messagebox.showinfo("No Models", "No trained models found. Please train models first.")
            return

        # Create analysis report
        report = []
        report.append("Two-Step Model Analysis Report")
        report.append("=============================")

        # Load validation data
        validation_data = None
        try:
            data_dir = 'data'
            if os.path.exists(data_dir):
                validation_files = [f for f in os.listdir(data_dir) 
                                  if (f.startswith('viscosity_data_') or f.startswith('Master_Viscosity_Data_')) 
                                  and f.endswith('.csv')]

                if validation_files:
                    # Use most recent file
                    validation_files.sort(reverse=True)
                    latest_file = os.path.join('data', validation_files[0])
                    validation_data = pd.read_csv(latest_file)
                    report.append(f"Using {os.path.basename(latest_file)} for validation")
    
                    # Prepare data for Arrhenius models
                    validation_data['temperature_kelvin'] = validation_data['temperature'] + 273.15
                    validation_data['inverse_temp'] = 1 / validation_data['temperature_kelvin']
                    validation_data['log_viscosity'] = np.log(validation_data['viscosity'])
    
                    # Add is_raw flag if missing
                    if 'is_raw' not in validation_data.columns:
                        validation_data['is_raw'] = validation_data['terpene'].isna() | (validation_data['terpene'] == '')
        
                    # Clean up terpene values
                    validation_data.loc[validation_data['terpene'].isna(), 'terpene'] = 'Raw'
                    validation_data.loc[validation_data['terpene'] == '', 'terpene'] = 'Raw'
            
                    # Create model_feature columns if primary features exist
                    if 'total_potency' in validation_data.columns:
                        validation_data['model_feature'] = validation_data['total_potency']
                    elif 'terpene_pct' in validation_data.columns:
                        validation_data['model_feature'] = validation_data['terpene_pct']
        except Exception as e:
            report.append(f"Error loading validation data: {str(e)}")

        # Define a helper function to check model structure
        def is_valid_model(model):
            return (isinstance(model, dict) and 
                    'baseline_model' in model and 
                    'residual_model' in model and
                    'baseline_features' in model and
                    'residual_features' in model)

        # Analyze standard models
        if standard_models_exist:
            report.append(f"\nStandard Models (terpene-based): {len(self.viscosity_models)}")
            report.append("-" * 40)
        
            # Check model format and offer to retrain if needed
            should_retrain = False
            if self.viscosity_models:
                sample_key = next(iter(self.viscosity_models))
                sample_model = self.viscosity_models[sample_key]
                if not is_valid_model(sample_model):
                    should_retrain = messagebox.askyesno("Retrain Models", 
                                                        "Your standard models have an incompatible format. Would you like to retrain them?")
                
            if should_retrain:
                self.train_unified_models()
                return  # Exit early as we'll reanalyze with new models
        
            # Process each model
            for model_key, model in self.viscosity_models.items():
                report.append(f"\nModel: {model_key}")
            
                try:
                    # Check if model has the correct structure
                    if not is_valid_model(model):
                        report.append("Error analyzing model: incompatible model format")
                        continue
                    
                    # Extract model components
                    baseline_model = model['baseline_model']
                    residual_model = model['residual_model']
                    baseline_features = model['baseline_features']
                    residual_features = model['residual_features']
            
                    # Get metadata
                    metadata = model.get('metadata', {})
                    primary_feature = metadata.get('primary_feature', 'terpene_pct')
        
                    # Analyze baseline model (Arrhenius)
                    report.append("1. Temperature baseline model (Arrhenius)")
        
                    if hasattr(baseline_model, 'coef_'):
                        # Extract Arrhenius parameters
                        coef = baseline_model.coef_[0]
                        intercept = baseline_model.intercept_
            
                        # Calculate activation energy
                        R = 8.314  # Gas constant (J/mol·K)
                        Ea = coef * R
                        Ea_kJ = Ea / 1000  # Convert to kJ/mol
            
                        report.append(f"  - Equation: log(viscosity) = {intercept:.4f} + {coef:.4f} * (1/T)")
                        report.append(f"  - Activation energy: {Ea_kJ:.2f} kJ/mol")
        
                    # Analyze residual model (terpene effects)
                    report.append("2. Residual model (terpene effects)")
                    report.append(f"  - Model type: {type(residual_model).__name__}")
        
                    if hasattr(residual_model, 'feature_importances_'):
                        # Report feature importance
                        importances = residual_model.feature_importances_
            
                        if len(importances) == len(residual_features):
                            report.append("  - Feature importances:")
                            for i, feature in enumerate(residual_features):
                                report.append(f"    * {feature}: {importances[i]:.4f}")
            
                        # Check terpene importance
                        if primary_feature in residual_features:
                            feature_idx = residual_features.index(primary_feature)
                            feature_importance = importances[feature_idx]
                
                            if feature_importance > 0.3:
                                report.append(f"  - {primary_feature} has HIGH importance in viscosity deviation")
                            elif feature_importance > 0.1:
                                report.append(f"  - {primary_feature} has MODERATE importance in viscosity deviation")
                            elif feature_importance > 0.01:
                                report.append(f"  - {primary_feature} has LOW importance in viscosity deviation")
                            else:
                                report.append(f"  - {primary_feature} has NEGLIGIBLE importance in viscosity deviation")
        
                    elif hasattr(residual_model, 'coef_'):
                        # For linear models
                        coefs = residual_model.coef_
                        intercept = residual_model.intercept_
            
                        if len(coefs) == len(residual_features):
                            report.append("  - Coefficients:")
                            report.append(f"    * Intercept: {intercept:.4f}")
                            for i, feature in enumerate(residual_features):
                                report.append(f"    * {feature}: {coefs[i]:.4f}")
        
                    # Validate with data if available
                    if validation_data is not None:
                        try:
                            # Extract media/terpene from model key
                            components = model_key.split('_', 1)
                            media = components[0]
                
                            # Handle general models
                            if len(components) > 1 and components[1] == 'general':
                                model_val_data = validation_data[validation_data['media'] == media].copy()
                    
                                # One-hot encode terpenes
                                model_val_data = pd.get_dummies(
                                    model_val_data, 
                                    columns=['terpene'],
                                    prefix=['terpene']
                                )
                            else:
                                # Extract terpene from key
                                terpene = components[1] if len(components) > 1 else 'Raw'
                                model_val_data = validation_data[
                                    (validation_data['media'] == media) & 
                                    (validation_data['terpene'] == terpene)
                                ].copy()
                
                            if len(model_val_data) >= 5:
                                report.append(f"\nValidation results ({len(model_val_data)} samples):")
                    
                                # Drop NaN values in key features
                                model_val_data = model_val_data.dropna(subset=['inverse_temp', 'log_viscosity'])
                            
                                # Ensure primary feature exists
                                if primary_feature not in model_val_data.columns:
                                    if primary_feature == 'terpene_pct' and 'terpene_pct' in validation_data.columns:
                                        model_val_data['terpene_pct'] = validation_data.loc[model_val_data.index, 'terpene_pct']
                            
                                # Filter out rows with NaN in primary feature
                                model_val_data = model_val_data.dropna(subset=[primary_feature])
                        
                                if len(model_val_data) < 5:
                                    report.append(f"Insufficient data after removing NaN values. Only {len(model_val_data)} samples left.")
                                    continue
                            
                                # Create model_feature column if needed
                                if 'model_feature' in residual_features and 'model_feature' not in model_val_data.columns:
                                    model_val_data['model_feature'] = model_val_data[primary_feature]
                    
                                # Step 1: Evaluate baseline model
                                X_baseline = model_val_data[baseline_features]
                                y_true = model_val_data['log_viscosity']
                    
                                # Get baseline predictions
                                baseline_preds = baseline_model.predict(X_baseline)
                    
                                # Calculate residuals
                                model_val_data['baseline_pred'] = baseline_preds
                                model_val_data['residual'] = y_true - baseline_preds
                    
                                # Step 2: Evaluate residual model
                                # Check if we have all required features
                                missing_features = [f for f in residual_features if f not in model_val_data.columns]
                            
                                # Better handling of missing features
                                if missing_features:
                                    if 'model_feature' in missing_features and primary_feature in model_val_data.columns:
                                        # Create model_feature explicitly from primary feature
                                        model_val_data['model_feature'] = model_val_data[primary_feature]
                                        missing_features.remove('model_feature')
                                
                                    if 'potency_terpene_ratio' in missing_features and 'total_potency' in model_val_data.columns and 'terpene_pct' in model_val_data.columns:
                                        # Create ratio with safety for division by zero
                                        model_val_data['potency_terpene_ratio'] = model_val_data['total_potency'] / model_val_data['terpene_pct'].clip(lower=0.01)
                                        missing_features.remove('potency_terpene_ratio')

                                    # If we still have missing features, report and skip
                                    if missing_features:
                                        report.append(f"Missing required features for validation: {', '.join(missing_features)}")
                                        continue
                    
                                # Extract features and residuals
                                X_residual = model_val_data[residual_features]
                                y_residual = model_val_data['residual']
                            
                                # Handle NaN values in the feature data
                                if X_residual.isna().any().any():
                                    # Use SimpleImputer to replace NaN values with mean
                                    imputer = SimpleImputer(strategy='mean')
                                    X_residual_values = imputer.fit_transform(X_residual)
                                    X_residual = pd.DataFrame(X_residual_values, 
                                                              index=X_residual.index, 
                                                              columns=X_residual.columns)
                    
                                # Get residual predictions
                                residual_preds = residual_model.predict(X_residual)
                    
                                # Calculate metrics for residual model
                                r2_residual = r2_score(y_residual, residual_preds)
                                mse_residual = mean_squared_error(y_residual, residual_preds)
                    
                                report.append(f"Residual model - MSE: {mse_residual:.2f}, R^2: {r2_residual:.4f}")
                    
                                # Combined prediction metrics
                                combined_preds = baseline_preds + residual_preds
                    
                                # Log scale metrics
                                r2_log = r2_score(y_true, combined_preds)
                                mse_log = mean_squared_error(y_true, combined_preds)
                    
                                # Original scale metrics
                                y_orig = np.exp(y_true)
                                preds_orig = np.exp(combined_preds)
                    
                                r2_orig = r2_score(y_orig, preds_orig)
                                mse_orig = mean_squared_error(y_orig, preds_orig)
                    
                                report.append(f"Log scale - MSE: {mse_log:.2f}, R^2: {r2_log:.4f}")
                                report.append(f"Original scale - MSE: {mse_orig:.2f}, R^2: {r2_orig:.4f}")
                    
                                if r2_orig < 0.5:
                                    report.append("WARNING: Poor validation performance (R^2 < 0.5)!")
                            else:
                                report.append(f"Insufficient validation data: {len(model_val_data)} samples")
                        except Exception as e:
                            report.append(f"Error validating model: {str(e)}")
                            report.append(traceback.format_exc())
                except Exception as e:
                    report.append(f"Error analyzing model: {str(e)}")
                    report.append(traceback.format_exc())

        # Analyze enhanced models using same approach
        if enhanced_models_exist:
            report.append(f"\nEnhanced Models (potency-based): {len(self.enhanced_viscosity_models)}")
            report.append("-" * 40)

            for model_key, model in self.enhanced_viscosity_models.items():
                report.append(f"\nModel: {model_key}")
    
                try:
                    # Check if model has the correct structure
                    if not is_valid_model(model):
                        report.append("Error analyzing model: incompatible model format")
                        continue
                    
                    # Extract model components
                    baseline_model = model['baseline_model']
                    residual_model = model['residual_model']
                    baseline_features = model['baseline_features']
                    residual_features = model['residual_features']
            
                    # Get metadata
                    metadata = model.get('metadata', {})
                    primary_feature = metadata.get('primary_feature', 'total_potency')
        
                    # Analyze baseline model (Arrhenius)
                    report.append("1. Temperature baseline model (Arrhenius)")
        
                    if hasattr(baseline_model, 'coef_'):
                        # Extract Arrhenius parameters
                        coef = baseline_model.coef_[0]
                        intercept = baseline_model.intercept_
            
                        # Calculate activation energy
                        R = 8.314  # Gas constant (J/mol·K)
                        Ea = coef * R
                        Ea_kJ = Ea / 1000  # Convert to kJ/mol
            
                        report.append(f"  - Equation: log(viscosity) = {intercept:.4f} + {coef:.4f} * (1/T)")
                        report.append(f"  - Activation energy: {Ea_kJ:.2f} kJ/mol")
        
                    # Analyze residual model (potency effects)
                    report.append("2. Residual model (potency and chemistry effects)")
                    report.append(f"  - Model type: {type(residual_model).__name__}")
        
                    if hasattr(residual_model, 'feature_importances_'):
                        # Report feature importance
                        importances = residual_model.feature_importances_
            
                        if len(importances) == len(residual_features):
                            report.append("  - Feature importances:")
                            for i, feature in enumerate(residual_features):
                                report.append(f"    * {feature}: {importances[i]:.4f}")
            
                        # Check potency importance
                        if primary_feature in residual_features and 'terpene_pct' in residual_features:
                            potency_idx = residual_features.index(primary_feature)
                            terpene_idx = residual_features.index('terpene_pct')
                
                            potency_imp = importances[potency_idx]
                            terpene_imp = importances[terpene_idx]
                
                            # Compare relative importance
                            if potency_imp > terpene_imp * 2:
                                report.append("  - Potency is MUCH MORE important than terpene percentage")
                            elif potency_imp > terpene_imp:
                                report.append("  - Potency is more important than terpene percentage")
                            elif terpene_imp > potency_imp * 2:
                                report.append("  - Terpene percentage is MUCH MORE important than potency")
                            elif terpene_imp > potency_imp:
                                report.append("  - Terpene percentage is more important than potency")
                            else:
                                report.append("  - Terpene percentage and potency have similar importance")
                            
                    # Same validation code as for standard models - reusing the same pattern
                    if validation_data is not None:
                        try:
                            # Extract media/terpene from model key
                            components = model_key.split('_', 1)
                            media = components[0]
                
                            # Handle general models
                            if len(components) > 1 and components[1] == 'general':
                                model_val_data = validation_data[validation_data['media'] == media].copy()
                    
                                # One-hot encode terpenes
                                model_val_data = pd.get_dummies(
                                    model_val_data, 
                                    columns=['terpene'],
                                    prefix=['terpene']
                                )
                            else:
                                # Extract terpene from key
                                terpene = components[1] if len(components) > 1 else 'Raw'
                                model_val_data = validation_data[
                                    (validation_data['media'] == media) & 
                                    (validation_data['terpene'] == terpene)
                                ].copy()
                
                            if len(model_val_data) >= 5:
                                report.append(f"\nValidation results ({len(model_val_data)} samples):")
                        
                                # Calculate potency from individual cannabinoids if total_potency is missing
                                if primary_feature == 'total_potency' and primary_feature not in model_val_data.columns:
                                    if 'd9_thc' in model_val_data.columns and 'd8_thc' in model_val_data.columns:
                                        model_val_data['total_potency'] = model_val_data[['d9_thc', 'd8_thc']].sum(axis=1, skipna=True)
                                        report.append("Calculated total_potency from individual cannabinoid values")
                                    else:
                                        report.append("Cannot calculate total_potency - missing cannabinoid data")
                                        continue
                        
                                # Drop NaN values in key features
                                model_val_data = model_val_data.dropna(subset=['inverse_temp', 'log_viscosity'])
                            
                                # Ensure primary feature exists 
                                if primary_feature not in model_val_data.columns:
                                    if primary_feature == 'total_potency' and 'total_potency' in validation_data.columns:
                                        model_val_data['total_potency'] = validation_data.loc[model_val_data.index, 'total_potency']
                                
                                # Filter out rows with NaN in primary feature
                                model_val_data = model_val_data.dropna(subset=[primary_feature])
                        
                                if len(model_val_data) < 5:
                                    report.append(f"Insufficient data after removing NaN values. Only {len(model_val_data)} samples left.")
                                    continue
                            
                                # Create model_feature column if needed
                                if 'model_feature' in residual_features and 'model_feature' not in model_val_data.columns:
                                    model_val_data['model_feature'] = model_val_data[primary_feature]
                            
                                # Step 1: Evaluate baseline model
                                X_baseline = model_val_data[baseline_features]
                                y_true = model_val_data['log_viscosity']
                    
                                # Get baseline predictions
                                baseline_preds = baseline_model.predict(X_baseline)
                    
                                # Calculate residuals
                                model_val_data['baseline_pred'] = baseline_preds
                                model_val_data['residual'] = y_true - baseline_preds
                    
                                # Step 2: Evaluate residual model
                                # Check if we have all required features
                                missing_features = [f for f in residual_features if f not in model_val_data.columns]
                            
                                # Better handling of missing features
                                if missing_features:
                                    if 'model_feature' in missing_features and primary_feature in model_val_data.columns:
                                        # Create model_feature explicitly from primary feature
                                        model_val_data['model_feature'] = model_val_data[primary_feature]
                                        missing_features.remove('model_feature')
                                
                                    # If we still have missing features, report and skip
                                    if missing_features:
                                        report.append(f"Missing required features for validation: {', '.join(missing_features)}")
                                        continue
                        
                                # Extract features and residuals
                                X_residual = model_val_data[residual_features]
                                y_residual = model_val_data['residual']
                            
                                # Handle NaN values in the feature data
                                if X_residual.isna().any().any():
                                    # Use SimpleImputer to replace NaN values with mean
                                    imputer = SimpleImputer(strategy='mean')
                                    X_residual_values = imputer.fit_transform(X_residual)
                                    X_residual = pd.DataFrame(X_residual_values, 
                                                              index=X_residual.index, 
                                                              columns=X_residual.columns)
                    
                                # Get residual predictions
                                residual_preds = residual_model.predict(X_residual)
                    
                                # Calculate metrics for residual model
                                r2_residual = r2_score(y_residual, residual_preds)
                                mse_residual = mean_squared_error(y_residual, residual_preds)
                    
                                report.append(f"Residual model - MSE: {mse_residual:.2f}, R^2: {r2_residual:.4f}")
                    
                                # Combined prediction metrics
                                combined_preds = baseline_preds + residual_preds
                    
                                # Log scale metrics
                                r2_log = r2_score(y_true, combined_preds)
                                mse_log = mean_squared_error(y_true, combined_preds)
                    
                                # Original scale metrics
                                y_orig = np.exp(y_true)
                                preds_orig = np.exp(combined_preds)
                    
                                r2_orig = r2_score(y_orig, preds_orig)
                                mse_orig = mean_squared_error(y_orig, preds_orig)
                    
                                report.append(f"Log scale - MSE: {mse_log:.2f}, R^2: {r2_log:.4f}")
                                report.append(f"Original scale - MSE: {mse_orig:.2f}, R^2: {r2_orig:.4f}")
                    
                                if r2_orig < 0.5:
                                    report.append("WARNING: Poor validation performance (R^2 < 0.5)!")
                            else:
                                report.append(f"Insufficient validation data: {len(model_val_data)} samples")
                        except Exception as e:
                            report.append(f"Error validating model: {str(e)}")
                            report.append(traceback.format_exc())
                except Exception as e:
                    report.append(f"Error analyzing model: {str(e)}")
                    report.append(traceback.format_exc())

        # Analyze combined models
        if combined_models_exist:
            report.append(f"\nCombined Models (terpene + potency): {len(self.combined_viscosity_models)}")
            report.append("-" * 40)

            for model_key, model in self.combined_viscosity_models.items():
                report.append(f"\nModel: {model_key}")
    
                try:
                    # Check if model has the correct structure
                    if not is_valid_model(model):
                        report.append("Error analyzing model: incompatible model format")
                        continue
                    
                    # Extract model components
                    baseline_model = model['baseline_model']
                    residual_model = model['residual_model']
                    baseline_features = model['baseline_features']
                    residual_features = model['residual_features']
            
                    # Get metadata
                    metadata = model.get('metadata', {})
                    primary_feature = metadata.get('primary_feature', 'terpene_pct')
        
                    # Analyze baseline model (Arrhenius)
                    report.append("1. Temperature baseline model (Arrhenius)")
        
                    if hasattr(baseline_model, 'coef_'):
                        # Extract Arrhenius parameters
                        coef = baseline_model.coef_[0]
                        intercept = baseline_model.intercept_
            
                        # Calculate activation energy
                        R = 8.314  # Gas constant (J/mol·K)
                        Ea = coef * R
                        Ea_kJ = Ea / 1000  # Convert to kJ/mol
            
                        report.append(f"  - Equation: log(viscosity) = {intercept:.4f} + {coef:.4f} * (1/T)")
                        report.append(f"  - Activation energy: {Ea_kJ:.2f} kJ/mol")
        
                    # Analyze residual model (terpene effects)
                    report.append("2. Residual model (terpene effects)")
                    report.append(f"  - Model type: {type(residual_model).__name__}")
        
                    if hasattr(residual_model, 'feature_importances_'):
                        # Report feature importance
                        importances = residual_model.feature_importances_
            
                        if len(importances) == len(residual_features):
                            report.append("  - Feature importances:")
                            for i, feature in enumerate(residual_features):
                                report.append(f"    * {feature}: {importances[i]:.4f}")
            
                        # Check terpene importance
                        if primary_feature in residual_features:
                            feature_idx = residual_features.index(primary_feature)
                            feature_importance = importances[feature_idx]
                
                            if feature_importance > 0.3:
                                report.append(f"  - {primary_feature} has HIGH importance in viscosity deviation")
                            elif feature_importance > 0.1:
                                report.append(f"  - {primary_feature} has MODERATE importance in viscosity deviation")
                            elif feature_importance > 0.01:
                                report.append(f"  - {primary_feature} has LOW importance in viscosity deviation")
                            else:
                                report.append(f"  - {primary_feature} has NEGLIGIBLE importance in viscosity deviation")
                
                    # Same validation code as for standard and enhanced models
                    if validation_data is not None:
                        try:
                            # Extract media/terpene from model key
                            components = model_key.split('_', 1)
                            media = components[0]
                
                            # Handle general models
                            if len(components) > 1 and components[1] == 'general':
                                model_val_data = validation_data[validation_data['media'] == media].copy()
                    
                                # One-hot encode terpenes
                                model_val_data = pd.get_dummies(
                                    model_val_data, 
                                    columns=['terpene'],
                                    prefix=['terpene']
                                )
                            else:
                                # Extract terpene from key
                                terpene = components[1] if len(components) > 1 else 'Raw'
                                model_val_data = validation_data[
                                    (validation_data['media'] == media) & 
                                    (validation_data['terpene'] == terpene)
                                ].copy()
                
                            if len(model_val_data) >= 5:
                                report.append(f"\nValidation results ({len(model_val_data)} samples):")
                    
                                # Drop NaN values in key features
                                model_val_data = model_val_data.dropna(subset=['inverse_temp', 'log_viscosity'])
                            
                                # Ensure primary feature exists
                                if primary_feature not in model_val_data.columns:
                                    if primary_feature == 'terpene_pct' and 'terpene_pct' in validation_data.columns:
                                        model_val_data['terpene_pct'] = validation_data.loc[model_val_data.index, 'terpene_pct']
                            
                                # Filter out rows with NaN in primary feature
                                model_val_data = model_val_data.dropna(subset=[primary_feature])
                        
                                if len(model_val_data) < 5:
                                    report.append(f"Insufficient data after removing NaN values. Only {len(model_val_data)} samples left.")
                                    continue
                            
                                # Create model_feature column if needed
                                if 'model_feature' in residual_features and 'model_feature' not in model_val_data.columns:
                                    model_val_data['model_feature'] = model_val_data[primary_feature]
                    
                                # Step 1: Evaluate baseline model
                                X_baseline = model_val_data[baseline_features]
                                y_true = model_val_data['log_viscosity']
                    
                                # Get baseline predictions
                                baseline_preds = baseline_model.predict(X_baseline)
                    
                                # Calculate residuals
                                model_val_data['baseline_pred'] = baseline_preds
                                model_val_data['residual'] = y_true - baseline_preds
                    
                                # Step 2: Evaluate residual model
                                # Get valid features
                                valid_features = [f for f in residual_features if f in model_val_data.columns]
                        
                                # Check if we have all required features
                                missing_features = [f for f in residual_features if f not in model_val_data.columns]
                            
                                # Better handling of missing features
                                if missing_features:
                                    if 'model_feature' in missing_features and primary_feature in model_val_data.columns:
                                        # Create model_feature explicitly from primary feature
                                        model_val_data['model_feature'] = model_val_data[primary_feature]
                                        missing_features.remove('model_feature')
                                
                                    # If we still have missing features, report and skip
                                    if missing_features:
                                        report.append(f"Missing required features for validation: {', '.join(missing_features)}")
                                        continue
                    
                                # Extract features and residuals
                                X_residual = model_val_data[residual_features]
                                y_residual = model_val_data['residual']
                            
                                # Handle NaN values in the feature data
                                if X_residual.isna().any().any():
                                    # Use SimpleImputer to replace NaN values with mean
                                    imputer = SimpleImputer(strategy='mean')
                                    X_residual_values = imputer.fit_transform(X_residual)
                                    X_residual = pd.DataFrame(X_residual_values, 
                                                              index=X_residual.index, 
                                                              columns=X_residual.columns)
                    
                                # Get residual predictions
                                residual_preds = residual_model.predict(X_residual)
                    
                                # Calculate metrics for residual model
                                r2_residual = r2_score(y_residual, residual_preds)
                                mse_residual = mean_squared_error(y_residual, residual_preds)
                    
                                report.append(f"Residual model - MSE: {mse_residual:.2f}, R^2: {r2_residual:.4f}")
                    
                                # Combined prediction metrics
                                combined_preds = baseline_preds + residual_preds
                    
                                # Log scale metrics
                                r2_log = r2_score(y_true, combined_preds)
                                mse_log = mean_squared_error(y_true, combined_preds)
                    
                                # Original scale metrics
                                y_orig = np.exp(y_true)
                                preds_orig = np.exp(combined_preds)
                    
                                r2_orig = r2_score(y_orig, preds_orig)
                                mse_orig = mean_squared_error(y_orig, preds_orig)
                    
                                report.append(f"Log scale - MSE: {mse_log:.2f}, R^2: {r2_log:.4f}")
                                report.append(f"Original scale - MSE: {mse_orig:.2f}, R^2: {r2_orig:.4f}")
                    
                                if r2_orig < 0.5:
                                    report.append("WARNING: Poor validation performance (R^2 < 0.5)!")
                            else:
                                report.append(f"Insufficient validation data: {len(model_val_data)} samples")
                        except Exception as e:
                            report.append(f"Error validating model: {str(e)}")
                            report.append(traceback.format_exc())
                except Exception as e:
                    report.append(f"Error analyzing model: {str(e)}")
                    report.append(traceback.format_exc())
    
        # Add recommendations
        report.append("\nRecommendations:")
        report.append("---------------")
        report.append("1. Focus on residual model R² values for terpene effect analysis")
        report.append("2. Compare baseline activation energies across media types")
        report.append("3. Consider using Ridge regression for small datasets")
        report.append("4. For better models, standardize features before training")
        report.append("5. Use log-scale metrics for model evaluation")

        # Print report to console
        print("\n".join(report))

        # Show dialog if requested
        if show_dialog:
            report_window = Toplevel(self.root)
            report_window.title("Two-Step Model Analysis")
            report_window.geometry("700x500")

            Label(report_window, text="Two-Step Model Analysis", 
                  font=("Arial", 14, "bold")).pack(pady=10)

            text_frame = Frame(report_window)
            text_frame.pack(fill="both", expand=True, padx=10, pady=10)

            scrollbar = Scrollbar(text_frame)
            scrollbar.pack(side="right", fill="y")

            text_widget = Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
            text_widget.pack(side="left", fill="both", expand=True)

            scrollbar.config(command=text_widget.yview)

            text_widget.insert("1.0", "\n".join(report))
            text_widget.config(state="disabled")

            Button(report_window, text="Close", 
                   command=report_window.destroy).pack(pady=10)

        return report

    def ensure_model_compatibility(self):
        """
        Check if models have the expected structure and offer to retrain if not.
        This function helps transition to the new model format.
        """
        from tkinter import messagebox
    
        # Define what a valid model structure looks like
        def is_valid_model(model):
            return (isinstance(model, dict) and 
                    'baseline_model' in model and 
                    'residual_model' in model and
                    'baseline_features' in model and
                    'residual_features' in model)
    
        # Check standard models
        retraining_needed = False
        if hasattr(self, 'viscosity_models') and self.viscosity_models:
            # Sample check - just look at the first model
            sample_key = next(iter(self.viscosity_models))
            sample_model = self.viscosity_models[sample_key]
        
            if not is_valid_model(sample_model):
                retraining_needed = True
    
        # Similar checks for enhanced and combined models
        if hasattr(self, 'enhanced_viscosity_models') and self.enhanced_viscosity_models:
            sample_key = next(iter(self.enhanced_viscosity_models))
            sample_model = self.enhanced_viscosity_models[sample_key]
        
            if not is_valid_model(sample_model):
                retraining_needed = True
    
        if hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models:
            sample_key = next(iter(self.combined_viscosity_models))
            sample_model = self.combined_viscosity_models[sample_key]
        
            if not is_valid_model(sample_model):
                retraining_needed = True
    
        # If any models have incompatible format, offer to retrain
        if retraining_needed:
            if messagebox.askyesno("Model Format", 
                                 "Some models appear to use an older format. Would you like to retrain them with the new format?"):
                # Clear old models
                if hasattr(self, 'viscosity_models'):
                    self.viscosity_models = {}
                if hasattr(self, 'enhanced_viscosity_models'):
                    self.enhanced_viscosity_models = {}
                if hasattr(self, 'combined_viscosity_models'):
                    self.combined_viscosity_models = {}
                
                # Retrain models in the new format
                self.train_unified_models()
                return True
    
        return False

    def filter_and_analyze_specific_combinations(self):
        """
        Analyze and build models for specific media-terpene combinations with potency integration.
        Shows a configuration window and only starts analysis when the Run Analysis button is clicked.
        Only processes combinations that have terpene percentage, viscosity, temperature, and potency data.
        """
        import tkinter as tk
        from tkinter import ttk, StringVar, DoubleVar, Frame, Label, Scale, HORIZONTAL, Toplevel, Text, Scrollbar

        

        # Create the main window without starting analysis
        progress_window = Toplevel(self.root)
        progress_window.title("Arrhenius Analysis Configuration")
        progress_window.geometry("800x600")  # Initial window size
        progress_window.transient(self.root)
        progress_window.grab_set()
    
        # Create main layout frames
        top_frame = Frame(progress_window)
        top_frame.pack(fill="x", padx=10, pady=5)
    
        Label(top_frame, text="Arrhenius Analysis with Potency Integration", 
              font=("Arial", 14, "bold")).pack(pady=10)
    
        # Add a frame for configuration controls
        config_frame = Frame(top_frame)
        config_frame.pack(fill="x", padx=10, pady=5)
    
        # Add potency configuration
        Label(config_frame, text="Total Potency (%) for analysis:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    
        potency_var = DoubleVar(value=80.0)  # Default value of 80%
        potency_slider = Scale(config_frame, variable=potency_var, from_=60.0, to=95.0, 
                              orient=HORIZONTAL, length=200, resolution=0.5)
        potency_slider.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    
        # Add a dropdown to select the potency analysis mode
        potency_mode_var = StringVar(value="fixed")
        potency_modes = ["fixed", "variable"]
    
        Label(config_frame, text="Potency Analysis Mode:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        potency_mode_dropdown = ttk.Combobox(
            config_frame, 
            textvariable=potency_mode_var,
            values=potency_modes,
            state="readonly",
            width=10
        )
        potency_mode_dropdown.grid(row=0, column=3, padx=5, pady=5, sticky="w")
    
        # Text area for showing progress
        text_frame = Frame(progress_window)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
        scrollbar = Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
    
        text_widget = Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_widget.yview)
    
        # Function to add text to the widget
        def add_text(message):
            text_widget.insert("end", message + "\n")
            text_widget.see("end")
            progress_window.update_idletasks()  # Force UI update
    
        # Add initial message
        add_text("Configure the potency settings above and click 'Run Analysis' to start.")
        add_text("Analysis will only include combinations with complete data for:")
        add_text("- Terpene percentage")
        add_text("- Viscosity")
        add_text("- Temperature")
        add_text("- Potency (for chemistry-aware models)")
    
        # Create control buttons at the bottom
        button_frame = Frame(progress_window)
        button_frame.pack(pady=10)
    
        # Define function for the background thread to run the analysis
        def run_analysis_thread(potency_value, potency_mode):
            try:
                # Import required modules
                import glob
                import threading
                import pandas as pd
                import numpy as np
                import matplotlib
                # Use Agg backend for matplotlib when running in thread
                matplotlib.use('Agg')
                import matplotlib.pyplot as plt
                from scipy import stats
                from scipy.optimize import curve_fit
                import os
                import pickle
                import math
                from sklearn.metrics import r2_score
     
                # Function to sanitize filenames
                def sanitize_filename(name):
                    """Replace invalid filename characters with underscores."""
                    invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ']
                    sanitized = name
                    for char in invalid_chars:
                        sanitized = sanitized.replace(char, '_')
                    return sanitized
            
                # First try to load all model types if they're not already loaded
                if not hasattr(self, 'viscosity_models') or not self.viscosity_models:
                    self.load_standard_models()
                if not hasattr(self, 'enhanced_viscosity_models') or not self.enhanced_viscosity_models:
                    self.load_enhanced_models()
                if not hasattr(self, 'combined_viscosity_models') or not self.combined_viscosity_models:
                    # Check if the combined models loading function exists
                    if hasattr(self, 'load_combined_models'):
                        self.load_combined_models()
                    else:
                        # Try to load combined models directly
                        combined_model_path = 'models/viscosity_models_combined.pkl'
                        if os.path.exists(combined_model_path):
                            try:
                                with open(combined_model_path, 'rb') as f:
                                    self.combined_viscosity_models = pickle.load(f)
                            except Exception as e:
                                add_text(f"Error loading combined models: {str(e)}")
                                self.combined_viscosity_models = {}
                        else:
                            self.combined_viscosity_models = {}
    
                # Check which models are available to use
                using_combined_models = hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models
                using_enhanced_models = hasattr(self, 'enhanced_viscosity_models') and self.enhanced_viscosity_models
                using_standard_models = hasattr(self, 'viscosity_models') and self.viscosity_models
    
                # Determine which models to use (in priority order: combined, enhanced, standard)
                if using_combined_models:
                    models_to_analyze = self.combined_viscosity_models
                    add_text("Using combined models with both terpene and potency data.")
                elif using_enhanced_models:
                    models_to_analyze = self.enhanced_viscosity_models
                    add_text("Using enhanced models with potency data.")
                elif using_standard_models:
                    models_to_analyze = self.viscosity_models
                    add_text("Using standard models with terpene data.")
                else:
                    # Fall back to checking file paths directly if attributes don't exist
                    if os.path.exists('models/viscosity_models_combined.pkl'):
                        try:
                            with open('models/viscosity_models_combined.pkl', 'rb') as f:
                                combined_models = pickle.load(f)
                                if combined_models:
                                    self.combined_viscosity_models = combined_models
                                    models_to_analyze = combined_models
                                    using_combined_models = True
                                    add_text("Using combined models loaded from file.")
                                    using_enhanced_models = False
                        except Exception as e:
                            add_text(f"Error loading combined models file: {str(e)}")
            
                    if not using_combined_models and os.path.exists('models/viscosity_models_potency.pkl'):
                        try:
                            with open('models/viscosity_models_potency.pkl', 'rb') as f:
                                enhanced_models = pickle.load(f)
                                if enhanced_models:
                                    self.enhanced_viscosity_models = enhanced_models
                                    models_to_analyze = enhanced_models
                                    using_enhanced_models = True
                                    add_text("Using potency models loaded from file.")
                        except Exception as e:
                            add_text(f"Error loading potency models file: {str(e)}")
            
                    if not using_combined_models and not using_enhanced_models and os.path.exists('models/viscosity_models.pkl'):
                        try:
                            with open('models/viscosity_models.pkl', 'rb') as f:
                                standard_models = pickle.load(f)
                                if standard_models:
                                    self.viscosity_models = standard_models
                                    models_to_analyze = standard_models
                                    using_standard_models = True
                                    add_text("Using standard models loaded from file.")
                        except Exception as e:
                            add_text(f"Error loading standard models file: {str(e)}")
        
                # Check if any models were found
                if not using_combined_models and not using_enhanced_models and not using_standard_models:
                    add_text("No trained models found. Please train models first.")
                    run_button.config(state="normal")  # Re-enable run button
                    return
            
                # Create plots directory if it doesn't exist
                os.makedirs('plots', exist_ok=True)
            
                # Load training data to identify combinations with complete data
                training_data = None
                try:
                    # Get all data files
                    data_files = glob.glob('data/Master_Viscosity_Data_*.csv')
                    if data_files:
                        # Sort by modification time (newest first)
                        data_files.sort(key=os.path.getmtime, reverse=True)
                        latest_file = data_files[0]
                        training_data = pd.read_csv(latest_file)
                        add_text(f"Loaded training data from {latest_file}")
                
                        # Check if potency data is available
                        has_potency = 'total_potency' in training_data.columns
                        if has_potency:
                            add_text(f"Training data includes potency information!")
                            add_text(f"Potency range: {training_data['total_potency'].min():.1f}% - {training_data['total_potency'].max():.1f}%")
                        else:
                            add_text("Warning: Training data does not include potency information.")
                    else:
                        add_text("No training data files found.")
                except Exception as e:
                    add_text(f"Error loading training data: {str(e)}")
            
                # Function to check if a media-terpene combination has complete data
                def has_complete_data(media, terpene):
                    """Check if this combination has all required data fields."""
                    if training_data is None:
                        # No data to check against, assume it has complete data
                        return True
                
                    # Filter training data for this combination
                    combo_data = training_data[
                        (training_data['media'] == media) & 
                        (training_data['terpene'] == terpene)
                    ]
                
                    if combo_data.empty:
                        return False
                
                    # Check for required columns
                    required_columns = ['terpene_pct', 'temperature', 'viscosity']
                
                    # If using enhanced models, also require potency data
                    if using_enhanced_models or using_combined_models:
                        required_columns.append('total_potency')
                
                    # Check if any rows have all required fields
                    complete_rows = combo_data.dropna(subset=required_columns)
                
                    # Return True if there's at least 5 samples with complete data
                    return len(complete_rows) >= 5
            
                # Arrhenius function: ln(viscosity) = ln(A) + (Ea/R)*(1/T)
                def arrhenius_function(x, a, b):
                    """x is 1/T (inverse temperature in Kelvin)"""
                    return a + b * x
            
                # Temperature range for analysis
                temperature_range = np.linspace(20, 70, 11)  # 20C to 70C
            
                # Define potency range based on mode
                if potency_mode == "variable":
                    # Calculate a range around the selected potency value
                    # Start with -10% to +10% from the selected value, in 5% increments
                    center_potency = potency_value
                    offsets = [-10, -5, 0, 5, 10]
                    potency_range = [max(20, center_potency + offset) for offset in offsets]
                    # Ensure no values exceed 100%
                    potency_range = [min(100, p) for p in potency_range]
                    # Remove duplicates and sort
                    potency_range = sorted(list(set(potency_range)))
    
                    add_text(f"Performing variable potency analysis with levels: {', '.join([f'{p:.1f}%' for p in potency_range])}")
                else:
                    # Fixed potency mode
                    potency_range = [potency_value]
                    add_text(f"Using fixed potency value of {potency_value:.1f}%")
            
                # Count successful models
                successful_models = 0
                successful_combinations = []
            
                # Process each model
                for model_key, model in models_to_analyze.items():
                    try:
                        # Extract media and terpene from model key
                        if using_enhanced_models and "_with_chemistry" in model_key:
                            display_key = model_key.replace("_with_chemistry", "")
                            media, terpene = display_key.split('_', 1)
                            components = display_key.split('_')
                        else:
                            media, terpene = model_key.split('_', 1)
                            components = model_key.split('_')
                    
                        # Handle general models better
                        if len(components) > 1:
                            media = components[0]
                            terpene = components[1]
                            is_general = (terpene == 'general')
                        else:
                            media = components[0]
                            terpene = 'Raw'
                            is_general = False

                        add_text(f"\nAnalyzing {media}/{terpene}...")

                        # For general models, select a representative terpene for visualization
                        if is_general:
                            # Find a representative terpene from the data
                            if training_data is not None:
                                media_data = training_data[training_data['media'] == media]
                                if not media_data.empty:
                                    # Count terpenes and pick most common that's not 'Raw'
                                    terpene_counts = media_data['terpene'].value_counts()
                                    valid_terpenes = [t for t in terpene_counts.index if t != 'Raw' and terpene_counts[t] >= 5]
                                    if valid_terpenes:
                                        visualization_terpene = valid_terpenes[0]
                                        add_text(f"Using {visualization_terpene} as representative terpene for {media} general model")
                                        terpene = visualization_terpene  # Use this for visualization
                    
                        # Check if this combination has complete data
                        if not has_complete_data(media, terpene):
                            add_text(f"Skipping {media}/{terpene} - incomplete data")
                            continue
                    
                        # Find typical Terpene percentage for this combo from data
                        terpene_pct = 1.0  # Default
                        if training_data is not None:
                            combo_data = training_data[
                                (training_data['media'] == media) & 
                                (training_data['terpene'] == terpene)
                            ]
                            if 'terpene_pct' in combo_data.columns and not combo_data.empty:
                                terpene_pct = 100 * combo_data['terpene_pct'].median()
                                if pd.isna(terpene_pct) or terpene_pct <= 0:
                                    terpene_pct = 1.0
                    
                        add_text(f"Using terpene percentage of {terpene_pct:.2f}% for analysis")
                        
                        # Variable potency analysis if requested
                        if potency_mode == "variable" and using_enhanced_models:
                            # Create figure for variable potency analysis
                            plt.figure(figsize=(12, 10))
                        
                            # Create subplots
                            ax1 = plt.subplot(211)
                            ax2 = plt.subplot(212)
                        
                            # Store activation energies
                            activation_energies = []
                        
                            # Plot for each potency level
                            for potency in potency_range:
                                # Calculate temperatures in Kelvin
                                temperatures_kelvin = temperature_range + 273.15
                                inverse_temp = 1 / temperatures_kelvin
    
                                # Convert percentage to decimal (e.g., 80% -> 0.8)
                                # This is the correct scale for the model (no further scaling needed)
                                decimal_potency = potency / 100.0
    
                                add_text(f"  • Using potency {potency:.1f}% (decimal: {decimal_potency:.4f})")
    
                                # Debug the model inputs and outputs
                                add_text(f"  • Debug: terpene_pct={terpene_pct}, temp=25.0, potency={decimal_potency}")
                                try:
                                    viscosity_at_25C = self.predict_model_viscosity(model, 100*(1-decimal_potency), 25.0, decimal_potency)
                                    add_text(f"  • At 25°C: predicted viscosity = {viscosity_at_25C:.0f} cP")
        
                                    # Get predictions for each temperature
                                    predicted_visc = []
                                    for temp in temperature_range:
                                        visc = self.predict_model_viscosity(model, 100*(1-decimal_potency), temp, decimal_potency)
                                        predicted_visc.append(visc)
                            
                                    predicted_visc = np.array(predicted_visc)
                            
                                    # Filter invalid values
                                    valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                                    if not any(valid_indices):
                                        add_text(f"No valid predictions for {potency:.1f}% potency. Skipping.")
                                        continue
                            
                                    # Extract valid data
                                    inv_temp_valid = inverse_temp[valid_indices]
                                    predicted_visc_valid = predicted_visc[valid_indices]
                            
                                    # Calculate ln(viscosity)
                                    ln_visc = np.log(predicted_visc_valid)
                            
                                    # Fit Arrhenius equation
                                    params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                                    a, b = params
                            
                                    # Calculate activation energy
                                    R = 8.314  # Gas constant
                                    Ea = b * R
                                    Ea_kJ = Ea / 1000  # Convert to kJ/mol
                            
                                    # Store for comparison
                                    activation_energies.append((potency, Ea_kJ))
                            
                                    # Calculate predicted values
                                    ln_visc_pred = arrhenius_function(inv_temp_valid, a, b)
                            
                                    # Calculate R-squared
                                    r2 = r2_score(ln_visc, ln_visc_pred)
                            
                                    # Plot with distinct colors
                                    ax1.semilogy(temperature_range[valid_indices], predicted_visc_valid,
                                            'o-', label=f'Potency {potency:.1f}% (25C = {viscosity_at_25C:.0f} cP)')
        
                            
                                    ax2.scatter(inv_temp_valid, ln_visc,
                                             label=f'Potency {potency:.1f}%')
                                    ax2.plot(inv_temp_valid, ln_visc_pred, '--',
                                            label=f'Fit {potency:.1f}% (Ea={Ea_kJ:.1f} kJ/mol)')

                                    add_text(f"  • At {potency:.1f}% potency: viscosity at 25C = {viscosity_at_25C:.0f} cP")

                                except Exception as e:
                                    add_text(f" -error in prediction: {str(e)}")
                                    add_text(" - skipping this potency level due to prediction error")


                        
                            # Configure plots
                            ax1.set_xlabel('Temperature (C)')
                            ax1.set_ylabel('Viscosity (cP)')
                            ax1.set_title(f'Viscosity vs Temperature for {media}/{terpene}\nTerpene: {terpene_pct:.2f}%')
                            ax1.grid(True)
                            ax1.legend()
                        
                            ax2.set_xlabel('1/T (K⁻¹)')
                            ax2.set_ylabel('ln(Viscosity)')
                            ax2.set_title('Arrhenius Plots for Different Potency Levels')
                            ax2.grid(True)
                            ax2.legend()
                        
                            plt.tight_layout()
                        
                            # Save plot
                            plot_path = f'plots/Arrhenius_{sanitize_filename(media)}_{sanitize_filename(terpene)}_variable_potency.png'
                            plt.savefig(plot_path)
                            plt.close()
                        
                            # Create potency vs activation energy plot
                            if len(activation_energies) > 1:
                                plt.figure(figsize=(8, 6))
                                potency_values, ea_values = zip(*activation_energies)
                            
                                plt.plot(potency_values, ea_values, 'o-', linewidth=2)
                                plt.xlabel('Total Potency (%)')
                                plt.ylabel('Activation Energy (kJ/mol)')
                                plt.title(f'Effect of Potency on Activation Energy\n{media}/{terpene}')
                                plt.grid(True)
                            
                                # Add trendline
                                if len(potency_values) > 2:
                                    z = np.polyfit(potency_values, ea_values, 1)
                                    p = np.poly1d(z)
                                    plt.plot(potency_values, p(potency_values), "r--",
                                           label=f"Trend: {z[0]:.2f}x + {z[1]:.2f}")
                                    plt.legend()
                            
                                # Save plot
                                potency_plot_path = f'plots/Potency_Effect_{sanitize_filename(media)}_{sanitize_filename(terpene)}.png'
                                plt.savefig(potency_plot_path)
                                plt.close()
                            
                                add_text(f"Variable potency analysis complete")
                                add_text(f"  • Plot saved to: {plot_path}")
                                add_text(f"  • Potency effect plot saved to: {potency_plot_path}")
                            
                                # Report trend
                                if len(potency_values) > 2:
                                    if z[0] > 0:
                                        add_text(f"  • Trend: Activation energy increases by {z[0]:.2f} kJ/mol per 1% increase in potency")
                                    else:
                                        add_text(f"  • Trend: Activation energy decreases by {abs(z[0]):.2f} kJ/mol per 1% increase in potency")
                                
                                    slope_significance = abs(z[0]) / (max(ea_values) - min(ea_values)) * 100
                                    if slope_significance < 5:
                                        add_text("  • Potency has minimal effect on temperature sensitivity")
                                    elif slope_significance < 15:
                                        add_text("  • Potency has moderate effect on temperature sensitivity")
                                    else:
                                        add_text("  • Potency has significant effect on temperature sensitivity")
                        
                            # Skip standard analysis
                            successful_models += 1
                            successful_combinations.append((media, terpene))
                            continue
                    
                        # --- Standard single-potency analysis ---
                    
                        # Calculate temperatures
                        temperatures_kelvin = temperature_range + 273.15
                        inverse_temp = 1 / temperatures_kelvin
                    
                        # Get predictions
                        predicted_visc = []
                        for temp in temperature_range:
                            if using_enhanced_models or using_combined_models:
                                visc = self.predict_model_viscosity(model, terpene_pct, temp, potency_value)
                            else:
                                visc = self.predict_model_viscosity(model, terpene_pct, temp)
                            predicted_visc.append(visc)
                    
                        predicted_visc = np.array(predicted_visc)
                    
                        # Filter invalid values
                        valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                        if not any(valid_indices):
                            add_text(f"No valid predictions for {media}/{terpene}. Skipping.")
                            continue
                    
                        # Extract valid data
                        inv_temp_valid = inverse_temp[valid_indices]
                        predicted_visc_valid = predicted_visc[valid_indices]
                    
                        # Calculate ln(viscosity)
                        ln_visc = np.log(predicted_visc_valid)
                    
                        # Fit Arrhenius equation
                        params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                        a, b = params
                    
                        # Calculate activation energy
                        R = 8.314  # Gas constant
                        Ea = b * R
                        Ea_kJ = Ea / 1000  # Convert to kJ/mol
                    
                        # Calculate pre-exponential factor
                        A = np.exp(a)
                    
                        # Calculate predicted values
                        ln_visc_pred = arrhenius_function(inv_temp_valid, a, b)
                    
                        # Calculate R-squared
                        r2 = r2_score(ln_visc, ln_visc_pred)
                    
                        # Generate plot
                        plt.figure(figsize=(10, 8))
                    
                        # Create subplots
                        plt.subplot(211)
                        plt.scatter(temperature_range[valid_indices], predicted_visc_valid, color='blue', label='Predicted viscosity')
                        plt.yscale('log')
                        plt.xlabel('Temperature (C)')
                        plt.ylabel('Viscosity (cP)')
                        if using_enhanced_models:
                            plt.title(f'Viscosity vs Temperature for {media}/{terpene}\nTerpene: {terpene_pct:.2f}%, Potency: {potency_value:.1f}%')
                        else:
                            plt.title(f'Viscosity vs Temperature for {media}/{terpene}\nTerpene: {terpene_pct:.2f}%')
                        plt.grid(True)
                    
                        # Arrhenius plot
                        plt.subplot(212)
                        plt.scatter(inv_temp_valid, ln_visc, color='blue', label='ln(Viscosity)')
                        plt.plot(inv_temp_valid, ln_visc_pred, 'r-', label=f'Arrhenius fit (R^2 = {r2:.4f})')
                        plt.xlabel('1/T (K⁻¹)')
                        plt.ylabel('ln(Viscosity)')
                        plt.title(f'Arrhenius Plot: Ea = {Ea_kJ:.2f} kJ/mol, ln(A) = {a:.2f}')
                        plt.grid(True)
                        plt.legend()
                    
                        plt.tight_layout()
                    
                        # Save plot
                        if using_enhanced_models:
                            plot_path = f'plots/Arrhenius_{sanitize_filename(media)}_{sanitize_filename(terpene)}_potency{int(potency_value)}.png'
                        else:
                            plot_path = f'plots/Arrhenius_{sanitize_filename(media)}_{sanitize_filename(terpene)}.png'
                        plt.savefig(plot_path)
                        plt.close()
                    
                        # Report results
                        add_text(f"Results for {media}/{terpene}:")
                        add_text(f"  • Activation energy (Ea): {Ea_kJ:.2f} kJ/mol")
                        add_text(f"  • Pre-exponential factor ln(A): {a:.2f}")
                        add_text(f"  • Arrhenius equation: ln(viscosity) = {a:.2f} + {b:.2f}*(1/T)")
                        add_text(f"  • R-squared: {r2:.4f}")
                        add_text(f"  • Plot saved to: {plot_path}")
                    
                        # Categorize activation energy
                        if Ea_kJ < 20:
                            add_text("  • Low activation energy: less temperature-sensitive")
                        elif Ea_kJ < 40:
                            add_text("  • Medium activation energy: moderately temperature-sensitive")
                        else:
                            add_text("  • High activation energy: highly temperature-sensitive")
                    
                        successful_models += 1
                        successful_combinations.append((media, terpene))
                    
                    except Exception as e:
                        add_text(f"Error analyzing {model_key}: {str(e)}")
                        import traceback
                        traceback_str = traceback.format_exc()
                        print(f"Detailed error: {traceback_str}")
            
                # Summary
                add_text(f"\nAnalysis complete! Successfully analyzed {successful_models} models.")
                add_text(f"Plot files are saved in the 'plots' directory.")
            
                # Generate comparison plots if we have successful models
                if successful_models > 0:
                    try:
                        # Generate activation energy comparison plot
                        self.generate_fixed_activation_energy_comparison(
                            models_to_analyze, potency_value, using_enhanced_models, 
                            successful_combinations, add_text
                        )
                        add_text("Generated comparison plot of activation energies.")
                    except Exception as e:
                        add_text(f"Error generating comparison plot: {str(e)}")
                        import traceback
                        traceback_str = traceback.format_exc()
                        print(f"Comparison plot error: {traceback_str}")
            
                # Re-enable the run button
                self.root.after(0, lambda: run_button.config(state="normal"))
            
            except Exception as e:
                add_text(f"Error in analysis thread: {str(e)}")
                import traceback
                traceback_str = traceback.format_exc()
                print(f"Thread error: {traceback_str}")
                # Re-enable button
                self.root.after(0, lambda: run_button.config(state="normal"))
    
        # Add the Run Analysis button
        def start_analysis():
            # Disable the run button while analysis is running
            import threading
            run_button.config(state="disabled")
            add_text("\nStarting analysis...")
        
            # Start the analysis in a background thread
            analysis_thread = threading.Thread(
                target=lambda: run_analysis_thread(
                    potency_var.get(), 
                    potency_mode_var.get()
                )
            )
            analysis_thread.daemon = True
            analysis_thread.start()
    
        run_button = ttk.Button(
            button_frame, 
            text="Run Analysis",
            command=start_analysis
        )
        run_button.pack(side="left", padx=10)
    
        # Add a close button
        ttk.Button(
            button_frame,
            text="Close",
            command=progress_window.destroy
        ).pack(side="left", padx=10)

    def generate_fixed_activation_energy_comparison(self,models, potency_value, using_enhanced_models, valid_combinations, log_func=None):
        """
        Generate a fixed version of the activation energy comparison plot.
        Only includes combinations that have been successfully analyzed.
    
        Args:
            models: Dictionary of trained models
            potency_value: Potency value used for predictions
            using_enhanced_models: Flag indicating if enhanced models are being used
            valid_combinations: List of (media, terpene) tuples that were successfully analyzed
            log_func: Optional function to log messages
        """
        import numpy as np
        import matplotlib
        matplotlib.use('Agg')  # Use non-interactive backend
        import matplotlib.pyplot as plt
        from scipy.optimize import curve_fit
        import os
    
        def sanitize_filename(name):
            """Replace invalid filename characters with underscores."""
            invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ']
            sanitized = name
            for char in invalid_chars:
                sanitized = sanitized.replace(char, '_')
            return sanitized
    
        # Filter models to only include validated combinations
        filtered_models = {}
        for model_key, model in models.items():
            # Extract media and terpene
            if using_enhanced_models and "_with_chemistry" in model_key:
                base_key = model_key.replace("_with_chemistry", "")
                media, terpene = base_key.split('_', 1)
            else:
                media, terpene = model_key.split('_', 1)
        
            # Check if this is a valid combination
            if (media, terpene) in valid_combinations:
                filtered_models[model_key] = model
    
        if not filtered_models:
            if log_func:
                log_func("No valid models for comparison plot")
            return
    
        # Store results for comparison
        media_types = set()
        results = []
    
        # Arrhenius function
        def arrhenius_function(x, a, b):
            return a + b * x
    
        # Temperature range
        temperature_range = np.linspace(20, 70, 11)
        temperatures_kelvin = temperature_range + 273.15
        inverse_temp = 1 / temperatures_kelvin
    
        # R constant
        R = 8.314  # Gas constant
    
        # Process each model
        for model_key, model in filtered_models.items():
            try:
                # Extract media and terpene
                if using_enhanced_models and "_with_chemistry" in model_key:
                    display_key = model_key.replace("_with_chemistry", "")
                    media, terpene = display_key.split('_', 1)
                else:
                    media, terpene = model_key.split('_', 1)
            
                media_types.add(media)
            
                # Default terpene percentage
                terpene_pct = 5.0
            
                # Generate predictions
                predicted_visc = []
                for temp in temperature_range:
                    visc = self.predict_model_viscosity(model, terpene_pct, temp, potency_value)
                    predicted_visc.append(visc)
            
                predicted_visc = np.array(predicted_visc)
            
                # Filter invalid values
                valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                if not any(valid_indices):
                    print(f"Warning: No valid predictions for {model_key}")
                    continue
            
                inv_temp_valid = inverse_temp[valid_indices]
                predicted_visc_valid = predicted_visc[valid_indices]
            
                # Calculate ln(viscosity)
                ln_visc = np.log(predicted_visc_valid)
            
                # Fit Arrhenius equation
                params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                a, b = params
            
                # Calculate activation energy
                Ea = b * R
                Ea_kJ = Ea / 1000
            
                # Store result
                results.append({
                    'media': media,
                    'terpene': terpene,
                    'Ea_kJ': Ea_kJ,
                    'ln_A': a,
                    'potency': potency_value if using_enhanced_models else None
                })
        
            except Exception as e:
                print(f"Error processing {model_key}: {e}")
    
        if not results:
            if log_func:
                log_func("No valid results for comparison plot")
            return
    
        # Convert to DataFrame
        import pandas as pd
        results_df = pd.DataFrame(results)
    
        # Sort by activation energy
        results_df = results_df.sort_values('Ea_kJ', ascending=False)
    
        # Limit to top and bottom results if many entries
        if len(results_df) > 50:
            top_entries = results_df.head(25)
            bottom_entries = results_df.tail(25)
            results_df = pd.concat([top_entries, bottom_entries])
    
        # Create figure
        plt.figure(figsize=(15, max(8, len(results_df) * 0.25)))
    
        # Create positions for bars
        positions = np.arange(len(results_df))
        bar_height = 0.6
    
        # Create colormap for media types
        media_list = sorted(list(media_types))
        colors = plt.cm.tab20(np.linspace(0, 1, len(media_list)))
        color_map = dict(zip(media_list, colors))
    
        # Create horizontal bars
        bars = plt.barh(
            positions, 
            results_df['Ea_kJ'], 
            height=bar_height,
            color=[color_map[media] for media in results_df['media']]
        )
    
        # Add labels with smaller font
        plt.yticks(positions, [f"{t[:15]}... ({m})" for t, m in zip(results_df['terpene'], results_df['media'])], fontsize=6)
    
        plt.xlabel('Activation Energy (kJ/mol)', fontsize=12)
        if using_enhanced_models and potency_value is not None:
            plt.title(f'Activation Energy Comparison by Media-Terpene Combination\nPotency: {potency_value:.1f}%', fontsize=14)
        else:
            plt.title('Activation Energy Comparison by Media-Terpene Combination', fontsize=14)
    
        # Add legend for media types
        from matplotlib.patches import Patch
        legend_elements = [Patch(facecolor=color_map[media], label=media) for media in media_list]
        plt.legend(handles=legend_elements, loc='upper right', fontsize=8)
    
        # Add value labels
        for i, bar in enumerate(bars):
            plt.text(
                bar.get_width() + 0.5, 
                bar.get_y() + bar.get_height()/2, 
                f"{results_df['Ea_kJ'].iloc[i]:.1f}", 
                va='center',
                fontsize=6
            )
    
        plt.tight_layout()
    
        # Save plot
        try:
            if using_enhanced_models and potency_value is not None:
                plot_path = f'plots/Activation_Energy_Comparison_Potency{int(potency_value)}.png'
            else:
                plot_path = 'plots/Activation_Energy_Comparison.png'
            plt.savefig(plot_path, dpi=300)
            if log_func:
                log_func(f"Comparison plot saved to: {plot_path}")
        except Exception as e:
            if log_func:
                log_func(f"Error saving comparison plot: {e}")
    
        plt.close()

    def embed_in_frame(self, parent_frame):
        """
        Embed the calculator interface in a parent frame instead of creating a new window.
    
        Args:
            parent_frame: The frame to embed the calculator in
        """
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(parent_frame)
        self.notebook.pack(fill='both', expand=True)
    
        # Create the tabs
        self.create_calculator_tab(self.notebook)
        self.create_advanced_tab(self.notebook)
        self.create_measure_tab(self.notebook)
    
        # Bind tab change event
        self.notebook.bind("<<NotebookTabChanged>>", 
                          lambda e: self.update_advanced_tab_fields())
    
        # Since we're embedding, we won't create buttons for functions that are in the menu
        # This allows the menu-driven approach similar to the screenshot
    
        return self.notebook

    def analyze_chemical_importance(self):
        """
        Analyze and visualize the importance of chemical properties in viscosity models.
        """
        # Import required libraries
        import matplotlib.pyplot as plt
        import numpy as np
        import os
        from tkinter import Toplevel, Label, Frame
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.figure import Figure
    
        # Check if enhanced models exist
        enhanced_model_path = 'models/viscosity_models_with_chemistry.pkl'
        if not os.path.exists(enhanced_model_path):
            messagebox.showinfo(
                "No Enhanced Models",
                "No enhanced models with chemical properties found.\n\n"
                "Please train models with chemical properties first."
            )
            return
    
        # Load enhanced models
        try:
            import pickle
            with open(enhanced_model_path, 'rb') as f:
                enhanced_models = pickle.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load enhanced models: {str(e)}")
            return
    
        # Create window for the analysis
        analysis_window = Toplevel(self.root)
        analysis_window.title("Chemical Properties Importance Analysis")
        analysis_window.geometry("800x800")  # Increased height
        analysis_window.transient(self.root)
    
        # Add title
        Label(
            analysis_window, 
            text="Impact of Chemical Properties on Viscosity",
            font=("Arial", 16, "bold")
        ).pack(pady=10)
    
        # Create a frame for the plots
        frame = Frame(analysis_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
    
        # Create figure with subplots
        fig = Figure(figsize=(10, 12))  # Taller figure
    
        # Get all unique features and media types
        all_features = set()
        media_types = set()
    
        for model_key, model_data in enhanced_models.items():
            # Extract media type from the model key
            media = model_key.split('_')[0]  
            media_types.add(media)
        
            # Add features
            if isinstance(model_data, dict) and 'features' in model_data:
                all_features.update(model_data['features'])
    
        # Remove temperature as it's always included
        if 'temperature' in all_features:
            all_features.remove('temperature')
    
        # Create bar plot for feature importance
        ax1 = fig.add_subplot(211)
    
        # Calculate average importance for each feature by media type
        media_list = sorted(list(media_types))
        feature_list = sorted(list(all_features))
    
        # Create arrays to store importance values
        importance_data = {media: {feature: [] for feature in feature_list} for media in media_list}
    
        # Collect importance values
        for model_key, model_data in enhanced_models.items():
            # Extract media type
            media = model_key.split('_')[0]
        
            # Get the model and its features with proper type checking
            if isinstance(model_data, dict) and 'model' in model_data:
                model = model_data['model']
                features = model_data['features']
            else:
                model = model_data
                features = ['terpene_pct', 'temperature', 'total_potency']
        
            # Skip if model doesn't have feature_importances_
            if not hasattr(model, 'feature_importances_'):
                continue
            
            importances = model.feature_importances_
        
            # Map importances to features
            for i, feature in enumerate(features):
                if feature in feature_list:  # Skip temperature
                    importance_data[media][feature].append(importances[i])
    
        # Calculate averages
        avg_importances = {media: {feature: np.mean(values) if values else 0 
                                  for feature, values in media_data.items()}
                          for media, media_data in importance_data.items()}
    
        # Plot bar chart
        bar_width = 0.8 / len(media_list)
        x = np.arange(len(feature_list))
    
        for i, media in enumerate(media_list):
            values = [avg_importances[media][feature] for feature in feature_list]
            ax1.bar(x + i * bar_width, values, bar_width, label=media)
    
        ax1.set_xlabel('Chemical Property')
        ax1.set_ylabel('Average Importance')
        ax1.set_title('Importance of Chemical Properties by Media Type')
        ax1.set_xticks(x + bar_width * (len(media_list) - 1) / 2)
        ax1.set_xticklabels(feature_list)
        
        ax1.legend()
    
        # Create heatmap showing feature importance across models
        ax2 = fig.add_subplot(212)
    
        # Get all model keys and organize by media type
        model_keys_by_media = {media: [] for media in media_list}
        for model_key in enhanced_models.keys():
            media = model_key.split('_')[0]
            model_keys_by_media[media].append(model_key)
    
        # Create heatmap data
        heatmap_data = []
        ylabels = []
    
        # Limit the number of models per media type to prevent overcrowding
        MAX_MODELS_PER_MEDIA = 5
    
        for media in media_list:
            # Sort and limit model keys
            media_models = sorted(model_keys_by_media[media])[:MAX_MODELS_PER_MEDIA]
        
            for model_key in media_models:
                model_data = enhanced_models[model_key]
            
                # Skip if not a dict with 'features' and 'model'
                if not isinstance(model_data, dict) or 'features' not in model_data or 'model' not in model_data:
                    continue
                
                features = model_data['features']
                model = model_data['model']
            
                # Skip if model doesn't have feature_importances_
                if not hasattr(model, 'feature_importances_'):
                    continue
                
                importances = model.feature_importances_
            
                # Map features to importances
                row = []
                for feature in feature_list:
                    if feature in features:
                        idx = features.index(feature)
                        row.append(importances[idx])
                    else:
                        row.append(0)
            
                heatmap_data.append(row)
            
                # Create shorter label with media and terpene
                parts = model_key.split('_')
                if len(parts) >= 2:
                    # Truncate terpene name if too long
                    terpene = parts[1][:10] + '..' if len(parts[1]) > 12 else parts[1]
                    label = f"{parts[0]}: {terpene}"
                else:
                    label = model_key
                ylabels.append(label)
    
        # Create heatmap
        if heatmap_data:
            # Set appropriate figure size for the number of models
            fig_height = min(0.3 * len(heatmap_data) + 6, 12)
            fig.set_figheight(fig_height)
        
            im = ax2.imshow(heatmap_data, cmap='viridis', aspect='auto')
        
            # Add colorbar
            fig.colorbar(im, ax=ax2)
        
            # Set labels
            ax2.set_xticks(np.arange(len(feature_list)))
            ax2.set_yticks(np.arange(len(ylabels)))
            ax2.set_xticklabels(feature_list)
            ax2.set_yticklabels(ylabels)
            ax2.tick_params(axis='both', which='major', labelsize=5)
        
            # Ensure ylabels are readable
            plt.setp(ax2.get_yticklabels(), fontsize=8)
        
            # Rotate x labels for better readability
            plt.setp(ax2.get_xticklabels(), rotation=45, ha="right", rotation_mode="anchor")
        
            ax2.set_title("Feature Importance Heatmap by Model")
    
        # Adjust layout
        fig.tight_layout()
    
        # Create a canvas to display the figure
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
        # Add a button to close the window
        ttk.Button(
            analysis_window,
            text="Close",
            command=analysis_window.destroy
        ).pack(pady=10)

    def train_unified_models(self, data=None):
        """
        Trains viscosity models using a two-step approach with improved data classification:
        1. First models the temperature-viscosity relationship (Arrhenius)
        2. Then models the residuals using available features (both potency and terpene when possible)
        """
        import os
        import pandas as pd
        import numpy as np
        import pickle
        import threading
        from sklearn.ensemble import RandomForestRegressor
        from sklearn.linear_model import LinearRegression, Ridge
        from tkinter import Toplevel, StringVar, Frame, Label, messagebox, OptionMenu

        # Create configuration window
        config_window = Toplevel(self.root)
        config_window.title("Train Two-Step Models")
        config_window.geometry("500x300")
        config_window.transient(self.root)
        config_window.grab_set()
        config_window.configure(bg=APP_BACKGROUND_COLOR)

        # Center the window
        self.gui.center_window(config_window)

        # Configuration variables
        model_type_var = StringVar(value="RandomForest")
    
        # Create a frame for options
        options_frame = Frame(config_window, bg=APP_BACKGROUND_COLOR, padx=20, pady=20)
        options_frame.pack(fill="both", expand=True)

        # Model type selection
        Label(options_frame, text="Residual Model Type:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=0, column=0, sticky="w", pady=10)

        model_types = ["RandomForest", "Ridge", "LinearRegression"]
        model_dropdown = OptionMenu(options_frame, model_type_var, *model_types)
        model_dropdown.grid(row=0, column=1, sticky="w", pady=10)

        # Status label
        status_label = Label(options_frame, text="Click Train Models to begin training with combined features",
                         bg=APP_BACKGROUND_COLOR, fg="white", font=FONT)
        status_label.grid(row=4, column=0, columnspan=2, pady=20)

        # Function to start training
        def start_training():
            status_label.config(text="Training models... Please wait.")
            config_window.update_idletasks()

            # Store configuration
            config = {
                "model_type": model_type_var.get()
            }

            # Start training in a thread
            training_thread = threading.Thread(
                target=lambda: train_models_thread(config, status_label, config_window)
            )
            training_thread.daemon = True
            training_thread.start()
    
        # Training thread function
        def train_models_thread(config, status_label, window):
            from sklearn.impute import SimpleImputer

            try:
                # Load data if not provided
                nonlocal data
                if data is None:
                    specific_file = './data_scraper/Master_Viscosity_Data_processed.csv'
            
                    try:
                        if os.path.exists(specific_file):
                            data = pd.read_csv(specific_file)
                            self.root.after(0, lambda: status_label.config(
                                text=f"Loaded {len(data)} data points..."
                            ))
                        else:
                            self.root.after(0, lambda: messagebox.showerror(
                                "Error", f"Master data file not found: {specific_file}"
                            ))
                            window.after(0, window.destroy)
                            return
                    except Exception as e:
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error", f"Failed to load data: {str(e)}"
                        ))
                        window.after(0, window.destroy)
                        return
        
                # Clean and prepare data
                self.root.after(0, lambda: status_label.config(text="Performing initial data cleaning..."))
                data_cleaned = data.copy()
        
                # Convert columns to numeric and handle errors
                num_columns = ['terpene_pct', 'temperature', 'viscosity', 'd9_thc', 'd8_thc', 'thca', 'total_potency']
                for col in num_columns:
                    if col in data_cleaned.columns:
                        data_cleaned[col] = pd.to_numeric(data_cleaned[col], errors='coerce')
        
                # Clean terpene values - safe replacement
                data_cleaned.loc[data_cleaned['terpene'].isna(), 'terpene'] = 'Raw'
                data_cleaned.loc[data_cleaned['terpene'] == '', 'terpene'] = 'Raw'
            
                # Add is_raw flag
                data_cleaned['is_raw'] = (data_cleaned['terpene'] == 'Raw')
            
                # Fill missing terpene_pct with 0 for raw data
                data_cleaned.loc[data_cleaned['is_raw'] & data_cleaned['terpene_pct'].isna(), 'terpene_pct'] = 0.0
                
                if 'total_potency' in data_cleaned.columns and 'terpene_pct' in data_cleaned.columns:
                    # Calculate theoretical maximum terpene percentage
                    data_cleaned['theoretical_max_terpene'] = 1.0 - data_cleaned['total_potency']
        
                    # Calculate how close the formulation is to theoretical maximum
                    data_cleaned['terpene_headroom'] = data_cleaned['theoretical_max_terpene'] - data_cleaned['terpene_pct']
        
                    # Flag physically impossible formulations (allowing for small measurement error)
                    data_cleaned['physically_valid'] = data_cleaned['terpene_pct'] <= (1.05 * data_cleaned['theoretical_max_terpene'])
        
                    # Calculate ratio as a proportion of theoretical maximum
                    data_cleaned['terpene_max_ratio'] = data_cleaned['terpene_pct'] / data_cleaned['theoretical_max_terpene'].clip(lower=0.01)
        
                    # Log information about the constraints
                    valid_pct = 100 * data_cleaned['physically_valid'].mean()
                    self.root.after(0, lambda p=valid_pct: status_label.config(
                        text=f"Added physical constraints. {p:.1f}% of formulations are physically valid."
                    ))


                # Calculate total_potency if missing
                if 'total_potency' in data_cleaned.columns:
                    missing_potency = data_cleaned['total_potency'].isna()
                    if missing_potency.any():
                        self.root.after(0, lambda: status_label.config(text="Calculating missing potency values..."))
                        # Get individual cannabinoid columns
                        cannabinoid_cols = [col for col in ['d9_thc', 'd8_thc', 'thca'] if col in data_cleaned.columns]
                    
                        if cannabinoid_cols:
                            # Calculate sum of available cannabinoids (skip NaN)
                            cannabinoid_sum = data_cleaned[cannabinoid_cols].sum(axis=1, skipna=True)
                            data_cleaned.loc[missing_potency, 'total_potency'] = cannabinoid_sum[missing_potency]
                        
                            self.root.after(0, lambda: status_label.config(
                                text=f"Calculated potency for {missing_potency.sum()} records"
                            ))
            
                # Apply Arrhenius transformation
                data_cleaned['temperature_kelvin'] = data_cleaned['temperature'] + 273.15
                data_cleaned['inverse_temp'] = 1 / data_cleaned['temperature_kelvin']
                data_cleaned['log_viscosity'] = np.log(data_cleaned['viscosity'])
            
                # -- Feature Selection and Classification Logic --
                self.root.after(0, lambda: status_label.config(text="Classifying data points by available features..."))
            
                # Identify points with different feature combinations
                has_potency = ~data_cleaned['total_potency'].isna() if 'total_potency' in data_cleaned.columns else pd.Series(False, index=data_cleaned.index)
                has_terpene = ~data_cleaned['terpene_pct'].isna()
            
                # NEW CODE: Add validation for physical constraints
                if 'total_potency' in data_cleaned.columns and 'terpene_pct' in data_cleaned.columns:
                    # Check for physically impossible formulations
                    impossible_mask = data_cleaned['terpene_pct'] > (1.0 - data_cleaned['total_potency'] + 0.05)
                    impossible_count = impossible_mask.sum()
    
                    if impossible_count > 0:
                        # Log a warning but don't remove the data
                        self.root.after(0, lambda c=impossible_count: status_label.config(
                            text=f"Warning: {c} data points exceed physical constraints (terpene% + potency% > 100%). "
                                 f"These may contain measurement errors but will still be used for training."
                        ))
        
                        # Optionally, you could add a flag for these points
                        data_cleaned['potential_error'] = impossible_mask

                # Create three distinct datasets:
                # 1. Combined: Points with both potency AND terpene data
                combined_mask = has_potency & has_terpene
                combined_data = data_cleaned[combined_mask].copy()
                combined_data['feature_type'] = 'combined'
            
                # 2. Potency-only: Points with ONLY potency data (no terpene data)
                potency_only_mask = has_potency & ~has_terpene
                potency_data = data_cleaned[potency_only_mask].copy()
                potency_data['feature_type'] = 'potency'
            
                # 3. Terpene-only: Points with ONLY terpene data (no potency data)
                terpene_only_mask = ~has_potency & has_terpene
                terpene_data = data_cleaned[terpene_only_mask].copy()
                terpene_data['feature_type'] = 'terpene'
            
                # Data stats for information
                self.root.after(0, lambda: status_label.config(
                    text=f"Found {len(combined_data)} points with both features, "
                         f"{len(potency_data)} with only potency, "
                         f"{len(terpene_data)} with only terpene"
                ))
            
                # Drop rows with NaN in critical columns from all datasets
                critical_cols = ['temperature', 'viscosity', 'inverse_temp', 'log_viscosity']
                combined_data = combined_data.dropna(subset=critical_cols + ['total_potency', 'terpene_pct'])
                potency_data = potency_data.dropna(subset=critical_cols + ['total_potency'])
                terpene_data = terpene_data.dropna(subset=critical_cols + ['terpene_pct'])
            
                self.root.after(0, lambda: status_label.config(
                    text=f"After NaN removal: {len(combined_data)} combined, "
                         f"{len(potency_data)} potency-only, "
                         f"{len(terpene_data)} terpene-only"
                ))
            
                # Model creation function
                def build_residual_model(config):
                    if config["model_type"] == "RandomForest":
                        return RandomForestRegressor(
                            n_estimators=100,
                            max_depth=3,
                            min_samples_leaf=5,
                            random_state=42
                        )
                    elif config["model_type"] == "Ridge":
                        return Ridge(alpha=1.0)
                    else:
                        return LinearRegression()
            
                # Initialize model dictionaries
                combined_models = {}  # Uses both features
                potency_models = {}   # Uses only potency
                terpene_models = {}   # Uses only terpene
            
                # Process media types for baseline models
                self.root.after(0, lambda: status_label.config(
                    text="Creating temperature baseline models..."
                ))
            
                # Combine all datasets for baseline modeling
                all_clean_data = pd.concat([combined_data, potency_data, terpene_data])
            
                media_types = all_clean_data['media'].unique()
                baseline_models = {}
            
                # Step 1: Create temperature baseline models for each media type
                for media in media_types:
                    media_data = all_clean_data[all_clean_data['media'] == media].copy()
            
                    if len(media_data) < 10:
                        continue
                
                    # Data is already cleaned for NaNs in critical columns
                    X_temp = media_data[['inverse_temp']]
                    y_temp = media_data['log_viscosity']
            
                    # Create and fit baseline model
                    temp_model = LinearRegression()
                    temp_model.fit(X_temp, y_temp)
            
                    baseline_models[media] = temp_model
            
                    # Calculate and store baseline predictions and residuals
                    baseline_preds = temp_model.predict(X_temp)
                    media_data['baseline_prediction'] = baseline_preds
                    media_data['residual'] = y_temp - baseline_preds
                
                    
                    def update_indices(df):
                        df.loc[df.index.intersection(media_data.index), 
                               ['baseline_prediction', 'residual']] = media_data.loc[
                                   df.index.intersection(media_data.index), 
                                   ['baseline_prediction', 'residual']]
                        # No return needed as dataframes are modified in-place

                    # Then call the function as before
                    update_indices(combined_data)
                    update_indices(potency_data)
                    update_indices(terpene_data)
            
                # Step 2A: Create combined-feature residual models
                if not combined_data.empty:
                    self.root.after(0, lambda: status_label.config(
                        text="Training combined-feature residual models..."
                    ))
                
                    # Process each media/terpene combination
                    for idx, (media, terpene) in enumerate(combined_data.groupby(['media', 'terpene']).groups.keys()):
                        # Update progress
                        progress = f"Processing combined model {idx+1}/{len(combined_data.groupby(['media', 'terpene']))} - {media}/{terpene}"
                        self.root.after(0, lambda p=progress: status_label.config(text=p))
                
                        # Skip if no baseline model
                        if media not in baseline_models:
                            continue
                    
                        # Filter data for this combination
                        combo_data = combined_data[
                            (combined_data['media'] == media) & 
                            (combined_data['terpene'] == terpene)
                        ]
                
                        # Skip if not enough data
                        if len(combo_data) < 5:
                            continue
                
                        # Get the baseline model
                        baseline_model = baseline_models[media]
                
                        # Create residual model using both features
                        combined_features = [
                            'total_potency', 
                            'terpene_pct', 
                            'is_raw',
                            'theoretical_max_terpene',   # Add the theoretical maximum
                            'terpene_headroom',          # Add the headroom
                            'terpene_max_ratio'          # Add the proportion of maximum
                        ]
                        # Make sure these columns exist in the data
                        for feature in combined_features:
                            if feature not in combo_data.columns:
                                combo_data[feature] = 0.0
                        # Add potency-terpene ratio as an interaction feature
                        combo_data['potency_terpene_ratio'] = combo_data['total_potency'] / combo_data['terpene_pct'].clip(lower=0.01)
                        combined_features.append('potency_terpene_ratio')

                        # Make sure residual column exists and has no NaNs
                        combo_data = combo_data.dropna(subset=['residual'])
                        if len(combo_data) < 5:
                            continue
                    
                        X_combined = combo_data[combined_features]
                        y_combined = combo_data['residual']
                
                        # Train residual model
                        combined_model = build_residual_model(config)
                        combined_model.fit(X_combined, y_combined)
                
                        # Store model with metadata
                        model_key = f"{media}_{terpene}"
                        combined_models[model_key] = {
                            'baseline_model': baseline_model,
                            'residual_model': combined_model,
                            'baseline_features': ['inverse_temp'],
                            'residual_features': combined_features,  # Now includes physical constraints
                            'metadata': {
                                'use_arrhenius': True,
                                'temperature_feature': 'inverse_temp',
                                'target_feature': 'log_viscosity',
                                'use_two_step': True,
                                'feature_type': 'combined',
                                'primary_features': ['total_potency', 'terpene_pct'],
                                'physical_features': ['theoretical_max_terpene', 'terpene_headroom', 'terpene_max_ratio']
                            }
                        }
                
                    # Create general combined models for each media type
                    self.root.after(0, lambda: status_label.config(
                        text="Training general combined models..."
                    ))
                
                    for media in media_types:
                        if media not in baseline_models:
                            continue
                    
                        # Get data for this media with both features
                        media_data = combined_data[combined_data['media'] == media]
                    
                        if len(media_data) < 10:
                            continue
                    
                        # One-hot encode terpenes for general model
                        terpene_counts = media_data['terpene'].value_counts()
                        valid_terpenes = terpene_counts[terpene_counts >= 2].index.tolist()
                    
                        # Filter data with valid terpenes
                        filtered_data = media_data[media_data['terpene'].isin(valid_terpenes)]
                        if len(filtered_data) < 10:
                            continue
                    
                        # Create one-hot encoding
                        encoded_data = pd.get_dummies(
                            filtered_data,
                            columns=['terpene'],
                            prefix=['terpene']
                        )
                    
                        # Get terpene_X columns
                        terpene_cols = [c for c in encoded_data.columns if c.startswith('terpene_')]
                    
                        # Define features for general model
                        gen_features = [
                            'total_potency', 
                            'terpene_pct', 
                            'is_raw',
                            'theoretical_max_terpene',  
                            'terpene_headroom',        
                            'terpene_max_ratio'
                        ] + terpene_cols
                    
                        # Make sure residual column exists and has no NaNs
                        encoded_data = encoded_data.dropna(subset=['residual'])
                        if len(encoded_data) < 10:
                            continue
                    
                        X_gen = encoded_data[gen_features]
                        y_gen = encoded_data['residual']

                        # Check for NaN values in features
                        if X_gen.isna().any().any():
                            # First, save a copy of the column names
                            original_columns = X_gen.columns.tolist()
    
                            # Replace completely empty columns with zeros instead of using the imputer on them
                            for col in X_gen.columns:
                                if X_gen[col].isna().all():
                                    X_gen[col] = 0
    
                            # Now use imputer only on columns that have some values
                            columns_with_values = [col for col in X_gen.columns if not X_gen[col].isna().all()]
                            columns_to_impute = X_gen[columns_with_values]
    
                            if not columns_to_impute.empty and columns_to_impute.isna().any().any():
                                # Create an imputer for columns with at least some non-NaN values
                                imputer = SimpleImputer(strategy='mean')
                                imputed_values = imputer.fit_transform(columns_to_impute)
        
                                # Put the imputed values back
                                X_gen.loc[:, columns_with_values] = imputed_values

                        # Now fit the model with the cleaned data
                        gen_model = build_residual_model(config)
                        gen_model.fit(X_gen, y_gen)
                    
                        # Store general model
                        combined_models[f"{media}_general"] = {
                            'baseline_model': baseline_models[media],
                            'residual_model': gen_model,
                            'baseline_features': ['inverse_temp'],
                            'residual_features': gen_features,
                            'metadata': {
                                'use_arrhenius': True,
                                'temperature_feature': 'inverse_temp',
                                'target_feature': 'log_viscosity',
                                'use_two_step': True,
                                'feature_type': 'combined',
                                'primary_features': ['total_potency', 'terpene_pct']
                            }
                        }
            
                # Step 2B: Create potency-only residual models
                if not potency_data.empty:
                    self.root.after(0, lambda: status_label.config(
                        text="Training potency-only residual models..."
                    ))
                
                    # Process each media/terpene combination
                    for idx, (media, terpene) in enumerate(potency_data.groupby(['media', 'terpene']).groups.keys()):
                        # Update progress
                        progress = f"Processing potency model {idx+1}/{len(potency_data.groupby(['media', 'terpene']))} - {media}/{terpene}"
                        self.root.after(0, lambda p=progress: status_label.config(text=p))
                
                        # Skip if no baseline model
                        if media not in baseline_models:
                            continue
                    
                        # Filter data for this combination
                        combo_data = potency_data[
                            (potency_data['media'] == media) & 
                            (potency_data['terpene'] == terpene)
                        ]
                
                        # Skip if not enough data
                        if len(combo_data) < 5:
                            continue
                
                        # Get the baseline model
                        baseline_model = baseline_models[media]
                
                        # Create potency-based residual model
                        pot_features = ['total_potency', 'is_raw']
                    
                        # Make sure residual column exists and has no NaNs
                        combo_data = combo_data.dropna(subset=['residual'])
                        if len(combo_data) < 5:
                            continue
                    
                        X_pot = combo_data[pot_features]
                        y_pot = combo_data['residual']
                
                        # Train residual model
                        pot_model = build_residual_model(config)
                        pot_model.fit(X_pot, y_pot)
                
                        # Store model with metadata
                        model_key = f"{media}_{terpene}"
                        potency_models[model_key] = {
                            'baseline_model': baseline_model,
                            'residual_model': pot_model,
                            'baseline_features': ['inverse_temp'],
                            'residual_features': pot_features,
                            'metadata': {
                                'use_arrhenius': True,
                                'temperature_feature': 'inverse_temp',
                                'target_feature': 'log_viscosity',
                                'use_two_step': True,
                                'feature_type': 'potency',
                                'primary_feature': 'total_potency'
                            }
                        }
                
                    # Create general potency models for each media type
                    self.root.after(0, lambda: status_label.config(
                        text="Training general potency models..."
                    ))
                
                    for media in media_types:
                        if media not in baseline_models:
                            continue
                    
                        # Get data for this media with potency values
                        media_data = potency_data[potency_data['media'] == media]
                    
                        if len(media_data) < 10:
                            continue
                    
                        # One-hot encode terpenes for general model
                        terpene_counts = media_data['terpene'].value_counts()
                        valid_terpenes = terpene_counts[terpene_counts >= 2].index.tolist()
                    
                        # Filter data with valid terpenes
                        filtered_data = media_data[media_data['terpene'].isin(valid_terpenes)]
                        if len(filtered_data) < 10:
                            continue
                    
                        # Create one-hot encoding
                        encoded_data = pd.get_dummies(
                            filtered_data,
                            columns=['terpene'],
                            prefix=['terpene']
                        )
                    
                        # Get terpene_X columns
                        terpene_cols = [c for c in encoded_data.columns if c.startswith('terpene_')]
                    
                        # Define features for general model
                        gen_features = ['total_potency', 'is_raw'] + terpene_cols
                    
                        # Make sure residual column exists and has no NaNs
                        encoded_data = encoded_data.dropna(subset=['residual'])
                        if len(encoded_data) < 10:
                            continue
                    
                        X_gen = encoded_data[gen_features]
                        y_gen = encoded_data['residual']

                        # Check for NaN values in features
                        
                        if X_gen.isna().any().any():
                            # First, save a copy of the column names
                            original_columns = X_gen.columns.tolist()
    
                            # Replace completely empty columns with zeros instead of using the imputer on them
                            for col in X_gen.columns:
                                if X_gen[col].isna().all():
                                    X_gen[col] = 0
    
                            # Now use imputer only on columns that have some values
                            columns_with_values = [col for col in X_gen.columns if not X_gen[col].isna().all()]
                            columns_to_impute = X_gen[columns_with_values]
    
                            if not columns_to_impute.empty and columns_to_impute.isna().any().any():
                                # Create an imputer for columns with at least some non-NaN values
                                imputer = SimpleImputer(strategy='mean')
                                imputed_values = imputer.fit_transform(columns_to_impute)
        
                                # Put the imputed values back
                                X_gen.loc[:, columns_with_values] = imputed_values

                        # Now fit the model with the cleaned data
                        gen_model = build_residual_model(config)
                        gen_model.fit(X_gen, y_gen)
                    
                        # Store general model
                        potency_models[f"{media}_general"] = {
                            'baseline_model': baseline_models[media],
                            'residual_model': gen_model,
                            'baseline_features': ['inverse_temp'],
                            'residual_features': gen_features,
                            'metadata': {
                                'use_arrhenius': True,
                                'temperature_feature': 'inverse_temp',
                                'target_feature': 'log_viscosity',
                                'use_two_step': True,
                                'feature_type': 'potency',
                                'primary_feature': 'total_potency'
                            }
                        }
            
                # Step 2C: Create terpene-only residual models
                if not terpene_data.empty:
                    self.root.after(0, lambda: status_label.config(
                        text="Training terpene-only residual models..."
                    ))
                
                    # Process each media/terpene combination
                    for idx, (media, terpene) in enumerate(terpene_data.groupby(['media', 'terpene']).groups.keys()):
                        # Update progress
                        progress = f"Processing terpene model {idx+1}/{len(terpene_data.groupby(['media', 'terpene']))} - {media}/{terpene}"
                        self.root.after(0, lambda p=progress: status_label.config(text=p))
                
                        # Skip if no baseline model
                        if media not in baseline_models:
                            continue
                    
                        # Filter data for this combination
                        combo_data = terpene_data[
                            (terpene_data['media'] == media) & 
                            (terpene_data['terpene'] == terpene)
                        ]
                
                        # Skip if not enough data
                        if len(combo_data) < 5:
                            continue
                
                        # Get the baseline model
                        baseline_model = baseline_models[media]
                
                        # Create terpene-based residual model
                        terp_features = ['terpene_pct', 'is_raw']
                    
                        # Make sure residual column exists and has no NaNs
                        combo_data = combo_data.dropna(subset=['residual'])
                        if len(combo_data) < 5:
                            continue
                    
                        X_terp = combo_data[terp_features]
                        y_terp = combo_data['residual']
                
                        # Train residual model
                        terp_model = build_residual_model(config)
                        terp_model.fit(X_terp, y_terp)
                
                        # Store model with metadata
                        model_key = f"{media}_{terpene}"
                        terpene_models[model_key] = {
                            'baseline_model': baseline_model,
                            'residual_model': terp_model,
                            'baseline_features': ['inverse_temp'],
                            'residual_features': terp_features,
                            'metadata': {
                                'use_arrhenius': True,
                                'temperature_feature': 'inverse_temp',
                                'target_feature': 'log_viscosity',
                                'use_two_step': True,
                                'feature_type': 'terpene',
                                'primary_feature': 'terpene_pct'
                            }
                        }
                
                    # Create general terpene models for each media type
                    self.root.after(0, lambda: status_label.config(
                        text="Training general terpene models..."
                    ))
                
                    for media in media_types:
                        if media not in baseline_models:
                            continue
                    
                        # Get data for this media with terpene values
                        media_data = terpene_data[terpene_data['media'] == media]
                    
                        if len(media_data) < 10:
                            continue
                    
                        # One-hot encode terpenes for general model
                        terpene_counts = media_data['terpene'].value_counts()
                        valid_terpenes = terpene_counts[terpene_counts >= 2].index.tolist()
                    
                        # Filter data with valid terpenes
                        filtered_data = media_data[media_data['terpene'].isin(valid_terpenes)]
                        if len(filtered_data) < 10:
                            continue
                    
                        # Create one-hot encoding
                        encoded_data = pd.get_dummies(
                            filtered_data,
                            columns=['terpene'],
                            prefix=['terpene']
                        )
                    
                        # Get terpene_X columns
                        terpene_cols = [c for c in encoded_data.columns if c.startswith('terpene_')]
                    
                        # Define features for general model
                        gen_features = ['terpene_pct', 'is_raw'] + terpene_cols
                    
                        # Make sure residual column exists and has no NaNs
                        encoded_data = encoded_data.dropna(subset=['residual'])
                        if len(encoded_data) < 10:
                            continue
                    
                        X_gen = encoded_data[gen_features]
                        y_gen = encoded_data['residual']

                        # Check for NaN values in features
                        if X_gen.isna().any().any():
                            # First, save a copy of the column names
                            original_columns = X_gen.columns.tolist()
    
                            # Replace completely empty columns with zeros instead of using the imputer on them
                            for col in X_gen.columns:
                                if X_gen[col].isna().all():
                                    X_gen[col] = 0
    
                            # Now use imputer only on columns that have some values
                            columns_with_values = [col for col in X_gen.columns if not X_gen[col].isna().all()]
                            columns_to_impute = X_gen[columns_with_values]
    
                            if not columns_to_impute.empty and columns_to_impute.isna().any().any():
                                # Create an imputer for columns with at least some non-NaN values
                                imputer = SimpleImputer(strategy='mean')
                                imputed_values = imputer.fit_transform(columns_to_impute)
        
                                # Put the imputed values back
                                X_gen.loc[:, columns_with_values] = imputed_values

                        # Now fit the model with the cleaned data
                        gen_model = build_residual_model(config)
                        gen_model.fit(X_gen, y_gen)
                    
                        # Store general model
                        terpene_models[f"{media}_general"] = {
                            'baseline_model': baseline_models[media],
                            'residual_model': gen_model,
                            'baseline_features': ['inverse_temp'],
                            'residual_features': gen_features,
                            'metadata': {
                                'use_arrhenius': True,
                                'temperature_feature': 'inverse_temp',
                                'target_feature': 'log_viscosity',
                                'use_two_step': True,
                                'feature_type': 'terpene',
                                'primary_feature': 'terpene_pct'
                            }
                        }
            
                # Save models
                os.makedirs('models', exist_ok=True)
            
                # Save combined models (new!)
                if combined_models:
                    self.root.after(0, lambda: status_label.config(
                        text=f"Saving {len(combined_models)} combined-feature models..."
                    ))
                    with open('models/viscosity_models_combined.pkl', 'wb') as f:
                        pickle.dump(combined_models, f)
                    # Store in a new attribute
                    self.combined_viscosity_models = combined_models
                
                # Save potency-based models
                if potency_models:
                    self.root.after(0, lambda: status_label.config(
                        text=f"Saving {len(potency_models)} potency-only models..."
                    ))
                    with open('models/viscosity_models_potency.pkl', 'wb') as f:
                        pickle.dump(potency_models, f)
                    self.enhanced_viscosity_models = potency_models
            
                # Save terpene-based models
                if terpene_models:
                    self.root.after(0, lambda: status_label.config(
                        text=f"Saving {len(terpene_models)} terpene-only models..."
                    ))
                    with open('models/viscosity_models.pkl', 'wb') as f:
                        pickle.dump(terpene_models, f)
                    self.viscosity_models = terpene_models
            
                # Show success message
                message = f"Training complete!\n\n"
                message += f"Created {len(combined_models)} combined-feature models\n"
                message += f"Created {len(potency_models)} potency-only models\n"
                message += f"Created {len(terpene_models)} terpene-only models\n\n"
            
                total_points = len(all_clean_data)
                total_original = len(data)
                message += f"Data cleaning: Used {total_points} clean points out of {total_original} total points"
            
                self.root.after(0, lambda: messagebox.showinfo("Success", message))
            
                # Close window
                window.after(0, window.destroy)
        
            except Exception as e:
                import traceback
                error_msg = f"Training error: {str(e)}\n\n{traceback.format_exc()}"
                print(error_msg)
                self.root.after(0, lambda: messagebox.showerror("Error", f"Training failed: {str(e)}"))
                window.after(0, window.destroy)

        # Create button frame
        button_frame = Frame(config_window, bg=APP_BACKGROUND_COLOR)
        button_frame.pack(pady=10)

        # Add buttons
        ttk.Button(button_frame, text="Train Models", command=start_training).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=config_window.destroy).pack(side="left", padx=10)

    def predict_viscosity(self, model_info, inputs):
        """
        Predict viscosity using the two-step model approach.
        Works with combined, potency-only, and terpene-only models.

        Args:
            model_info (dict): Model info with baseline and residual models
            inputs (dict): Dictionary of input values for prediction

        Returns:
            float: Predicted viscosity
        """
        import numpy as np
        import pandas as pd
        # Extract models and features
        baseline_model = model_info['baseline_model']
        residual_model = model_info['residual_model']
        baseline_features = model_info.get('baseline_features', ['inverse_temp'])
        residual_features = model_info.get('residual_features', [])

        # Get metadata about transformations
        metadata = model_info.get('metadata', {})
        use_arrhenius = metadata.get('use_arrhenius', True)
        feature_type = metadata.get('feature_type', 'terpene')

        # Create baseline feature vector
        baseline_vector = []

        for feature in baseline_features:
            if feature == 'inverse_temp' and 'inverse_temp' not in inputs and 'temperature' in inputs:
                # Calculate inverse_temp from temperature
                temperature_kelvin = inputs['temperature'] + 273.15
                inverse_temp = 1 / temperature_kelvin
                baseline_vector.append(inverse_temp)
            elif feature in inputs:
                baseline_vector.append(inputs[feature])
            else:
                baseline_vector.append(0)  # Default value

        # Get baseline prediction
        baseline_prediction = baseline_model.predict([baseline_vector])[0]

        # Create residual feature vector
        residual_vector = []

        for feature in residual_features:
            if feature in inputs:
                residual_vector.append(inputs[feature])
            elif feature.startswith('terpene_') and not feature in ['terpene_pct']:
                # Handle terpene one-hot encoding for general models
                terpene_name = feature.replace('terpene_', '')
                if 'terpene' in inputs and inputs['terpene'] == terpene_name:
                    residual_vector.append(1)
                else:
                    residual_vector.append(0)
            else:
                residual_vector.append(0)  # Default value

        if 'potency_terpene_ratio' in residual_features and 'total_potency' in residual_inputs and 'terpene_pct' in residual_inputs:
            residual_inputs['potency_terpene_ratio'] = residual_inputs['total_potency'] / max(0.01, residual_inputs['terpene_pct'])

        # Convert to numpy array for imputation if needed
        residual_vector = np.array(residual_vector).reshape(1, -1)
    
        # Check if residual model expects more features than we provided
        # This could happen with general models that have one-hot encoded features
        if hasattr(residual_model, 'n_features_in_') and residual_model.n_features_in_ != residual_vector.shape[1]:
            # Handle mismatch by padding with zeros or other strategy
            if residual_vector.shape[1] < residual_model.n_features_in_:
                padding = np.zeros((1, residual_model.n_features_in_ - residual_vector.shape[1]))
                residual_vector = np.hstack((residual_vector, padding))
    
        # Check for NaN values in residual vector
        if np.isnan(residual_vector).any():
            # Use SimpleImputer to replace NaN values with mean (or other strategy)
            imputer = SimpleImputer(strategy='mean')
            # Since we only have one sample, use mean=0 for simplicity
            imputer = SimpleImputer(strategy='constant', fill_value=0)
            residual_vector = imputer.fit_transform(residual_vector)

        # Convert list to DataFrame for consistent handling with training code
        residual_df = pd.DataFrame([residual_vector], columns=residual_features)

        # Handle NaNs exactly as you do in training
        for col in residual_df.columns:
            if residual_df[col].isna().all():
                residual_df[col] = 0

        # Apply imputer only to columns with some values
        columns_with_values = [col for col in residual_df.columns if not residual_df[col].isna().all()]
        columns_to_impute = residual_df[columns_with_values]

        if not columns_to_impute.empty and columns_to_impute.isna().any().any():
            imputer = SimpleImputer(strategy='mean')
            imputed_values = imputer.fit_transform(columns_to_impute)
            residual_df.loc[:, columns_with_values] = imputed_values

        # Convert back to numpy array for prediction
        residual_vector = residual_df.values[0]

        # Get residual prediction
        residual_prediction = residual_model.predict([residual_vector])[0]

        # Combine predictions
        combined_prediction = baseline_prediction + residual_prediction

        # Transform back if using Arrhenius
        if use_arrhenius:
            return np.exp(combined_prediction)
        else:
            return combined_prediction
        
    def load_standard_models(self):
        """Load standard viscosity models from disk with format verification."""
        import pickle
        import os
        
        model_path = 'models/viscosity_models.pkl'
        if os.path.exists(model_path):
            try:
                with open(model_path, 'rb') as f:
                    models = pickle.load(f)
                
                    # Quick structure check on first model
                    if models:
                        sample_key = next(iter(models))
                        sample_model = models[sample_key]
                    
                        # Check if this uses the new structure
                        if (isinstance(sample_model, dict) and 
                            'baseline_model' in sample_model and 
                            'residual_model' in sample_model):
                            self._viscosity_models = models
                        else:
                            # Old format detected
                            print("Warning: Old model format detected. Consider retraining models.")
                            if hasattr(self, '_viscosity_models'):
                                # Keep current models if they exist
                                if not self._viscosity_models:
                                    self._viscosity_models = {}
                            else:
                                self._viscosity_models = {}
                    else:
                        if hasattr(self, '_viscosity_models') and self._viscosity_models:
                            # Keep existing models
                            pass
                        else:
                            self._viscosity_models = {}
                    
            except Exception as e:
                print(f"Error loading standard models: {e}")
                self._viscosity_models = {}
        else:
            if hasattr(self, '_viscosity_models') and self._viscosity_models:
                # Keep existing models
                pass
            else:
                self._viscosity_models = {}
        
    def load_enhanced_models(self):
        """Load enhanced viscosity models with chemistry from disk with format verification."""
        import pickle
        import os
    
        model_path = 'models/viscosity_models_potency.pkl'
        if os.path.exists(model_path):
            try:
                with open(model_path, 'rb') as f:
                    models = pickle.load(f)
                
                    # Quick structure check on first model
                    if models:
                        sample_key = next(iter(models))
                        sample_model = models[sample_key]
                    
                        # Check if this uses the new structure
                        if (isinstance(sample_model, dict) and 
                            'baseline_model' in sample_model and 
                            'residual_model' in sample_model):
                            self._enhanced_viscosity_models = models
                        else:
                            # Old format detected
                            print("Warning: Old enhanced model format detected. Consider retraining models.")
                            if hasattr(self, '_enhanced_viscosity_models'):
                                # Keep current models if they exist
                                if not self._enhanced_viscosity_models:
                                    self._enhanced_viscosity_models = {}
                            else:
                                self._enhanced_viscosity_models = {}
                    else:
                        if hasattr(self, '_enhanced_viscosity_models') and self._enhanced_viscosity_models:
                            # Keep existing models
                            pass
                        else:
                            self._enhanced_viscosity_models = {}
                    
            except Exception as e:
                print(f"Error loading enhanced models: {e}")
                self._enhanced_viscosity_models = {}
        else:
            if hasattr(self, '_enhanced_viscosity_models') and self._enhanced_viscosity_models:
                # Keep existing models
                pass
            else:
                self._enhanced_viscosity_models = {}

    def load_combined_models(self):
        """Load combined viscosity models (terpene + potency) from disk with format verification."""
        import pickle
        import os
    
        model_path = 'models/viscosity_models_combined.pkl'
        if os.path.exists(model_path):
            try:
                with open(model_path, 'rb') as f:
                    models = pickle.load(f)
                
                    # Quick structure check on first model
                    if models:
                        sample_key = next(iter(models))
                        sample_model = models[sample_key]
                    
                        # Check if this uses the new structure
                        if (isinstance(sample_model, dict) and 
                            'baseline_model' in sample_model and 
                            'residual_model' in sample_model):
                            self._combined_viscosity_models = models
                        else:
                            # Old format detected
                            print("Warning: Old combined model format detected. Consider retraining models.")
                            if hasattr(self, '_combined_viscosity_models'):
                                # Keep current models if they exist
                                if not self._combined_viscosity_models:
                                    self._combined_viscosity_models = {}
                            else:
                                self._combined_viscosity_models = {}
                    else:
                        if hasattr(self, '_combined_viscosity_models') and self._combined_viscosity_models:
                            # Keep existing models
                            pass
                        else:
                            self._combined_viscosity_models = {}
                    
            except Exception as e:
                print(f"Error loading combined models: {e}")
                self._combined_viscosity_models = {}
        else:
            if hasattr(self, '_combined_viscosity_models') and self._combined_viscosity_models:
                # Keep existing models
                pass
            else:
                self._combined_viscosity_models = {}

    def calculate_terpene_percentage(self):
        """
        Calculate the terpene percentage needed to achieve target viscosity
        using the best available model (combined, potency, or terpene).
        """
        try:


            # First load all models if not loaded
            if not hasattr(self, 'viscosity_models') or not self.viscosity_models:
                self.load_standard_models()
            if not hasattr(self, 'enhanced_viscosity_models') or not self.enhanced_viscosity_models:
                self.load_enhanced_models()
            if not hasattr(self, 'combined_viscosity_models') or not self.combined_viscosity_models:
                self.load_combined_models()


            # Extract input values
            media = self.media_var.get()
            terpene = self.terpene_var.get() or "Raw"
            target_viscosity = float(self.target_viscosity_var.get())
            mass_of_oil = float(self.mass_of_oil_var.get())
        
            # Get potency values (if available)
            potency = self._total_potency_var.get()
            # Add constraint check and warning
            if potency > 0:
                theoretical_max_terpene = 100 * (1 - potency/100)
                # Warn if optimization finds a value beyond physical possibility
                if exact_value > theoretical_max_terpene:
                    messagebox.showinfo(
                        "Physical Constraint Notice",
                        f"The calculated terpene percentage ({exact_value:.1f}%) exceeds the "
                        f"theoretical maximum ({theoretical_max_terpene:.1f}%) for a formulation "
                        f"with {potency:.1f}% potency.\n\n"
                        f"Consider either reducing potency or accepting a higher viscosity."
                    )
            d9_thc = self._d9_thc_var.get()
            d8_thc = self._d8_thc_var.get()
        
            # Calculate total potency if not provided directly
            if potency == 0 and (d9_thc > 0 or d8_thc > 0):
                potency = d9_thc + d8_thc
        
            # Determine which model to use based on available data
            have_potency = potency > 0
            have_terpene = True  # We're trying to calculate terpene %, so we assume we'll add it
        
            # Find the best model in priority order:
            # 1. Combined model (if both potency and terpene are available)
            # 2. Potency-only model (if only potency is available)
            # 3. Terpene-only model (fallback)
        
            specific_key = f"{media}_{terpene}"
            general_key = f"{media}_general"
        
            model_info = None
            model_type = "unknown"
        
            # Check combined models first if both features are available
            if have_potency and have_terpene and hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models:
                combined_models = self.combined_viscosity_models
            
                # Try specific model first, then general
                if specific_key in combined_models:
                    model_info = combined_models[specific_key]
                    model_type = "combined"
                elif general_key in combined_models:
                    model_info = combined_models[general_key]
                    model_type = "combined"
        
            # Check potency models if no combined model was found
            if model_info is None and have_potency and hasattr(self, 'enhanced_viscosity_models') and self.enhanced_viscosity_models:
                potency_models = self.enhanced_viscosity_models
            
                # Try specific model first, then general
                if specific_key in potency_models:
                    model_info = potency_models[specific_key]
                    model_type = "potency"
                elif general_key in potency_models:
                    model_info = potency_models[general_key]
                    model_type = "potency"
        
            # Finally, check terpene models as a fallback
            if model_info is None and hasattr(self, 'viscosity_models') and self.viscosity_models:
                terpene_models = self.viscosity_models
            
                # Try specific model first, then general
                if specific_key in terpene_models:
                    model_info = terpene_models[specific_key]
                    model_type = "terpene"
                elif general_key in terpene_models:
                    model_info = terpene_models[general_key]
                    model_type = "terpene"
        
            # If no model found, raise error
            if model_info is None:
                raise ValueError(f"No suitable model found for {media}/{terpene}. Please train models first.")
            
            # Get the metadata
            metadata = model_info.get('metadata', {})
        
            # Use optimization to find optimal terpene percentage
            from scipy.optimize import minimize_scalar
        
            def objective(terpene_pct):
                # Prepare inputs for prediction
                inputs = {
                    'temperature': 25.0,  # Fixed reference temperature
                    'is_raw': 0,
                    'terpene': terpene
                }
            
                # Set feature values based on model type
                if model_type == "combined":
                    # For combined models, set both features
                    inputs['total_potency'] = potency
                    inputs['terpene_pct'] = terpene_pct/100.0  # Convert to decimal
                elif model_type == "potency":
                    # For potency models, set only potency
                    inputs['total_potency'] = potency
                else:
                    # For terpene models, set only terpene_pct
                    inputs['terpene_pct'] = terpene_pct/100.0  # Convert to decimal
            
                # Add terpene one-hot encoding if needed for general models
                if model_key.endswith('_general'):
                    # Get all terpene features from the model
                    terpene_features = [f for f in model_info['residual_features'] 
                                       if f.startswith('terpene_') and f != 'terpene_pct']
        
                    # Add one-hot encoding for each terpene feature
                    for feature in terpene_features:
                        feature_terpene = feature.replace('terpene_', '')
                        inputs[feature] = 1 if feature_terpene == terpene else 0
    
                # Get predicted viscosity
                predicted_viscosity = self.predict_model_viscosity(model_info, inputs)
                return abs(predicted_viscosity - target_viscosity)

            # Find optimal terpene percentage (bounded between 0.1% and 15%)
            result = minimize_scalar(objective, bounds=(0.1, 15.0), method='bounded')
            exact_value = result.x
        
            # Calculate mass values
            exact_mass = mass_of_oil * (exact_value / 100)
            start_percent = min(exact_value * 1.1, 15.0)  # Cap at 15%
            start_mass = mass_of_oil * (start_percent / 100)
        
            # Update result variables
            self.exact_percent_var.set(f"{exact_value:.1f}%")
            self.exact_mass_var.set(f"{exact_mass:.2f}g")
            self.start_percent_var.set(f"{start_percent:.1f}%")
            self.start_mass_var.set(f"{start_mass:.2f}g")
        
            # Show confirmation with model type information
            if specific_key in globals()[f"{model_type}_models"]:
                model_key = specific_key
            else:
                model_key = general_key
            
            model_type_display = {
                "combined": "combined (potency + terpene)",
                "potency": "potency-based",
                "terpene": "terpene-based"
            }.get(model_type, model_type)
        
            messagebox.showinfo(
                "Calculation Complete", 
                f"Calculation performed using {model_type_display} model: {model_key}\n\n"
                f"For {exact_value:.1f}% terpenes, estimated viscosity: {target_viscosity:.1f}"
            )
        
        except Exception as e:
            import traceback
            traceback_str = traceback.format_exc()
            print(f"Error during calculation: {e}\n{traceback_str}")
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")

    def predict_model_viscosity(self, model_obj, terpene_pct, temp, pot=None):
        """
        Helper method to prepare inputs and call predict_viscosity with enhanced potency handling.
    
        Args:
            model_obj: The viscosity model object
            terpene_pct: Terpene percentage (0-100)
            temp: Temperature in Celsius
            pot: Potency percentage (0-100), optional
    
        Returns:
            float: Predicted viscosity
        """
        import numpy as np
        import pandas as pd
        from sklearn.impute import SimpleImputer
    
        try:
            # Extract model components
            baseline_model = model_obj['baseline_model']
            residual_model = model_obj['residual_model']
            baseline_features = model_obj.get('baseline_features', ['inverse_temp'])
            residual_features = model_obj.get('residual_features', [])
        
            # Get metadata 
            metadata = model_obj.get('metadata', {})
            feature_type = metadata.get('feature_type', 'terpene')
        
            # Calculate temperature in Kelvin and inverse temperature
            temp_kelvin = temp + 273.15
            inverse_temp = 1 / temp_kelvin
        
            # Debug information to help troubleshoot
            debug_info = {}
            # Prepare baseline input for temperature effect
            baseline_df = pd.DataFrame({baseline_features[0]: [inverse_temp]})
        
            # Get baseline prediction (temperature effect only)
            log_visc_baseline = baseline_model.predict(baseline_df)[0]
            debug_info['log_visc_baseline'] = log_visc_baseline
        
            # Prepare residual feature inputs
            residual_inputs = {}
        
            # Handle terpene percentage - ensure decimal format (0-1)
            terp_decimal = terpene_pct / 100.0 if terpene_pct > 1.0 else terpene_pct
            if 'terpene_pct' in residual_features:
                residual_inputs['terpene_pct'] = terp_decimal
                debug_info['terpene_pct'] = terp_decimal
        
            # Handle potency - ensure decimal format (0-1)
            if pot is not None and 'total_potency' in residual_features:
                # Convert from percentage (0-100) to decimal (0-1) if needed
                pot_decimal = pot / 100.0 if pot > 1.0 else pot
                residual_inputs['total_potency'] = pot_decimal
                debug_info['total_potency'] = pot_decimal
            elif 'total_potency' in residual_features:
                # If potency not provided but required, estimate from terpene
                est_potency = max(0.7, min(0.99, 1.0 - terp_decimal))  # Capped at reasonable range
                residual_inputs['total_potency'] = est_potency
                debug_info['total_potency'] = est_potency
        
            # Add is_raw flag
            if 'is_raw' in residual_features:
                residual_inputs['is_raw'] = 0  # Assume not raw
        
            # Handle model_feature (special case)
            if 'model_feature' in residual_features:
                # Use whichever feature is primary for this model
                if 'total_potency' in residual_inputs and feature_type in ['potency', 'combined']:
                    residual_inputs['model_feature'] = residual_inputs['total_potency']
                    debug_info['model_feature'] = residual_inputs['model_feature']
                elif 'terpene_pct' in residual_inputs:
                    residual_inputs['model_feature'] = residual_inputs['terpene_pct']
                    debug_info['model_feature'] = residual_inputs['model_feature']
                
            # Always calculate potency-terpene ratio, even if not explicitly in features
            # This ensures continuity in predictions and fills important derived values
            if 'total_potency' in residual_inputs and 'terpene_pct' in residual_inputs:
                # Use a less extreme denominator to avoid huge ratio jumps
                ratio = residual_inputs['total_potency'] / max(0.05, residual_inputs['terpene_pct'])
                residual_inputs['potency_terpene_ratio'] = ratio
                debug_info['potency_terpene_ratio'] = ratio
            
                # Also add squared terms for better polynomial approximation
                residual_inputs['potency_squared'] = residual_inputs['total_potency'] ** 2
                residual_inputs['terpene_squared'] = residual_inputs['terpene_pct'] ** 2
                debug_info['polynomial_terms_added'] = True
        
            # Handle one-hot encoded terpene features
            for feature in residual_features:
                if feature.startswith('terpene_') and feature != 'terpene_pct':
                    # All features default to 0
                    residual_inputs[feature] = 0
        
            # Convert to DataFrame with all required features
            residual_df = pd.DataFrame([residual_inputs])
        
            # Add missing features with zeros
            for feature in residual_features:
                if feature not in residual_df.columns:
                    residual_df[feature] = 0
        
            # Ensure features are in the right order
            residual_df = residual_df[residual_features]

            # Add verbose debug for feature sets
            debug_info['features_provided'] = list(residual_inputs.keys())
            debug_info['features_required'] = residual_features
        
            # Handle NaN values if present
            if residual_df.isna().any().any():
                imputer = SimpleImputer(strategy='constant', fill_value=0)
                residual_df = pd.DataFrame(
                    imputer.fit_transform(residual_df),
                    columns=residual_features
                )
        
            # Get residual prediction
            log_visc_residual = residual_model.predict(residual_df)[0]
            debug_info['log_visc_residual'] = log_visc_residual
        
            # Combine predictions and convert back from log scale
            log_visc_total = log_visc_baseline + log_visc_residual
            viscosity = np.exp(log_visc_total)
            debug_info['final_viscosity'] = viscosity
        
            return viscosity
        
        except Exception as e:
            import traceback
            print(f"Error in predict_model_viscosity: {str(e)}")
            print(traceback.format_exc())
            return float('nan')

    def analyze_potency_sensitivity(self, model_key, model):
        """
        Analyze a model's sensitivity to potency changes
    
        Args:
            model_key: Key identifying the model
            model: The model object to analyze
        
        Returns:
            dict: Analysis results
        """
        import numpy as np
        import pandas as pd
    
        # Extract residual model and features
        residual_model = model['residual_model']
        residual_features = model['residual_features']
    
        # Check if 'total_potency' is in features
        if 'total_potency' not in residual_features:
            return {"result": "No potency feature in this model"}
    
        # Get feature importance if available
        if hasattr(residual_model, 'feature_importances_'):
            feature_importances = residual_model.feature_importances_
            feature_dict = dict(zip(residual_features, feature_importances))
        
            potency_importance = feature_dict.get('total_potency', 0)
        
            if potency_importance < 0.01:
                return {
                    "result": f"Very low potency importance: {potency_importance:.4f}",
                    "suggestion": "Model doesn't use potency for predictions"
                }
    
        # Test with range of potency values at fixed temp and terpene
        test_temp = 25.0
        test_terpene = 5.0  # Use a standard value
    
        test_potencies = np.linspace(0.7, 0.95, 6)  # Test more values
        predictions = []
    
        for pot in test_potencies:
            pred = self.predict_model_viscosity(model, test_terpene, test_temp, pot)
            predictions.append(pred)
    
        # Check for variation in predictions
        min_pred = min(predictions)
        max_pred = max(predictions)
    
        if min_pred == max_pred:
            return {
                "result": "No variation in predictions across potency values",
                "suggestion": "Model doesn't respond to potency changes"
            }
    
        variation_pct = (max_pred - min_pred) / min_pred * 100
    
        if variation_pct < 5:
            result = {
                "result": f"Low variation ({variation_pct:.1f}%) in predictions across potency values",
                "predictions": dict(zip([f"{p*100:.1f}%" for p in test_potencies], predictions)),
                "suggestion": "Retrain model with more diverse potency data"
            }
        else:
            result = {
                "result": f"Good variation ({variation_pct:.1f}%) in predictions across potency values",
                "predictions": dict(zip([f"{p*100:.1f}%" for p in test_potencies], predictions))
            }
    
        return result

    def diagnose_models(self):
        """Diagnose issues with feature importance in models"""
        print("\nModel Feature Importance Analysis")
        print("=================================")
    
        if hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models:
            models = self.combined_viscosity_models
            print(f"Analyzing {len(models)} combined models:")
        
            for model_key, model in models.items():
                print(f"\nModel: {model_key}")
            
                # Extract residual model
                residual_model = model['residual_model']
                residual_features = model['residual_features']
            
                # Check feature importance
                if hasattr(residual_model, 'feature_importances_'):
                    importances = residual_model.feature_importances_
                    for i, feature in enumerate(residual_features):
                        if i < len(importances):  # Ensure index is valid
                            importance = importances[i]
                            print(f"  * {feature}: {importance:.6f}")
                        
                            # Flag problems
                            if feature == 'total_potency' and importance < 0.01:
                                print("    WARNING: Potency has extremely low importance")
                            elif feature == 'total_potency' and importance < 0.1:
                                print("    WARNING: Potency has low importance")
                        
                elif hasattr(residual_model, 'coef_'):
                    # For linear models
                    coefs = residual_model.coef_
                    for i, feature in enumerate(residual_features):
                        if i < len(coefs):  # Ensure index is valid
                            coef = coefs[i] if len(coefs.shape) == 1 else coefs[0, i]
                            print(f"  * {feature} coefficient: {coef:.6f}")
                        
                            # Flag problems
                            if feature == 'total_potency' and abs(coef) < 0.01:
                                print("    WARNING: Potency has extremely low coefficient")
                            elif feature == 'total_potency' and abs(coef) < 0.1:
                                print("    WARNING: Potency has low coefficient")
                else:
                    print("  WARNING: No feature importance information available")
                
                # Test potency variation
                potencies = [70, 75, 80, 85, 90]
                viscosities = []
            
                for pot in potencies:
                    visc = self.predict_model_viscosity(model, 5.0, 25.0, pot)
                    viscosities.append(visc)
                
                # Check for variation
                if len(set([round(v, 2) for v in viscosities])) == 1:
                    print("  CRITICAL ERROR: Model shows no response to potency variation")
                else:
                    min_visc = min(viscosities)
                    max_visc = max(viscosities)
                    variation = (max_visc - min_visc) / min_visc * 100
                    print(f"  * Potency variation effect: {variation:.2f}% change in viscosity")
                
                    # Print values
                    for i, pot in enumerate(potencies):
                        print(f"    - {pot}%: {viscosities[i]:.0f} cP")

    def analyze_model_feature_response(self, model_key=None, model=None):
        """
        Analyze how a model responds to changes in feature values
        """
        import pandas as pd
        import numpy as np
        import matplotlib.pyplot as plt
        from tkinter import Toplevel, Label, Frame
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    
        # If no specific model provided, use a model from combined models
        if model is None:
            if model_key is None:
                # Use first model in combined models
                if hasattr(self, 'combined_viscosity_models') and self.combined_viscosity_models:
                    model_key = next(iter(self.combined_viscosity_models))
                    model = self.combined_viscosity_models[model_key]
                else:
                    print("No combined models available for analysis")
                    return
            else:
                # Try to find the specified model
                if hasattr(self, 'combined_viscosity_models') and model_key in self.combined_viscosity_models:
                    model = self.combined_viscosity_models[model_key]
                else:
                    print(f"Model '{model_key}' not found in combined models")
                    return
    
        # Extract model components
        if not isinstance(model, dict) or 'residual_model' not in model:
            print("Invalid model structure for analysis")
            return
    
        # Create analysis window
        analysis_window = Toplevel(self.root)
        analysis_window.title(f"Model Feature Response Analysis: {model_key}")
        analysis_window.geometry("800x600")
    
        # Create figure for plots
        fig = plt.Figure(figsize=(10, 8), tight_layout=True)
    
        # 1. Potency Response Plot
        ax1 = fig.add_subplot(221)
        potencies = np.linspace(0.7, 1.0, 7)  # Range as decimal
        predictions = []
    
        # Fixed values for other parameters
        temperature = 25
        terpene_pct = 5.0
    
        for pot in potencies:
            # Convert potency to percentage for prediction function
            pot_pct = pot * 100
            pred = self.predict_model_viscosity(model, terpene_pct, temperature, pot_pct)
            predictions.append(pred)
    
        ax1.plot(potencies * 100, predictions, 'o-', linewidth=2)
        ax1.set_xlabel('Potency (%)')
        ax1.set_ylabel('Viscosity (cP)')
        ax1.set_title('Viscosity vs Potency')
        ax1.grid(True)
    
        # Calculate responsiveness
        if max(predictions) > min(predictions):
            potency_response = (max(predictions) - min(predictions)) / min(predictions) * 100
            ax1.annotate(f"Δ: {potency_response:.1f}%", 
                         xy=(0.05, 0.95), xycoords='axes fraction',
                         bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))
    
        # 2. Terpene Response Plot
        ax2 = fig.add_subplot(222)
        terpenes = np.linspace(1.0, 20.0, 10)  # Range in percentage
        predictions = []
    
        # Fixed potency (80%)
        potency = 80.0
    
        for terp in terpenes:
            pred = self.predict_model_viscosity(model, terp, temperature, potency)
            predictions.append(pred)
    
        ax2.plot(terpenes, predictions, 'o-', linewidth=2, color='green')
        ax2.set_xlabel('Terpene (%)')
        ax2.set_ylabel('Viscosity (cP)')
        ax2.set_title('Viscosity vs Terpene %')
        ax2.grid(True)
    
        # Calculate responsiveness
        if max(predictions) > min(predictions):
            terpene_response = (max(predictions) - min(predictions)) / min(predictions) * 100
            ax2.annotate(f"Δ: {terpene_response:.1f}%", 
                         xy=(0.05, 0.95), xycoords='axes fraction',
                         bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))
    
        # 3. Temperature Response Plot
        ax3 = fig.add_subplot(223)
        temperatures = np.linspace(20, 70, 11)  # Range in Celsius
        predictions = []
    
        # Fixed values
        terpene_pct = 5.0
        potency = 80.0
    
        for temp in temperatures:
            pred = self.predict_model_viscosity(model, terpene_pct, temp, potency)
            predictions.append(pred)
    
        ax3.semilogy(temperatures, predictions, 'o-', linewidth=2, color='red')
        ax3.set_xlabel('Temperature (°C)')
        ax3.set_ylabel('Viscosity (cP)')
        ax3.set_title('Viscosity vs Temperature')
        ax3.grid(True)
    
        # Calculate responsiveness
        if max(predictions) > min(predictions):
            temp_response = (max(predictions) - min(predictions)) / min(predictions) * 100
            ax3.annotate(f"Δ: {temp_response:.1f}%", 
                         xy=(0.05, 0.95), xycoords='axes fraction',
                         bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))
    
        # 4. Feature Importance
        ax4 = fig.add_subplot(224)
    
        residual_model = model['residual_model']
        residual_features = model.get('residual_features', [])
    
        if hasattr(residual_model, 'feature_importances_'):
            importances = residual_model.feature_importances_
            if len(importances) == len(residual_features):
                # Create bar chart of feature importances
                y_pos = range(len(residual_features))
                ax4.barh(y_pos, importances, align='center')
                ax4.set_yticks(y_pos)
                ax4.set_yticklabels(residual_features)
                ax4.set_xlabel('Importance')
                ax4.set_title('Feature Importance')
        elif hasattr(residual_model, 'coef_'):
            coefs = residual_model.coef_
            if len(coefs) == len(residual_features):
                # Create bar chart of coefficients
                y_pos = range(len(residual_features))
                ax4.barh(y_pos, np.abs(coefs), align='center')
                ax4.set_yticks(y_pos)
                ax4.set_yticklabels(residual_features)
                ax4.set_xlabel('|Coefficient|')
                ax4.set_title('Feature Coefficients')
        else:
            ax4.text(0.5, 0.5, "No feature importance data available", 
                     ha='center', va='center', transform=ax4.transAxes)
    
        # Create canvas to display figure
        canvas = FigureCanvasTkAgg(fig, master=analysis_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
        # Create button to close window
        from tkinter import Button
        Button(analysis_window, text="Close", command=analysis_window.destroy).pack(pady=10)

    def preprocess_training_data(self, data):
        """Fill missing potency or terpene values using inverse relationship"""
        processed_data = data.copy()
    
        # For rows with missing potency but available terpene percentage
        potency_missing = processed_data['total_potency'].isna()
        terpene_available = ~processed_data['terpene_pct'].isna()
    
        # Apply the inverse relationship: potency = 1 - terpene_pct (as decimal)
        mask = potency_missing & terpene_available
        if mask.any():
            processed_data.loc[mask, 'total_potency'] = 1.0 - (processed_data.loc[mask, 'terpene_pct'] / 100.0)
            print(f"Filled {mask.sum()} missing potency values using inverse relationship")
    
        # For rows with missing terpene but available potency
        terpene_missing = processed_data['terpene_pct'].isna()
        potency_available = ~processed_data['total_potency'].isna()
    
        # Apply the inverse relationship: terpene_pct = (1 - potency) * 100
        mask = terpene_missing & potency_available
        if mask.any():
            processed_data.loc[mask, 'terpene_pct'] = (1.0 - processed_data.loc[mask, 'total_potency']) * 100.0
            print(f"Filled {mask.sum()} missing terpene values using inverse relationship")
    
        return processed_data

    def create_potency_demo_model(self):
        """Create a demonstration model with very clear potency effects"""
        import numpy as np
        import pandas as pd
        from sklearn.linear_model import Ridge
    
        print("Creating enhanced potency demo model with strong effects...")
    
        # Define baseline model for temperature effects (Arrhenius)
        baseline_model = Ridge(alpha=1.0)
        baseline_log_visc = np.log(10000)  # Reference viscosity at 25°C
        activation_energy = 12000  # Controls temperature sensitivity
        gas_constant = 8.314  # Physical constant
    
        # Train baseline model on temperature relationship
        temps = np.array([20, 25, 30, 40, 50, 60])
        temps_kelvin = temps + 273.15
        inverse_temps = 1 / temps_kelvin
        baseline_log_viscs = baseline_log_visc + (activation_energy / gas_constant) * (1/298.15 - inverse_temps)
        X_baseline = pd.DataFrame({'inverse_temp': inverse_temps})
        y_baseline = baseline_log_viscs
        baseline_model.fit(X_baseline, y_baseline)
    
        # Create residual model with very strong potency effect
        residual_model = Ridge(alpha=0.01)  # Lower alpha for stronger fitting
    
        # Generate synthetic training data with extreme potency effect
        # Create more potency values for smoother curve
        potencies = np.linspace(0.7, 0.95, 10)  # More granular potency range
        terpene_pcts = np.linspace(0.03, 0.20, 7)  # Wider terpene range
    
        potency_vals = []
        terpene_vals = []
        pt_ratio_vals = []  # New feature: potency/terpene ratio
        residual_vals = []
    
        for pot in potencies:
            for terp in terpene_pcts:
                # STRONG potency effect - exponential relationship
                # Higher potency = exponentially higher viscosity
                potency_effect = 3.0 * np.exp(pot * 2.0) - 15.0
            
                # Standard terpene effect - inverse relationship
                # Higher terpene = lower viscosity
                terpene_effect = -5.0 * terp
            
                # Interaction effect (potency matters more at lower terpene %)
                interaction = -2.0 * pot * terp
            
                # Combined effect
                combined_effect = potency_effect + terpene_effect + interaction
            
                # Store values
                potency_vals.append(pot)
                terpene_vals.append(terp)
                pt_ratio_vals.append(pot / max(0.01, terp))  # Avoid division by zero
                residual_vals.append(combined_effect)
    
        # Create training dataframe
        X_residual = pd.DataFrame({
            'total_potency': potency_vals,
            'terpene_pct': terpene_vals,
            'potency_terpene_ratio': pt_ratio_vals,
            'is_raw': [0] * len(potency_vals)
        })
        y_residual = np.array(residual_vals)
    
        # Train residual model
        residual_model.fit(X_residual, y_residual)
    
        # Create model dictionary
        demo_model = {
            'baseline_model': baseline_model,
            'residual_model': residual_model,
            'baseline_features': ['inverse_temp'],
            'residual_features': ['total_potency', 'terpene_pct', 'potency_terpene_ratio', 'is_raw'],
            'metadata': {
                'use_arrhenius': True,
                'temperature_feature': 'inverse_temp',
                'target_feature': 'log_viscosity',
                'use_two_step': True,
                'feature_type': 'combined',
                'primary_features': ['total_potency', 'terpene_pct']
            }
        }
    
        # Test the model with different potency values
        print("Testing enhanced demo model with varying potency:")
        for pot in [70, 75, 80, 85, 90]:
            visc = self.predict_model_viscosity(demo_model, 5.0, 25.0, pot)
            print(f"  • Potency {pot}%: viscosity = {visc:.0f} cP")
    
        print("\nTesting with varying terpene percentages (at 80% potency):")
        for terp in [3, 5, 7, 10, 15]:
            visc = self.predict_model_viscosity(demo_model, terp, 25.0, 80)
            print(f"  • Terpene {terp}%: viscosity = {visc:.0f} cP")
    
        # Store in combined models dictionary
        if hasattr(self, 'combined_viscosity_models'):
            self.combined_viscosity_models['Enhanced_PotencyDemo'] = demo_model
            print("Enhanced demo model added as 'Enhanced_PotencyDemo'")
        else:
            self.combined_viscosity_models = {'Enhanced_PotencyDemo': demo_model}
    
        return demo_model
"""
viscosity_calculator.py
Module for calculating terpene percentages based on viscosity.
This version uses lazy loading to minimize startup time and memory usage.
"""
import os
import json
import tkinter as tk
from tkinter import ttk, StringVar, DoubleVar, IntVar, Toplevel, Frame, Label, Entry, Button, messagebox

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
        """Lazy load the viscosity models"""
        if self._viscosity_models is None:
            self._viscosity_models = self.load_models()
        return self._viscosity_models
    
    @viscosity_models.setter
    def viscosity_models(self, value):
        self._viscosity_models = value
        
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

        # Initialize total potency variable
        self.total_potency_var = DoubleVar(value=0.0)
    
        # Row 5: Total Potency
        Label(form_frame, text="Total Potency (%):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=5, column=0, sticky="w", pady=5)
    
        potency_entry = Entry(form_frame, textvariable=self.total_potency_var, width=15)
        potency_entry.grid(row=5, column=1, sticky="w", padx=5, pady=5)
        
        # Separator for results
        ttk.Separator(form_frame, orient='horizontal').grid(row=5, column=0, columnspan=4, sticky="ew", pady=15)
        
        # Results section - Row 6: Results header
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
            command=self.calculate_viscosity
        )
        calculate_btn.pack(padx=(0, 5))

        # Add a new button for calculating with potency
        calculate_with_potency_btn = ttk.Button(
            button_row1,
            text="Calculate with Potency",
            command=self.calculate_viscosity_with_potency
        )
        calculate_with_potency_btn.pack(padx=(5, 0))

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
        These are placeholder values - replace with your actual data.
        """
        # These are just example values - replace with actual values based on your data
        viscosity_map = {
            "D8": 18000000,
            "D9": 12000000,
            "Liquid Diamonds": 3000000,
            "Other": 500000 
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
        Calculate the second step amount of terpenes to add and
        predict the final viscosity.
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
            
            # Calculate how much the viscosity dropped per percent of terpenes
            step1_percent = (step1_amount / mass_of_oil) * 100
            viscosity_drop_per_percent = (raw_oil_viscosity - step1_viscosity) / step1_percent
            
            # Calculate how much more percentage we need
            percent_needed = (raw_oil_viscosity - target_viscosity) / viscosity_drop_per_percent
            percent_needed -= step1_percent  # Subtract what we've already added
            
            # Ensure we don't get negative percentages
            percent_needed = max(0, percent_needed)
            
            # Calculate amount for step 2
            step2_amount = (percent_needed / 100) * mass_of_oil
            
            # Update the UI
            self.step2_amount_var.set(f"{step2_amount:.2f}g")
            
            # Calculate expected final viscosity
            expected_viscosity = raw_oil_viscosity - (viscosity_drop_per_percent * (step1_percent + percent_needed))
            self.expected_viscosity_var.set(f"{expected_viscosity:.2f}")
            
            messagebox.showinfo("Step 2", 
                               f"Add an additional {step2_amount:.2f}g of {self.terpene_var.get()} terpenes.\n"
                               f"Mix thoroughly and then measure the final viscosity at 25C.\n"
                               f"The expected final viscosity is {expected_viscosity:.2f}.")
            
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")
    
    def calculate_viscosity(self):
        """
        Calculate terpene percentage based on target viscosity using trained models.
        """
        try:
            # Extract input values
            media = self.media_var.get()
            media_brand = self.media_brand_var.get()
            terpene = self.terpene_var.get()
            terpene_brand = self.terpene_brand_var.get()
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
        
            # Load the appropriate model for this media/terpene combination
            model_key = f"{media}_{terpene}"
        
            if model_key in self.viscosity_models:
                # Lazy-load scipy here since it's only needed for this calculation
                from scipy.optimize import fsolve
                import numpy as np
                
                model = self.viscosity_models[model_key]
            
                # For numerical solution, find terpene percentage that gives target viscosity
                def objective(terpene_pct):
                    # Predict viscosity at 25°C with given terpene percentage
                    predicted_viscosity = model.predict([[terpene_pct, 25.0]])[0]
                    return predicted_viscosity - target_viscosity
            
                # Solve for terpene percentage (start with a guess of 5%)
                exact_percent = fsolve(objective, 5.0)[0]
            
                # Ensure the percentage is reasonable
                exact_percent = max(0.1, min(15.0, exact_percent))
                exact_mass = mass_of_oil * (exact_percent / 100)
            
                # Suggested starting point (slightly higher)
                start_percent = exact_percent * 1.1
                start_mass = mass_of_oil * (start_percent / 100)
            
                # Update result variables
                self.exact_percent_var.set(f"{exact_percent:.1f}%")
                self.exact_mass_var.set(f"{exact_mass:.2f}g")
                self.start_percent_var.set(f"{start_percent:.1f}%")
                self.start_mass_var.set(f"{start_mass:.2f}g")
            
            else:
                # No model available, use iterative method
                messagebox.askyesno("No Model Available", 
                                  "No prediction model found for this combination.\n\n"
                                  "Would you like to use the Iterative Method instead?")
    
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")

    def save_formulation(self):
        """
        Save the terpene formulation to a database file.
        """
        try:
            # Verify we have all required data
            step2_viscosity_text = self.step2_viscosity_var.get()
            
            if not step2_viscosity_text:
                messagebox.showinfo("Input Needed", 
                                   "Please enter the final measured viscosity from Step 2 before saving.")
                return
            
            # Get all formulation data
            formulation = {
                "media": self.media_var.get(),
                "media_brand": self.media_brand_var.get(),
                "terpene": self.terpene_var.get(),
                "terpene_brand": self.terpene_brand_var.get(),
                "target_viscosity": float(self.target_viscosity_var.get()),
                "step1_amount": float(self.step1_amount_var.get().replace('g', '')),
                "step1_viscosity": float(self.step1_viscosity_var.get()),
                "step2_amount": float(self.step2_amount_var.get().replace('g', '')),
                "step2_viscosity": float(step2_viscosity_text),
                "expected_viscosity": float(self.expected_viscosity_var.get()),
                "total_oil_mass": float(self.mass_of_oil_var.get()),
                "total_terpene_mass": float(self.step1_amount_var.get().replace('g', '')) + 
                                     float(self.step2_amount_var.get().replace('g', '')),
                "total_terpene_percent": ((float(self.step1_amount_var.get().replace('g', '')) + 
                                         float(self.step2_amount_var.get().replace('g', ''))) / 
                                         float(self.mass_of_oil_var.get())) * 100
            }
            
            # Add to database
            key = f"{formulation['media']}_{formulation['media_brand']}_{formulation['terpene']}_{formulation['terpene_brand']}"
            
            if key not in self.formulation_db:
                self.formulation_db[key] = []
            
            self.formulation_db[key].append(formulation)
            
            # Save database to file
            self.save_formulation_database()
            
            messagebox.showinfo("Success", 
                               f"Formulation saved successfully!\n\n"
                               f"Total terpene percentage: {formulation['total_terpene_percent']:.2f}%\n"
                               f"Total terpene mass: {formulation['total_terpene_mass']:.2f}g")
            
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            messagebox.showerror("Save Error", f"An error occurred: {str(e)}")

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
        Save the measurements in a standardized CSV format for machine learning.
        
        Args:
            measurements (dict): The measurements data structure
        """
        # Lazy import pandas only when needed
        import pandas as pd
        import os
        import datetime
        
        # Create the data directory if it doesn't exist
        data_dir = 'data'
        try:
            if not os.path.exists(data_dir):
                os.makedirs(data_dir, exist_ok=True)
        except Exception as e:
            raise Exception(f"Failed to create data directory: {str(e)}")
        
        # Generate a timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create rows for the CSV
        rows = []
        
        media = measurements['media']
        terpene = measurements['terpene']
        terpene_pct = measurements['terpene_pct']
        
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
                        viscosity_float = float(viscosity)
                        row = {
                            'media': media,
                            'terpene': terpene,
                            'terpene_pct': terpene_pct,
                            'temperature': temperature,
                            'speed': speed,
                            'torque': torque,
                            'viscosity': viscosity_float
                        }
                        rows.append(row)
                    except ValueError as e:
                        print(f"Warning: Could not convert viscosity value '{viscosity}' to float: {e}")
            
            # Add the average if available
            avg_viscosity = temp_block.get('average_viscosity', '')
            if avg_viscosity and avg_viscosity.strip():
                try:
                    avg_viscosity_float = float(avg_viscosity)
                    row = {
                        'media': media,
                        'terpene': terpene,
                        'terpene_pct': terpene_pct,
                        'temperature': temperature,
                        'speed': speed,
                        'torque': temp_block.get('average_torque', ''),
                        'viscosity': avg_viscosity_float,
                        'is_average': True
                    }
                    rows.append(row)
                except ValueError as e:
                    print(f"Warning: Could not convert average viscosity value '{avg_viscosity}' to float: {e}")
        
        # Create a DataFrame and save to CSV
        if rows:
            df = pd.DataFrame(rows)
            filename = f"data/viscosity_data_{timestamp}.csv"
            df.to_csv(filename, index=False)
            return filename
        else:
            print("No valid data rows to save to CSV")
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
        Allow the user to upload CSV data for training viscosity models.
    
        The CSV should have columns: media, terpene, terpene_pct, temperature, viscosity
        """
        from tkinter import filedialog
        import pandas as pd
        import os
    
        # Prompt user to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select CSV with Viscosity Training Data",
            filetypes=[("CSV files", "*.csv")]
        )
    
        if not file_path:
            return None  # User canceled
    
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
        
            # Save a copy of the data in our data directory
            os.makedirs('data', exist_ok=True)
            backup_path = f"data/viscosity_data_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv"
            data.to_csv(backup_path, index=False)
        
            messagebox.showinfo("Success", 
                              f"Loaded {len(data)} data points from {os.path.basename(file_path)}.\n"
                              f"A copy has been saved to {backup_path}")
        
            return data
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
            return None

    def train_models_from_data(self, data=None):
        """
        Train viscosity prediction models using uploaded data or existing data.
        
        Args:
            data (pd.DataFrame, optional): DataFrame with training data. If None,
                                        attempts to use data from saved CSV files.
        """
        # Lazy imports - only needed when training models
        import threading
        import glob
        import traceback
        import pandas as pd
        from tkinter import Toplevel
        
        # If no data provided, try to load from saved files
        if data is None:
            data_files = glob.glob('data/viscosity_data_*.csv')
            if not data_files:
                messagebox.showerror("Error", "No training data available. Please upload data first.")
                return
        
            # Load and combine all data files
            data_frames = []
            for file in data_files:
                try:
                    df = pd.read_csv(file)
                    data_frames.append(df)
                except Exception as e:
                    print(f"Error loading {file}: {str(e)}")
                    continue
        
            if not data_frames:
                messagebox.showerror("Error", "Failed to load any training data.")
                return
            
            data = pd.concat(data_frames, ignore_index=True)

        # Show a progress dialog
        progress_window = Toplevel(self.root)
        progress_window.title("Training Models")
        progress_window.geometry("300x150")
        progress_window.transient(self.root)
        progress_window.grab_set()

        progress_label = Label(progress_window, text="Training models...", font=('Arial', 12))
        progress_label.pack(pady=20)

        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill='x', padx=20)
        progress_bar.start()

        # Create a new thread for training to avoid freezing the UI
        def train_thread():
            try:
                print("Starting model training...")
        
                # Try to import required libraries
                try:
                    import numpy as np
                    from sklearn.model_selection import train_test_split
                    from sklearn.linear_model import LinearRegression
                    from sklearn.preprocessing import PolynomialFeatures
                    from sklearn.ensemble import RandomForestRegressor
                    from sklearn.pipeline import make_pipeline
                    print("Successfully imported required ML libraries")
                except ImportError as e:
                    error_msg = str(e)
                    print(f"Error importing required libraries: {error_msg}")
                    self.root.after(0, lambda msg=error_msg: messagebox.showerror(
                        "Library Error", 
                        f"Required library not found: {msg}\nMake sure scikit-learn is installed."
                    ))
                    return
                
                # Import the build_viscosity_models function only when needed
                # This would normally be a separate function, but we'll define it inline for this example
                def build_viscosity_models(data):
                    """
                    Build multiple regression models for viscosity prediction and 
                    select the best one based on cross-validation.
                    """
                    # Print initial data shape for debugging
                    print(f"Initial data shape in build_viscosity_models: {data.shape}")
                    
                    # Extract features and target - data should already be cleaned
                    X = data[['terpene_pct', 'temperature']]
                    y = data['viscosity']
                    
                    # Double-check that we don't have any NaN values
                    if X.isna().any().any() or y.isna().any():
                        print("Warning: NaN values found in features or target. Cleaning data...")
                        # Get indices of rows with NaN values
                        rows_with_nan = X.index[X.isna().any(axis=1)].union(y.index[y.isna()])
                        print(f"Dropping {len(rows_with_nan)} rows with NaN values")
                        # Drop these rows
                        X = X.drop(rows_with_nan)
                        y = y.drop(rows_with_nan)
                    
                    # Check that we have enough data for splitting
                    if len(X) < 5:  # Need at least a few samples for meaningful train/test split
                        print(f"WARNING: Only {len(X)} samples after cleaning. Using all data for training.")
                        # Create a simple model using all data
                        model = LinearRegression()
                        model.fit(X, y)
                        return model, {"small_data": True}

                    # Ensure at least 5 samples in test set, or 20% of data, whichever is larger
                    min_test_size = max(5, int(0.2 * len(X)))
                    test_size = min(min_test_size / len(X), 0.4)  # Cap at 40% max
                    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=test_size, random_state=42)
                    
                    print(f"Training set size: {len(X_train)}, Test set size: {len(X_test)}")

                    # Create models
                    models = {
                        'linear': LinearRegression(),
                        'polynomial': make_pipeline(PolynomialFeatures(2), LinearRegression()),
                        'random_forest': RandomForestRegressor(
                            n_estimators=50,        # Fewer trees
                            max_depth=5,            # Limit tree depth
                            min_samples_leaf=3,     # Require more samples per leaf
                            random_state=42
                        )
                    }

                    # Evaluate models
                    results = {}
                    best_score = -float('inf')
                    best_model_name = None
                    best_model = None

                    for name, model in models.items():
                        try:
                            print(f"Training {name} model...")
                            # Try to fit the model
                            model.fit(X_train, y_train)
                            
                            # Test performance - this is more reliable with small datasets
                            test_score = model.score(X_test, y_test)
                            results[name] = {'test_score': test_score}
                            print(f"{name} model test R² score: {test_score:.4f}")
                            
                            # Track the best model
                            if test_score > best_score:
                                best_score = test_score
                                best_model_name = name
                                best_model = model
                                
                            # Only try cross-validation if we have enough data
                            if len(X_train) >= 10:
                                cv = min(5, len(X_train))  # Don't use more folds than samples
                                scores = cross_val_score(model, X_train, y_train, cv=cv)
                                results[name]['mean_cv_score'] = np.mean(scores)
                                results[name]['std_cv_score'] = np.std(scores)
                                print(f"{name} model CV score: {np.mean(scores):.4f} ± {np.std(scores):.4f}")
                            
                        except Exception as e:
                            print(f"Error training {name} model: {str(e)}")
                            results[name] = {'error': str(e)}

                    # Use the best model if we found one
                    if best_model is not None:
                        print(f"Best model: {best_model_name} with test score {best_score:.4f}")
                        return best_model, results
                    else:
                        # Fallback to a simple model
                        print("WARNING: No models were successfully trained. Using fallback model.")
                        fallback = LinearRegression()
                        fallback.fit(X, y)  # Use all data
                        return fallback, {"fallback": True}
                
                # Step 1: Clean up the data - drop any unnamed or empty columns
                print(f"Original data columns: {data.columns.tolist()}")
                data_cleaned = data.drop(columns=[col for col in data.columns if 'Unnamed:' in col], errors='ignore')
                print(f"Columns after dropping unnamed: {data_cleaned.columns.tolist()}")
                
                # Step 2: Ensure correct data types for numeric columns
                print("Converting columns to numeric types...")
                for col in ['terpene_pct', 'temperature', 'viscosity']:
                    # Print original data type and some sample values
                    print(f"Column {col} before conversion - dtype: {data_cleaned[col].dtype}")
                    print(f"Sample values before conversion: {data_cleaned[col].head().tolist()}")
                    
                    # Convert to numeric, using coerce to handle non-numeric values
                    data_cleaned[col] = pd.to_numeric(data_cleaned[col], errors='coerce')
                    
                    # Print new data type and NaN counts
                    print(f"Column {col} after conversion - dtype: {data_cleaned[col].dtype}")
                    print(f"NaN values after conversion: {data_cleaned[col].isna().sum()} ({data_cleaned[col].isna().mean()*100:.1f}%)")
                
                # Train a model for each unique media-terpene combination
                media_terpene_combos = data_cleaned[['media', 'terpene']].drop_duplicates()
                models_dict = {}
                
                print(f"Found {len(media_terpene_combos)} unique media/terpene combinations to process")
            
                for idx, row in media_terpene_combos.iterrows():
                    media = row['media']
                    terpene = row['terpene']
                    
                    # Filter data for this combination
                    combo_data = data_cleaned[(data_cleaned['media'] == media) & (data_cleaned['terpene'] == terpene)]
                    print(f"\nProcessing {media}/{terpene} combination: {len(combo_data)} samples")
                    
                    # Skip if we don't have enough data to start with
                    if len(combo_data) < 5:  # Need at least 5 samples for minimal training
                        print(f"Skipping {media}/{terpene}: Not enough initial samples ({len(combo_data)})")
                        continue
                    
                    # Clean data - handle missing values, but check if we have any data left
                    combo_data_clean = combo_data.dropna(subset=['terpene_pct', 'temperature', 'viscosity'])
                    
                    if len(combo_data_clean) < 3:  # Too few samples after cleaning
                        print(f"Skipping {media}/{terpene}: Not enough clean samples ({len(combo_data_clean)})")
                        continue
                    
                    # Print a summary of the data
                    print(f"  Clean data shape: {combo_data_clean.shape}")
                    print(f"  Terpene % range: {combo_data_clean['terpene_pct'].min()}-{combo_data_clean['terpene_pct'].max()}")
                    print(f"  Temperature range: {combo_data_clean['temperature'].min()}-{combo_data_clean['temperature'].max()}")
                    print(f"  Viscosity range: {combo_data_clean['viscosity'].min()}-{combo_data_clean['viscosity'].max()}")
                    
                    # Train the model using clean data
                    try:
                        model, results = build_viscosity_models(combo_data_clean)
                        
                        # If we got a valid model
                        if model is not None:
                            # Store the model
                            model_key = f"{media}_{terpene}"
                            models_dict[model_key] = model
                            print(f"Successfully trained model for {media}/{terpene}")
                    except Exception as e:
                        print(f"Error training model for {media}/{terpene}: {str(e)}")
                        traceback.print_exc()
            
                # Save the trained models
                if models_dict:
                    # Import pickle only when needed
                    import pickle
                    import os
                    os.makedirs('models', exist_ok=True)
                    with open('models/viscosity_models.pkl', 'wb') as f:
                        pickle.dump(models_dict, f)
                    self.viscosity_models = models_dict  # Update current models
                
                    # Update UI in the main thread
                    self.root.after(0, lambda: messagebox.showinfo("Success", 
                                                            f"Successfully trained {len(models_dict)} models."))
                else:
                    self.root.after(0, lambda: messagebox.showwarning("Warning", 
                                                                "No models could be trained. Check data quality."))
                
            except Exception as e:
                error_message = str(e)
                print(f"Error in training thread: {error_message}")
                traceback.print_exc()
                self.root.after(0, lambda msg=error_message: messagebox.showerror("Error", 
                                                            f"Error training models: {msg}"))
            finally:
                # Close the progress window
                print("Training thread completed")
                self.root.after(0, progress_window.destroy)

        # Start the training thread
        training_thread = threading.Thread(target=train_thread)
        training_thread.daemon = True
        training_thread.start()

    def analyze_models(self, show_dialog=True):
        """Analyze trained viscosity models to identify potential issues."""
        # Import modules only when needed
        import pandas as pd
        import matplotlib.pyplot as plt
        import numpy as np
        from tkinter import Toplevel, Text, Scrollbar, Label, Frame, Button
        import os
        import pickle
    
        # Check if standard models exist
        standard_models_exist = hasattr(self, 'viscosity_models') and self.viscosity_models
    
        # Check if enhanced models exist
        enhanced_models_exist = False
        enhanced_models = {}
        try:
            if os.path.exists('models/viscosity_models_with_potency.pkl'):
                with open('models/viscosity_models_with_potency.pkl', 'rb') as f:
                    enhanced_models = pickle.load(f)
                    enhanced_models_exist = bool(enhanced_models)
                    if not hasattr(self, 'enhanced_viscosity_models'):
                        self.enhanced_viscosity_models = enhanced_models
        except Exception as e:
            print(f"Error loading enhanced models: {e}")
    
        if not standard_models_exist and not enhanced_models_exist:
            messagebox.showinfo("No Models", "No trained models found. Please train models first.")
            return
    
        # Create a detailed analysis report
        report = []
        report.append("Model Analysis Report")
        report.append("==================")
    
        # Load validation data if available
        validation_data = None
        validation_files = []
    
        try:
            data_dir = 'data'
            if os.path.exists(data_dir):
                validation_files = [f for f in os.listdir(data_dir) if f.startswith('viscosity_data_') or f.startswith('Master_Viscosity_Data_') and f.endswith('.csv')]
        except Exception as e:
            report.append(f"Error checking for validation data: {str(e)}")
    
        if validation_files:
            try:
                # Use the most recent data file for validation
                validation_files.sort(reverse=True)
                latest_file = os.path.join('data', validation_files[0])
                validation_data = pd.read_csv(latest_file)
                report.append(f"Using {os.path.basename(latest_file)} for validation")
            
                # Check if validation data has potency information
                has_potency_validation = 'total_potency' in validation_data.columns
                if has_potency_validation:
                    report.append("Validation data includes potency information")
                else:
                    report.append("Validation data does not include potency information")
            except Exception as e:
                report.append(f"Error loading validation data: {str(e)}")
    
        # Analyze standard models
        if standard_models_exist:
            report.append(f"\nStandard Models (without potency): {len(self.viscosity_models)}")
            report.append("-" * 40)
        
            for model_key, model in self.viscosity_models.items():
                report.append(f"\nModel: {model_key}")
            
                # Get model type
                model_type = type(model).__name__
                if hasattr(model, 'steps'):
                    # Handle pipeline models
                    for name, step in model.steps:
                        if 'classifier' in name or 'regressor' in name:
                            model_type = type(step).__name__
                            break
            
                report.append(f"Model type: {model_type}")
            
                # Extract and show model parameters for different model types
                if hasattr(model, 'feature_importances_'):
                    # For Random Forest and tree-based models
                    importances = model.feature_importances_
                    report.append(f"Feature importances: terpene_pct={importances[0]:.4f}, temperature={importances[1]:.4f}")
                
                    if hasattr(model, 'n_estimators'):
                        report.append(f"Number of trees: {model.n_estimators}")
                        report.append(f"Max depth: {model.max_depth if model.max_depth else 'None (unlimited)'}")
            
                elif hasattr(model, 'coef_'):
                    # For linear models
                    coefs = model.coef_
                    intercept = model.intercept_
                    report.append(f"Coefficients: {coefs}")
                    report.append(f"Intercept: {intercept}")
            
                # Check for possible overfitting with Random Forest
                if 'RandomForest' in model_type:
                    report.append("\nPotential overfitting checking:")
                    report.append("Random Forest models often achieve perfect R² scores (1.0) when they overfit.")
                    report.append("For small datasets, this is especially common and means the model may")
                    report.append("memorize the training data rather than learn generalizable patterns.")
            
                # Validate model if validation data is available
                if validation_data is not None:
                    try:
                        # Filter validation data for this model's media/terpene combination
                        media, terpene = model_key.split('_', 1)
                        model_validation_data = validation_data[
                            (validation_data['media'] == media) & 
                            (validation_data['terpene'] == terpene)
                        ]
                    
                        # If we have validation data for this model
                        if len(model_validation_data) >= 5:
                            # Clean data
                            model_validation_data = model_validation_data.dropna(
                                subset=['terpene_pct', 'temperature', 'viscosity']
                            )
                        
                            if len(model_validation_data) >= 3:
                                # Extract features and target
                                X_val = model_validation_data[['terpene_pct', 'temperature']]
                                y_val = model_validation_data['viscosity']
                            
                                # Get predictions
                                y_pred = model.predict(X_val)
                            
                                # Calculate metrics
                                from sklearn.metrics import mean_squared_error, r2_score
                                mse = mean_squared_error(y_val, y_pred)
                                r2 = r2_score(y_val, y_pred)
                            
                                report.append(f"\nValidation results:")
                                report.append(f"Validation samples: {len(model_validation_data)}")
                                report.append(f"MSE: {mse:.2f}")
                                report.append(f"R²: {r2:.4f}")
                            
                                if r2 < 0.5:
                                    report.append("WARNING: Poor validation performance (R² < 0.5)!")
                    except Exception as e:
                        report.append(f"Error validating model: {str(e)}")
    
        # Analyze enhanced models with potency
        if enhanced_models_exist:
            report.append(f"\nEnhanced Models (with potency): {len(enhanced_models)}")
            report.append("-" * 40)
        
            for model_key, model in enhanced_models.items():
                report.append(f"\nModel: {model_key}")
            
                # Get model type
                model_type = type(model).__name__
                if hasattr(model, 'steps'):
                    # Handle pipeline models
                    for name, step in model.steps:
                        if 'classifier' in name or 'regressor' in name:
                            model_type = type(step).__name__
                            break
            
                report.append(f"Model type: {model_type}")
            
                # Extract and show model parameters
                if hasattr(model, 'feature_importances_'):
                    importances = model.feature_importances_
                    if len(importances) == 3:
                        report.append(f"Feature importances:")
                        report.append(f"  - terpene_pct: {importances[0]:.4f}")
                        report.append(f"  - temperature: {importances[1]:.4f}")
                        report.append(f"  - total_potency: {importances[2]:.4f}")
                    
                        # Analyze relative importance of potency
                        potency_importance = importances[2]
                        terpene_importance = importances[0]
                        if potency_importance > terpene_importance:
                            report.append(f"NOTE: Potency has higher importance than terpene percentage!")
                    
                        # Check for potential feature dominance
                        max_importance = max(importances)
                        if max_importance > 0.7:
                            dominant_feature = ['terpene_pct', 'temperature', 'total_potency'][np.argmax(importances)]
                            report.append(f"WARNING: Feature '{dominant_feature}' dominates with {max_importance:.4f} importance")
                    else:
                        report.append(f"Feature importances: {importances}")
                
                    if hasattr(model, 'n_estimators'):
                        report.append(f"Number of trees: {model.n_estimators}")
                        report.append(f"Max depth: {model.max_depth if model.max_depth else 'None (unlimited)'}")
            
                elif hasattr(model, 'coef_'):
                    # For linear models
                    coefs = model.coef_
                    intercept = model.intercept_
                    if len(coefs) == 3:
                        report.append(f"Coefficients:")
                        report.append(f"  - terpene_pct: {coefs[0]:.4f}")
                        report.append(f"  - temperature: {coefs[1]:.4f}")
                        report.append(f"  - total_potency: {coefs[2]:.4f}")
                        report.append(f"Intercept: {intercept}")
                    else:
                        report.append(f"Coefficients: {coefs}")
                        report.append(f"Intercept: {intercept}")
            
                # Validate enhanced model if validation data includes potency
                if validation_data is not None and 'total_potency' in validation_data.columns:
                    try:
                        # Extract media/terpene from model key (handle the "_with_potency" suffix)
                        if "_with_potency" in model_key:
                            base_key = model_key.replace("_with_potency", "")
                        else:
                            base_key = model_key
                        
                        media, terpene = base_key.split('_', 1)
                    
                        model_validation_data = validation_data[
                            (validation_data['media'] == media) & 
                            (validation_data['terpene'] == terpene)
                        ]
                    
                        # If we have validation data for this model
                        if len(model_validation_data) >= 5:
                            # Clean data
                            model_validation_data = model_validation_data.dropna(
                                subset=['terpene_pct', 'temperature', 'total_potency', 'viscosity']
                            )
                        
                            if len(model_validation_data) >= 3:
                                # Extract features and target
                                X_val = model_validation_data[['terpene_pct', 'temperature', 'total_potency']]
                                y_val = model_validation_data['viscosity']
                            
                                # Get predictions
                                y_pred = model.predict(X_val)
                            
                                # Calculate metrics
                                from sklearn.metrics import mean_squared_error, r2_score
                                mse = mean_squared_error(y_val, y_pred)
                                r2 = r2_score(y_val, y_pred)
                            
                                report.append(f"\nValidation results:")
                                report.append(f"Validation samples: {len(model_validation_data)}")
                                report.append(f"MSE: {mse:.2f}")
                                report.append(f"R²: {r2:.4f}")
                            
                                if r2 < 0.5:
                                    report.append("WARNING: Poor validation performance (R² < 0.5)!")
                    except Exception as e:
                        report.append(f"Error validating enhanced model: {str(e)}")
    
        # Compare standard vs enhanced models if both exist
        if standard_models_exist and enhanced_models_exist:
            report.append("\nComparison of Standard vs Enhanced Models:")
            report.append("-" * 42)
        
            # Find pairs of models for the same media/terpene combo
            common_keys = []
            for enh_key in enhanced_models.keys():
                if "_with_potency" in enh_key:
                    std_key = enh_key.replace("_with_potency", "")
                    if std_key in self.viscosity_models:
                        common_keys.append((std_key, enh_key))
        
            if common_keys:
                report.append(f"Found {len(common_keys)} pairs of models for comparison")
            
                if validation_data is not None and 'total_potency' in validation_data.columns:
                    report.append("Comparing performance on validation data:")
                
                    for std_key, enh_key in common_keys:
                        report.append(f"\nComparing {std_key} vs {enh_key}:")
                    
                        try:
                            # Get the models
                            std_model = self.viscosity_models[std_key]
                            enh_model = enhanced_models[enh_key]
                        
                            # Extract media/terpene
                            media, terpene = std_key.split('_', 1)
                        
                            # Filter validation data for this combination
                            model_validation_data = validation_data[
                                (validation_data['media'] == media) & 
                                (validation_data['terpene'] == terpene)
                            ]
                        
                            if len(model_validation_data) >= 5:
                                # Clean data - only keep rows with all required columns
                                model_validation_data = model_validation_data.dropna(
                                    subset=['terpene_pct', 'temperature', 'total_potency', 'viscosity']
                                )
                            
                                if len(model_validation_data) >= 3:
                                    # Prepare features and target
                                    X_std = model_validation_data[['terpene_pct', 'temperature']]
                                    X_enh = model_validation_data[['terpene_pct', 'temperature', 'total_potency']]
                                    y_true = model_validation_data['viscosity']
                                
                                    # Get predictions
                                    y_pred_std = std_model.predict(X_std)
                                    y_pred_enh = enh_model.predict(X_enh)
                                
                                    # Calculate metrics
                                    from sklearn.metrics import mean_squared_error, r2_score
                                    mse_std = mean_squared_error(y_true, y_pred_std)
                                    r2_std = r2_score(y_true, y_pred_std)
                                
                                    mse_enh = mean_squared_error(y_true, y_pred_enh)
                                    r2_enh = r2_score(y_true, y_pred_enh)
                                
                                    report.append(f"Standard model: MSE = {mse_std:.2f}, R² = {r2_std:.4f}")
                                    report.append(f"Enhanced model: MSE = {mse_enh:.2f}, R² = {r2_enh:.4f}")
                                
                                    # Calculate improvement
                                    mse_improvement = ((mse_std - mse_enh) / mse_std) * 100
                                    r2_improvement = ((r2_enh - r2_std) / max(0.001, abs(r2_std))) * 100
                                
                                    report.append(f"MSE improvement: {mse_improvement:.1f}%")
                                    report.append(f"R² improvement: {r2_improvement:.1f}%")
                                
                                    if mse_enh < mse_std:
                                        report.append("The enhanced model with potency performs BETTER")
                                    else:
                                        report.append("The enhanced model with potency performs WORSE")
                        except Exception as e:
                            report.append(f"Error comparing models: {str(e)}")
            else:
                report.append("No matching pairs of standard and enhanced models found")
    
        # Add recommendations
        report.append("\nGeneral Recommendations:")
        report.append("------------------------")
        report.append("1. Collect more training data to improve model generalization")
        report.append("2. Use simpler models like LinearRegression for small datasets")
        report.append("3. For RandomForest, limit complexity with max_depth, min_samples_leaf")
        report.append("4. Consider cross-validation when training to better detect overfitting")
        report.append("5. Create a separate test dataset to validate models after training")
    
        if enhanced_models_exist:
            report.append("\nRecommendations for Models with Potency:")
            report.append("--------------------------------------")
            report.append("1. Compare performance of standard vs. enhanced models")
            report.append("2. Verify if including potency improves prediction accuracy")
            report.append("3. For production use, prefer the model type with better validation scores")
            report.append("4. If potency data is inconsistent, consider using standard models")
    
        # Print to console
        print("\n".join(report))
    
        # Show in dialog if requested
        if show_dialog:
            report_window = Toplevel(self.root)
            report_window.title("Model Analysis Report")
            report_window.geometry("700x500")
        
            Label(report_window, text="Model Analysis Report", font=("Arial", 14, "bold")).pack(pady=10)
        
            text_frame = Frame(report_window)
            text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
            scrollbar = Scrollbar(text_frame)
            scrollbar.pack(side="right", fill="y")
        
            text_widget = Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
            text_widget.pack(side="left", fill="both", expand=True)
        
            scrollbar.config(command=text_widget.yview)
        
            text_widget.insert("1.0", "\n".join(report))
            text_widget.config(state="disabled")
        
            Button(report_window, text="Close", command=report_window.destroy).pack(pady=10)
    
        return report

    def filter_and_analyze_specific_combinations(self):
        """
        Analyze and build models for specific media-terpene combinations.
        This function performs Arrhenius analysis to determine the relationship
        between temperature and viscosity for different combinations.
        """
        # Display progress window immediately to show the user something is happening
        progress_window = Toplevel(self.root)
        progress_window.title("Analyzing Specific Combinations")
        progress_window.geometry("700x500")  # Larger window for more text
        progress_window.transient(self.root)
        progress_window.grab_set()
    
        # Import modules - lazily loaded only when this function is called
        import threading
    
        # Create a background thread to do the heavy lifting
        def bg_thread():
            try:
                # Import all required modules inside the thread to avoid blocking UI
                import pandas as pd
                import numpy as np
                import matplotlib.pyplot as plt
                from tkinter import Text, Scrollbar, Label, Frame, Canvas, Entry, StringVar, DoubleVar
                import os
                import glob
                from scipy import stats
                from sklearn.linear_model import LinearRegression
                from scipy.optimize import curve_fit
                import pickle
                import math
                from sklearn.metrics import r2_score
            
                # Create scrollable text area for live updates
                Label(progress_window, text="Arrhenius Analysis of Temperature-Viscosity Relationship", 
                      font=("Arial", 14, "bold")).pack(pady=10)
            
                # Add a frame for potency specification
                potency_frame = Frame(progress_window)
                potency_frame.pack(fill="x", padx=10, pady=5)
            
                potency_var = DoubleVar(value=80.0)  # Default value of 80%
            
                Label(potency_frame, text="Total Potency (%) for analysis:").pack(side="left", padx=5)
                potency_entry = Entry(potency_frame, textvariable=potency_var, width=10)
                potency_entry.pack(side="left", padx=5)
            
                text_frame = Frame(progress_window)
                text_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
                scrollbar = Scrollbar(text_frame)
                scrollbar.pack(side="right", fill="y")
            
                text_widget = Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
                text_widget.pack(side="left", fill="both", expand=True)
                scrollbar.config(command=text_widget.yview)
            
                # Function to add text to the widget from any thread
                def add_text(message):
                    self.root.after(0, lambda: text_widget.insert("end", message + "\n"))
                    self.root.after(0, lambda: text_widget.see("end"))
            
                # Add a close button
                self.root.after(0, lambda: Button(progress_window, text="Close", 
                                            command=progress_window.destroy).pack(pady=10))
            
                add_text("Starting Arrhenius analysis of viscosity-temperature relationship...")
            
                # Determine if enhanced models exist
                using_enhanced_models = False
                if os.path.exists('models/viscosity_models_with_potency.pkl'):
                    try:
                        with open('models/viscosity_models_with_potency.pkl', 'rb') as f:
                            enhanced_models = pickle.load(f)
                            if enhanced_models:
                                using_enhanced_models = True
                                self.enhanced_viscosity_models = enhanced_models
                                add_text("Using enhanced models with potency data.")
                    except Exception as e:
                        add_text(f"Error loading enhanced models: {str(e)}")
            
                if not using_enhanced_models:
                    if hasattr(self, 'viscosity_models') and self.viscosity_models:
                        add_text("Using standard models without potency data.")
                    else:
                        add_text("No trained models found. Please train models first.")
                        return
            
                # Get the models to use
                models_to_analyze = self.enhanced_viscosity_models if using_enhanced_models else self.viscosity_models
            
                # Get a representative potency value (from entry or default)
                representative_potency = potency_var.get()
                add_text(f"Using representative potency value of {representative_potency}% for analysis")
            
                # Create plots directory if it doesn't exist
                os.makedirs('plots', exist_ok=True)
            
                # Load training data to extract typical terpene percentages
                training_data = None
                try:
                    # Get all data files
                    data_files = glob.glob('data/viscosity_data_*.csv') + glob.glob('data/Master_Viscosity_Data_*.csv')
                    if data_files:
                        # Sort by modification time (newest first)
                        data_files.sort(key=os.path.getmtime, reverse=True)
                        # Use the most recent file
                        latest_file = data_files[0]
                        training_data = pd.read_csv(latest_file)
                        add_text(f"Loaded training data from {latest_file}")
                except Exception as e:
                    add_text(f"Error loading training data: {str(e)}")
            
                # Function to predict viscosity based on model type
                def predict_viscosity(model, terpene_pct, temperature, potency=None):
                    """Predict viscosity based on model type and available features."""
                    try:
                        if using_enhanced_models and potency is not None:
                            # Enhanced model with potency
                            return model.predict([[terpene_pct, temperature, potency]])[0]
                        else:
                            # Standard model without potency
                            return model.predict([[terpene_pct, temperature]])[0]
                    except Exception as e:
                        add_text(f"Error predicting viscosity: {str(e)}")
                        return np.nan
            
                # Arrhenius function: ln(viscosity) = ln(A) + (Ea/R)*(1/T)
                def arrhenius_function(x, a, b):
                    # x is 1/T (inverse temperature in Kelvin)
                    # a is ln(A) where A is the pre-exponential factor
                    # b is Ea/R where Ea is activation energy and R is gas constant
                    return a + b * x
            
                # Extend the temperature range for prediction
                temperature_range = np.linspace(20, 70, 11)  # 20°C to 70°C in 11 steps
            
                # Counter for successful models
                successful_models = 0
            
                # Process each model
                for model_key, model in models_to_analyze.items():
                    try:
                        # Extract media and terpene from the model key
                        if using_enhanced_models and "_with_potency" in model_key:
                            # Remove the "_with_potency" suffix for display purposes
                            display_key = model_key.replace("_with_potency", "")
                            media, terpene = display_key.split('_', 1)
                        else:
                            media, terpene = model_key.split('_', 1)
                    
                        add_text(f"\nAnalyzing {media}/{terpene}...")
                    
                        # Find typical terpene percentage for this combination from training data
                        terpene_pct = 5.0  # Default value
                        if training_data is not None:
                            combo_data = training_data[
                                (training_data['media'] == media) & 
                                (training_data['terpene'] == terpene)
                            ]
                            if 'terpene_pct' in combo_data.columns and not combo_data.empty:
                                # Use median as a robust measure
                                terpene_pct = combo_data['terpene_pct'].median()
                                if pd.isna(terpene_pct) or terpene_pct <= 0:
                                    terpene_pct = 5.0  # Default if invalid
                    
                        add_text(f"Using terpene percentage of {terpene_pct:.2f}% for analysis")
                    
                        # Generate viscosity predictions across temperature range
                        temperatures_kelvin = temperature_range + 273.15  # Convert °C to K
                        inverse_temp = 1 / temperatures_kelvin  # 1/T for Arrhenius plot
                    
                        predicted_visc = []
                        for temp in temperature_range:
                            if using_enhanced_models:
                                visc = predict_viscosity(model, terpene_pct, temp, representative_potency)
                            else:
                                visc = predict_viscosity(model, terpene_pct, temp)
                            predicted_visc.append(visc)
                    
                        predicted_visc = np.array(predicted_visc)
                    
                        # Filter out invalid values
                        valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                        if not any(valid_indices):
                            add_text(f"No valid viscosity predictions for {media}/{terpene}. Skipping.")
                            continue
                    
                        inv_temp_valid = inverse_temp[valid_indices]
                        predicted_visc_valid = predicted_visc[valid_indices]
                    
                        # Calculate natural log of viscosity
                        ln_visc = np.log(predicted_visc_valid)
                    
                        # Fit Arrhenius equation
                        params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                        a, b = params
                    
                        # Calculate activation energy (Ea = b * R)
                        R = 8.314  # Gas constant in J/(mol·K)
                        Ea = b * R  # Activation energy in J/mol
                        Ea_kJ = Ea / 1000  # Convert to kJ/mol
                    
                        # Calculate pre-exponential factor
                        A = np.exp(a)
                    
                        # Calculate predicted values from the fitted model
                        ln_visc_pred = arrhenius_function(inv_temp_valid, a, b)
                    
                        # Calculate R-squared
                        r2 = r2_score(ln_visc, ln_visc_pred)
                    
                        # Generate plot
                        plt.figure(figsize=(10, 8))
                    
                        # Create two subplots
                        plt.subplot(211)
                        plt.scatter(temperature_range[valid_indices], predicted_visc_valid, color='blue', label='Predicted viscosity')
                        plt.yscale('log')
                        plt.xlabel('Temperature (°C)')
                        plt.ylabel('Viscosity (cP)')
                        plt.title(f'Viscosity vs Temperature for {media}/{terpene}\nTerpene: {terpene_pct:.2f}%')
                        plt.grid(True)
                    
                        # Arrhenius plot
                        plt.subplot(212)
                        plt.scatter(inv_temp_valid, ln_visc, color='blue', label='ln(Viscosity)')
                        plt.plot(inv_temp_valid, ln_visc_pred, 'r-', label=f'Arrhenius fit (R² = {r2:.4f})')
                        plt.xlabel('1/T (K⁻¹)')
                        plt.ylabel('ln(Viscosity)')
                        plt.title(f'Arrhenius Plot: Ea = {Ea_kJ:.2f} kJ/mol, ln(A) = {a:.2f}')
                        plt.grid(True)
                        plt.legend()
                    
                        plt.tight_layout()
                    
                        # Save the plot
                        plot_path = f'plots/Arrhenius_{media}_{terpene}.png'
                        plt.savefig(plot_path)
                        plt.close()
                    
                        # Update report
                        add_text(f"Results for {media}/{terpene}:")
                        add_text(f"  • Activation energy (Ea): {Ea_kJ:.2f} kJ/mol")
                        add_text(f"  • Pre-exponential factor ln(A): {a:.2f}")
                        add_text(f"  • Arrhenius equation: ln(viscosity) = {a:.2f} + {b:.2f}*(1/T)")
                        add_text(f"  • R-squared: {r2:.4f}")
                        add_text(f"  • Plot saved to: {plot_path}")
                    
                        # Categorize the activation energy
                        if Ea_kJ < 20:
                            add_text("  • Low activation energy: less temperature-sensitive")
                        elif Ea_kJ < 40:
                            add_text("  • Medium activation energy: moderately temperature-sensitive")
                        else:
                            add_text("  • High activation energy: highly temperature-sensitive")
                    
                        successful_models += 1
                    
                    except Exception as e:
                        add_text(f"Error analyzing {model_key}: {str(e)}")
            
                # Summary
                add_text(f"\nAnalysis complete! Successfully analyzed {successful_models} models.")
                add_text(f"Plot files are saved in the 'plots' directory.")
            
                if successful_models > 0:
                    # Generate summary plot comparing activation energies
                    self.generate_activation_energy_comparison(models_to_analyze, representative_potency, using_enhanced_models)
                    add_text("Generated comparison plot of activation energies.")
            
            except Exception as e:
                import traceback
                traceback_str = traceback.format_exc()
                add_text(f"Error in analysis thread: {e}\n{traceback_str}")
                self.root.after(0, lambda: messagebox.showerror("Error", f"Analysis failed: {str(e)}"))
    
        # Start background thread
        thread = threading.Thread(target=bg_thread)
        thread.daemon = True
        thread.start()

    def generate_activation_energy_comparison(self, models, potency_value=None, using_enhanced_models=False):
        """
        Generate a comparison plot of activation energies for different media-terpene combinations.
        """
        import numpy as np
        import matplotlib.pyplot as plt
        from scipy.optimize import curve_fit
        import os
    
        # Create a figure
        plt.figure(figsize=(12, 8))
    
        # Arrhenius function: ln(viscosity) = ln(A) + (Ea/R)*(1/T)
        def arrhenius_function(x, a, b):
            return a + b * x
    
        # Temperature range
        temperature_range = np.linspace(20, 70, 11)  # 20°C to 70°C in 11 steps
        temperatures_kelvin = temperature_range + 273.15  # Convert °C to K
        inverse_temp = 1 / temperatures_kelvin  # 1/T for Arrhenius plot
    
        # R constant
        R = 8.314  # Gas constant in J/(mol·K)
    
        # Store results for comparison
        media_types = set()
        results = []
    
        # Process each model
        for model_key, model in models.items():
            try:
                # Extract media and terpene from the model key
                if using_enhanced_models and "_with_potency" in model_key:
                    # Remove the "_with_potency" suffix
                    display_key = model_key.replace("_with_potency", "")
                    media, terpene = display_key.split('_', 1)
                else:
                    media, terpene = model_key.split('_', 1)
            
                media_types.add(media)
            
                # Default terpene percentage
                terpene_pct = 5.0
            
                # Generate viscosity predictions
                predicted_visc = []
                for temp in temperature_range:
                    if using_enhanced_models and potency_value is not None:
                        visc = model.predict([[terpene_pct, temp, potency_value]])[0]
                    else:
                        visc = model.predict([[terpene_pct, temp]])[0]
                    predicted_visc.append(visc)
            
                predicted_visc = np.array(predicted_visc)
            
                # Filter out invalid values
                valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                if not any(valid_indices):
                    continue
            
                inv_temp_valid = inverse_temp[valid_indices]
                predicted_visc_valid = predicted_visc[valid_indices]
            
                # Calculate natural log of viscosity
                ln_visc = np.log(predicted_visc_valid)
            
                # Fit Arrhenius equation
                params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                a, b = params
            
                # Calculate activation energy
                Ea = b * R  # Activation energy in J/mol
                Ea_kJ = Ea / 1000  # Convert to kJ/mol
            
                # Store result
                results.append({
                    'media': media,
                    'terpene': terpene,
                    'Ea_kJ': Ea_kJ,
                    'ln_A': a
                })
            
            except Exception as e:
                print(f"Error processing {model_key}: {e}")
    
        if not results:
            return
    
        # Convert to DataFrame
        import pandas as pd
        results_df = pd.DataFrame(results)
    
        # Create bar chart grouped by media type
        media_list = list(media_types)
        colors = plt.cm.tab10(np.linspace(0, 1, len(media_list)))
        color_map = dict(zip(media_list, colors))
    
        # Sort by activation energy
        results_df = results_df.sort_values('Ea_kJ', ascending=False)
    
        # Plot bar chart
        ax = plt.subplot(111)
    
        # Create positions for bars
        positions = np.arange(len(results_df))
        bar_height = 0.8
    
        # Create bars with colors based on media type
        bars = ax.barh(
            positions, 
            results_df['Ea_kJ'], 
            height=bar_height,
            color=[color_map[media] for media in results_df['media']]
        )
    
        # Add labels
        ax.set_yticks(positions)
        ax.set_yticklabels([f"{row['terpene']} ({row['media']})" for _, row in results_df.iterrows()])
        ax.set_xlabel('Activation Energy (kJ/mol)')
        ax.set_title('Activation Energy Comparison by Media-Terpene Combination')
    
        # Add a legend for media types
        from matplotlib.patches import Patch
        legend_elements = [Patch(facecolor=color_map[media], label=media) for media in media_list]
        ax.legend(handles=legend_elements, loc='upper right')
    
        # Add value labels
        for i, bar in enumerate(bars):
            ax.text(
                bar.get_width() + 0.5, 
                bar.get_y() + bar.get_height()/2, 
                f"{results_df['Ea_kJ'].iloc[i]:.1f}", 
                va='center'
            )
    
        plt.tight_layout()
    
        # Save the plot
        plot_path = 'plots/Activation_Energy_Comparison.png'
        plt.savefig(plot_path)
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

    # Add to viscosity_calculator.py

    def train_models_with_potency(self, data=None):
        """
        Train viscosity prediction models including total potency as a feature.
    
        Args:
            data (pd.DataFrame, optional): DataFrame with training data. If None,
                                         attempts to use data from saved CSV files.
        """
        # Import required libraries for model training
        import glob
        import pandas as pd
        import numpy as np
        from sklearn.model_selection import train_test_split
        from sklearn.linear_model import LinearRegression
        from sklearn.preprocessing import PolynomialFeatures
        from sklearn.ensemble import RandomForestRegressor
        from sklearn.pipeline import make_pipeline
        import pickle
        import os
    
        # If no data provided, load from saved files
        if data is None:
            data_files = glob.glob('data/viscosity_data_*.csv')
            if not data_files:
                messagebox.showerror("Error", "No training data available. Please upload data first.")
                return
        
            # Load and combine all data files
            data_frames = []
            for file in data_files:
                try:
                    df = pd.read_csv(file)
                    data_frames.append(df)
                except Exception as e:
                    print(f"Error loading {file}: {str(e)}")
                    continue
        
            if not data_frames:
                messagebox.showerror("Error", "Failed to load any training data.")
                return
        
            data = pd.concat(data_frames, ignore_index=True)
    
        # Show a progress dialog
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Training Models with Potency")
        progress_window.geometry("400x200")
        progress_window.transient(self.root)
        progress_window.grab_set()
    
        progress_label = tk.Label(
            progress_window, 
            text="Training enhanced models with potency data...",
            font=('Arial', 12)
        )
        progress_label.pack(pady=20)
    
        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill='x', padx=20)
        progress_bar.start()
    
        # Define function to update the label
        def update_label(text):
            progress_label.config(text=text)
            progress_window.update_idletasks()
    
        # Start training in a new thread
        def train_thread():
            try:
                update_label("Preprocessing data...")
            
                # Check if total_potency column exists
                has_potency = 'total_potency' in data.columns
            
                # Group data by media and terpene type
                media_terpene_combos = data[['media', 'terpene']].drop_duplicates()
            
                # Models dictionary to store results
                models_dict = {}
                enhanced_models_dict = {}  # For models with potency
            
                update_label("Training models for each media/terpene combination...")
            
                # Process each combination
                for idx, row in media_terpene_combos.iterrows():
                    media = row['media']
                    terpene = row['terpene']
                
                    # Filter data for this combination
                    combo_data = data[(data['media'] == media) & (data['terpene'] == terpene)]
                
                    # Skip if not enough data
                    if len(combo_data) < 5:
                        continue
                
                    # Basic model (just terpene_pct and temperature)
                    if 'terpene_pct' in combo_data.columns and 'temperature' in combo_data.columns:
                        # Standard model features
                        basic_features = combo_data[['terpene_pct', 'temperature']].dropna()
                        if len(basic_features) >= 5:
                            # Get corresponding viscosity values
                            basic_target = combo_data.loc[basic_features.index, 'viscosity']
                        
                            # Train a model
                            model = RandomForestRegressor(
                                n_estimators=50,
                                max_depth=5,
                                min_samples_leaf=3,
                                random_state=42
                            )
                            model.fit(basic_features, basic_target)
                        
                            # Store the model
                            model_key = f"{media}_{terpene}"
                            models_dict[model_key] = model
                
                    # Enhanced model with potency if available
                    if has_potency and 'terpene_pct' in combo_data.columns and 'temperature' in combo_data.columns:
                        # Enhanced features including potency
                        enhanced_features = combo_data[['terpene_pct', 'temperature', 'total_potency']].dropna()
                        if len(enhanced_features) >= 5:
                            # Get corresponding viscosity values
                            enhanced_target = combo_data.loc[enhanced_features.index, 'viscosity']
                        
                            # Train an enhanced model
                            enhanced_model = RandomForestRegressor(
                                n_estimators=50,
                                max_depth=5,
                                min_samples_leaf=3,
                                random_state=42
                            )
                            enhanced_model.fit(enhanced_features, enhanced_target)
                        
                            # Store the enhanced model
                            model_key = f"{media}_{terpene}_with_potency"
                            enhanced_models_dict[model_key] = enhanced_model
            
                update_label("Saving models...")
            
                # Save standard models
                if models_dict:
                    os.makedirs('models', exist_ok=True)
                    with open('models/viscosity_models.pkl', 'wb') as f:
                        pickle.dump(models_dict, f)
                    self.viscosity_models = models_dict
            
                # Save enhanced models with potency
                if enhanced_models_dict:
                    os.makedirs('models', exist_ok=True)
                    with open('models/viscosity_models_with_potency.pkl', 'wb') as f:
                        pickle.dump(enhanced_models_dict, f)
                    self.enhanced_viscosity_models = enhanced_models_dict
            
                # Close progress window and show success message
                progress_window.after(0, progress_window.destroy)
                messagebox.showinfo(
                    "Success", 
                    f"Successfully trained {len(models_dict)} standard models and "
                    f"{len(enhanced_models_dict)} enhanced models with potency."
                )
            
            except Exception as e:
                import traceback
                traceback_str = traceback.format_exc()
                print(f"Error in training thread: {e}\n{traceback_str}")
                progress_window.after(0, progress_window.destroy)
                messagebox.showerror("Error", f"Error training models: {str(e)}")
    
        # Start the training thread
        import threading
        training_thread = threading.Thread(target=train_thread)
        training_thread.daemon = True
        training_thread.start()

    def calculate_viscosity_with_potency(self):
        """
        Calculate terpene percentage based on target viscosity using models that include potency data.
        """
        try:
            # Extract input values
            media = self.media_var.get()
            media_brand = self.media_brand_var.get()
            terpene = self.terpene_var.get()
            terpene_brand = self.terpene_brand_var.get()
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
        
            # Add new UI element to get total potency
            total_potency = None
            if hasattr(self, 'total_potency_var'):
                try:
                    total_potency = float(self.total_potency_var.get())
                except (ValueError, tk.TclError):
                    pass
        
            # First try to load enhanced models
            enhanced_models = {}
            try:
                with open('models/viscosity_models_with_potency.pkl', 'rb') as f:
                    enhanced_models = pickle.load(f)
            except:
                enhanced_models = {}
        
            # Load standard models as fallback
            model_key = f"{media}_{terpene}"
            enhanced_key = f"{model_key}_with_potency"
        
            if enhanced_key in enhanced_models and total_potency is not None:
                # Use enhanced model with potency
                from scipy.optimize import fsolve
            
                model = enhanced_models[enhanced_key]
        
                # For numerical solution, find terpene percentage that gives target viscosity
                def objective(terpene_pct):
                    # Predict viscosity at 25°C with given terpene & potency percentage
                    predicted_viscosity = model.predict([[terpene_pct, 25.0, total_potency]])[0]
                    return predicted_viscosity - target_viscosity
        
                # Solve for terpene percentage (start with a guess of 5%)
                exact_percent = fsolve(objective, 5.0)[0]
        
                # Ensure the percentage is reasonable
                exact_percent = max(0.1, min(15.0, exact_percent))
                exact_mass = mass_of_oil * (exact_percent / 100)
        
                # Suggested starting point (slightly higher)
                start_percent = exact_percent * 1.1
                start_mass = mass_of_oil * (start_percent / 100)
        
                # Update result variables
                self.exact_percent_var.set(f"{exact_percent:.1f}%")
                self.exact_mass_var.set(f"{exact_mass:.2f}g")
                self.start_percent_var.set(f"{start_percent:.1f}%")
                self.start_mass_var.set(f"{start_mass:.2f}g")
            
                # Indicate that enhanced model was used
                messagebox.showinfo("Calculation Complete", 
                                  "Calculation performed using enhanced model with potency data.")
            
            elif model_key in self.viscosity_models:
                # Fallback to standard model without potency
                self.calculate_viscosity()  # Use the original method
            
                if total_potency is not None:
                    messagebox.showinfo("Notice", 
                                     "Enhanced model with potency not available for this combination. "
                                     "Used standard model instead.")
            else:
                # No model available, use iterative method
                messagebox.askyesno("No Model Available", 
                                  "No prediction model found for this combination.\n\n"
                                  "Would you like to use the Iterative Method instead?")

        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred: {e}")



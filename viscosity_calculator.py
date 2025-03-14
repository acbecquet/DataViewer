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
        self.d9_thc_var = DoubleVar(value=0.0)
        self.d8_thc_var = DoubleVar(value=0.0)
        
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

        # Row for d9-THC
        Label(form_frame, text="d9-THC (%):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=6, column=0, sticky="w", pady=5)

        d9_thc_entry = Entry(form_frame, textvariable=self.d9_thc_var, width=15)
        d9_thc_entry.grid(row=6, column=1, sticky="w", padx=5, pady=5)

        # Row for d8-THC  
        Label(form_frame, text="d8-THC (%):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=6, column=2, sticky="w", pady=5)

        d8_thc_entry = Entry(form_frame, textvariable=self.d8_thc_var, width=15)
        d8_thc_entry.grid(row=6, column=3, sticky="w", padx=5, pady=5)

        # Adjust the rows of later elements accordingly
        # Separator for results - update row number
        ttk.Separator(form_frame, orient='horizontal').grid(row=7, column=0, columnspan=4, sticky="ew", pady=15)

        # Results section header - update row number
        Label(form_frame, text="Results:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=(FONT[0], FONT[1], "bold"), anchor="w").grid(row=8, column=0, sticky="w", pady=5)
        
        
        # Row 7: Exact % and Exact Mass
        Label(form_frame, text="Exact %:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=9, column=0, sticky="w", pady=3)
        
        exact_percent_label = Label(form_frame, textvariable=self.exact_percent_var, 
                              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        exact_percent_label.grid(row=9, column=1, sticky="w", pady=3)
        
        Label(form_frame, text="Exact Mass:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=9, column=2, sticky="w", pady=3)
        
        exact_mass_label = Label(form_frame, textvariable=self.exact_mass_var, 
                           bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        exact_mass_label.grid(row=9, column=3, sticky="w", pady=3)
        
        # Row 8: Start % and Start Mass
        Label(form_frame, text="Start %:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=10, column=0, sticky="w", pady=3)
        
        start_percent_label = Label(form_frame, textvariable=self.start_percent_var, 
                              bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        start_percent_label.grid(row=10, column=1, sticky="w", pady=3)
        
        Label(form_frame, text="Start Mass:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=10, column=2, sticky="w", pady=3)
        
        start_mass_label = Label(form_frame, textvariable=self.start_mass_var, 
                           bg=APP_BACKGROUND_COLOR, fg="#90EE90", font=FONT)
        start_mass_label.grid(row=10, column=3, sticky="w", pady=3)
        
        # Create button frame for organized rows of buttons
        button_frame = Frame(form_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.grid(row=11, column=0, columnspan=4, pady=10)

        # Create first row of buttons
        button_row1 = Frame(button_frame, bg=APP_BACKGROUND_COLOR)
        button_row1.pack(fill="x", pady=(0, 5))  # Add some space between rows

        # Calculate button
        calculate_btn = ttk.Button(
            button_row1,
            text="Calculate",
            command=self.calculate_viscosity
        )
        calculate_btn.pack(padx=5,pady=5)

        # Add a new button for calculating with chemistry
        calculate_with_chemistry_btn = ttk.Button(
            button_row1,
            text="Calculate with Chemistry",
            command=self.calculate_viscosity_with_chemistry
        )
        calculate_with_chemistry_btn.pack(padx=5,pady=5)

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
            data_files = glob.glob('data/Master_Viscosity_Data_processed.csv')
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
        progress_window.geometry("200x100")
        progress_window.transient(self.root)
        progress_window.grab_set()

        progress_label = Label(progress_window, text="Training models...", font=('Arial', 12))
        progress_label.pack(pady=10)

        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill='x', padx=10)
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
            if os.path.exists('models/viscosity_models_with_chemistry.pkl'):
                with open('models/viscosity_models_with_chemistry.pkl', 'rb') as f:
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
                        # Extract media/terpene from model key (handle the "_with_chemistry" suffix)
                        if "_with_chemistry" in model_key:
                            base_key = model_key.replace("_with_chemistry", "")
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
                if "_with_chemistry" in enh_key:
                    std_key = enh_key.replace("_with_chemistry", "")
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
            report_window.geometry("800x1200")
        
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
        Analyze and build models for specific media-terpene combinations with potency integration.
        This function performs Arrhenius analysis to determine the relationship
        between temperature and viscosity for different combinations, while
        accounting for the effects of potency on viscosity.
        """
        # Display progress window immediately to show the user something is happening
        progress_window = Toplevel(self.root)
        progress_window.title("Analyzing Specific Combinations")
        progress_window.geometry("800x1200")  # Larger window for more text
        progress_window.transient(self.root)
        progress_window.grab_set()

        # Define a function to sanitize filenames
        def sanitize_filename(name):
            """Replace invalid filename characters with underscores."""
            # Replace common invalid filename characters
            invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ']
            sanitized = name
            for char in invalid_chars:
                sanitized = sanitized.replace(char, '_')
            return sanitized

        # Import modules - lazily loaded only when this function is called
        import threading

        # Create a background thread to do the heavy lifting
        def bg_thread():
            try:
                # Import all required modules inside the thread to avoid blocking UI
                import pandas as pd
                import numpy as np
                import matplotlib.pyplot as plt
                from tkinter import Text, Scrollbar, Label, Frame, Canvas, Entry, StringVar, DoubleVar, Scale, HORIZONTAL
                import os
                import glob
                from scipy import stats
                from sklearn.linear_model import LinearRegression
                from scipy.optimize import curve_fit
                import pickle
                import math
                from sklearn.metrics import r2_score
                from matplotlib.figure import Figure
                from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    
                # Create scrollable text area for live updates
                Label(progress_window, text="Arrhenius Analysis with Potency Integration", 
                      font=("Arial", 14, "bold")).pack(pady=10)
    
                # Add a frame for potency specification with a slider for better UX
                potency_frame = Frame(progress_window)
                potency_frame.pack(fill="x", padx=10, pady=5)
    
                potency_var = DoubleVar(value=80.0)  # Default value of 80%
    
                Label(potency_frame, text="Total Potency (%) for analysis:").pack(side="left", padx=5)
        
                # Add a slider for potency selection
                potency_slider = Scale(potency_frame, variable=potency_var, from_=60.0, to=95.0, 
                                      orient=HORIZONTAL, length=200, resolution=0.5)
                potency_slider.pack(side="left", padx=5, fill="x", expand=True)
        
                # Add a dropdown to select the potency analysis mode
                potency_mode_var = StringVar(value="fixed")
                potency_modes = ["fixed", "variable"]
        
                Label(potency_frame, text="Potency Analysis Mode:").pack(side="left", padx=5)
                potency_mode_dropdown = ttk.Combobox(
                    potency_frame, 
                    textvariable=potency_mode_var,
                    values=potency_modes,
                    state="readonly",
                    width=10
                )
                potency_mode_dropdown.pack(side="left", padx=5)
    
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
    
                add_text("Starting Arrhenius analysis of viscosity-temperature relationship with potency integration...")
    
                # Determine if enhanced models exist
                using_enhanced_models = False
                if os.path.exists('models/viscosity_models_with_chemistry.pkl'):
                    try:
                        with open('models/viscosity_models_with_chemistry.pkl', 'rb') as f:
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
    
                # Get a representative potency value (from slider)
                representative_potency = potency_var.get()
                add_text(f"Using reference potency value of {representative_potency}% for analysis")
    
                # Create plots directory if it doesn't exist
                os.makedirs('plots', exist_ok=True)
    
                # Load training data to extract typical terpene percentages and potency ranges
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
                
                        # Check if training data has potency information
                        if 'total_potency' in training_data.columns:
                            add_text(f"Training data includes potency information!")
                            add_text(f"Potency range: {training_data['total_potency'].min():.1f}% - {training_data['total_potency'].max():.1f}%")
                        else:
                            add_text("Warning: Training data does not include potency information.")
                except Exception as e:
                    add_text(f"Error loading training data: {str(e)}")
    
                # Function to predict viscosity based on model type
                def predict_viscosity(model, terpene_pct, temperature, potency=None, d9_thc=None, d8_thc=None):
                    """
                    Predict viscosity based on model type and available features.

                    This function intelligently handles models with different feature requirements
                    by examining each model's expected inputs and formatting prediction data accordingly.

                    Args:
                        model: The trained model object or dict containing model information
                        terpene_pct: Terpene percentage value
                        temperature: Temperature in degrees C
                        potency: Total potency percentage (optional)
                        d9_thc: Delta-9 THC percentage (optional)
                        d8_thc: Delta-8 THC percentage (optional)
    
                    Returns:
                        float: Predicted viscosity value
                    """
                    try:
                        # Extract the actual model if it's in a dictionary
                        if isinstance(model, dict) and 'model' in model:
                            # Get the feature list if available
                            features = model.get('features', None)
                            actual_model = model['model']
                        else:
                            actual_model = model
                            features = None
    
                        # Determine what features the model expects
                        if hasattr(actual_model, 'n_features_in_'):
                            # For scikit-learn models that have n_features_in_ attribute
                            n_features = actual_model.n_features_in_
                        elif features is not None:
                            # If features list is provided in the model dict
                            n_features = len(features)
                        else:
                            # Default to basic features if we can't determine
                            n_features = 2
    
                        # Create prediction array based on the number of features the model expects
                        if n_features == 2:
                            # Basic model with only terpene_pct and temperature
                            X = [[terpene_pct, temperature]]
                        elif n_features == 3:
                            # Model with terpene_pct, temperature, and total_potency
                            if potency is None:
                                potency = 80.0  # Default value if none provided
                            X = [[terpene_pct, temperature, potency]]
                        elif n_features == 5:
                            # Model with all features: terpene_pct, temperature, potency, d9_thc, d8_thc
                            if potency is None:
                                potency = 80.0
                            if d9_thc is None:
                                d9_thc = potency * 0.85  # Estimate d9 as 85% of total if not provided
                            if d8_thc is None:
                                d8_thc = 0.0
                            X = [[terpene_pct, temperature, potency, d9_thc, d8_thc]]
                        else:
                            # For any other feature count, try to build a sensible array
                            X = [[terpene_pct, temperature]]
                            if n_features > 2 and potency is not None:
                                X[0].append(potency)
                            if n_features > 3 and d9_thc is not None:
                                X[0].append(d9_thc)
                            if n_features > 4 and d8_thc is not None:
                                X[0].append(d8_thc)
        
                            # Pad with zeros if still not enough features
                            while len(X[0]) < n_features:
                                X[0].append(0.0)
    
                        # Make prediction with properly formatted input
                        return actual_model.predict(X)[0]
    
                    except Exception as e:
                        print(f"Error predicting viscosity: {e}")
                        import traceback
                        traceback.print_exc()
                        return float('nan')  # Return NaN for invalid predictions
    
                # Arrhenius function: ln(viscosity) = ln(A) + (Ea/R)*(1/T)
                def arrhenius_function(x, a, b):
                    # x is 1/T (inverse temperature in Kelvin)
                    # a is ln(A) where A is the pre-exponential factor
                    # b is Ea/R where Ea is activation energy and R is gas constant
                    return a + b * x
    
                # Extend the temperature range for prediction
                temperature_range = np.linspace(20, 70, 11)  # 20°C to 70°C in 11 steps
    
                # Define potency ranges for variable potency analysis
                if potency_mode_var.get() == "variable":
                    if training_data is not None and 'total_potency' in training_data.columns:
                        # Use range from actual data
                        potency_min = max(50, training_data['total_potency'].min())
                        potency_max = min(99, training_data['total_potency'].max())
                    else:
                        # Default range if no data available
                        potency_min = 50
                        potency_max = 99
            
                    potency_range = np.linspace(potency_min, potency_max, 4)  # 4 potency values for comparison
                    add_text(f"Performing variable potency analysis with levels: {', '.join([f'{p:.1f}%' for p in potency_range])}")
                else:
                    # Fixed potency mode - just use the selected value
                    potency_range = [representative_potency]
    
                # Counter for successful models
                successful_models = 0
    
                # Process each model
                for model_key, model in models_to_analyze.items():
                    try:
                        # Extract media and terpene from the model key
                        if using_enhanced_models and "_with_chemistry" in model_key:
                            # Remove the "_with_chemistry" suffix for display purposes
                            display_key = model_key.replace("_with_chemistry", "")
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
                                terpene_pct = 100 * combo_data['terpene_pct'].median() 
                                if pd.isna(terpene_pct) or terpene_pct <= 0:
                                    terpene_pct = 5.0  # Default if invalid
            
                        add_text(f"Using terpene percentage of {terpene_pct:.2f}% for analysis")
                
                        # Check if we're doing variable potency analysis
                        if potency_mode_var.get() == "variable" and using_enhanced_models:
                            # Create a figure for variable potency analysis
                            plt.figure(figsize=(12, 10))
                    
                            # Create two subplots - one for viscosity vs temp, one for ln(visc) vs 1/T
                            ax1 = plt.subplot(211)
                            ax2 = plt.subplot(212)
                    
                            # Store activation energies for different potency levels
                            activation_energies = []
                    
                            # Plot for each potency level
                            for potency in potency_range:
                                # Generate viscosity predictions across temperature range
                                temperatures_kelvin = temperature_range + 273.15  # Convert °C to K
                                inverse_temp = 1 / temperatures_kelvin  # 1/T for Arrhenius plot
                        
                                predicted_visc = []
                                for temp in temperature_range:
                                    visc = predict_viscosity(model, terpene_pct, temp, potency)
                                    predicted_visc.append(visc)
                        
                                predicted_visc = np.array(predicted_visc)
                        
                                # Filter out invalid values
                                valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                                if not any(valid_indices):
                                    add_text(f"No valid viscosity predictions for {media}/{terpene} at {potency:.1f}% potency. Skipping.")
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
                        
                                # Store for comparison
                                activation_energies.append((potency, Ea_kJ))
                        
                                # Calculate predicted values from the fitted model
                                ln_visc_pred = arrhenius_function(inv_temp_valid, a, b)
                        
                                # Calculate R-squared
                                r2 = r2_score(ln_visc, ln_visc_pred)
                        
                                # Plot for this potency level with distinct colors
                                ax1.semilogy(temperature_range[valid_indices], predicted_visc_valid, 
                                         'o-', label=f'Potency {potency:.1f}%')
                        
                                ax2.scatter(inv_temp_valid, ln_visc, 
                                        label=f'Potency {potency:.1f}%')
                                ax2.plot(inv_temp_valid, ln_visc_pred, '--', 
                                       label=f'Fit {potency:.1f}% (Ea={Ea_kJ:.1f} kJ/mol)')
                    
                            # Configure plots
                            ax1.set_xlabel('Temperature (°C)')
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
                    
                            # Save the variable potency plot
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
                        
                                # Add a trendline
                                if len(potency_values) > 2:
                                    z = np.polyfit(potency_values, ea_values, 1)
                                    p = np.poly1d(z)
                                    plt.plot(potency_values, p(potency_values), "r--", 
                                           label=f"Trend: {z[0]:.2f}x + {z[1]:.2f}")
                                    plt.legend()
                        
                                # Save the potency influence plot
                                potency_plot_path = f'plots/Potency_Effect_{sanitize_filename(media)}_{sanitize_filename(terpene)}.png'
                                plt.savefig(potency_plot_path)
                                plt.close()
                        
                                add_text(f"Variable potency analysis complete for {media}/{terpene}")
                                add_text(f"  • Plot saved to: {plot_path}")
                                add_text(f"  • Potency effect plot saved to: {potency_plot_path}")
                        
                                # Report the trend
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
                    
                            # Skip the standard analysis for this model when doing variable potency analysis
                            successful_models += 1
                            continue
                
                        # --- Standard single-potency analysis for all models ---
                
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
                        if using_enhanced_models:
                            plt.title(f'Viscosity vs Temperature for {media}/{terpene}\nTerpene: {terpene_pct:.2f}%, Potency: {representative_potency:.1f}%')
                        else:
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
                        if using_enhanced_models:
                            plot_path = f'plots/Arrhenius_{sanitize_filename(media)}_{sanitize_filename(terpene)}_potency{int(representative_potency)}.png'
                        else:
                            plot_path = f'plots/Arrhenius_{sanitize_filename(media)}_{sanitize_filename(terpene)}.png'
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
        Includes potency information when available.
    
        Args:
            models: Dictionary of trained models
            potency_value: Potency value to use for enhanced models
            using_enhanced_models: Flag indicating if enhanced models with potency are being used
        """
        import numpy as np
        import matplotlib.pyplot as plt
        from scipy.optimize import curve_fit
        from scipy import stats
        import os

        # Define function to sanitize filenames
        def sanitize_filename(name):
            """Replace invalid filename characters with underscores."""
            # Replace common invalid filename characters
            invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ']
            sanitized = name
            for char in invalid_chars:
                sanitized = sanitized.replace(char, '_')
            return sanitized
    
        # Create a figure
        plt.figure(figsize=(12, 10))
    
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
    
        # Default values for additional chemical properties
        if using_enhanced_models and potency_value is not None:
            d9_thc_value = potency_value * 0.85  # Estimate d9-THC as 85% of potency
            d8_thc_value = 0.0  # Default to zero
        else:
            d9_thc_value = 0.0
            d8_thc_value = 0.0
    
        # Define a consistent prediction function
        def predict_model_viscosity(model_obj, terpene_pct, temp, pot=None):
            """Intelligently predict viscosity handling different feature sets"""
            try:
                if isinstance(model_obj, dict) and 'model' in model_obj:
                    features = model_obj.get('features', None)
                    actual_model = model_obj['model']
                else:
                    actual_model = model_obj
                    features = None
            
                # Determine expected feature count
                if hasattr(actual_model, 'n_features_in_'):
                    n_features = actual_model.n_features_in_
                elif features is not None:
                    n_features = len(features)
                else:
                    n_features = 2  # Default
            
                # Create appropriate input array
                if n_features == 2:
                    return actual_model.predict([[terpene_pct, temp]])[0]
                elif n_features == 3:
                    return actual_model.predict([[terpene_pct, temp, pot or 80.0]])[0]
                elif n_features == 5:
                    return actual_model.predict([[terpene_pct, temp, pot or 80.0, d9_thc_value, d8_thc_value]])[0]
                else:
                    # Create flexible input array for other feature counts
                    X = [terpene_pct, temp]
                    if n_features > 2:
                        X.append(pot or 80.0)
                    if n_features > 3:
                        X.append(d9_thc_value)
                    if n_features > 4:
                        X.append(d8_thc_value)
                
                    # Pad with zeros if needed
                    while len(X) < n_features:
                        X.append(0.0)
                
                    return actual_model.predict([X])[0]
            except Exception as e:
                print(f"Prediction error: {e}")
                return np.nan
    
        # Process each model
        for model_key, model in models.items():
            try:
                # Extract media and terpene from the model key
                if using_enhanced_models and "_with_chemistry" in model_key:
                    # Remove the "_with_chemistry" suffix
                    display_key = model_key.replace("_with_chemistry", "")
                    media, terpene = display_key.split('_', 1)
                else:
                    media, terpene = model_key.split('_', 1)
            
                media_types.add(media)
            
                # Default terpene percentage
                terpene_pct = 5.0
            
                # Generate viscosity predictions
                predicted_visc = []
                for temp in temperature_range:
                    visc = predict_model_viscosity(model, terpene_pct, temp, potency_value)
                    predicted_visc.append(visc)
            
                predicted_visc = np.array(predicted_visc)
            
                # Filter out invalid values
                valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                if not any(valid_indices):
                    print(f"Warning: No valid predictions for {model_key}")
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
                    'ln_A': a,
                    'potency': potency_value if using_enhanced_models else None
                })
            
            except Exception as e:
                print(f"Error processing {model_key}: {e}")
                import traceback
                traceback.print_exc()
    
        if not results:
            print("No valid results generated. Cannot create comparison plot.")
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
        ax1 = plt.subplot(211)
    
        # Create positions for bars
        positions = np.arange(len(results_df))
        bar_height = 0.8
    
        # Create bars with colors based on media type
        bars = ax1.barh(
            positions, 
            results_df['Ea_kJ'], 
            height=bar_height,
            color=[color_map[media] for media in results_df['media']]
        )
    
        # Add labels
        ax1.set_yticks(positions)
        ax1.set_yticklabels([f"{row['terpene']} ({row['media']})" for _, row in results_df.iterrows()])
        ax1.set_xlabel('Activation Energy (kJ/mol)')
    
        if using_enhanced_models and potency_value is not None:
            ax1.set_title(f'Activation Energy Comparison by Media-Terpene Combination\nPotency: {potency_value:.1f}%')
        else:
            ax1.set_title('Activation Energy Comparison by Media-Terpene Combination')
    
        # Add a legend for media types
        from matplotlib.patches import Patch
        legend_elements = [Patch(facecolor=color_map[media], label=media) for media in media_list]
        ax1.legend(handles=legend_elements, loc='upper right')
    
        # Add value labels
        for i, bar in enumerate(bars):
            ax1.text(
                bar.get_width() + 0.5, 
                bar.get_y() + bar.get_height()/2, 
                f"{results_df['Ea_kJ'].iloc[i]:.1f}", 
                va='center'
            )
    
        # Create a scatter plot showing ln(A) vs Ea correlation (compensation effect)
        ax2 = plt.subplot(212)
    
        # Extract data for scatter plot
        Ea_values = results_df['Ea_kJ']
        lnA_values = results_df['ln_A']
    
        # Create scatter plot with points colored by media type
        for media in media_list:
            media_indices = results_df['media'] == media
            ax2.scatter(
                Ea_values[media_indices], 
                lnA_values[media_indices],
                color=color_map[media],
                label=media,
                s=80,
                alpha=0.7
            )
    
        # Add text labels for each point
        for i, row in results_df.iterrows():
            ax2.annotate(
                row['terpene'],
                (row['Ea_kJ'], row['ln_A']),
                xytext=(5, 0),
                textcoords='offset points',
                fontsize=8
            )
    
        # Add a trend line if there are enough points
        if len(results_df) > 2:
            try:
                # Calculate linear fit
                slope, intercept, r_value, p_value, std_err = stats.linregress(Ea_values, lnA_values)
            
                # Calculate endpoints for the line
                x_min, x_max = min(Ea_values), max(Ea_values)
                x_line = np.array([x_min, x_max])
                y_line = intercept + slope * x_line
            
                # Plot the line
                ax2.plot(x_line, y_line, 'k--', 
                         label=f'Enthalpy-Entropy Compensation\ny = {slope:.2f}x + {intercept:.2f}, R² = {r_value**2:.2f}')
            except Exception as e:
                print(f"Error creating trend line: {e}")
                # Continue without the trend line
    
        ax2.set_xlabel('Activation Energy, Ea (kJ/mol)')
        ax2.set_ylabel('ln(A)')
        ax2.set_title('Enthalpy-Entropy Compensation Effect')
        ax2.grid(True, linestyle='--', alpha=0.7)
        ax2.legend(loc='upper left')
    
        plt.tight_layout()
    
        # Save the plot with potency information in the filename if relevant
        try:
            if using_enhanced_models and potency_value is not None:
                plot_path = f'plots/Activation_Energy_Comparison_Potency{int(potency_value)}.png'
            else:
                plot_path = 'plots/Activation_Energy_Comparison.png'
            plt.savefig(plot_path)
            print(f"Plot saved to: {plot_path}")
        except Exception as e:
            print(f"Error saving plot: {e}")
    
        plt.close()
    
        # Additionally, if we're using enhanced models, create a plot showing potency vs viscosity
        # for a fixed temperature and terpene percentage
        if using_enhanced_models and isinstance(potency_value, (int, float)):
            try:
                # Create a new figure for potency vs viscosity at fixed temperature
                plt.figure(figsize=(10, 8))
            
                # Pick a fixed temperature (25°C is common reference)
                fixed_temp = 25.0
            
                # Pick a single model to analyze (preferably one with good performance)
                # Just use the first valid model we find
                example_model = None
                model_key = None
                for mk, model in models.items():
                    if isinstance(model, dict) and 'model' in model:
                        example_model = model['model']
                        model_key = mk
                        break
                    elif hasattr(model, 'predict'):
                        example_model = model
                        model_key = mk
                        break
            
                if example_model is not None:
                    # Extract media/terpene from the model key
                    if "_with_chemistry" in model_key:
                        display_key = model_key.replace("_with_chemistry", "")
                        media, terpene = display_key.split('_', 1)
                    else:
                        media, terpene = model_key.split('_', 1)
                
                    # Generate a range of potency values
                    potency_range = np.linspace(60, 95, 15)
                
                    # For a few typical terpene percentages
                    terpene_percentages = [3.0, 5.0, 7.0]
                
                    for terpene_pct in terpene_percentages:
                        # Generate viscosity predictions across potency range
                        predicted_visc = []
                        for pot in potency_range:
                            visc = predict_model_viscosity(
                                {'model': example_model},
                                terpene_pct, 
                                fixed_temp, 
                                pot
                            )
                            predicted_visc.append(visc)
                    
                        # Filter out any NaN values
                        valid_indices = ~np.isnan(predicted_visc)
                        if not any(valid_indices):
                            continue
                        
                        # Plot this terpene percentage with only valid values
                        plt.plot(
                            potency_range[valid_indices], 
                            np.array(predicted_visc)[valid_indices], 
                            'o-', 
                            label=f'Terpene {terpene_pct:.1f}%', 
                            linewidth=2
                        )
                
                    plt.xlabel('Total Potency (%)')
                    plt.ylabel('Viscosity (cP)')
                    plt.title(f'Effect of Potency on Viscosity at {fixed_temp}°C\n{media}/{terpene}')
                    plt.grid(True)
                    plt.legend()
                
                    # Save the plot
                    try:
                        if using_enhanced_models and potency_value is not None:
                            plot_path = f'plots/Activation_Energy_Comparison_Potency{int(potency_value)}.png'
                        else:
                            plot_path = 'plots/Activation_Energy_Comparison.png'
                        plt.savefig(plot_path)
                        print(f"Plot saved to: {plot_path}")
                    except Exception as e:
                        print(f"Error saving plot: {e}")

                    plt.close()
            except Exception as e:
                print(f"Error creating potency vs viscosity plot: {e}")
                import traceback
                traceback.print_exc()

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

    def train_models_with_chemistry(self, data=None):
        """
        Train enhanced viscosity prediction models using all available chemical properties.
    
        Args:
            data (pd.DataFrame, optional): DataFrame with training data. If None,
                                         attempts to use data from saved CSV files.
        """
        import glob
        import os
        import pandas as pd
        import numpy as np
        import pickle
        import threading
        from sklearn.ensemble import RandomForestRegressor
    
        # Create progress window - keep existing code
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Training Enhanced Viscosity Models")
        progress_window.geometry("400x200")
        progress_window.transient(self.root)
        progress_window.grab_set()
    
        progress_label = tk.Label(
            progress_window, 
            text="Training enhanced models with chemical properties...",
            font=('Arial', 12)
        )
        progress_label.pack(pady=20)
    
        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill='x', padx=20)
        progress_bar.start()
    
        def update_label(text):
            progress_label.config(text=text)
            progress_window.update_idletasks()
    
        # Define training thread
        def train_thread():
            nonlocal data
            try:
                update_label("Loading and preprocessing data...")
            
                # If no data provided, load from files
                if data is None:
                    # Load Master_Viscosity_Data files first, as they have the most complete data
                    master_files = glob.glob('data/Master_Viscosity_Data_*.csv')
                    regular_files = glob.glob('data/viscosity_data_*.csv')
                
                    if not master_files and not regular_files:
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error", "No training data available. Please extract media database data first."
                        ))
                        progress_window.after(0, progress_window.destroy)
                        return
                
                    data_frames = []
                    # Load master files first
                    for file in master_files:
                        try:
                            df = pd.read_csv(file)
                            data_frames.append(df)
                            print(f"Loaded master file: {file}")
                        except Exception as e:
                            print(f"Error loading {file}: {str(e)}")
                
                    # Then load regular files
                    for file in regular_files:
                        try:
                            df = pd.read_csv(file)
                            data_frames.append(df)
                            print(f"Loaded data file: {file}")
                        except Exception as e:
                            print(f"Error loading {file}: {str(e)}")
                
                    if not data_frames:
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error", "Failed to load any training data."
                        ))
                        progress_window.after(0, progress_window.destroy)
                        return
                
                    data = pd.concat(data_frames, ignore_index=True)
            
                update_label("Checking chemical properties...")
            
                # Check which chemical properties are available
                chem_properties = []
                if 'total_potency' in data.columns:
                    chem_properties.append('total_potency')
                    print(f"Found total_potency column with {data['total_potency'].count()} non-null values")
            
                if 'terpene_pct' in data.columns:
                    chem_properties.append('terpene_pct')
                    print(f"Found terpene_pct column with {data['terpene_pct'].count()} non-null values")
            
                if 'd9_thc' in data.columns:
                    chem_properties.append('d9_thc')
                    print(f"Found d9_thc column with {data['d9_thc'].count()} non-null values")
            
                if 'd8_thc' in data.columns:
                    chem_properties.append('d8_thc')
                    print(f"Found d8_thc column with {data['d8_thc'].count()} non-null values")
            
                # Standard models will use just temperature and terpene_pct
                standard_features = ['terpene_pct', 'temperature']
            
                # Enhanced models will use all available chemical properties
                enhanced_features = ['terpene_pct', 'temperature'] + chem_properties
                # Remove duplicates (terpene_pct might be in both lists)
                enhanced_features = list(dict.fromkeys(enhanced_features))
            
                print(f"Standard features: {standard_features}")
                print(f"Enhanced features: {enhanced_features}")
            
                # Clean up the data
                update_label("Cleaning data...")
                data_cleaned = data.copy()
            
                # Convert columns to numeric
                numeric_columns = ['temperature', 'viscosity'] + chem_properties
                for col in numeric_columns:
                    if col in data_cleaned.columns:
                        data_cleaned[col] = pd.to_numeric(data_cleaned[col], errors='coerce')
            
                # Fill missing values with appropriate defaults
                # For chemical properties, 0 might be appropriate for missing values
                for col in chem_properties:
                    if col in data_cleaned.columns:
                        data_cleaned[col] = data_cleaned[col].fillna(0)
            
                # Train models by media/terpene combination
                update_label("Training models...")
            
                media_terpene_combos = data_cleaned[['media', 'terpene']].drop_duplicates()
                standard_models = {}
                enhanced_models = {}
            
                for idx, row in media_terpene_combos.iterrows():
                    media = row['media']
                    terpene = row['terpene']
                
                    print(f"Processing {media}/{terpene} combination...")
                
                    # Filter data for this combination
                    combo_data = data_cleaned[
                        (data_cleaned['media'] == media) & 
                        (data_cleaned['terpene'] == terpene)
                    ]
                
                    # Skip if not enough data
                    if len(combo_data) < 5:
                        print(f"Skipping {media}/{terpene} - insufficient data ({len(combo_data)} samples)")
                        continue
                
                    # Train standard model (temperature and terpene_pct only)
                    try:
                        # Extract features and target
                        X_std = combo_data[standard_features].dropna()
                        if len(X_std) >= 5:
                            # Get corresponding viscosity values
                            y_std = combo_data.loc[X_std.index, 'viscosity']
                        
                            # Train a model
                            model = RandomForestRegressor(
                                n_estimators=100,
                                max_depth=15,
                                min_samples_leaf=3,
                                random_state=42
                            )
                            model.fit(X_std, y_std)
                        
                            # Store the model
                            model_key = f"{media}_{terpene}"
                            standard_models[model_key] = model
                            print(f"Trained standard model for {media}/{terpene} with {len(X_std)} samples")
                    except Exception as e:
                        print(f"Error training standard model for {media}/{terpene}: {str(e)}")
                
                    # Train enhanced model with chemical properties
                    # Only if we have at least one chemical property to use
                    if len(enhanced_features) > len(standard_features):
                        try:
                            # Filter features that exist in this dataset
                            valid_features = [f for f in enhanced_features if f in combo_data.columns]
                            X_enh = combo_data[valid_features].dropna()
                        
                            if len(X_enh) >= 5:
                                # Get corresponding viscosity values
                                y_enh = combo_data.loc[X_enh.index, 'viscosity']
                            
                                # Train enhanced model
                                model = RandomForestRegressor(
                                    n_estimators=100,
                                    max_depth=15,
                                    min_samples_leaf=3,
                                    random_state=42
                                )
                                model.fit(X_enh, y_enh)
                            
                                # Store the model
                                model_key = f"{media}_{terpene}_with_chemistry"
                                enhanced_models[model_key] = {
                                    'model': model,
                                    'features': valid_features  # Save which features this model uses
                                }
                                print(f"Trained enhanced model for {media}/{terpene} with {len(X_enh)} samples")
                                print(f"  Features used: {valid_features}")
                        except Exception as e:
                            print(f"Error training enhanced model for {media}/{terpene}: {str(e)}")
            
                # Save models
                update_label("Saving models...")
            
                # Create models directory if it doesn't exist
                os.makedirs('models', exist_ok=True)
            
                # Save standard models
                if standard_models:
                    with open('models/viscosity_models.pkl', 'wb') as f:
                        pickle.dump(standard_models, f)
                    print(f"Saved {len(standard_models)} standard models")
                    self.viscosity_models = standard_models
            
                # Save enhanced models with chemical properties
                if enhanced_models:
                    with open('models/viscosity_models_with_chemistry.pkl', 'wb') as f:
                        pickle.dump(enhanced_models, f)
                    print(f"Saved {len(enhanced_models)} enhanced models")
                    self.enhanced_viscosity_models = enhanced_models
            
                # Close progress window and show success message
                progress_window.after(0, progress_window.destroy)
            
                # Report results to user
                message = f"Training complete!\n\n"
                message += f"Standard models: {len(standard_models)}\n"
                message += f"Enhanced models: {len(enhanced_models)}\n\n"
            
                if len(enhanced_models) > 0:
                    sample_model = list(enhanced_models.values())[0]
                    message += f"Enhanced models use features: {sample_model['features']}"
            
                self.root.after(0, lambda: messagebox.showinfo("Success", message))
            
            except Exception as e:
                import traceback
                traceback_str = traceback.format_exc()
                print(f"Error in training thread: {e}\n{traceback_str}")
                progress_window.after(0, progress_window.destroy)
                error_msg = str(e) # capture error
                self.root.after(0, lambda: messagebox.showerror("Error", f"Training failed: {str(e)}"))
    
        # Start the training thread
        training_thread = threading.Thread(target=train_thread)
        training_thread.daemon = True
        training_thread.start()

    def calculate_viscosity_with_chemistry(self):
        """
        Calculate terpene percentage based on target viscosity using models with chemical properties.
        """
        try:
            # Extract input values
            media = self.media_var.get()
            media_brand = self.media_brand_var.get()
            terpene = self.terpene_var.get()
            terpene_brand = self.terpene_brand_var.get()
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
        
            # Get total potency
            total_potency = 0.0
            if hasattr(self, 'total_potency_var'):
                try:
                    total_potency = float(self.total_potency_var.get())
                except (ValueError, tk.TclError):
                    total_potency = 0.0
        
            try:
                d9_thc = float(self.d9_thc_var.get())
            except (ValueError, tk.TclError):
                # Fallback to estimation if the input is invalid
                d9_thc = total_potency * 0.85  # Assuming d9-THC is about 85% of total potency

            try:
                d8_thc = float(self.d8_thc_var.get())
            except (ValueError, tk.TclError):
                d8_thc = 0.0  # Default value
        
            # Load enhanced models
            enhanced_models = {}
            enhanced_model_path = 'models/viscosity_models_with_chemistry.pkl'
            if os.path.exists(enhanced_model_path):
                try:
                    with open(enhanced_model_path, 'rb') as f:
                        enhanced_models = pickle.load(f)
                except Exception as e:
                    print(f"Error loading enhanced models: {e}")
                    enhanced_models = {}
        
            # Fallback to standard model if needed
            model_key = f"{media}_{terpene}"
            enhanced_key = f"{media}_{terpene}_with_chemistry"
        
            # Import required packages here to maintain lazy loading
            from scipy.optimize import fsolve
            import numpy as np
        
            # Check if we have an enhanced model
            if enhanced_key in enhanced_models:
                # Get the model and its features
                model_info = enhanced_models[enhanced_key]

                if isinstance(model_info, dict) and 'model' in model_info:
                    model = model_info['model']
                    features = model_info['features']
                else:
                    model = model_info
                    features = ['terpene_pct', 'temperature', 'total_potency', 'd9_thc', 'd8_thc']
                    features = [f for f in features if f in ['terpene_pct', 'temperature','total_potency','d9_thc','d8_thc']]

                model = model_info['model']
                features = model_info['features']
            
                # Check if total_potency is one of the features the model uses
                if 'total_potency' in features and total_potency > 0:
                    # Create a function to find the terpene percentage that gives target viscosity
                    def objective(terpene_pct):
                        # Create input array based on model features
                        X = np.zeros(len(features))
                        for i, feature in enumerate(features):
                            if feature == 'terpene_pct':
                                X[i] = terpene_pct
                            elif feature == 'temperature':
                                X[i] = 25.0  # Fixed reference temperature
                            elif feature == 'total_potency':
                                X[i] = total_potency
                            elif feature == 'd9_thc':
                                X[i] = d9_thc
                            elif feature == 'd8_thc':
                                X[i] = d8_thc
                    
                        # Predict viscosity with these inputs
                        predicted_viscosity = model.predict([X])[0]
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
                
                    # Show which features were used
                    features_text = ", ".join(features)
                    messagebox.showinfo(
                        "Calculation Complete", 
                        f"Calculation performed using enhanced model with chemical properties.\n\n"
                        f"Features used: {features_text}\n\n"
                        f"For {exact_percent:.1f}% terpenes, estimated viscosity: {target_viscosity:.1f}"
                    )
                    return
        
            # No enhanced model, or total_potency not available
            # Fall back to existing methods
            if model_key in self.viscosity_models:
                # Use the standard model
                self.calculate_viscosity()
                messagebox.showinfo(
                    "Notice", 
                    "Enhanced model with chemical properties not available for this combination.\n"
                    "Used standard model instead."
                )
            else:
                # No models available - suggest iterative method
                result = messagebox.askyesno(
                    "No Model Available", 
                    "No prediction model found for this combination.\n\n"
                    "Would you like to use the Iterative Method instead?"
                )
                if result:
                    # Switch to the iterative tab
                    self.notebook.select(1)  # Assuming tab index 1 is the iterative method tab
    
        except Exception as e:
            import traceback
            traceback_str = traceback.format_exc()
            print(f"Error during calculation: {e}\n{traceback_str}")
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")


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
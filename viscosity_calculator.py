"""
viscosity_calculator.py
Module for calculating terpene percentages based on viscosity.
"""
import tkinter as tk
from tkinter import ttk, Toplevel, StringVar, DoubleVar, Frame, Label, Entry, Button, messagebox, IntVar
import json
import os
from utils import FONT, APP_BACKGROUND_COLOR, BUTTON_COLOR

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
        self.formulation_db = self.load_formulation_database()
        
        # Store notebook reference
        self.notebook = None

        self.viscosity_models = self.load_models()
        
    def show_calculator(self):
        """
        Show the viscosity calculator window with tabbed interface.
        
        Returns:
            tk.Toplevel: The calculator window
        """
        # Create a new top-level window
        calculator_window = Toplevel(self.root)
        calculator_window.title("Calculate Terpene % for Viscosity")
        calculator_window.geometry("550x420")
        calculator_window.resizable(False, False)
        calculator_window.configure(bg=APP_BACKGROUND_COLOR)
        
        # Make the window modal
        calculator_window.transient(self.root)
        calculator_window.grab_set()
        
        # Center the window on the screen
        self.gui.center_window(calculator_window, 550, 420)
        
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
    
    def create_calculator_tab(self, notebook):
        """
        Create the main calculator tab.
        
        Args:
            notebook (ttk.Notebook): The notebook widget to add the tab to
        """
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
        
        # Create button frame for a single row of buttons
        button_frame = Frame(form_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.grid(row=9, column=0, columnspan=4, pady=10)

        # Calculate button
        calculate_btn = ttk.Button(
            button_frame,
            text="Calculate",
            command=self.calculate_viscosity
        )
        calculate_btn.pack(side="left", padx=(0, 5))

        # Upload Data button
        upload_btn = ttk.Button(
            button_frame,
            text="Upload Data",
            command=self.upload_training_data
        )
        upload_btn.pack(side="left", padx=5)

        # Train Models button
        train_btn = ttk.Button(
            button_frame,
            text="Train Models",
            command=self.train_models_from_data
        )
        train_btn.pack(side="left", padx=5)

        analyze_btn = ttk.Button(
            button_frame,
            text="Analyze Data",
            command=self.analyze_training_data
        )
        analyze_btn.pack(side="left", padx=5)

    
    def create_advanced_tab(self, notebook):
        """
        Create the advanced tab for iterative viscosity calculation with balanced spacing.
        """
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
    
        # Media selection frame
        media_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        media_frame.pack(fill='x', pady=5)
    
        Label(media_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).pack(side="left", padx=(5, 5))
    
        media_dropdown = ttk.Combobox(
            media_frame, 
            textvariable=self.media_var,
            values=self.media_options,
            state="readonly",
            width=15
        )
        media_dropdown.pack(side="left", padx=(0, 10))
    
        Label(media_frame, text="Media Brand:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT).pack(side="left", padx=(5, 5))
    
        media_brand_entry = Entry(media_frame, textvariable=self.media_brand_var, width=15)
        media_brand_entry.pack(side="left")
    
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

    def create_temperature_block(self, parent, temperature):
        """
        Create a block for a temperature with a table for 3 runs
    
        Args:
            parent: Parent frame to add the block to
            temperature: Temperature value for this block
        """
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
            Entry(table_frame, textvariable=torque_var, width=col_widths[1]).grid(
                row=run+1, column=1, sticky="nsew", padx=1, pady=1)
        
            # Create entry for viscosity
            Entry(table_frame, textvariable=visc_var, width=col_widths[2]).grid(
                row=run+1, column=2, sticky="nsew", padx=1, pady=1)
    
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
        import math
    
        for temp, _ in self.temperature_blocks:
            # Get the torque and viscosity values for this temperature
            torque_values = []
            visc_values = []
        
            for run in range(3):
                # Find the corresponding variables for this temperature
                for t, torque_var in self.torque_vars[run]:
                    if t == temp:
                        try:
                            torque_value = torque_var.get().strip()
                            if torque_value:  # Check if not empty
                                torque = float(torque_value)
                                torque_values.append(torque)
                        except ValueError:
                            pass
            
                for t, visc_var in self.viscosity_vars[run]:
                    if t == temp:
                        try:
                            visc_value = visc_var.get().strip()
                            if visc_value:  # Check if not empty
                                visc = float(visc_value)
                                visc_values.append(visc)
                        except ValueError:
                            pass
        
            # Calculate averages if we have values
            if torque_values:
                avg_torque = sum(torque_values) / len(torque_values)
                # Find the average torque variable for this temperature
                for t, avg_var in self.avg_torque_vars:
                    if t == temp:
                        avg_var.set(f"{avg_torque:.1f}")
                        break
        
            if visc_values:
                avg_visc = sum(visc_values) / len(visc_values)
                # Find the average viscosity variable for this temperature
                for t, avg_var in self.avg_visc_vars:
                    if t == temp:
                        avg_var.set(f"{avg_visc:.1f}")
                        break
    
        # Show a message to confirm calculation
        if self.temperature_blocks:
            messagebox.showinfo("Calculation Complete", "Averages have been calculated successfully.")

    def save_block_measurements(self):
        """Save the block-based viscosity measurements to the database"""
        import datetime
        # Create a data structure to save
        measurements = {
            "media": self.media_var.get(),
            "media_brand": self.media_brand_var.get(),
            "timestamp": datetime.datetime.now().isoformat(),
            "temperature_data": []
        }
    
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
    
        # Add measurements to the database
        key = f"{measurements['media']}_{measurements['media_brand']}_raw_viscosity_blocks"
    
        if key not in self.formulation_db:
            self.formulation_db[key] = []
    
        self.formulation_db[key].append(measurements)
    
        # Save to file
        self.save_formulation_database()
    
        messagebox.showinfo("Success", "Viscosity measurements saved successfully!")

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
                model = self.viscosity_models[model_key]
            
                # For numerical solution, find terpene percentage that gives target viscosity
                from scipy.optimize import fsolve
            
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
            
                # ... rest of your code for the iterative method
    
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")

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


# ------------ Regression Model -------------

    def collect_structured_data(self):
        """
        Collect viscosity data in a structured manner following Design of Experiments
        principles to minimize experiments while maximizing information gain.
        """
        # Define a grid of test points 
        test_points = []
    
        # For each media type
        for media in self.media_options:
            # Test at several terpene percentages
            for terpene_pct in [0.0, 0.5, 1.0, 2.0, 3.0, 5.0, 7.0, 10.0]:
                # Test at different temperatures
                for temp in [25, 30, 40, 50]:
                    test_points.append({
                        'media': media,
                        'terpene_pct': terpene_pct,
                        'temperature': temp
                    })
    
        return test_points

    def build_viscosity_models(self, data):
        """
        Build multiple regression models for viscosity prediction and 
        select the best one based on cross-validation.
        """
        import numpy as np
        import pandas as pd
        from sklearn.model_selection import train_test_split, cross_val_score
        from sklearn.linear_model import LinearRegression
        from sklearn.preprocessing import PolynomialFeatures
        from sklearn.ensemble import RandomForestRegressor
        from sklearn.pipeline import make_pipeline

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

        # Split data - using a smaller test size if we have limited data
        test_size = min(0.2, 1/len(X))  # Ensure at least 1 test sample, but no more than 20%
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=test_size, random_state=42)
        
        print(f"Training set size: {len(X_train)}, Test set size: {len(X_test)}")

        # Create models
        models = {
            'linear': LinearRegression(),
            'polynomial': make_pipeline(PolynomialFeatures(2), LinearRegression()),
            'random_forest': RandomForestRegressor(n_estimators=100, random_state=42)
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
                    cv = min(5, len(X_train))  # Don't use more folds than we have samples
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
        
    def arrhenius_model(self, data):
        """
        Build a physics-based model using the Arrhenius equation for viscosity's
        temperature dependence, and a polynomial relationship with terpene percentage.
        """
        import numpy as np
        from scipy.optimize import curve_fit

        # Define Arrhenius-based model function
        def viscosity_model(X, A, B, C, D, E):
            """
            X[:, 0] = terpene_pct
            X[:, 1] = temperature in Kelvin
        
            Model: μ = (A + B*T + C*T²) * exp(D/T) * (1 - E*terpene_pct)
            """
            terpene_pct = X[:, 0]
            T = X[:, 1] + 273.15  # Convert to Kelvin
        
            return (A + B*T + C*T**2) * np.exp(D/T) * (1 - E*terpene_pct)
    
        # Prepare data
        X = np.column_stack([data['terpene_pct'], data['temperature']])
        y = data['viscosity']
    
        # Initial parameter guess
        initial_params = [1.0, 0.0, 0.0, 1000.0, 0.1]
    
        # Fit the model
        params, covariance = curve_fit(viscosity_model, X, y, p0=initial_params)
    
        # Create a prediction function
        def predict_viscosity(terpene_pct, temperature):
            X_pred = np.array([[terpene_pct, temperature]])
            return viscosity_model(X_pred, *params)[0]
    
        return predict_viscosity, params

    def save_models(self, models_dict):
        """
        Save trained viscosity prediction models to disk.
        
        Args:
            models_dict (dict): Dictionary of trained models for different 
                            media/terpene combinations
        """
        import pickle
        import os
        
        # Create the models directory if it doesn't exist
        os.makedirs("models", exist_ok=True)
        
        with open('models/viscosity_models.pkl', 'wb') as f:
            pickle.dump(models_dict, f)
        
        print(f"Saved {len(models_dict)} models to models/viscosity_models.pkl")
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
        import pandas as pd
        import os
        import glob
        import traceback
        
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

        progress_label = Label(progress_window, text="Training models...", font=FONT)
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
                    self.root.after(0, lambda msg=error_msg: messagebox.showerror("Library Error", 
                                                        f"Required library not found: {msg}\n"
                                                        f"Make sure scikit-learn is installed."))
                    return
                    
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
                        model, results = self.build_viscosity_models(combo_data_clean)
                        
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
                    self.save_models(models_dict)
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
        import threading
        training_thread = threading.Thread(target=train_thread)
        training_thread.daemon = True
        training_thread.start()
    def analyze_training_data(self, data=None, show_dialog=True):
        """Analyze the training data CSV to identify issues."""
        import pandas as pd
        import matplotlib.pyplot as plt
        from tkinter import Toplevel, Text, Scrollbar, Label
        
        if data is None:
            data = self.upload_training_data()
            if data is None:
                return
        
        # Create a detailed analysis report
        report = []
        report.append(f"CSV Analysis Report")
        report.append(f"==================")
        report.append(f"Total rows: {len(data)}")
        report.append(f"Columns: {', '.join(data.columns)}")
        report.append(f"")
        
        # Check for NaN values
        for col in data.columns:
            nan_count = data[col].isna().sum()
            nan_percent = (nan_count / len(data)) * 100
            report.append(f"{col}: {nan_count} NaN values ({nan_percent:.2f}%)")
        
        report.append(f"")
        
        # Check unique values for categorical columns
        for col in ['media', 'terpene', 'spindle', 'speed']:
            if col in data.columns:
                unique_vals = data[col].dropna().unique()
                report.append(f"{col} unique values: {', '.join(str(x) for x in unique_vals)}")
        
        report.append(f"")
        
        # Check numeric ranges
        for col in ['terpene_pct', 'temperature', 'viscosity']:
            if col in data.columns:
                non_nan_data = data[col].dropna()
                report.append(f"{col} range: {non_nan_data.min()} to {non_nan_data.max()}")
        
        report.append(f"")
        
        # Check for usable combinations
        media_terpene_combos = data[['media', 'terpene']].drop_duplicates()
        report.append(f"Found {len(media_terpene_combos)} unique media/terpene combinations:")
        
        usable_combos = 0
        for _, row in media_terpene_combos.iterrows():
            media = row['media']
            terpene = row['terpene']
            combo_data = data[(data['media'] == media) & (data['terpene'] == terpene)]
            clean_combo = combo_data.dropna(subset=['terpene_pct', 'temperature', 'viscosity'])
            status = "USABLE" if len(clean_combo) >= 5 else "NOT ENOUGH DATA"
            if status == "USABLE":
                usable_combos += 1
            report.append(f"  {media}/{terpene}: {len(combo_data)} samples, {len(clean_combo)} clean samples - {status}")
        
        report.append(f"")
        report.append(f"Summary: {usable_combos} usable combinations out of {len(media_terpene_combos)}")
        
        # Print to console
        print("\n".join(report))
        
        # Show in dialog if requested
        if show_dialog:
            report_window = Toplevel(self.root)
            report_window.title("CSV Analysis Report")
            report_window.geometry("600x400")
            
            Label(report_window, text="CSV Analysis Report", font=("Arial", 14, "bold")).pack(pady=10)
            
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
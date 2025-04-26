import tkinter as tk
from tkinter import ttk, Frame, Label, Entry, StringVar, DoubleVar, Canvas
from tkinter import messagebox

# Import constants from core module
from .core import APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT

class UI_Methods:
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
        ttk.Button(button_frame, text="Save Measurements", command=self.save_block_measurements).pack(side="right", padx=5)	

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
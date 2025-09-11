# views/calculators/viscosity_view.py
"""
views/calculators/viscosity_view.py
Viscosity calculator view components.
This contains the UI components from viscosity_gui.py and viscosity_calculator/ui.py.
"""

import tkinter as tk
from tkinter import ttk, messagebox, Frame, Label, Entry, StringVar, DoubleVar, Canvas, Scrollbar
from typing import Optional, Callable, Dict, Any, List
import os


class ViscosityView:
    """Viscosity calculator view for displaying calculator interfaces."""
    
    def __init__(self, parent: tk.Widget, controller: Optional[Any] = None):
        """Initialize the viscosity calculator view."""
        self.parent = parent
        self.controller = controller
        self.window: Optional[tk.Toplevel] = None
        self.notebook: Optional[ttk.Notebook] = None
        
        # UI Variables for calculator tab
        self.media_var = StringVar()
        self.terpene_var = StringVar()
        self.media_brand_var = StringVar()
        self.terpene_brand_var = StringVar()
        self.sample_mass_var = StringVar()
        self.target_viscosity_var = StringVar()
        self.result_var = StringVar()
        self.exact_mass_var = StringVar()
        self.start_percent_var = StringVar()
        self.start_mass_var = StringVar()
        
        # UI Variables for advanced tab
        self.step1_amount_var = StringVar()
        self.step1_viscosity_var = StringVar()
        self.step2_amount_var = StringVar()
        self.step2_viscosity_var = StringVar()
        self.expected_viscosity_var = StringVar()
        
        # UI Variables for measure tab
        self.temperature_blocks = []
        self.speed_vars = []
        self.torque_vars = [[] for _ in range(3)]  # 3 runs for each temperature
        self.viscosity_vars = [[] for _ in range(3)]
        
        # Media options
        self.media_options = ["D8", "D9", "Liquid Diamonds", "Other"]
        
        # Callbacks (set by controller)
        self.on_calculate_terpene: Optional[Callable] = None
        self.on_calculate_step1: Optional[Callable] = None
        self.on_calculate_step2: Optional[Callable] = None
        self.on_save_formulation: Optional[Callable] = None
        self.on_upload_training_data: Optional[Callable] = None
        self.on_train_standard_models: Optional[Callable] = None
        self.on_train_enhanced_models: Optional[Callable] = None
        self.on_analyze_models: Optional[Callable] = None
        self.on_arrhenius_analysis: Optional[Callable] = None
        self.on_save_measurements: Optional[Callable] = None
        self.on_add_temperature_block: Optional[Callable] = None
        
        print("DEBUG: ViscosityView initialized")
    
    def show_calculator_window(self) -> tk.Toplevel:
        """Show the viscosity calculator as a standalone window."""
        if self.window and self.window.winfo_exists():
            self.window.lift()
            return self.window
        
        # Constants
        APP_BACKGROUND_COLOR = '#0504AA'
        
        # Create window
        self.window = tk.Toplevel(self.parent)
        self.window.title("Viscosity Calculator")
        self.window.geometry("550x500")
        self.window.resizable(True, True)
        self.window.configure(bg=APP_BACKGROUND_COLOR)
        
        # Make modal
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # Center window
        self._center_window(self.window, 550, 500)
        
        # Create menu
        self._create_calculator_menu()
        
        # Create main container
        main_container = ttk.Frame(self.window)
        main_container.pack(fill='both', expand=True, padx=8, pady=8)
        
        # Embed calculator in container
        self.embed_calculator(main_container)
        
        # Bind close event
        self.window.protocol("WM_DELETE_WINDOW", self._on_window_close)
        
        print("DEBUG: ViscosityView - calculator window shown")
        return self.window
    
    def show_embedded_calculator(self, parent_frame: tk.Widget) -> ttk.Notebook:
        """Show the calculator embedded in a parent frame."""
        return self.embed_calculator(parent_frame)
    
    def embed_calculator(self, parent_frame: tk.Widget) -> ttk.Notebook:
        """Embed the calculator interface in a parent frame."""
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(parent_frame)
        self.notebook.pack(fill='both', expand=True)
        
        # Create the tabs
        self.create_calculator_tab()
        self.create_advanced_tab()
        self.create_measure_tab()
        
        # Bind tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        print("DEBUG: ViscosityView - calculator embedded successfully")
        return self.notebook
    
    def _create_calculator_menu(self):
        """Create the calculator window menu."""
        if not self.window:
            return
        
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Upload Viscosity Data", command=self._on_upload_training_data)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Model menu
        model_menu = tk.Menu(menubar, tearoff=0)
        model_menu.add_command(label="Train Standard Models", command=self._on_train_standard_models)
        model_menu.add_command(label="Train Enhanced Models with Potency", command=self._on_train_enhanced_models)
        model_menu.add_command(label="Analyze Models", command=self._on_analyze_models)
        model_menu.add_command(label="Arrhenius Analysis", command=self._on_arrhenius_analysis)
        menubar.add_cascade(label="Model", menu=model_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Help", command=self._on_help)
        help_menu.add_command(label="About", command=self._on_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        print("DEBUG: ViscosityView - menu created")
    
    def create_calculator_tab(self):
        """Create the main calculator tab."""
        if not self.notebook:
            return
        
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = 'white'
        
        tab1 = Frame(self.notebook, bg=APP_BACKGROUND_COLOR)
        self.notebook.add(tab1, text="Calculator")
        
        # Create form frame
        form_frame = Frame(tab1, bg=APP_BACKGROUND_COLOR, padx=20, pady=10)
        form_frame.pack(fill='both', expand=True)
        
        # Configure grid weights
        for i in range(4):
            form_frame.columnconfigure(i, weight=1)
        
        # Row 0: Title and explanation
        explanation = Label(form_frame, 
                           text="Calculate terpene % based on viscosity formula.\nIf no formula is known, use the 'Iterative Method' tab.",
                           bg=APP_BACKGROUND_COLOR, font=FONT, justify="center")
        explanation.grid(row=0, column=0, columnspan=4, pady=(0, 10))
        
        # Row 1: Media and Terpene
        Label(form_frame, text="Media:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=1, column=0, sticky="w", pady=5)
        
        media_dropdown = ttk.Combobox(form_frame, textvariable=self.media_var,
                                     values=self.media_options, state="readonly", width=15)
        media_dropdown.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        media_dropdown.current(0)
        
        Label(form_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=1, column=2, sticky="w", pady=5)
        
        terpene_entry = Entry(form_frame, textvariable=self.terpene_var, width=15)
        terpene_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        
        # Row 2: Media Brand and Terpene Brand
        Label(form_frame, text="Media Brand:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=2, column=0, sticky="w", pady=5)
        
        media_brand_entry = Entry(form_frame, textvariable=self.media_brand_var, width=15)
        media_brand_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        Label(form_frame, text="Terpene Brand:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=2, column=2, sticky="w", pady=5)
        
        terpene_brand_entry = Entry(form_frame, textvariable=self.terpene_brand_var, width=15)
        terpene_brand_entry.grid(row=2, column=3, sticky="w", padx=5, pady=5)
        
        # Row 3: Sample Mass and Target Viscosity
        Label(form_frame, text="Sample Mass (g):", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=3, column=0, sticky="w", pady=5)
        
        sample_mass_entry = Entry(form_frame, textvariable=self.sample_mass_var, width=15)
        sample_mass_entry.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        Label(form_frame, text="Target Viscosity:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=3, column=2, sticky="w", pady=5)
        
        target_viscosity_entry = Entry(form_frame, textvariable=self.target_viscosity_var, width=15)
        target_viscosity_entry.grid(row=3, column=3, sticky="w", padx=5, pady=5)
        
        # Row 4: Result display
        Label(form_frame, text="Terpene %:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=4, column=0, sticky="w", pady=5)
        
        result_label = Label(form_frame, textvariable=self.result_var, 
                           bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT)
        result_label.grid(row=4, column=1, sticky="w", pady=5)
        
        # Row 5: Exact Mass and Start Mass
        Label(form_frame, text="Exact Mass:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=5, column=0, sticky="w", pady=5)
        
        exact_mass_label = Label(form_frame, textvariable=self.exact_mass_var, 
                               bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT)
        exact_mass_label.grid(row=5, column=1, sticky="w", pady=5)
        
        Label(form_frame, text="Start %:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=5, column=2, sticky="w", pady=5)
        
        start_percent_label = Label(form_frame, textvariable=self.start_percent_var, 
                                  bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT)
        start_percent_label.grid(row=5, column=3, sticky="w", pady=5)
        
        # Row 6: Start Mass
        Label(form_frame, text="Start Mass:", bg=APP_BACKGROUND_COLOR, font=FONT, anchor="w").grid(
            row=6, column=0, sticky="w", pady=5)
        
        start_mass_label = Label(form_frame, textvariable=self.start_mass_var, 
                               bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT)
        start_mass_label.grid(row=6, column=1, sticky="w", pady=5)
        
        # Row 7: Calculate button
        button_frame = Frame(form_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.grid(row=7, column=0, columnspan=4, pady=20)
        
        calculate_btn = ttk.Button(button_frame, text="Calculate", command=self._on_calculate_terpene)
        calculate_btn.pack(padx=5, pady=5)
        
        print("DEBUG: ViscosityView - calculator tab created")
    
    def create_advanced_tab(self):
        """Create the advanced/iterative method tab."""
        if not self.notebook:
            return
        
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = 'white'
        
        tab2 = Frame(self.notebook, bg=APP_BACKGROUND_COLOR)
        self.notebook.add(tab2, text="Iterative Method")
        
        # Create form frame
        form_frame = Frame(tab2, bg=APP_BACKGROUND_COLOR, padx=20, pady=10)
        form_frame.pack(fill='both', expand=True)
        
        # Configure column weights
        for i in range(3):
            form_frame.columnconfigure(i, weight=1)
        
        # Row 0: Title and explanation
        explanation = Label(form_frame, 
                           text="Use this method when no formula is available.\nFollow the steps and measure viscosity at each stage.",
                           bg=APP_BACKGROUND_COLOR, font=FONT, justify="center")
        explanation.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # Create steps frame for better organization
        steps_frame = Frame(form_frame, bg=APP_BACKGROUND_COLOR, relief="ridge", bd=1)
        steps_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=10, padx=10)
        steps_frame.columnconfigure(1, weight=1)
        
        # Step 1 row
        Label(steps_frame, text="Step 1: Add", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=FONT).grid(row=0, column=0, sticky="e", padx=(5, 0), pady=3)
        Label(steps_frame, textvariable=self.step1_amount_var, 
              bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT).grid(row=0, column=1, pady=3)
        Label(steps_frame, text="of terpenes", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=FONT).grid(row=0, column=2, sticky="w", padx=(0, 5), pady=3)
        
        # Viscosity input row 1
        visc_frame1 = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        visc_frame1.grid(row=1, column=0, columnspan=3, sticky="ew", pady=3)
        visc_frame1.columnconfigure(0, weight=1)
        visc_frame1.columnconfigure(1, weight=1)
        
        Label(visc_frame1, text="Viscosity @ 25C:", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=FONT).grid(row=0, column=0, sticky="e", padx=(0, 5))
        Entry(visc_frame1, textvariable=self.step1_viscosity_var, width=10).grid(row=0, column=1, sticky="w")
        
        # Step 2 row
        Label(steps_frame, text="Step 2: Add", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=FONT).grid(row=2, column=0, sticky="e", padx=(5, 0), pady=3)
        Label(steps_frame, textvariable=self.step2_amount_var, 
              bg=APP_BACKGROUND_COLOR, fg="#00b539", font=FONT).grid(row=2, column=1, pady=3)
        Label(steps_frame, text="of terpenes", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=FONT).grid(row=2, column=2, sticky="w", padx=(0, 5), pady=3)
        
        # Viscosity input row 2
        visc_frame2 = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        visc_frame2.grid(row=3, column=0, columnspan=3, sticky="ew", pady=3)
        visc_frame2.columnconfigure(0, weight=1)
        visc_frame2.columnconfigure(1, weight=1)
        
        Label(visc_frame2, text="Viscosity @ 25C:", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=FONT).grid(row=0, column=0, sticky="e", padx=(0, 5))
        Entry(visc_frame2, textvariable=self.step2_viscosity_var, width=10).grid(row=0, column=1, sticky="w")
        
        # Expected Viscosity display
        expected_frame = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        expected_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=8)
        expected_frame.columnconfigure(0, weight=1)
        
        expected_container = Frame(expected_frame, bg=APP_BACKGROUND_COLOR)
        expected_container.grid(row=0, column=0)
        
        Label(expected_container, text="Expected Viscosity: ", bg=APP_BACKGROUND_COLOR, fg="black", 
              font=(FONT[0], FONT[1], "bold")).pack(side="left")
        Label(expected_container, textvariable=self.expected_viscosity_var, 
              bg=APP_BACKGROUND_COLOR, fg="#00b539", font=(FONT[0], FONT[1], "bold")).pack(side="left")
        
        # Buttons
        button_frame = Frame(steps_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=5)
        
        calculate_btn1 = ttk.Button(button_frame, text="Calculate Step 1", 
                                   command=self._on_calculate_step1, width=16)
        calculate_btn1.pack(side="left", padx=5, expand=True)
        
        calculate_btn2 = ttk.Button(button_frame, text="Calculate Step 2", 
                                   command=self._on_calculate_step2, width=16)
        calculate_btn2.pack(side="left", padx=5, expand=True)
        
        save_btn = ttk.Button(button_frame, text="Save", 
                             command=self._on_save_formulation, width=8)
        save_btn.pack(side="left", padx=5, expand=True)
        
        print("DEBUG: ViscosityView - advanced tab created")
    
    def create_measure_tab(self):
        """Create the measurement tab with temperature blocks."""
        if not self.notebook:
            return
        
        FONT = ('Arial', 12)
        APP_BACKGROUND_COLOR = 'white'
        
        tab3 = Frame(self.notebook, bg=APP_BACKGROUND_COLOR)
        self.notebook.add(tab3, text="Measure")
        
        # Main container
        main_frame = Frame(tab3, bg=APP_BACKGROUND_COLOR)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title section
        title_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        title_frame.pack(fill='x', pady=(0, 10))
        
        Label(title_frame, text="Raw Viscosity Measurement", 
              bg=APP_BACKGROUND_COLOR, fg="black", font=(FONT[0], FONT[1]+2, "bold")).pack(pady=(0, 2))
        
        # Input fields frame
        input_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        input_frame.pack(fill='x', pady=5)
        
        # Configure grid columns
        for i in range(6):
            input_frame.columnconfigure(i, weight=1)
        
        # Media and Terpene inputs
        Label(input_frame, text="Media:", bg=APP_BACKGROUND_COLOR, font=FONT).grid(
            row=0, column=0, sticky="w", padx=5, pady=2)
        media_measure_dropdown = ttk.Combobox(input_frame, values=self.media_options, 
                                            state="readonly", width=12)
        media_measure_dropdown.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        Label(input_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, font=FONT).grid(
            row=0, column=2, sticky="w", padx=5, pady=2)
        terpene_measure_entry = Entry(input_frame, width=12)
        terpene_measure_entry.grid(row=0, column=3, sticky="w", padx=5, pady=2)
        
        Label(input_frame, text="% Terpene:", bg=APP_BACKGROUND_COLOR, font=FONT).grid(
            row=0, column=4, sticky="w", padx=5, pady=2)
        percent_entry = Entry(input_frame, width=10)
        percent_entry.grid(row=0, column=5, sticky="w", padx=5, pady=2)
        
        # Scrollable area for temperature blocks
        scroll_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        scroll_frame.pack(fill='both', expand=True, pady=10)
        
        canvas = Canvas(scroll_frame, bg=APP_BACKGROUND_COLOR, height=300)
        scrollbar = Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        
        # Container for all temperature blocks
        blocks_frame = Frame(canvas, bg=APP_BACKGROUND_COLOR)
        
        # Configure scrolling
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create window in canvas for the blocks
        canvas_window = canvas.create_window((0, 0), window=blocks_frame, anchor="nw")
        
        # Update scrollregion when blocks_frame changes size
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
        
        blocks_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=canvas.winfo_width()))
        
        # Initialize variables
        self.temperature_blocks = []
        self.speed_vars = []
        self.torque_vars = [[] for _ in range(3)]
        self.viscosity_vars = [[] for _ in range(3)]
        
        # Default temperatures
        default_temps = [25, 30, 40, 50]
        
        # Create temperature blocks
        for temp in default_temps:
            self.create_temperature_block(blocks_frame, temp)
        
        # Button frame
        button_frame = Frame(main_frame, bg=APP_BACKGROUND_COLOR)
        button_frame.pack(fill='x', pady=10)
        
        ttk.Button(button_frame, text="Add Block", 
                  command=lambda: self._on_add_temperature_block(blocks_frame)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Save Measurements", 
                  command=self._on_save_measurements).pack(side="right", padx=5)
        
        print("DEBUG: ViscosityView - measure tab created")
    
    def create_temperature_block(self, parent, temperature):
        """Create a temperature measurement block."""
        FONT = ('Arial', 10)
        APP_BACKGROUND_COLOR = 'white'
        
        # Create block frame
        block_frame = Frame(parent, bg=APP_BACKGROUND_COLOR, relief="ridge", bd=2)
        block_frame.pack(fill="x", padx=5, pady=5)
        
        # Temperature header
        temp_label = Label(block_frame, text=f"Temperature: {temperature}°C", 
                          bg=APP_BACKGROUND_COLOR, font=(FONT[0], FONT[1], "bold"))
        temp_label.grid(row=0, column=0, columnspan=8, pady=5)
        
        # Headers
        headers = ["Run", "Speed (RPM)", "Torque (%)", "Viscosity (cP)", 
                  "Run", "Speed (RPM)", "Torque (%)", "Viscosity (cP)"]
        for i, header in enumerate(headers):
            Label(block_frame, text=header, bg=APP_BACKGROUND_COLOR, font=FONT).grid(
                row=1, column=i, padx=2, pady=2)
        
        # Create entry fields for 3 runs (displayed in 2 columns)
        speed_vars = []
        torque_vars = []
        viscosity_vars = []
        
        for run in range(3):
            # Determine row and column offset
            if run < 2:  # First two runs in top row
                row_offset = 2
                col_offset = run * 4
            else:  # Third run in bottom row, centered
                row_offset = 3
                col_offset = 2
            
            # Run number
            Label(block_frame, text=str(run + 1), bg=APP_BACKGROUND_COLOR, font=FONT).grid(
                row=row_offset, column=col_offset, padx=2, pady=2)
            
            # Speed entry
            speed_var = StringVar()
            speed_entry = Entry(block_frame, textvariable=speed_var, width=8)
            speed_entry.grid(row=row_offset, column=col_offset + 1, padx=2, pady=2)
            speed_vars.append(speed_var)
            
            # Torque entry
            torque_var = StringVar()
            torque_entry = Entry(block_frame, textvariable=torque_var, width=8)
            torque_entry.grid(row=row_offset, column=col_offset + 2, padx=2, pady=2)
            torque_vars.append(torque_var)
            
            # Viscosity entry
            viscosity_var = StringVar()
            viscosity_entry = Entry(block_frame, textvariable=viscosity_var, width=8)
            viscosity_entry.grid(row=row_offset, column=col_offset + 3, padx=2, pady=2)
            viscosity_vars.append(viscosity_var)
        
        # Store variables
        self.temperature_blocks.append({
            'temperature': temperature,
            'frame': block_frame,
            'speed_vars': speed_vars,
            'torque_vars': torque_vars,
            'viscosity_vars': viscosity_vars
        })
        
        print(f"DEBUG: ViscosityView - temperature block created for {temperature}°C")
    
    # Event handlers that delegate to controller
    def _on_calculate_terpene(self):
        """Handle calculate terpene percentage action."""
        if self.on_calculate_terpene:
            data = {
                'media': self.media_var.get(),
                'terpene': self.terpene_var.get(),
                'media_brand': self.media_brand_var.get(),
                'terpene_brand': self.terpene_brand_var.get(),
                'sample_mass': self.sample_mass_var.get(),
                'target_viscosity': self.target_viscosity_var.get()
            }
            self.on_calculate_terpene(data)
            print("DEBUG: ViscosityView - calculate terpene triggered")
    
    def _on_calculate_step1(self):
        """Handle calculate step 1 action."""
        if self.on_calculate_step1:
            self.on_calculate_step1()
            print("DEBUG: ViscosityView - calculate step 1 triggered")
    
    def _on_calculate_step2(self):
        """Handle calculate step 2 action."""
        if self.on_calculate_step2:
            self.on_calculate_step2()
            print("DEBUG: ViscosityView - calculate step 2 triggered")
    
    def _on_save_formulation(self):
        """Handle save formulation action."""
        if self.on_save_formulation:
            self.on_save_formulation()
            print("DEBUG: ViscosityView - save formulation triggered")
    
    def _on_upload_training_data(self):
        """Handle upload training data action."""
        if self.on_upload_training_data:
            self.on_upload_training_data()
            print("DEBUG: ViscosityView - upload training data triggered")
    
    def _on_train_standard_models(self):
        """Handle train standard models action."""
        if self.on_train_standard_models:
            self.on_train_standard_models()
            print("DEBUG: ViscosityView - train standard models triggered")
    
    def _on_train_enhanced_models(self):
        """Handle train enhanced models action."""
        if self.on_train_enhanced_models:
            self.on_train_enhanced_models()
            print("DEBUG: ViscosityView - train enhanced models triggered")
    
    def _on_analyze_models(self):
        """Handle analyze models action."""
        if self.on_analyze_models:
            self.on_analyze_models()
            print("DEBUG: ViscosityView - analyze models triggered")
    
    def _on_arrhenius_analysis(self):
        """Handle Arrhenius analysis action."""
        if self.on_arrhenius_analysis:
            self.on_arrhenius_analysis()
            print("DEBUG: ViscosityView - Arrhenius analysis triggered")
    
    def _on_save_measurements(self):
        """Handle save measurements action."""
        if self.on_save_measurements:
            # Collect measurement data from all temperature blocks
            measurements = []
            for block in self.temperature_blocks:
                block_data = {
                    'temperature': block['temperature'],
                    'measurements': []
                }
                for i in range(3):  # 3 runs per temperature
                    run_data = {
                        'run': i + 1,
                        'speed': block['speed_vars'][i].get(),
                        'torque': block['torque_vars'][i].get(),
                        'viscosity': block['viscosity_vars'][i].get()
                    }
                    block_data['measurements'].append(run_data)
                measurements.append(block_data)
            
            self.on_save_measurements(measurements)
            print("DEBUG: ViscosityView - save measurements triggered")
    
    def _on_add_temperature_block(self, blocks_frame):
        """Handle add temperature block action."""
        if self.on_add_temperature_block:
            self.on_add_temperature_block(blocks_frame)
            print("DEBUG: ViscosityView - add temperature block triggered")
    
    def _on_tab_changed(self, event=None):
        """Handle notebook tab change."""
        print("DEBUG: ViscosityView - tab changed")
        # Update advanced tab fields based on calculator tab
        if hasattr(self, 'update_advanced_tab_fields'):
            self.update_advanced_tab_fields()
    
    def _on_help(self):
        """Show help dialog."""
        help_text = (
            "Viscosity Calculator Help\n\n"
            "This application helps you calculate terpene percentages based on viscosity.\n\n"
            "Calculator Tab: Calculate terpene percentage using mathematical models.\n"
            "Iterative Method Tab: Use a two-step iterative process for finding optimal terpene percentage.\n"
            "Measure Tab: Record viscosity measurements at different temperatures.\n\n"
            "Use the Models menu to train machine learning models on your data."
        )
        messagebox.showinfo("Help", help_text)
        print("DEBUG: ViscosityView - help shown")
    
    def _on_about(self):
        """Show about dialog."""
        messagebox.showinfo("About", "Viscosity Calculator\nVersion 1.0\nDeveloped by Charlie Becquet")
        print("DEBUG: ViscosityView - about shown")
    
    def _on_window_close(self):
        """Handle window close event."""
        if self.window:
            self.window.destroy()
            self.window = None
            print("DEBUG: ViscosityView - window closed")
    
    def _center_window(self, window: tk.Toplevel, width: int = None, height: int = None):
        """Center a window on the screen."""
        window.update_idletasks()
        w = width or window.winfo_width()
        h = height or window.winfo_height()
        x = (window.winfo_screenwidth() - w) // 2
        y = (window.winfo_screenheight() - h) // 2
        window.geometry(f"{w}x{h}+{x}+{y}")
        print(f"DEBUG: ViscosityView - window centered at {w}x{h}+{x}+{y}")
    
    # UI Update Methods
    def update_result(self, result: str):
        """Update the calculation result display."""
        self.result_var.set(result)
        print(f"DEBUG: ViscosityView - result updated: {result}")
    
    def update_exact_mass(self, mass: str):
        """Update the exact mass display."""
        self.exact_mass_var.set(mass)
        print(f"DEBUG: ViscosityView - exact mass updated: {mass}")
    
    def update_start_percent(self, percent: str):
        """Update the start percentage display."""
        self.start_percent_var.set(percent)
        print(f"DEBUG: ViscosityView - start percent updated: {percent}")
    
    def update_start_mass(self, mass: str):
        """Update the start mass display."""
        self.start_mass_var.set(mass)
        print(f"DEBUG: ViscosityView - start mass updated: {mass}")
    
    def update_step1_amount(self, amount: str):
        """Update step 1 amount display."""
        self.step1_amount_var.set(amount)
        print(f"DEBUG: ViscosityView - step 1 amount updated: {amount}")
    
    def update_step2_amount(self, amount: str):
        """Update step 2 amount display."""
        self.step2_amount_var.set(amount)
        print(f"DEBUG: ViscosityView - step 2 amount updated: {amount}")
    
    def update_expected_viscosity(self, viscosity: str):
        """Update expected viscosity display."""
        self.expected_viscosity_var.set(viscosity)
        print(f"DEBUG: ViscosityView - expected viscosity updated: {viscosity}")
    
    def get_calculator_data(self) -> Dict[str, str]:
        """Get all data from the calculator tab."""
        return {
            'media': self.media_var.get(),
            'terpene': self.terpene_var.get(),
            'media_brand': self.media_brand_var.get(),
            'terpene_brand': self.terpene_brand_var.get(),
            'sample_mass': self.sample_mass_var.get(),
            'target_viscosity': self.target_viscosity_var.get()
        }
    
    def get_advanced_data(self) -> Dict[str, str]:
        """Get all data from the advanced tab."""
        return {
            'step1_viscosity': self.step1_viscosity_var.get(),
            'step2_viscosity': self.step2_viscosity_var.get()
        }
    
    def get_measurement_data(self) -> List[Dict[str, Any]]:
        """Get all measurement data from the measure tab."""
        measurements = []
        for block in self.temperature_blocks:
            block_data = {
                'temperature': block['temperature'],
                'measurements': []
            }
            for i in range(3):
                run_data = {
                    'run': i + 1,
                    'speed': block['speed_vars'][i].get(),
                    'torque': block['torque_vars'][i].get(),
                    'viscosity': block['viscosity_vars'][i].get()
                }
                block_data['measurements'].append(run_data)
            measurements.append(block_data)
        return measurements
    
    def clear_all_fields(self):
        """Clear all input fields in all tabs."""
        # Calculator tab
        self.media_var.set("")
        self.terpene_var.set("")
        self.media_brand_var.set("")
        self.terpene_brand_var.set("")
        self.sample_mass_var.set("")
        self.target_viscosity_var.set("")
        self.result_var.set("")
        self.exact_mass_var.set("")
        self.start_percent_var.set("")
        self.start_mass_var.set("")
        
        # Advanced tab
        self.step1_amount_var.set("")
        self.step1_viscosity_var.set("")
        self.step2_amount_var.set("")
        self.step2_viscosity_var.set("")
        self.expected_viscosity_var.set("")
        
        # Measurement tab - clear all temperature blocks
        for block in self.temperature_blocks:
            for var_list in [block['speed_vars'], block['torque_vars'], block['viscosity_vars']]:
                for var in var_list:
                    var.set("")
        
        print("DEBUG: ViscosityView - all fields cleared")
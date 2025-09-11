# views/main_window.py
"""
views/main_window.py
Main application window view.
This contains the UI layout from main_gui.py without business logic.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, Callable, List
import queue


class MainWindow:
    """Main application window view."""
    
    def __init__(self, controller):
        """Initialize the main window."""
        self.controller = controller
        self.root = tk.Tk()
        
        # UI Components
        self.top_frame: Optional[ttk.Frame] = None
        self.display_frame: Optional[ttk.Frame] = None
        self.bottom_frame: Optional[ttk.Frame] = None
        self.dynamic_frame: Optional[ttk.Frame] = None
        self.table_frame: Optional[ttk.Frame] = None
        self.plot_frame: Optional[ttk.Frame] = None
        self.image_frame: Optional[ttk.Frame] = None
        self.controls_frame: Optional[ttk.Frame] = None
        
        # UI Variables
        self.selected_sheet = tk.StringVar()
        self.selected_plot_type = tk.StringVar(value="TPM")
        self.file_dropdown_var = tk.StringVar()
        self.crop_enabled = tk.BooleanVar(value=False)
        self.crop_enable = tk.BooleanVar(value=False)  # For backwards compatibility
        
        # Menu and widgets
        self.menubar: Optional[tk.Menu] = None
        self.drop_down_menu: Optional[ttk.Combobox] = None
        self.file_dropdown: Optional[ttk.Combobox] = None
        self.plot_dropdown: Optional[ttk.Combobox] = None
        
        # Report handling
        self.report_thread = None
        self.report_queue = queue.Queue()
        
        # Callback placeholders (set by controller)
        self.on_file_selected: Optional[Callable] = None
        self.on_sheet_selected: Optional[Callable] = None
        self.on_plot_type_changed: Optional[Callable] = None
        self.on_new_file: Optional[Callable] = None
        self.on_load_excel: Optional[Callable] = None
        self.on_load_vap3: Optional[Callable] = None
        self.on_save_vap3: Optional[Callable] = None
        self.on_browse_database: Optional[Callable] = None
        self.on_load_from_database: Optional[Callable] = None
        self.on_viscosity_calculator: Optional[Callable] = None
        self.on_test_report: Optional[Callable] = None
        self.on_full_report: Optional[Callable] = None
        self.on_view_raw_data: Optional[Callable] = None
        self.on_trend_analysis: Optional[Callable] = None
        
        print("DEBUG: MainWindow initialized")
        print(f"DEBUG: Connected to controller: {type(controller).__name__}")
    
    def setup_window(self):
        """Set up the main window properties."""
        self.root.title("DataViewer Application")
        self.root.geometry("1200x800")
        self.root.minsize(1200, 800)
        
        # Maximize the window
        try:
            self.root.state('zoomed')  # Windows
        except:
            try:
                self.root.attributes('-zoomed', True)  # Linux
            except:
                pass  # macOS or fallback
        
        # Set icon if available
        try:
            icon_path = get_resource_path('resources/ccell_icon.png')
            self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
            pass
        except:
            pass
        
        # Configure colors and styles
        self._configure_styles()
        
        # Bind window events
        self.root.bind("<Configure>", self.on_window_resize)
        self.root.protocol("WM_DELETE_WINDOW", self.on_app_close)
        
        print("DEBUG: MainWindow setup completed")
    
    def _configure_styles(self):
        """Configure window styles and colors."""
        # Use constants from utils when available
        APP_BACKGROUND_COLOR = '#EFEFEF'
        BUTTON_COLOR = '#E0E0E0'
        FONT = ("Arial", 10)
        
        style = ttk.Style()
        self.root.configure(bg=APP_BACKGROUND_COLOR)
        style.configure('TLabel', background=APP_BACKGROUND_COLOR, font=FONT)
        style.configure('TButton', background=BUTTON_COLOR, font=FONT, padding=6)
        style.configure('TCombobox', font=FONT)
        style.map('TCombobox', background=[('readonly', APP_BACKGROUND_COLOR)])
        
        # Set colors for all widgets
        for widget in self.root.winfo_children():
            try:
                widget.configure(bg=APP_BACKGROUND_COLOR)
            except Exception:
                continue
        
        print("DEBUG: MainWindow styles configured")
    
    def create_layout(self):
        """Create the main window layout."""
        print("DEBUG: MainWindow creating layout")
        
        # Create main frames
        self._create_frames()
        
        # Create menu
        self._create_menu()
        
        # Create controls
        self._create_controls()
        
        print("DEBUG: MainWindow layout created")
    
    def _create_frames(self):
        """Create the main frame structure."""
        # Top frame for controls
        self.top_frame = ttk.Frame(self.root, height=60)
        self.top_frame.pack(side="top", fill="x", padx=5, pady=5)
        self.top_frame.pack_propagate(False)
        
        # Display frame for main content
        self.display_frame = ttk.Frame(self.root)
        self.display_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)
        
        # Dynamic frame inside display frame
        self.dynamic_frame = ttk.Frame(self.display_frame)
        self.dynamic_frame.pack(fill="both", expand=True)
        
        # Bottom frame for images
        self.bottom_frame = ttk.Frame(self.root, height=150)
        self.bottom_frame.pack(side="bottom", fill="x", padx=5, pady=5)
        self.bottom_frame.pack_propagate(False)
        self.bottom_frame.grid_propagate(False)
        
        # Image frame within bottom frame
        self.image_frame = ttk.Frame(self.bottom_frame, height=150)
        self.image_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.image_frame.pack_propagate(False)
        self.image_frame.grid_propagate(False)
        
        print("DEBUG: MainWindow frames created")
    
    def _create_menu(self):
        """Create the application menu."""
        self.menubar = tk.Menu(self.root)
        
        # File menu
        filemenu = tk.Menu(self.menubar, tearoff=0)
        filemenu.add_command(label="New", command=self._on_new_file)
        filemenu.add_command(label="Load Excel", command=self._on_load_excel)
        filemenu.add_command(label="Load VAP3", command=self._on_load_vap3)
        filemenu.add_separator()
        filemenu.add_command(label="Save As VAP3", command=self._on_save_vap3)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.on_app_close)
        self.menubar.add_cascade(label="File", menu=filemenu)
        
        # View menu
        viewmenu = tk.Menu(self.menubar, tearoff=0)
        viewmenu.add_command(label="View Raw Data", command=self._on_view_raw_data)
        viewmenu.add_command(label="Trend Analysis", command=self._on_trend_analysis)
        self.menubar.add_cascade(label="View", menu=viewmenu)
        
        # Database menu
        dbmenu = tk.Menu(self.menubar, tearoff=0)
        dbmenu.add_command(label="Browse Database", command=self._on_browse_database)
        dbmenu.add_command(label="Load from Database", command=self._on_load_from_database)
        dbmenu.add_command(label="Update Database", command=self._on_update_database)
        self.menubar.add_cascade(label="Database", menu=dbmenu)
        
        # Data Collection menu
        datacollectionmenu = tk.Menu(self.menubar, tearoff=0)
        datacollectionmenu.add_command(label="Collect Data", command=self._on_collect_data)
        datacollectionmenu.add_command(label="Sensory Data Collection", command=self._on_sensory_data_collection)
        datacollectionmenu.add_command(label="Sample Comparison", command=self._on_sample_comparison)
        self.menubar.add_cascade(label="Data Collection", menu=datacollectionmenu)
        
        # Calculate menu
        calculatemenu = tk.Menu(self.menubar, tearoff=0)
        calculatemenu.add_command(label="Viscosity", command=self._on_viscosity_calculator)
        self.menubar.add_cascade(label="Calculate", menu=calculatemenu)
        
        # Reports menu
        reportmenu = tk.Menu(self.menubar, tearoff=0)
        reportmenu.add_command(label="Generate Test Report", command=self._on_test_report)
        reportmenu.add_command(label="Generate Full Report", command=self._on_full_report)
        self.menubar.add_cascade(label="Reports", menu=reportmenu)
        
        # Help menu
        helpmenu = tk.Menu(self.menubar, tearoff=0)
        helpmenu.add_command(label="Help", command=self._on_help)
        helpmenu.add_separator()
        helpmenu.add_command(label="About", command=self._on_about)
        self.menubar.add_cascade(label="Help", menu=helpmenu)
        
        self.root.config(menu=self.menubar)
        print("DEBUG: MainWindow menu created")
    
    def _create_controls(self):
        """Create control widgets in the top frame."""
        if not self.top_frame:
            return
        
        # Left side controls
        left_controls = ttk.Frame(self.top_frame)
        left_controls.pack(side="left", fill="x", expand=True)
        
        # File dropdown
        ttk.Label(left_controls, text="File:").pack(side="left", padx=(5, 2))
        self.file_dropdown = ttk.Combobox(left_controls, textvariable=self.file_dropdown_var, 
                                         state="readonly", width=30)
        self.file_dropdown.pack(side="left", padx=(0, 10))
        self.file_dropdown.bind('<<ComboboxSelected>>', self._on_file_dropdown_change)
        
        # Sheet dropdown
        ttk.Label(left_controls, text="Sheet:").pack(side="left", padx=(5, 2))
        self.drop_down_menu = ttk.Combobox(left_controls, textvariable=self.selected_sheet, 
                                          state="readonly", width=25)
        self.drop_down_menu.pack(side="left", padx=(0, 10))
        self.drop_down_menu.bind('<<ComboboxSelected>>', self._on_sheet_dropdown_change)
        
        # Plot type dropdown (will be shown/hidden based on sheet type)
        ttk.Label(left_controls, text="Plot:").pack(side="left", padx=(5, 2))
        self.plot_dropdown = ttk.Combobox(left_controls, textvariable=self.selected_plot_type,
                                         state="readonly", width=15)
        self.plot_dropdown.pack(side="left", padx=(0, 10))
        self.plot_dropdown.bind('<<ComboboxSelected>>', self._on_plot_type_change)
        
        # Right side controls
        right_controls = ttk.Frame(self.top_frame)
        right_controls.pack(side="right")
        
        # Crop toggle
        crop_checkbox = ttk.Checkbutton(right_controls, text="Enable Image Cropping", 
                                       variable=self.crop_enabled)
        crop_checkbox.pack(side="right", padx=5)
        
        print("DEBUG: MainWindow controls created")
    
    # Event handlers that delegate to controller
    def _on_new_file(self):
        """Handle new file menu action."""
        if self.on_new_file:
            self.on_new_file()
    
    def _on_load_excel(self):
        """Handle load Excel menu action."""
        if self.on_load_excel:
            self.on_load_excel()
    
    def _on_load_vap3(self):
        """Handle load VAP3 menu action."""
        if self.on_load_vap3:
            self.on_load_vap3()
    
    def _on_save_vap3(self):
        """Handle save VAP3 menu action."""
        if self.on_save_vap3:
            self.on_save_vap3()
    
    def _on_browse_database(self):
        """Handle browse database menu action."""
        if self.on_browse_database:
            self.on_browse_database()
    
    def _on_load_from_database(self):
        """Handle load from database menu action."""
        if self.on_load_from_database:
            self.on_load_from_database()
    
    def _on_update_database(self):
        """Handle update database menu action."""
        if hasattr(self, 'on_update_database') and self.on_update_database:
            self.on_update_database()
    
    def _on_collect_data(self):
        """Handle collect data menu action."""
        if hasattr(self, 'on_collect_data') and self.on_collect_data:
            self.on_collect_data()
    
    def _on_sensory_data_collection(self):
        """Handle sensory data collection menu action."""
        if hasattr(self, 'on_sensory_data_collection') and self.on_sensory_data_collection:
            self.on_sensory_data_collection()
    
    def _on_sample_comparison(self):
        """Handle sample comparison menu action."""
        if hasattr(self, 'on_sample_comparison') and self.on_sample_comparison:
            self.on_sample_comparison()
    
    def _on_viscosity_calculator(self):
        """Handle viscosity calculator menu action."""
        if self.on_viscosity_calculator:
            self.on_viscosity_calculator()
    
    def _on_test_report(self):
        """Handle test report menu action."""
        if self.on_test_report:
            self.on_test_report()
    
    def _on_full_report(self):
        """Handle full report menu action."""
        if self.on_full_report:
            self.on_full_report()
    
    def _on_view_raw_data(self):
        """Handle view raw data menu action."""
        if self.on_view_raw_data:
            self.on_view_raw_data()
    
    def _on_trend_analysis(self):
        """Handle trend analysis menu action."""
        if self.on_trend_analysis:
            self.on_trend_analysis()
    
    def _on_help(self):
        """Display help dialog."""
        messagebox.showinfo("Help", 
                           "This program is designed to be used with excel data according to the SDR Standardized Testing Template.\n\n"
                           "Click 'Generate Test Report' to create an excel report of a single test, or click 'Generate Full Report' "
                           "to generate both an excel file and powerpoint file of all the contents within the file.")
    
    def _on_about(self):
        """Display about dialog."""
        messagebox.showinfo("About", "SDR DataViewer Beta Version 3.0\nDeveloped by Charlie Becquet")
    
    def _on_file_dropdown_change(self, event=None):
        """Handle file dropdown selection change."""
        if self.on_file_selected:
            selected_file = self.file_dropdown_var.get()
            self.on_file_selected(selected_file)
            print(f"DEBUG: MainWindow - file dropdown changed to: {selected_file}")
    
    def _on_sheet_dropdown_change(self, event=None):
        """Handle sheet dropdown selection change."""
        if self.on_sheet_selected:
            selected_sheet = self.selected_sheet.get()
            self.on_sheet_selected(selected_sheet)
            print(f"DEBUG: MainWindow - sheet dropdown changed to: {selected_sheet}")
    
    def _on_plot_type_change(self, event=None):
        """Handle plot type dropdown selection change."""
        if self.on_plot_type_changed:
            selected_plot_type = self.selected_plot_type.get()
            self.on_plot_type_changed(selected_plot_type)
            print(f"DEBUG: MainWindow - plot type changed to: {selected_plot_type}")
    
    def on_window_resize(self, event):
        """Handle window resize events."""
        if event.widget == self.root:
            self.constrain_plot_width()
            print("DEBUG: MainWindow - window resize handled")
    
    def on_app_close(self):
        """Handle application close event."""
        try:
            # Cleanup any resources
            if hasattr(self, 'report_thread') and self.report_thread and self.report_thread.is_alive():
                print("DEBUG: MainWindow - waiting for report thread to finish...")
                self.report_thread.join(timeout=2)
            
            print("DEBUG: MainWindow - application closing")
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            print(f"DEBUG: MainWindow - error during close: {e}")
            self.root.quit()
    
    # UI Update Methods
    def show_startup_menu(self):
        """Display a startup menu with 'New' and 'Load' options."""
        startup_menu = tk.Toplevel(self.root)
        startup_menu.title("Welcome")
        startup_menu.geometry("300x150")
        startup_menu.transient(self.root)
        startup_menu.grab_set()
        
        APP_BACKGROUND_COLOR = '#EFEFEF'
        FONT = ("Arial", 10)
        
        startup_menu.configure(bg=APP_BACKGROUND_COLOR)
        
        label = ttk.Label(startup_menu, text="Welcome to DataViewer by SDR!", 
                         font=FONT, background=APP_BACKGROUND_COLOR, foreground="white")
        label.pack(pady=10)
        
        # New Button
        new_button = ttk.Button(startup_menu, text="New", command=lambda: self._handle_startup_new(startup_menu))
        new_button.pack(pady=5)
        
        # Load Button
        load_button = ttk.Button(startup_menu, text="Load", command=lambda: self._handle_startup_load(startup_menu))
        load_button.pack(pady=5)
        
        self.center_window(startup_menu)
        print("DEBUG: MainWindow - startup menu shown")
    
    def _handle_startup_new(self, startup_menu):
        """Handle new button from startup menu."""
        startup_menu.destroy()
        if self.on_new_file:
            self.on_new_file()
    
    def _handle_startup_load(self, startup_menu):
        """Handle load button from startup menu."""
        startup_menu.destroy()
        if self.on_load_excel:
            self.on_load_excel()
    
    def center_window(self, window: tk.Toplevel, width: Optional[int] = None, height: Optional[int] = None):
        """Center a given Tkinter window on the screen."""
        window.update_idletasks()
        
        window_width = width or window.winfo_width()
        window_height = height or window.winfo_height()
        
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        position_x = (screen_width - window_width) // 2
        position_y = (screen_height - window_height) // 2
        
        window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
        print(f"DEBUG: MainWindow - centered window: {window_width}x{window_height}+{position_x}+{position_y}")
    
    def update_file_dropdown(self, files: List[str]):
        """Update the file dropdown with available files."""
        if self.file_dropdown:
            current_selection = self.file_dropdown_var.get()
            self.file_dropdown['values'] = files
            
            # Restore selection if still valid
            if current_selection in files:
                self.file_dropdown_var.set(current_selection)
            elif files:
                self.file_dropdown_var.set(files[0])
            else:
                self.file_dropdown_var.set("")
            
            print(f"DEBUG: MainWindow - file dropdown updated with {len(files)} files")
    
    def update_sheet_dropdown(self, sheets: List[str]):
        """Update the sheet dropdown with available sheets."""
        if self.drop_down_menu:
            current_selection = self.selected_sheet.get()
            self.drop_down_menu['values'] = sheets
            
            # Restore selection if still valid
            if current_selection in sheets:
                self.selected_sheet.set(current_selection)
            elif sheets:
                self.selected_sheet.set(sheets[0])
            else:
                self.selected_sheet.set("")
            
            print(f"DEBUG: MainWindow - sheet dropdown updated with {len(sheets)} sheets")
    
    def update_plot_dropdown(self, plot_options: List[str], show: bool = True):
        """Update the plot dropdown with available plot options."""
        if self.plot_dropdown:
            if show and plot_options:
                self.plot_dropdown['values'] = plot_options
                current_selection = self.selected_plot_type.get()
                
                # Restore selection if still valid, otherwise use first option
                if current_selection not in plot_options and plot_options:
                    self.selected_plot_type.set(plot_options[0])
                
                self.plot_dropdown.pack(side="left", padx=(0, 10))
                print(f"DEBUG: MainWindow - plot dropdown shown with {len(plot_options)} options")
            else:
                self.plot_dropdown.pack_forget()
                print("DEBUG: MainWindow - plot dropdown hidden")
    
    def clear_dynamic_frame(self):
        """Clear all children widgets from the dynamic frame."""
        if self.dynamic_frame:
            for widget in self.dynamic_frame.winfo_children():
                widget.destroy()
            print("DEBUG: MainWindow - dynamic frame cleared")
    
    def setup_dynamic_frames(self, is_plotting_sheet: bool = False):
        """Create frames inside the dynamic_frame based on sheet type."""
        if not self.dynamic_frame:
            return
        
        # Clear previous widgets
        self.clear_dynamic_frame()
        
        # Get dynamic frame height
        window_height = self.root.winfo_height()
        top_height = self.top_frame.winfo_height() if self.top_frame else 60
        bottom_height = self.bottom_frame.winfo_height() if self.bottom_frame else 150
        padding = 20
        display_height = window_height - top_height - bottom_height - padding
        display_height = max(display_height, 100)
        
        if is_plotting_sheet:
            # Use grid layout for precise control of width proportions
            self.dynamic_frame.columnconfigure(0, weight=5)  # Table column
            self.dynamic_frame.columnconfigure(1, weight=5)  # Plot column
            self.dynamic_frame.rowconfigure(0, weight=1)
            
            # Table takes exactly 50% width
            self.table_frame = ttk.Frame(self.dynamic_frame)
            self.table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
            
            # Plot takes remaining 50% width
            self.plot_frame = ttk.Frame(self.dynamic_frame)
            self.plot_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)
            
            self.constrain_plot_width()
            print("DEBUG: MainWindow - plotting sheet layout created")
        else:
            # Non-plotting sheets use the full space
            self.table_frame = ttk.Frame(self.dynamic_frame, height=display_height)
            self.table_frame.pack(fill="both", expand=True, padx=5, pady=5)
            self.plot_frame = None
            print("DEBUG: MainWindow - non-plotting sheet layout created")
    
    def constrain_plot_width(self):
        """Ensure the plot doesn't exceed 50% of the window width."""
        if not hasattr(self, 'plot_frame') or not self.plot_frame:
            return
        
        window_width = self.root.winfo_width()
        max_plot_width = window_width // 2
        
        if max_plot_width > 50:
            self.plot_frame.config(width=max_plot_width)
            self.plot_frame.grid_propagate(False)
            
            if hasattr(self, 'table_frame') and self.table_frame:
                self.table_frame.config(width=max_plot_width)
                self.table_frame.grid_propagate(False)
        
        print(f"DEBUG: MainWindow - plot width constrained to {max_plot_width}px")
    
    def get_table_frame(self) -> Optional[ttk.Frame]:
        """Get the table frame for displaying data tables."""
        return self.table_frame
    
    def get_plot_frame(self) -> Optional[ttk.Frame]:
        """Get the plot frame for displaying plots."""
        return self.plot_frame
    
    def get_image_frame(self) -> Optional[ttk.Frame]:
        """Get the image frame for displaying images."""
        return self.image_frame
    
    def run(self):
        """Start the main window event loop."""
        print("DEBUG: MainWindow - starting main loop")
        self.root.mainloop()
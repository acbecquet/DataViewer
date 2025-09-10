# views/main_window.py
"""
views/main_window.py
Main application window view.
This will contain the UI layout from main_gui.py without business logic.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, Callable
from controllers.main_controller import MainController


class MainWindow:
    """Main application window view."""
    
    def __init__(self, controller: MainController):
        """Initialize the main window."""
        self.controller = controller
        self.root = tk.Tk()
        
        # UI Components
        self.top_frame: Optional[ttk.Frame] = None
        self.display_frame: Optional[ttk.Frame] = None
        self.bottom_frame: Optional[ttk.Frame] = None
        self.table_frame: Optional[ttk.Frame] = None
        self.plot_frame: Optional[ttk.Frame] = None
        self.image_frame: Optional[ttk.Frame] = None
        self.controls_frame: Optional[ttk.Frame] = None
        
        # UI Variables
        self.selected_sheet = tk.StringVar()
        self.selected_plot_type = tk.StringVar(value="TPM")
        self.file_dropdown_var = tk.StringVar()
        self.crop_enable = tk.BooleanVar(value=False)
        
        # Menu and widgets
        self.menubar: Optional[tk.Menu] = None
        self.drop_down_menu: Optional[ttk.Combobox] = None
        self.file_dropdown: Optional[ttk.Combobox] = None
        
        # Event callbacks (set by controller)
        self.on_file_selected: Optional[Callable] = None
        self.on_sheet_selected: Optional[Callable] = None
        self.on_plot_type_changed: Optional[Callable] = None
        
        print("DEBUG: MainWindow initialized")
        print(f"DEBUG: Connected to MainController")
    
    def setup_window(self):
        """Set up the main window properties."""
        self.root.title("DataViewer Application")
        self.root.geometry("1200x800")
        self.root.minsize(1200, 800)
        
        # Set icon if available
        try:
            # icon_path = get_resource_path('resources/ccell_icon.png')
            # self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
            pass
        except:
            pass
        
        # Configure colors and styles
        self._configure_styles()
        
        # Center window on screen
        self._center_window()
        
        print("DEBUG: MainWindow setup completed")
    
    def _configure_styles(self):
        """Configure window styles and colors."""
        # Placeholder for styling - would use utils.constants
        APP_BACKGROUND_COLOR = '#EFEFEF'
        BUTTON_COLOR = '#E0E0E0'
        
        style = ttk.Style()
        self.root.configure(bg=APP_BACKGROUND_COLOR)
        style.configure('TLabel', background=APP_BACKGROUND_COLOR)
        style.configure('TButton', background=BUTTON_COLOR, padding=6)
        style.configure('TCombobox', background=APP_BACKGROUND_COLOR)
        
        print("DEBUG: MainWindow styles configured")
    
    def _center_window(self):
        """Center the window on screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
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
        self.display_frame.pack(side="top", fill="both", expand=True, padx=5, pady=5)
        
        # Bottom frame for images
        self.bottom_frame = ttk.Frame(self.root, height=200)
        self.bottom_frame.pack(side="bottom", fill="x", padx=5, pady=5)
        self.bottom_frame.pack_propagate(False)
        
        # Subframes within display frame
        self.table_frame = ttk.Frame(self.display_frame)
        self.table_frame.pack(side="left", fill="both", expand=True)
        
        self.plot_frame = ttk.Frame(self.display_frame)
        self.plot_frame.pack(side="right", fill="both", expand=True)
        
        # Image frame within bottom frame
        self.image_frame = ttk.Frame(self.bottom_frame)
        self.image_frame.pack(side="left", fill="both", expand=True)
        
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
        filemenu.add_command(label="Exit", command=self._on_exit)
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
        self.menubar.add_cascade(label="Database", menu=dbmenu)
        
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
        
        # File dropdown
        file_label = ttk.Label(self.top_frame, text="File:")
        file_label.pack(side="left", padx=(5, 2))
        
        self.file_dropdown = ttk.Combobox(
            self.top_frame,
            textvariable=self.file_dropdown_var,
            state="readonly",
            width=30
        )
        self.file_dropdown.pack(side="left", padx=(0, 10))
        self.file_dropdown.bind("<<ComboboxSelected>>", self._on_file_dropdown_changed)
        
        # Sheet dropdown
        sheet_label = ttk.Label(self.top_frame, text="Test:")
        sheet_label.pack(side="left", padx=(5, 2))
        
        self.drop_down_menu = ttk.Combobox(
            self.top_frame,
            textvariable=self.selected_sheet,
            state="readonly",
            width=25
        )
        self.drop_down_menu.pack(side="left", padx=(0, 10))
        self.drop_down_menu.bind("<<ComboboxSelected>>", self._on_sheet_dropdown_changed)
        
        # Plot type dropdown
        plot_label = ttk.Label(self.top_frame, text="Plot:")
        plot_label.pack(side="left", padx=(5, 2))
        
        plot_dropdown = ttk.Combobox(
            self.top_frame,
            textvariable=self.selected_plot_type,
            state="readonly",
            values=["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"],
            width=20
        )
        plot_dropdown.pack(side="left", padx=(0, 10))
        plot_dropdown.bind("<<ComboboxSelected>>", self._on_plot_type_dropdown_changed)
        
        # Control buttons frame
        self.controls_frame = ttk.Frame(self.top_frame)
        self.controls_frame.pack(side="right", fill="x", padx=5)
        
        print("DEBUG: MainWindow controls created")
    
    # Menu callback stubs (delegate to controller)
    def _on_new_file(self):
        """Handle new file menu action."""
        print("DEBUG: MainWindow - New file requested")
        if self.controller:
            file_controller = self.controller.get_file_controller()
            if file_controller:
                # file_controller.create_new_file()
                pass
    
    def _on_load_excel(self):
        """Handle load Excel menu action."""
        print("DEBUG: MainWindow - Load Excel requested")
        if self.controller:
            file_controller = self.controller.get_file_controller()
            if file_controller:
                # file_controller.load_excel_file()
                pass
    
    def _on_load_vap3(self):
        """Handle load VAP3 menu action."""
        print("DEBUG: MainWindow - Load VAP3 requested")
        if self.controller:
            file_controller = self.controller.get_file_controller()
            if file_controller:
                # file_controller.load_vap3_file()
                pass
    
    def _on_save_vap3(self):
        """Handle save VAP3 menu action."""
        print("DEBUG: MainWindow - Save VAP3 requested")
        if self.controller:
            file_controller = self.controller.get_file_controller()
            if file_controller:
                # file_controller.save_vap3_file()
                pass
    
    def _on_exit(self):
        """Handle exit menu action."""
        print("DEBUG: MainWindow - Exit requested")
        if self.controller:
            self.controller.shutdown_application()
        self.root.quit()
    
    def _on_view_raw_data(self):
        """Handle view raw data menu action."""
        print("DEBUG: MainWindow - View raw data requested")
    
    def _on_trend_analysis(self):
        """Handle trend analysis menu action."""
        print("DEBUG: MainWindow - Trend analysis requested")
    
    def _on_browse_database(self):
        """Handle browse database menu action."""
        print("DEBUG: MainWindow - Browse database requested")
    
    def _on_load_from_database(self):
        """Handle load from database menu action."""
        print("DEBUG: MainWindow - Load from database requested")
    
    def _on_viscosity_calculator(self):
        """Handle viscosity calculator menu action."""
        print("DEBUG: MainWindow - Viscosity calculator requested")
        if self.controller:
            calc_controller = self.controller.get_calculation_controller()
            if calc_controller:
                # calc_controller.open_viscosity_calculator()
                pass
    
    def _on_test_report(self):
        """Handle test report menu action."""
        print("DEBUG: MainWindow - Test report requested")
        if self.controller:
            report_controller = self.controller.get_report_controller()
            if report_controller:
                # report_controller.generate_test_report()
                pass
    
    def _on_full_report(self):
        """Handle full report menu action."""
        print("DEBUG: MainWindow - Full report requested")
        if self.controller:
            report_controller = self.controller.get_report_controller()
            if report_controller:
                # report_controller.generate_full_report()
                pass
    
    def _on_help(self):
        """Handle help menu action."""
        messagebox.showinfo("Help", "DataViewer Application Help\n\nFor assistance, refer to the user manual.")
    
    def _on_about(self):
        """Handle about menu action."""
        messagebox.showinfo("About", "DataViewer Application v3.0\nDeveloped by Charlie Becquet")
    
    # Control callback stubs
    def _on_file_dropdown_changed(self, event):
        """Handle file dropdown selection."""
        print(f"DEBUG: MainWindow - File selected: {self.file_dropdown_var.get()}")
        if self.on_file_selected:
            self.on_file_selected(self.file_dropdown_var.get())
    
    def _on_sheet_dropdown_changed(self, event):
        """Handle sheet dropdown selection."""
        print(f"DEBUG: MainWindow - Sheet selected: {self.selected_sheet.get()}")
        if self.on_sheet_selected:
            self.on_sheet_selected(self.selected_sheet.get())
    
    def _on_plot_type_dropdown_changed(self, event):
        """Handle plot type dropdown selection."""
        print(f"DEBUG: MainWindow - Plot type selected: {self.selected_plot_type.get()}")
        if self.on_plot_type_changed:
            self.on_plot_type_changed(self.selected_plot_type.get())
    
    # Public methods for controller to call
    def update_file_dropdown(self, files: list):
        """Update the file dropdown with new files."""
        if self.file_dropdown:
            self.file_dropdown['values'] = files
            print(f"DEBUG: MainWindow updated file dropdown with {len(files)} files")
    
    def update_sheet_dropdown(self, sheets: list):
        """Update the sheet dropdown with new sheets."""
        if self.drop_down_menu:
            self.drop_down_menu['values'] = sheets
            print(f"DEBUG: MainWindow updated sheet dropdown with {len(sheets)} sheets")
    
    def set_callbacks(self, file_callback: Callable, sheet_callback: Callable, plot_callback: Callable):
        """Set callback functions for UI events."""
        self.on_file_selected = file_callback
        self.on_sheet_selected = sheet_callback
        self.on_plot_type_changed = plot_callback
        print("DEBUG: MainWindow callbacks set")
    
    def show_message(self, title: str, message: str, message_type: str = "info"):
        """Show a message dialog."""
        if message_type == "error":
            messagebox.showerror(title, message)
        elif message_type == "warning":
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)
    
    def run(self):
        """Start the GUI main loop."""
        print("DEBUG: MainWindow starting main loop")
        self.root.mainloop()
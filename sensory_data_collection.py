"""
sensory_data_collection.py
Sensory Data Collection Window for DataViewer Application
Developed by Charlie Becquet
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
import math
from datetime import datetime
import json
import os
from utils import APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT, debug_print
from claude_form_processor import ClaudeFormProcessor

def _lazy_import_cv2():
    """Lazy import opencv."""
    try:
        import cv2
        print("TIMING: Lazy loaded cv2 for image processing")
        return cv2
    except ImportError as e:
        print(f"Error importing cv2: {e}")
        messagebox.showerror("Missing Dependency", 
                            "OpenCV is required for image processing.\nPlease install: pip install opencv-python")
        return None

def _lazy_import_pil():
    """Lazy import PIL."""
    try:
        from PIL import Image, ImageTk
        print("TIMING: Lazy loaded PIL for image processing")
        return Image, ImageTk
    except ImportError as e:
        print(f"Error importing PIL: {e}")
        messagebox.showerror("Missing Dependency", 
                            "PIL is required for image processing.\nPlease install: pip install Pillow")
        return None, None

def _lazy_import_sklearn():
    """Lazy import scikit-learn."""
    try:
        from sklearn.cluster import DBSCAN
        print("TIMING: Lazy loaded sklearn for image processing")
        return DBSCAN
    except ImportError as e:
        print(f"Error importing sklearn: {e}")
        messagebox.showerror("Missing Dependency", 
                            "Scikit-learn is required for advanced image processing.\nPlease install: pip install scikit-learn")
        return None

def _lazy_import_pytesseract():
    """Lazy import pytesseract."""
    try:
        import pytesseract
        print("TIMING: Lazy loaded pytesseract for OCR")
        return pytesseract
    except ImportError as e:
        print(f"Error importing pytesseract: {e}")
        messagebox.showerror("Missing Dependency", 
                            "Pytesseract is required for text recognition.\nPlease install: pip install pytesseract")
        return None


class SensoryDataCollectionWindow:
    """Main window for sensory data collection and visualization."""
    
    def __init__(self, parent):
        self.parent = parent
        self.window = None
        self.data = {}
        self.samples = {}
        self.current_sample = None
        
        # Sensory metrics (5 attributes)
        self.metrics = [
            "Burnt Taste",
            "Vapor Volume", 
            "Overall Flavor",
            "Smoothness",
            "Overall Liking"
        ]

        self.current_mode = "collection"
        self.all_sessions_data = {}
        self.average_samples = {}
        debug_print("Initialized dual-mode functionality")
        
        # Header data fields
        self.header_fields = [
            "Assessor Name", 
            "Media",
            "Puff Length",
            "Date"
        ]
        debug_print("Updated header fields to 4 essential fields only")
        debug_print(f"New header fields: {self.header_fields}")
        # SOP text
        self.sop_text = """
SENSORY EVALUATION STANDARD OPERATING PROCEDURE

1. PREPARATION:
   - Ensure all cartridges are at room temperature
   - Use clean, odor-free environment, ideally in a fume hood

2. EVALUATION:
   - Take 2-3 moderate puffs per sample
   - Be sure to compare flavor with original oil if available
   - Rate each attribute on 1-9 scale (1=poor, 9=excellent)

3. SCALE INTERPRETATION:
   - Burnt Taste: 1=Very Burnt, 9=No Burnt Taste
   - Vapor Volume: 1=Very Low, 9=Very High
   - Overall Flavor: 1=Poor, 9=Excellent
   - Smoothness: 1=Very Harsh, 9=Very Smooth
   - Overall Liking: 1=Dislike Extremely, 9=Like Extremely

4. NOTES:
   - Record any unusual observations
   - Note any technical issues with samples
        """
    # Lazy import functions for image processing
    

    def show(self):
        """Create and display the sensory data collection window."""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Sensory Data Collection")
        self.window.resizable(True,True)
        self.window.minsize(1000,400)

        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        # CHANGED: Make window 75% of original size (0.9 * 0.75 = 0.675 of screen)
        window_width = min(int(screen_width * 0.675), 1350)  # 67.5% of screen, max 1350 (75% of 1800)
        window_height = min(int(screen_height * 0.675), 900)   # 67.5% of screen, max 900 (75% of 1200)

        
        self.window.geometry(f"{window_width}x{window_height}")
        self.window.configure(bg=APP_BACKGROUND_COLOR)
        self.window.transient(self.parent)
        
        # Create main layout
        self.setup_layout()
        self.setup_menu()
        
        # Center the window
        self.center_window()
        
    def center_window(self):
        """Center the window on screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_menu(self):
        """Create the menu bar."""
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Session", command=self.new_session)
        file_menu.add_command(label="Load Session", command=self.load_session)
        file_menu.add_command(label="Save Session", command=self.save_session)
        file_menu.add_separator()
        file_menu.add_command(label="Load from Image (ML)", command=self.load_from_image)
        file_menu.add_command(label="Load with AI (Claude)", command=self.load_from_image_with_ai)
        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Close", command=self.window.destroy)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Sample menu
        sample_menu = tk.Menu(menubar, tearoff=0)
        sample_menu.add_command(label="Add Sample", command=self.add_sample)
        sample_menu.add_command(label="Remove Sample", command=self.remove_sample)
        sample_menu.add_command(label="Clear Sample Data", command=self.clear_current_sample)
        sample_menu.add_separator()  # NEW
        sample_menu.add_command(label="Rename Current Sample", command=self.rename_current_sample)  # NEW
        sample_menu.add_command(label="Batch Rename Samples", command=self.batch_rename_samples)  # NEW
        menubar.add_cascade(label="Sample", menu=sample_menu)

        # Export menu (NEW)
        export_menu = tk.Menu(menubar, tearoff=0)
        export_menu.add_command(label="Save Plot as Image", command=self.save_plot_as_image)
        export_menu.add_command(label="Save Table as Image", command=self.save_table_as_image)
        export_menu.add_separator()
        export_menu.add_command(label="Generate PowerPoint Report", command=self.generate_powerpoint_report)
        menubar.add_cascade(label="Export", menu=export_menu)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Show SOP", command=self.show_sop)
        view_menu.add_command(label="Refresh Plot", command=self.update_plot)
        menubar.add_cascade(label="View", menu=view_menu)
        
        # ML menu (NEW)
        ml_menu = tk.Menu(menubar, tearoff=0)
        ml_menu.add_command(label="Create Training Structure", command=self.create_training_structure)
        ml_menu.add_command(label="Extract Training Data", command=self.extract_training_data)
        ml_menu.add_command(label="Check Data Balance", command=self.check_training_data_balance)
        ml_menu.add_separator()
        ml_menu.add_command(label="Train Model", command=self.train_ml_model)
        menubar.add_cascade(label="ML", menu=ml_menu)

    def setup_layout(self):
        """Create the main layout with scrollable panels."""
        print("DEBUG: Setting up main layout with scrollable areas")
    
        # Create main paned window (horizontal split)
        main_paned = ttk.PanedWindow(self.window, orient='horizontal')
        main_paned.pack(fill='both', expand=True, padx=5, pady=5)
    
        # Left panel with scrollable frame for data entry (40% width)
        left_canvas = tk.Canvas(main_paned, bg=APP_BACKGROUND_COLOR)
        left_scrollbar = ttk.Scrollbar(main_paned, orient="vertical", command=left_canvas.yview)
        self.left_frame = ttk.Frame(left_canvas)
    
        self.left_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
    
        left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
    
        # Add mouse wheel scrolling
        def _on_mousewheel_left(event):
            left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        left_canvas.bind("<MouseWheel>", _on_mousewheel_left)
    
        main_paned.add(left_canvas, weight=40)
    
        # Right panel for plot (60% width) - also scrollable
        right_canvas = tk.Canvas(main_paned, bg=APP_BACKGROUND_COLOR)
        right_scrollbar = ttk.Scrollbar(main_paned, orient="vertical", command=right_canvas.yview)
        self.right_frame = ttk.Frame(right_canvas)
    
        self.right_frame.bind(
            "<Configure>",
            lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all"))
        )
    
        right_canvas.create_window((0, 0), window=self.right_frame, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)
    
        # Add mouse wheel scrolling
        def _on_mousewheel_right(event):
            right_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        right_canvas.bind("<MouseWheel>", _on_mousewheel_right)
    
        main_paned.add(right_canvas, weight=60)
    
        print("DEBUG: Scrollable layout configured")
    
        # Setup each panel
        self.setup_data_entry_panel()
        self.setup_plot_panel()
        
    def setup_data_entry_panel(self):
        """Setup the left panel for data entry."""
        # Header section
        header_frame = ttk.LabelFrame(self.left_frame, text="Session Information", padding=10)
        header_frame.pack(fill='x', padx=5, pady=5)
        
        # Configure header_frame for centering
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1) 
        header_frame.grid_columnconfigure(2, weight=1)
        print("DEBUG: Configured header_frame for centered layout")
        
        self.header_vars = {}
        row = 0
        for field in self.header_fields:
            ttk.Label(header_frame, text=f"{field}:", font=FONT).grid(
                row=row, column=0, sticky='e', padx=5, pady=2)  # CHANGED: sticky='e' for right alignment
            debug_print(f"DEBUG: Added centered label for field: {field}")
            
            if field == "Date":
                var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
            elif field == "Gender":
                var = tk.StringVar()
                combo = ttk.Combobox(header_frame, textvariable=var, 
                                   values=["Male", "Female", "Other", "Prefer not to say"],
                                   state="readonly", width=20)
                combo.grid(row=row, column=1, sticky='w', padx=5, pady=2)
                self.header_vars[field] = var
                row += 1
                continue
            else:
                var = tk.StringVar()
                
            entry = ttk.Entry(header_frame, textvariable=var, font=FONT, width=20)
            entry.grid(row=row, column=1, sticky='w', padx=5, pady=2)
            self.header_vars[field] = var
            row += 1
            
        # ADDED: Mode switch button
        mode_button_frame = ttk.Frame(header_frame)
        mode_button_frame.grid(row=row, column=0, columnspan=2, pady=10)
        
        self.mode_button = ttk.Button(mode_button_frame, text="Switch to Comparison Mode", 
                                     command=self.toggle_mode, width=25)
        self.mode_button.pack()
        debug_print("Added mode switch button to header section")

        # Sample management section - SIMPLE NATURAL CENTERING
        sample_frame = ttk.LabelFrame(self.left_frame, text="Sample Management", padding=10)
        sample_frame.pack(fill='x', padx=5, pady=5)
        
        debug_print("Setting up sample management with simple centering")
        
        # ROW 1: Sample selection - simple center using expand
        sample_select_outer = ttk.Frame(sample_frame)
        sample_select_outer.pack(fill='x', pady=5)
        
        sample_select_frame = ttk.Frame(sample_select_outer)
        sample_select_frame.pack(expand=True)  # This centers it naturally
        
        ttk.Label(sample_select_frame, text="Current Sample:", font=FONT).pack(side='left')
        self.sample_var = tk.StringVar()
        self.sample_combo = ttk.Combobox(sample_select_frame, textvariable=self.sample_var,
                                        state="readonly", width=15)
        self.sample_combo.pack(side='left', padx=5)
        self.sample_combo.bind('<<ComboboxSelected>>', self.on_sample_changed)
        
        debug_print("Sample selection centered with expand=True")
        
        # ROW 2: Buttons - simple center using expand
        button_outer = ttk.Frame(sample_frame)
        button_outer.pack(fill='x', pady=5)
        
        button_frame = ttk.Frame(button_outer)
        button_frame.pack(expand=True)  # This centers it naturally
        
        ttk.Button(button_frame, text="Add Sample", 
                  command=self.add_sample).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Remove Sample", 
                  command=self.remove_sample).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Clear Data", 
                  command=self.clear_current_sample).pack(side='left', padx=2)
        
        debug_print("Buttons centered with expand=True")
                  
        # Sensory evaluation section
        eval_frame = ttk.LabelFrame(self.left_frame, text="Sensory Evaluation", padding=10)
        eval_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create rating scales for each metric
        self.rating_vars = {}
        for i, metric in enumerate(self.metrics):
            metric_frame = ttk.Frame(eval_frame)
            metric_frame.pack(fill='x', pady=4)
            
            # Metric label
            ttk.Label(metric_frame, text=f"{metric}:", font=FONT, width=15).pack(side='left')
            
            # Rating scale (1-9)
            self.rating_vars[metric] = tk.IntVar(value=5)
            scale = tk.Scale(metric_frame, from_=1, to=9, orient='horizontal',
                           variable=self.rating_vars[metric], font=FONT)
            scale.pack(side='left', fill='x', expand=True, padx=5)
            
            # Current value display
            value_label = ttk.Label(metric_frame, text="5", width=2)
            value_label.pack(side='right')
            
            # Update value display AND plot when scale changes (LIVE UPDATES)
            def update_live(val, label=value_label, var=self.rating_vars[metric], metric_name=metric):
                label.config(text=str(var.get()))
                self.auto_save_and_update()  # Add this method call
            scale.config(command=update_live)
            
        # Comments section
        comments_frame = ttk.Frame(eval_frame)
        comments_frame.pack(fill='x', pady=10)
        
        ttk.Label(comments_frame, text="Additional Comments:", font=FONT).pack(anchor='w')
        self.comments_text = tk.Text(comments_frame, height=4, font=FONT)
        self.comments_text.pack(fill='x', pady=2)
        
       # Auto-save comments when user types
        def on_comment_change(event=None):
            """Auto-save comments when user types."""
            current_sample = self.sample_var.get()
            if current_sample and current_sample in self.samples:
                comments = self.comments_text.get('1.0', tk.END).strip()
                self.samples[current_sample]['comments'] = comments
                print(f"DEBUG: Auto-saved comments for {current_sample}: '{comments[:50]}...'")

        # Bind to text change events
        self.comments_text.bind('<KeyRelease>', on_comment_change)
        self.comments_text.bind('<FocusOut>', on_comment_change)
        self.comments_text.bind('<Button-1>', lambda e: self.window.after(100, on_comment_change))
                  
    def setup_plot_panel(self):
        """Setup the right panel for spider plot visualization."""
        plot_frame = ttk.LabelFrame(self.right_frame, text="Sensory Profile Comparison", padding=10)
        plot_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Sample selection for plotting
        control_frame = ttk.Frame(plot_frame)
        control_frame.pack(fill='x', pady=5)
        
        ttk.Label(control_frame, text="Select Samples to Display:", font=FONT).pack(anchor='w')
        
        # Checkboxes frame (will be populated when samples are added)
        self.checkbox_frame = ttk.Frame(control_frame)
        self.checkbox_frame.pack(fill='x', pady=5)
        
        self.sample_checkboxes = {}
        
        # Plot canvas
        self.setup_plot_canvas(plot_frame)
        
    def setup_plot_canvas(self, parent):
        """Create the matplotlib canvas for the spider plot with better resizing."""
        print("DEBUG: Setting up enhanced plot canvas")
    
        # Create figure with responsive sizing - REDUCED to 75% of original size
        self.fig, self.ax = plt.subplots(figsize=(7.5, 6), subplot_kw=dict(projection='polar'))  # CHANGED: from (10, 8) to (7.5, 6)
        self.fig.patch.set_facecolor('white')
        debug_print("Created spider plot with 75% size: 7.5x6 inches (was 10x8)")
    
        self.fig.subplots_adjust(left=0.0, right=0.85, top=0.85, bottom=0.08)
        debug_print("Applied subplot adjustments for smaller plot")
            
        # Create canvas with scrollbars
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill='both', expand=True)
    
        # Create canvas
        self.canvas = FigureCanvasTkAgg(self.fig, canvas_frame)
        canvas_widget = self.canvas.get_tk_widget()
        canvas_widget.pack(fill='both', expand=True)
    
        # Add toolbar for additional functionality
        self.setup_plot_context_menu(canvas_widget)
        debug_print("Plot canvas setup complete with right-click context menu")
    
        print("DEBUG: Plot canvas setup complete with toolbar")
    
        # Initialize empty plot
        self.update_plot()
        
    def setup_plot_context_menu(self, canvas_widget):
        """Set up right-click context menu for plot with essential functionality."""
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
    
        # Create a hidden toolbar to access the navigation functions
        self.hidden_toolbar = NavigationToolbar2Tk(self.canvas, canvas_widget.master)
        self.hidden_toolbar.pack_forget()  # Hide the toolbar but keep functionality
    
        def show_context_menu(event):
            """Show simplified right-click context menu with essential plot options."""
            context_menu = tk.Menu(self.window, tearoff=0)
        
            # Essential options only
            context_menu.add_command(label="⚙️ Configure Plot", 
                                   command=self.hidden_toolbar.configure_subplots)
            context_menu.add_separator()
            context_menu.add_command(label="💾 Save Plot...", 
                                   command=self.hidden_toolbar.save_figure)
        
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()
    
        # Bind right-click to show context menu
        canvas_widget.bind("<Button-3>", show_context_menu)  # Right-click on Windows/Linux
        canvas_widget.bind("<Button-2>", show_context_menu)  # Right-click on Mac (sometimes)
    
        debug_print("Simplified right-click context menu set up for plot canvas")
        
    def toggle_mode(self):
        """Toggle between collection mode and comparison mode."""
        if self.current_mode == "collection":
            # Switch to comparison mode
            self.switch_to_comparison_mode()
        else:
            # Switch to collection mode
            self.switch_to_collection_mode()
            
    def switch_to_comparison_mode(self):
        """Switch to comparison mode - show averages across users."""
        self.current_mode = "comparison"
        self.mode_button.config(text="Switch to Collection Mode")
        
        # Gray out sensory evaluation panel
        self.disable_sensory_evaluation()
        
        # Load multiple sessions if not already loaded
        if not self.all_sessions_data:
            self.load_multiple_sessions()
        
        # Calculate averages
        self.calculate_sample_averages()
        
        # Update plot with averages
        self.update_comparison_plot()
        
        debug_print("Switched to comparison mode - showing averaged data across users")
        messagebox.showinfo("Comparison Mode", "Now showing averaged data across multiple users.\nSensory evaluation is disabled in this mode.")
        
    def switch_to_collection_mode(self):
        """Switch to collection mode - normal single user operation."""
        self.current_mode = "collection" 
        self.mode_button.config(text="Switch to Comparison Mode")
        
        # Re-enable sensory evaluation panel
        self.enable_sensory_evaluation()
        
        # Update plot with current user's data
        self.update_plot()
        
        debug_print("Switched to collection mode - showing single user data")
        messagebox.showinfo("Collection Mode", "Now showing single user data collection mode.\nSensory evaluation is enabled.")

    def disable_sensory_evaluation(self):
        """Gray out and disable all sensory evaluation controls."""
        # Find the sensory evaluation frame and disable all children
        for widget in self.left_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and widget.cget('text') == 'Sensory Evaluation':
                self.set_widget_state(widget, 'disabled')
                widget.configure(style='Disabled.TLabelframe')
        debug_print("Disabled sensory evaluation panel for comparison mode")
        
    def enable_sensory_evaluation(self):
        """Re-enable all sensory evaluation controls."""
        # Find the sensory evaluation frame and enable all children
        for widget in self.left_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and widget.cget('text') == 'Sensory Evaluation':
                self.set_widget_state(widget, 'normal')
                widget.configure(style='TLabelframe')
        debug_print("Enabled sensory evaluation panel for collection mode")

    def set_widget_state(self, parent, state):
        """Recursively set state of all child widgets."""
        try:
            parent.configure(state=state)
        except:
            pass  # Some widgets don't support state
            
        for child in parent.winfo_children():
            self.set_widget_state(child, state)

    def load_multiple_sessions(self):
        """Load multiple session files to calculate averages."""
        filenames = filedialog.askopenfilenames(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Select Multiple Session Files for Comparison"
        )
        
        if not filenames:
            return
            
        self.all_sessions_data = {}
        
        for filename in filenames:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)
                    
                # Extract assessor name for identification
                assessor_name = session_data.get('header', {}).get('Assessor Name', 'Unknown')
                if not assessor_name.strip():
                    assessor_name = f"User_{len(self.all_sessions_data) + 1}"
                    
                self.all_sessions_data[assessor_name] = session_data
                debug_print(f"Loaded session data for assessor: {assessor_name}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load {filename}: {e}")
                
        if self.all_sessions_data:
            messagebox.showinfo("Success", f"Loaded {len(self.all_sessions_data)} session(s) for comparison")
            debug_print(f"Successfully loaded {len(self.all_sessions_data)} sessions for comparison")
        else:
            messagebox.showwarning("Warning", "No valid session files were loaded")
            
    def calculate_sample_averages(self):
        """Calculate average ratings for each sample across all users."""
        if not self.all_sessions_data:
            return
            
        sample_totals = {}  # {sample_name: {metric: [values], 'count': n}}
        
        # Collect all values for each sample/metric combination
        for assessor, session_data in self.all_sessions_data.items():
            samples = session_data.get('samples', {})
            for sample_name, sample_data in samples.items():
                if sample_name not in sample_totals:
                    sample_totals[sample_name] = {metric: [] for metric in self.metrics}
                    sample_totals[sample_name]['comments'] = []
                    
                for metric in self.metrics:
                    if metric in sample_data:
                        sample_totals[sample_name][metric].append(sample_data[metric])
                        
                # Collect comments
                if 'comments' in sample_data and sample_data['comments'].strip():
                    sample_totals[sample_name]['comments'].append(f"{assessor}: {sample_data['comments']}")
        
        # Calculate averages
        self.average_samples = {}
        for sample_name, data in sample_totals.items():
            self.average_samples[sample_name] = {}
            for metric in self.metrics:
                if data[metric]:  # If we have values
                    avg_value = sum(data[metric]) / len(data[metric])
                    self.average_samples[sample_name][metric] = round(avg_value, 1)
                else:
                    self.average_samples[sample_name][metric] = 5  # Default
                    
            # Combine comments
            self.average_samples[sample_name]['comments'] = '\n'.join(data['comments'])
            
        debug_print(f"Calculated averages for {len(self.average_samples)} samples")
        
    def update_comparison_plot(self):
        """Update plot to show averaged data across users."""
        if not self.average_samples:
            return
            
        # Temporarily replace samples with averages for plotting
        original_samples = self.samples.copy()
        original_checkboxes = self.sample_checkboxes.copy()
        
        # Set up average samples for plotting
        self.samples = self.average_samples.copy()
        
        # Update sample checkboxes to show average samples
        self.sample_checkboxes = {}
        for sample_name in self.average_samples.keys():
            var = tk.BooleanVar(value=True)  # Show all by default
            self.sample_checkboxes[sample_name] = var
            
        # Update the checkbox display
        self.update_sample_checkboxes()
        
        # Update the plot
        self.create_spider_plot()
        
        debug_print("Updated plot with averaged comparison data")

    def rename_current_sample(self):
        """Rename the currently selected sample."""
        current_sample = self.sample_var.get()
        if not current_sample:
            messagebox.showwarning("No Sample Selected", "Please select a sample to rename.")
            return
    
        if current_sample not in self.samples:
            messagebox.showwarning("Sample Not Found", f"Sample '{current_sample}' not found.")
            return
    
        # Get new name from user
        new_name = tk.simpledialog.askstring(
            "Rename Sample", 
            f"Current name: {current_sample}\n\nEnter new name:",
            initialvalue=current_sample
        )
    
        if not new_name or new_name.strip() == "":
            return  # User cancelled or entered empty name
    
        new_name = new_name.strip()
    
        # Check if new name already exists
        if new_name in self.samples and new_name != current_sample:
            messagebox.showerror("Name Conflict", 
                               f"A sample named '{new_name}' already exists.\n"
                               f"Please choose a different name.")
            return
    
        # Perform the rename
        if new_name != current_sample:
            # Copy data to new key
            self.samples[new_name] = self.samples[current_sample]
        
            # Remove old key
            del self.samples[current_sample]
        
            # Update UI components
            self.update_sample_combo()
            self.update_sample_checkboxes()
        
            # Select the renamed sample
            self.sample_var.set(new_name)
            self.load_sample_data(new_name)
        
            # Update plot
            self.update_plot()
        
            debug_print(f"Renamed sample '{current_sample}' to '{new_name}'")
            messagebox.showinfo("Success", f"Sample renamed to '{new_name}'")

    def batch_rename_samples(self):
        """Rename multiple samples at once."""
        if not self.samples:
            messagebox.showwarning("No Samples", "No samples available to rename.")
            return
    
        # Create batch rename dialog
        rename_window = tk.Toplevel(self.window)
        rename_window.title("Batch Rename Samples")
        rename_window.geometry("600x400")
        rename_window.transient(self.window)
        rename_window.grab_set()
    
        # Center the window
        rename_window.update_idletasks()
        x = (rename_window.winfo_screenwidth() // 2) - (300)
        y = (rename_window.winfo_screenheight() // 2) - (200)
        rename_window.geometry(f"600x400+{x}+{y}")
    
        main_frame = ttk.Frame(rename_window, padding=10)
        main_frame.pack(fill='both', expand=True)
    
        ttk.Label(main_frame, text="Batch Rename Samples", font=('Arial', 14, 'bold')).pack(pady=5)
        ttk.Label(main_frame, text="Edit the names below, then click Apply Changes", font=('Arial', 10)).pack(pady=2)
    
        # Create scrollable frame for sample entries
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
    
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    
        # Store the name variables
        name_vars = {}
        original_names = list(self.samples.keys())
    
        # Create entry fields for each sample
        for i, sample_name in enumerate(original_names):
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill='x', pady=2, padx=5)
        
            ttk.Label(row_frame, text=f"Sample {i+1}:", width=10).pack(side='left')
        
            name_var = tk.StringVar(value=sample_name)
            name_vars[sample_name] = name_var
        
            entry = ttk.Entry(row_frame, textvariable=name_var, width=40)
            entry.pack(side='left', padx=5, fill='x', expand=True)
    
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
        # Quick action buttons
        quick_frame = ttk.Frame(main_frame)
        quick_frame.pack(fill='x', pady=10)
    
        def apply_prefix():
            prefix = tk.simpledialog.askstring("Add Prefix", "Enter prefix to add:")
            if prefix:
                for sample_name in original_names:
                    current = name_vars[sample_name].get()
                    name_vars[sample_name].set(f"{prefix}{current}")
    
        def apply_suffix():
            suffix = tk.simpledialog.askstring("Add Suffix", "Enter suffix to add:")
            if suffix:
                for sample_name in original_names:
                    current = name_vars[sample_name].get()
                    name_vars[sample_name].set(f"{current}{suffix}")
    
        def number_samples():
            base = tk.simpledialog.askstring("Number Samples", "Enter base name (e.g., 'Test'):")
            if base:
                for i, sample_name in enumerate(original_names):
                    name_vars[sample_name].set(f"{base} {i+1}")
    
        ttk.Button(quick_frame, text="Add Prefix", command=apply_prefix).pack(side='left', padx=5)
        ttk.Button(quick_frame, text="Add Suffix", command=apply_suffix).pack(side='left', padx=5)
        ttk.Button(quick_frame, text="Number Samples", command=number_samples).pack(side='left', padx=5)
    
        # Apply/Cancel buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
    
        def apply_changes():
            # Collect all new names
            new_names = {}
            for original_name in original_names:
                new_name = name_vars[original_name].get().strip()
                if not new_name:
                    messagebox.showerror("Empty Name", "All samples must have names.")
                    return
                new_names[original_name] = new_name
        
            # Check for duplicates
            name_counts = {}
            for new_name in new_names.values():
                name_counts[new_name] = name_counts.get(new_name, 0) + 1
        
            duplicates = [name for name, count in name_counts.items() if count > 1]
            if duplicates:
                messagebox.showerror("Duplicate Names", 
                                   f"The following names are used more than once:\n{', '.join(duplicates)}\n\n"
                                   f"Please ensure all names are unique.")
                return
        
            # Apply the changes
            new_samples = {}
            for original_name in original_names:
                new_name = new_names[original_name]
                new_samples[new_name] = self.samples[original_name]
        
            self.samples = new_samples
        
            # Update UI
            current_selection = self.sample_var.get()
            if current_selection in new_names:
                new_selection = new_names[current_selection]
            else:
                new_selection = list(self.samples.keys())[0] if self.samples else ""
        
            self.update_sample_combo()
            self.update_sample_checkboxes()
        
            if new_selection:
                self.sample_var.set(new_selection)
                self.load_sample_data(new_selection)
        
            self.update_plot()
        
            rename_window.destroy()
            messagebox.showinfo("Success", f"Successfully renamed {len(original_names)} samples.")
    
        ttk.Button(button_frame, text="Apply Changes", command=apply_changes).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=rename_window.destroy).pack(side='right', padx=5)


    def create_training_structure(self):
        """Create training data folder structure."""
        try:
            from ml_form_processor import MLFormProcessor, MLTrainingHelper
            processor = MLFormProcessor()
            trainer = MLTrainingHelper(processor)
            trainer.create_training_data_structure()
            messagebox.showinfo("Success", "Training data structure created!\nSee console for instructions.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create structure: {e}")

    def train_ml_model(self):
        """Train the ML model."""
        try:
            from ml_form_processor import MLFormProcessor, MLTrainingHelper
            processor = MLFormProcessor()
            trainer = MLTrainingHelper(processor)
        
            # Show training dialog
            result = messagebox.askyesno("Train Model", 
                                       "This will train the ML model using data in training_data/sensory_ratings/\n\n"
                                       "Make sure you have added training images first.\n\n"
                                       "Continue?")
            if result:
                model, history = trainer.train_model()
                messagebox.showinfo("Success", "Model training completed!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Training failed: {e}")

    def create_training_structure(self):
        """Create training data folder structure using the helper tool."""
        try:
            from training_data_helper import TrainingDataExtractor
            extractor = TrainingDataExtractor()
            extractor.create_training_data_structure()
            messagebox.showinfo("Success", 
                              "Training data structure created!\n\n"
                              "Next steps:\n"
                              "1. Place your scanned forms in a folder\n"
                              "2. Run the training data extractor\n"
                              "3. Label the regions when prompted\n"
                              "4. Train the model")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create structure: {e}")

    def extract_training_data(self):
        """Launch the training data extraction tool."""
        try:
            import subprocess
            import sys
        
            # Run the training data helper as a separate process
            result = messagebox.askyesno("Extract Training Data",
                                       "This will launch the training data extraction tool.\n\n"
                                       "Make sure you have your scanned forms ready.\n\n"
                                       "Continue?")
            if result:
                subprocess.run([sys.executable, "training_data_helper.py"])
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch extractor: {e}")

    def check_training_data_balance(self):
        """Check the balance of the current training dataset."""
        try:
            from training_data_helper import TrainingDataExtractor
            extractor = TrainingDataExtractor()
        
            # Capture the output and show in a dialog
            import io
            import contextlib
        
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                extractor.check_dataset_balance()
            
            balance_info = output.getvalue()
        
            # Show in a scrollable text dialog
            info_window = tk.Toplevel(self.window)
            info_window.title("Training Data Balance")
            info_window.geometry("400x300")
        
            text_widget = tk.Text(info_window, font=('Courier', 10))
            text_widget.pack(fill='both', expand=True, padx=10, pady=10)
            text_widget.insert('1.0', balance_info)
            text_widget.config(state='disabled')
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to check balance: {e}")

    def create_spider_plot(self):
        """Create a spider/radar plot of sensory data."""
        self.ax.clear()
        
        # Check if we have any data to plot
        if not self.samples:
            self.ax.text(0.5, 0.5, 'No samples to display\nAdd samples to begin evaluation', 
                        transform=self.ax.transAxes, ha='center', va='center', 
                        fontsize=14, color='gray')
            self.canvas.draw()
            return
            
        # Get selected samples for plotting
        selected_samples = []
        for sample_name, checkbox_var in self.sample_checkboxes.items():
            if checkbox_var.get() and sample_name in self.samples:
                selected_samples.append(sample_name)
                    
        if not selected_samples:
            self.ax.text(0.5, 0.5, 'No samples selected\nUse checkboxes to select samples', 
                        transform=self.ax.transAxes, ha='center', va='center', 
                        fontsize=14, color='gray')
            self.canvas.draw()
            return
        
        # Setup the spider plot
        num_metrics = len(self.metrics)
        angles = np.linspace(0, 2 * np.pi, num_metrics, endpoint=False).tolist()
        angles += angles[:1]  # Complete the circle
        
        # Set up the plot
        self.ax.set_theta_offset(np.pi / 2)
        self.ax.set_theta_direction(-1)
        self.ax.set_thetagrids(np.degrees(angles[:-1]), self.metrics, fontsize=10)
        self.ax.set_ylim(0, 9)  # Changed to 9
        self.ax.set_yticks(range(1, 10))  # Changed to 1-9
        self.ax.set_yticklabels(range(1, 10))
        self.ax.grid(True, alpha=0.3)
        
        # Colors for different samples
        colors = ['red', 'green', 'blue', 'orange', 'purple', 'brown', 'pink', 'cyan', 'magenta', 'lime']
        line_styles = ['-', '--', '-.', ':']
        
        # Plot each selected sample
        for i, sample_name in enumerate(selected_samples):
            sample_data = self.samples[sample_name]
            values = [sample_data.get(metric, 5) for metric in self.metrics]  # Default to 5 if no data
            values += values[:1]  # Complete the circle
            
            color = colors[i % len(colors)]
            line_style = line_styles[i % len(line_styles)]
            
            # Plot the line and markers
            self.ax.plot(angles, values, 'o', linewidth=2.5, label=sample_name, 
                        color=color, linestyle=line_style, markersize=8, alpha=0.8)
            # Fill the area
            self.ax.fill(angles, values, alpha=0.1, color=color)
            
        # Add legend
        if selected_samples:
            self.ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1.0), fontsize=10)
            
        # Set title
        self.ax.set_title('Sensory Profile Comparison', fontsize=14, fontweight='bold', pad=20, ha = 'center', y = 1.08)
        
        # Force canvas update
        self.canvas.draw_idle()
        
    def add_sample(self):
        """Add a new sample for evaluation."""
        import tkinter.simpledialog
        sample_name = tk.simpledialog.askstring("Add Sample", "Enter sample name:")
        if sample_name and sample_name not in self.samples:
            # Initialize sample data
            self.samples[sample_name] = {metric: 0 for metric in self.metrics}
            self.samples[sample_name]['comments'] = ''
            
            # Update UI
            self.update_sample_combo()
            self.update_sample_checkboxes()
            self.sample_var.set(sample_name)
            self.load_sample_data(sample_name)
            
            # force plot update
            self.update_plot()

            debug_print(f"Added sample: {sample_name}")
        elif sample_name in self.samples:
            messagebox.showwarning("Warning", "Sample name already exists!")
            
    def remove_sample(self):
        """Remove the currently selected sample."""
        current_sample = self.sample_var.get()
        if current_sample and current_sample in self.samples:
            if messagebox.askyesno("Confirm", f"Remove sample '{current_sample}'?"):
                del self.samples[current_sample]
                self.update_sample_combo()
                self.update_sample_checkboxes()
                
                # Select first available sample or clear
                if self.samples:
                    first_sample = list(self.samples.keys())[0]
                    self.sample_var.set(first_sample)
                    self.load_sample_data(first_sample)
                else:
                    self.sample_var.set('')
                    self.clear_form()
                    
                self.update_plot()
                debug_print(f"Removed sample: {current_sample}")
                
    def clear_current_sample(self):
        """Clear data for the current sample."""
        current_sample = self.sample_var.get()
        if current_sample and current_sample in self.samples:
            if messagebox.askyesno("Confirm", f"Clear data for '{current_sample}'?"):
                # Reset all ratings to 5 (neutral)
                for metric in self.metrics:
                    self.samples[current_sample][metric] = 5
                    self.rating_vars[metric].set(5)
                
                self.samples[current_sample]['comments'] = ''
                self.comments_text.delete('1.0', tk.END)
                
                self.update_plot()
                debug_print(f"Cleared data for sample: {current_sample}")
                
    def update_sample_combo(self):
        """Update the sample selection combo box."""
        sample_names = list(self.samples.keys())
        self.sample_combo['values'] = sample_names
        
    def update_sample_checkboxes(self):
        """Update the checkboxes for sample selection in plotting."""
        # Clear existing checkboxes
        for widget in self.checkbox_frame.winfo_children():
            widget.destroy()

        # ADDED: Show different header based on mode
        if self.current_mode == "comparison":
            header_text = "Select Averaged Samples to Display:"
        else:
            header_text = "Select Samples to Display:"
            
        # Update the header label
        for widget in self.checkbox_frame.master.winfo_children():
            if isinstance(widget, ttk.Label):
                widget.config(text=header_text)
                break
            
        self.sample_checkboxes = {}
        
        # Create checkboxes for each sample
        for i, sample_name in enumerate(self.samples.keys()):
            var = tk.BooleanVar(value=True)  # Default to checked
            checkbox = ttk.Checkbutton(self.checkbox_frame, text=sample_name, 
                                     variable=var, command=self.update_plot)
            checkbox.grid(row=i//3, column=i%3, sticky='w', padx=5, pady=2)
            self.sample_checkboxes[sample_name] = var
            
    def on_sample_changed(self, event=None):
        """Handle sample selection change."""
        selected_sample = self.sample_var.get()
        if selected_sample in self.samples:
            self.load_sample_data(selected_sample)
            
    def load_sample_data(self, sample_name):
        """Load data for the specified sample into the form."""
        if sample_name in self.samples:
            sample_data = self.samples[sample_name]
            
            # Load ratings
            for metric in self.metrics:
                value = sample_data.get(metric, 5)
                self.rating_vars[metric].set(value)
                
            # Load comments
            comments = sample_data.get('comments', '')
            self.comments_text.delete('1.0', tk.END)
            self.comments_text.insert('1.0', comments)
         
    def save_plot_as_image(self):
        """Save the current spider plot as an image file."""
        print("DEBUG: Starting plot image save")
    
        if not self.samples:
            messagebox.showwarning("Warning", "No samples to save! Please add samples first.")
            return
        
        try:
            filename = filedialog.asksaveasfilename(
                title="Save Plot as Image",
                defaultextension=".png",
                filetypes=[
                    ("PNG files", "*.png"),
                    ("JPG files", "*.jpg"),
                    ("PDF files", "*.pdf"),
                    ("SVG files", "*.svg"),
                    ("All files", "*.*")
                ]
            )
        
            if filename:
                # Ensure we have the latest plot
                self.update_plot()
            
                # Save the figure with high DPI for quality
                self.fig.savefig(filename, dpi=300, bbox_inches='tight', 
                               facecolor='white', edgecolor='none')
            
                print(f"DEBUG: Plot saved successfully to {filename}")
                messagebox.showinfo("Success", f"Plot saved successfully as {os.path.basename(filename)}")
            
        except Exception as e:
            print(f"DEBUG: Error saving plot: {e}")
            messagebox.showerror("Error", f"Failed to save plot: {str(e)}")

    def save_table_as_image(self):
        """Save the sensory data table as an image."""
        print("DEBUG: Starting table image save with comments")
    
        if not self.samples:
            messagebox.showwarning("Warning", "No data to save! Please add samples first.")
            return
        
        try:
            filename = filedialog.asksaveasfilename(
                title="Save Table as Image",
                defaultextension=".png",
                filetypes=[
                    ("PNG files", "*.png"),
                    ("JPG files", "*.jpg"),
                    ("PDF files", "*.pdf"),
                    ("All files", "*.*")
                ]
            )
        
            if filename:
                # Create table data with attributes as headers and samples as rows
                table_data = []
            
                # Header row with attributes + comments
                headers = ["Sample"] + self.metrics + ["Additional Comments"]
                table_data.append(headers)
            
                # Data rows - one per sample
                for sample_name, sample_data in self.samples.items():
                    row = [sample_name]
                    for metric in self.metrics:
                        row.append(str(sample_data.get(metric, "N/A")))
                    # Add comments - get from sample data or leave blank if empty
                    comments = sample_data.get("comments", "").strip()
                    row.append(comments if comments else "")
                    table_data.append(row)
            
                # Create figure for table with wider width to accommodate comments
                fig, ax = plt.subplots(figsize=(16, max(6, len(self.samples) * 0.5)))
                ax.axis('tight')
                ax.axis('off')
            
                # Create table
                table = ax.table(cellText=table_data[1:], colLabels=table_data[0],
                               cellLoc='center', loc='center')
                table.auto_set_font_size(False)
                table.set_fontsize(9)  # Slightly smaller font to fit more content
                table.scale(1.4, 2.2)  # Wider scale to accommodate comments
            
                # Style the table
                for i in range(len(headers)):
                    table[(0, i)].set_facecolor('#4CAF50')
                    table[(0, i)].set_text_props(weight='bold', color='white')
            
                # Make comments column wider and left-aligned
                comments_col_idx = len(headers) - 1
                for row_idx in range(len(table_data)):
                    if row_idx == 0:  # Header
                        continue
                    cell = table[(row_idx, comments_col_idx)]
                    cell.set_width(0.3)  # Make comments column wider
                    cell.set_text_props(ha='left', va='top', wrap=True)  # Left align and wrap text
            
                # Add title
                ax.set_title("Sensory Evaluation Results", fontsize=16, fontweight='bold', pad=20)
            
                # Save with high quality
                fig.savefig(filename, dpi=300, bbox_inches='tight', 
                           facecolor='white', edgecolor='none')
                plt.close(fig)
            
                print(f"DEBUG: Table with comments saved successfully to {filename}")
                messagebox.showinfo("Success", f"Table saved successfully as {os.path.basename(filename)}")
            
        except Exception as e:
            print(f"DEBUG: Error saving table: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to save table: {str(e)}")

    def generate_powerpoint_report(self):
        """Generate a PowerPoint report using the same template as generate_test_report."""
        print("DEBUG: Starting PowerPoint report generation")

        if not self.samples:
            messagebox.showwarning("Warning", "No data to export! Please add samples first.")
            return
    
        try:
            # Get save location
            filename = filedialog.asksaveasfilename(
                title="Save PowerPoint Report",
                defaultextension=".pptx",
                filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
            )
    
            if not filename:
                return
        
            print(f"DEBUG: Creating PowerPoint report at {filename}")
    
            # Import required modules for PowerPoint generation
            from pptx import Presentation
            from pptx.util import Inches, Pt
            from pptx.enum.text import PP_ALIGN
            from pptx.dml.color import RGBColor
            from datetime import datetime
            import tempfile
    
            # Create presentation with same template structure as existing reports
            prs = Presentation()
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)
    
            # Create main content slide
            main_slide = prs.slides.add_slide(prs.slide_layouts[6])
    
            # Add background and logo using existing template structure
            from resource_utils import get_resource_path
    
            background_path = get_resource_path("resources/ccell_background.png")
            if os.path.exists(background_path):
                main_slide.shapes.add_picture(background_path, Inches(0), Inches(0),
                                            width=prs.slide_width, height=prs.slide_height)
                print("DEBUG: Background added successfully")
            else:
                print("DEBUG: Background not found, using plain slide")
        
            logo_path = get_resource_path("resources/ccell_logo_full.png")
            if os.path.exists(logo_path):
                main_slide.shapes.add_picture(logo_path, Inches(11.21), Inches(0.43),
                                            width=Inches(1.57), height=Inches(0.53))
                print("DEBUG: Logo added successfully")
    
            # Add title
            title_shape = main_slide.shapes.add_textbox(Inches(0.45), Inches(-0.04), 
                                                       Inches(10.72), Inches(0.64))
            text_frame = title_shape.text_frame
            text_frame.clear()
    
            p = text_frame.add_paragraph()
            p.text = "Sensory Evaluation Report"
            p.font.name = "Montserrat"
            p.font.size = Pt(32)
            p.font.bold = True
    
            # Create table data with proper structure (attributes as headers + comments)
            table_data = []
            headers = ["Sample"] + self.metrics + ["Additional Comments"]
    
            # Add header data if available
            header_info = []
            for field in self.header_fields:
                if field in self.header_vars and self.header_vars[field].get():
                    header_info.append(f"{field}: {self.header_vars[field].get()}")
    
            # Add data rows with current comments (including any just typed)
            for sample_name, sample_data in self.samples.items():
                row = [sample_name]
                for metric in self.metrics:
                    row.append(str(sample_data.get(metric, "N/A")))
                # Include current comments
                comments = sample_data.get("comments", "").strip()
                row.append(comments if comments else "")
                table_data.append(row)
    
            # FIXED: Better layout - smaller table, larger plot area for legend
            if table_data:
                table_shape = main_slide.shapes.add_table(
                    len(table_data) + 1, len(headers),  # +1 for header row
                    Inches(0.45), Inches(1.5),
                    Inches(6.5), Inches(4.5)  # REDUCED: Smaller table width
                )
                table = table_shape.table
        
                # Set header row
                for col_idx, header in enumerate(headers):
                    cell = table.cell(0, col_idx)
                    cell.text = header
                    cell.text_frame.paragraphs[0].font.bold = True
                    cell.text_frame.paragraphs[0].font.size = Pt(10)  # Slightly smaller
        
                # Set data rows
                for row_idx, row_data in enumerate(table_data, 1):
                    for col_idx, cell_value in enumerate(row_data):
                        cell = table.cell(row_idx, col_idx)
                        cell.text = str(cell_value)
                    
                        # Special formatting for comments column
                        if col_idx == len(headers) - 1:  # Comments column
                            cell.text_frame.paragraphs[0].font.size = Pt(8)
                            # Set text alignment for comments
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.alignment = PP_ALIGN.LEFT
                        else:
                            cell.text_frame.paragraphs[0].font.size = Pt(9)
        
                print(f"DEBUG: Table with comments added - {len(table_data)} rows and {len(headers)} columns")
    
            # FIXED: Create plot with better legend positioning
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                plot_image_path = tmp_file.name
        
            try:
                # Create a special figure for PowerPoint with better legend handling
                fig_ppt, ax_ppt = plt.subplots(figsize=(8, 6), subplot_kw=dict(projection='polar'))
                fig_ppt.patch.set_facecolor('white')
            
                # Get selected samples for plotting
                selected_samples = []
                for sample_name, checkbox_var in self.sample_checkboxes.items():
                    if checkbox_var.get() and sample_name in self.samples:
                        selected_samples.append(sample_name)
            
                # If no samples selected, select all
                if not selected_samples:
                    selected_samples = list(self.samples.keys())
            
                if selected_samples:
                    # Setup the spider plot
                    num_metrics = len(self.metrics)
                    angles = np.linspace(0, 2 * np.pi, num_metrics, endpoint=False).tolist()
                    angles += angles[:1]  # Complete the circle
                
                    # Set up the plot
                    ax_ppt.set_theta_offset(np.pi / 2)
                    ax_ppt.set_theta_direction(-1)
                    ax_ppt.set_thetagrids(np.degrees(angles[:-1]), self.metrics, fontsize=10)
                    ax_ppt.set_ylim(0, 9)
                    ax_ppt.set_yticks(range(1, 10))
                    ax_ppt.set_yticklabels(range(1, 10))
                    ax_ppt.grid(True, alpha=0.3)
                
                    # Colors for different samples
                    colors = ['red', 'green', 'blue', 'orange', 'purple', 'brown', 'pink', 'cyan', 'magenta', 'lime']
                    line_styles = ['-', '--', '-.', ':']
                
                    # Plot each selected sample
                    for i, sample_name in enumerate(selected_samples):
                        sample_data = self.samples[sample_name]
                        values = [sample_data.get(metric, 5) for metric in self.metrics]
                        values += values[:1]  # Complete the circle
                    
                        color = colors[i % len(colors)]
                        line_style = line_styles[i % len(line_styles)]
                    
                        # Plot the line and markers
                        ax_ppt.plot(angles, values, 'o', linewidth=2.5, label=sample_name, 
                                   color=color, linestyle=line_style, markersize=8, alpha=0.8)
                        # Fill the area
                        ax_ppt.fill(angles, values, alpha=0.1, color=color)
                
                    # FIXED: Better legend positioning - inside the plot area
                    ax_ppt.legend(loc='upper right', bbox_to_anchor=(1.1, 1.1), fontsize=9)
                
                    # Set title
                    ax_ppt.set_title('Sensory Profile Comparison', fontsize=12, fontweight='bold', pad=15)
            
                # Save the PowerPoint-specific plot
                fig_ppt.savefig(plot_image_path, dpi=300, bbox_inches='tight',
                               facecolor='white', edgecolor='none')
                plt.close(fig_ppt)
        
                # FIXED: Better plot positioning and sizing - more space, legend won't be cut off
                main_slide.shapes.add_picture(plot_image_path, 
                                            Inches(7.2), Inches(1.5),    # MOVED: Further left
                                            Inches(5.8), Inches(4.5))    # INCREASED: Wider plot
                print("DEBUG: Plot with proper legend positioning added to PowerPoint slide")
        
            finally:
                # Clean up temporary file
                if os.path.exists(plot_image_path):
                    os.remove(plot_image_path)
    
            # Add header information as text box if available
            if header_info:
                info_shape = main_slide.shapes.add_textbox(Inches(0.45), Inches(6.2),
                                                         Inches(12.0), Inches(1.0))
                info_frame = info_shape.text_frame
                info_frame.clear()
        
                p = info_frame.add_paragraph()
                p.text = " | ".join(header_info)
                p.font.size = Pt(10)
                p.font.name = "Montserrat"
    
            # Save the presentation
            prs.save(filename)
            print(f"DEBUG: PowerPoint report with auto-saved comments and proper legend saved successfully to {filename}")
            messagebox.showinfo("Success", f"PowerPoint report saved successfully as {os.path.basename(filename)}")
    
        except Exception as e:
            print(f"DEBUG: Error generating PowerPoint report: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to generate PowerPoint report: {str(e)}")

    def save_current_sample(self):
        """Save the current form data to the selected sample."""
        current_sample = self.sample_var.get()
        if not current_sample:
            messagebox.showwarning("Warning", "No sample selected!")
            return
            
        if current_sample not in self.samples:
            # Create new sample if it doesn't exist
            self.samples[current_sample] = {}
            
        # Save ratings
        for metric in self.metrics:
            self.samples[current_sample][metric] = self.rating_vars[metric].get()
            
        # Save comments
        comments = self.comments_text.get('1.0', tk.END).strip()
        self.samples[current_sample]['comments'] = comments
        
        self.update_plot()
        debug_print(f"Saved data for sample: {current_sample}")
        messagebox.showinfo("Success", f"Data saved for {current_sample}")
        
    def clear_form(self):
        """Clear all form fields."""
        for metric in self.metrics:
            self.rating_vars[metric].set(5)
        self.comments_text.delete('1.0', tk.END)
        
    def update_plot(self):
        """Update the spider plot."""
        self.create_spider_plot()
        
    def show_sop(self):
        """Display the Standard Operating Procedure."""
        sop_window = tk.Toplevel(self.window)
        sop_window.title("Sensory Evaluation SOP")
        sop_window.geometry("600x500")
        sop_window.configure(bg='white')
        
        text_widget = tk.Text(sop_window, wrap='word', font=('Arial', 11), 
                             bg='white', fg='black', padx=20, pady=20)
        text_widget.pack(fill='both', expand=True)
        text_widget.insert('1.0', self.sop_text)
        text_widget.config(state='disabled')
        
        # Center the SOP window
        sop_window.update_idletasks()
        x = (sop_window.winfo_screenwidth() // 2) - (300)
        y = (sop_window.winfo_screenheight() // 2) - (250)
        sop_window.geometry(f"600x500+{x}+{y}")
        
    def new_session(self):
        """Start a new sensory evaluation session."""
        if self.samples and messagebox.askyesno("Confirm", "Start new session? Current data will be lost."):
            self.samples = {}
            self.sample_checkboxes = {}
            
            # Clear header fields
            for field, var in self.header_vars.items():
                if field == "Date":
                    var.set(datetime.now().strftime("%Y-%m-%d"))
                else:
                    var.set('')
                    
            self.update_sample_combo()
            self.update_sample_checkboxes()
            self.sample_var.set('')
            self.clear_form()
            self.update_plot()
            
            debug_print("Started new sensory evaluation session")
            
    def save_session(self):
        """Save the current session to a JSON file."""
        if not self.samples:
            messagebox.showwarning("Warning", "No data to save!")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Sensory Session"
        )
        
        if filename:
            try:
                session_data = {
                    'header': {field: var.get() for field, var in self.header_vars.items()},
                    'samples': self.samples,
                    'timestamp': datetime.now().isoformat()
                }
                
                with open(filename, 'w') as f:
                    json.dump(session_data, f, indent=2)
                    
                messagebox.showinfo("Success", f"Session saved to {filename}")
                debug_print(f"Saved sensory session to: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save session: {e}")
                
    def load_session(self):
        """Load a session from a JSON file."""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Load Sensory Session"
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)
                    
                # Load header data
                if 'header' in session_data:
                    for field, value in session_data['header'].items():
                        if field in self.header_vars:
                            self.header_vars[field].set(value)
                            
                # Load samples
                if 'samples' in session_data:
                    self.samples = session_data['samples']
                    self.update_sample_combo()
                    self.update_sample_checkboxes()
                    
                    # Select first sample
                    if self.samples:
                        first_sample = list(self.samples.keys())[0]
                        self.sample_var.set(first_sample)
                        self.load_sample_data(first_sample)
                        
                    self.update_plot()
                    
                messagebox.showinfo("Success", f"Session loaded from {filename}")
                debug_print(f"Loaded sensory session from: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load session: {e}")
                
    def export_to_excel(self):
        """Export the sensory data to an Excel file."""
        if not self.samples:
            messagebox.showwarning("Warning", "No data to export!")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Sensory Data"
        )
        
        if filename:
            try:
                # Create DataFrame for sensory data
                data_rows = []
                for sample_name, sample_data in self.samples.items():
                    row = {'Sample': sample_name}
                    
                    # Add header information
                    for field, var in self.header_vars.items():
                        row[field] = var.get()
                        
                    # Add sensory ratings
                    for metric in self.metrics:
                        row[metric] = sample_data.get(metric, 0)
                        
                    row['Comments'] = sample_data.get('comments', '')
                    data_rows.append(row)
                    
                df = pd.DataFrame(data_rows)
                
                # Save to Excel
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sensory Data', index=False)
                    
                    # Save spider plot as image
                    plot_filename = filename.replace('.xlsx', '_spider_plot.png')
                    self.fig.savefig(plot_filename, dpi=300, bbox_inches='tight')
                    
                messagebox.showinfo("Success", f"Data exported to {filename}\nSpider plot saved as {plot_filename}")
                debug_print(f"Exported sensory data to: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {e}")


    def auto_save_and_update(self):
        """Automatically save current sample data and update plot."""
        current_sample = self.sample_var.get()
        if current_sample and current_sample in self.samples:
            # Auto-save ratings
            for metric in self.metrics:
                self.samples[current_sample][metric] = self.rating_vars[metric].get()
        
            # FIXED: Also auto-save comments
            comments = self.comments_text.get('1.0', tk.END).strip()
            self.samples[current_sample]['comments'] = comments
        
            # Update plot immediately
            self.update_plot()
        
            print(f"DEBUG: Auto-saved all data for {current_sample}")

    def load_from_image(self):
        """Load sensory data from a scanned form image using ML."""
        # Check for ML dependencies
        try:
            import importlib.util
            tf_spec = importlib.util.find_spec('tensorflow')
            cv2_spec = importlib.util.find_spec('cv2')
        
            missing = []
            if tf_spec is None:
                missing.append('tensorflow')
            if cv2_spec is None:
                missing.append('opencv-python')
            
            if missing:
                missing_str = ', '.join(missing)
                messagebox.showwarning("Missing Dependencies", 
                                     f"ML image processing requires: {missing_str}\n\n"
                                     f"Install with: pip install {' '.join(missing)}")
                return
            
        except Exception as e:
            debug_print(f"Dependency check failed: {e}")
    
        filename = filedialog.askopenfilename(
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.tiff *.bmp"),
                ("All files", "*.*")
            ],
            title="Load Sensory Form Image"
        )
    
        if filename:
            try:
                # Show processing dialog
                progress_window = tk.Toplevel(self.window)
                progress_window.title("Processing Image with ML...")
                progress_window.geometry("300x100")
                progress_window.transient(self.window)
                progress_window.grab_set()
            
                progress_label = ttk.Label(progress_window, text="Loading ML model and processing...", font=FONT)
                progress_label.pack(expand=True)
            
                progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
                progress_bar.pack(fill='x', padx=20, pady=10)
                progress_bar.start()
            
                self.window.update()
            
                # Import and use ML processor (lazy loading)
                from ml_form_processor import MLFormProcessor
                processor = MLFormProcessor()
                extracted_data, processed_img = processor.process_form_image(filename)
            
                # Stop progress bar
                progress_bar.stop()
                progress_window.destroy()
            
                # Show preview dialog
                self.show_extraction_preview(extracted_data, processed_img, filename)
            
            except Exception as e:
                if 'progress_window' in locals():
                    progress_window.destroy()
                messagebox.showerror("Error", f"Failed to process image: {e}")
                debug_print(f"ML image processing error: {e}")
               
                
    def load_from_image_with_ai(self):
        """Load sensory data from a scanned form image using Claude AI."""
    
        # Check for required dependencies
        try:
            import anthropic
            import base64
            import io
            from PIL import Image
        except ImportError as e:
            messagebox.showerror("Missing Dependencies", 
                               f"AI image processing requires additional packages.\n\n"
                               f"Install with: pip install anthropic pillow\n\n"
                               f"Error: {e}")
            return

        # Get image file
        filename = filedialog.askopenfilename(
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.tiff *.bmp"),
                ("All files", "*.*")
            ],
            title="Load Sensory Form Image for AI Processing"
        )

        if not filename:
            return

        try:
            # Show processing dialog
            progress_window = tk.Toplevel(self.window)
            progress_window.title("Processing with Claude AI...")
            progress_window.geometry("350x120")
            progress_window.transient(self.window)
            progress_window.grab_set()
        
            progress_label = ttk.Label(progress_window, 
                                     text="Sending image to Claude AI for analysis...", 
                                     font=FONT)
            progress_label.pack(expand=True, pady=10)
        
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
            progress_bar.pack(fill='x', padx=20, pady=10)
            progress_bar.start()
        
            self.window.update()
        
            # Process with Claude AI
            ai_processor = ClaudeFormProcessor()
            extracted_data = ai_processor.process_form_image(filename)
        
            # Stop progress bar
            progress_bar.stop()
            progress_window.destroy()
        
            # Show results preview
            if extracted_data:
                self.show_ai_extraction_preview(extracted_data, filename)
            else:
                messagebox.showwarning("No Data", "No sensory data could be extracted from the image.")
            
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("AI Processing Error", f"Failed to process image with AI: {e}")
            debug_print(f"AI processing error: {e}")

    def show_ai_extraction_preview(self, extracted_data, filename):
        """Show a preview of AI-extracted data with editable sample names before loading."""
        preview_window = tk.Toplevel(self.window)
        preview_window.title("AI Extraction Results")
        preview_window.geometry("900x650")
        preview_window.transient(self.window)
        preview_window.grab_set()
    
        # Create layout
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        # Title
        title_label = ttk.Label(main_frame, 
                               text=f"Claude AI extracted data from: {os.path.basename(filename)}", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=5)
    
        # Confidence indicator
        confidence_label = ttk.Label(main_frame, 
                                   text="✓ AI Processing Complete - Review and edit data below", 
                                   font=('Arial', 10),
                                   foreground='green')
        confidence_label.pack(pady=2)
    
        # Sample name editing frame
        name_edit_frame = ttk.LabelFrame(main_frame, text="Edit Sample Names", padding=10)
        name_edit_frame.pack(fill='x', pady=5)
    
        # Store the sample name variables
        sample_name_vars = {}
        original_names = list(extracted_data.keys())
    
        # Create editable fields for sample names
        name_grid_frame = ttk.Frame(name_edit_frame)
        name_grid_frame.pack(fill='x')
    
        ttk.Label(name_grid_frame, text="Current Names → New Names:", font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=4, pady=5)
    
        for i, original_name in enumerate(original_names):
            row = (i // 2) + 1  # 2 samples per row
            col = (i % 2) * 2
        
            # Original name label
            ttk.Label(name_grid_frame, text=f"{original_name}:").grid(row=row, column=col, sticky='w', padx=5, pady=2)
        
            # Editable name entry
            name_var = tk.StringVar(value=original_name)
            sample_name_vars[original_name] = name_var
        
            name_entry = ttk.Entry(name_grid_frame, textvariable=name_var, width=20)
            name_entry.grid(row=row, column=col+1, sticky='w', padx=5, pady=2)
    
        # Quick rename buttons
        quick_rename_frame = ttk.Frame(name_edit_frame)
        quick_rename_frame.pack(fill='x', pady=5)
    
        def auto_increment_names():
            """Auto-generate sequential sample names."""
            base_name = tk.simpledialog.askstring("Base Name", "Enter base name (e.g., 'Sample'):")
            if base_name:
                for i, original_name in enumerate(original_names):
                    sample_name_vars[original_name].set(f"{base_name} {i+1}")
    
        def clear_sample_numbers():
            """Remove 'Sample X' and leave just numbers."""
            for original_name in original_names:
                current = sample_name_vars[original_name].get()
                if current.startswith('Sample '):
                    number = current.replace('Sample ', '')
                    sample_name_vars[original_name].set(number)
    
        ttk.Button(quick_rename_frame, text="Auto-Increment Names", command=auto_increment_names).pack(side='left', padx=5)
        ttk.Button(quick_rename_frame, text="Remove 'Sample' Prefix", command=clear_sample_numbers).pack(side='left', padx=5)
    
        # Data preview
        data_frame = ttk.LabelFrame(main_frame, text="Extracted Ratings (Live Preview)", padding=10)
        data_frame.pack(fill='both', expand=True, pady=5)
    
        # Create tree view for data
        columns = ['Sample'] + self.metrics
        tree = ttk.Treeview(data_frame, columns=columns, show='headings', height=8)
    
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120)
    
        # Function to update the tree view when names change
        def update_preview():
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)
        
            # Repopulate with current names
            for original_name in original_names:
                current_name = sample_name_vars[original_name].get()
                sample_data = extracted_data[original_name]
            
                values = [current_name]
                for metric in self.metrics:
                    values.append(sample_data.get(metric, 'N/A'))
                tree.insert('', 'end', values=values)
    
        # Bind name changes to update preview
        for name_var in sample_name_vars.values():
            name_var.trace('w', lambda *args: update_preview())
    
        # Initial population
        update_preview()
        
        # Add scrollbar to tree
        tree_scroll = ttk.Scrollbar(data_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
    
        tree.pack(side='left', fill='both', expand=True)
        tree_scroll.pack(side='right', fill='y')
    
        # Show any additional comments or notes
        if any('comments' in sample_data and sample_data['comments'] 
               for sample_data in extracted_data.values()):
            comments_frame = ttk.LabelFrame(main_frame, text="Additional Comments", padding=5)
            comments_frame.pack(fill='x', pady=5)
        
            for original_name, sample_data in extracted_data.items():
                if sample_data.get('comments'):
                    comment_text = f"{sample_name_vars[original_name].get()}: {sample_data['comments']}"
                    comment_label = ttk.Label(comments_frame, text=comment_text, font=('Arial', 9))
                    comment_label.pack(anchor='w')
    
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
    
        def load_ai_data():
            # Create new data structure with updated names
            final_data = {}
            for original_name in original_names:
                new_name = sample_name_vars[original_name].get().strip()
                if not new_name:  # If empty, use original name
                    new_name = original_name
            
                # Check for duplicate names
                if new_name in final_data:
                    messagebox.showerror("Duplicate Names", 
                                       f"Sample name '{new_name}' is used more than once.\n"
                                       f"Please ensure all sample names are unique.")
                    return
            
                final_data[new_name] = extracted_data[original_name]
        
            # Load the AI-extracted data with new names into the interface
            for sample_name, sample_data in final_data.items():
                self.samples[sample_name] = sample_data
            
            # Update UI
            self.update_sample_combo()
            self.update_sample_checkboxes()
        
            # Select first sample
            if self.samples:
                first_sample = list(self.samples.keys())[0]
                self.sample_var.set(first_sample)
                self.load_sample_data(first_sample)
            
            self.update_plot()
            preview_window.destroy()
        
            messagebox.showinfo("Success", 
                              f"Loaded {len(final_data)} samples from AI analysis!\n\n"
                              f"Sample names have been updated as specified.\n"
                              f"You can now review and adjust the ratings as needed.")
        
        def manual_edit():
            # Load data and keep preview open for comparison
            load_ai_data()
            # Don't close preview window so user can reference it
        
        ttk.Button(button_frame, text="Load Data with New Names", command=load_ai_data).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Load & Keep Preview", command=manual_edit).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=preview_window.destroy).pack(side='right', padx=5)

    def show_extraction_preview(self, extracted_data, processed_img, filename):
        """Show a preview of extracted data before loading."""
        preview_window = tk.Toplevel(self.window)
        preview_window.title("Extracted Data Preview")
        preview_window.geometry("800x600")
        preview_window.transient(self.window)
        preview_window.grab_set()
        
        # Create layout
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text=f"Data extracted from: {os.path.basename(filename)}", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=5)
        
        # Data preview
        data_frame = ttk.LabelFrame(main_frame, text="Extracted Ratings", padding=10)
        data_frame.pack(fill='both', expand=True, pady=5)
        
        # Create tree view for data
        columns = ['Sample'] + self.metrics
        tree = ttk.Treeview(data_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
            
        # Populate tree with extracted data
        for sample_name, sample_data in extracted_data.items():
            values = [sample_name]
            for metric in self.metrics:
                values.append(sample_data.get(metric, 'N/A'))
            tree.insert('', 'end', values=values)
            
        tree.pack(fill='both', expand=True)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        def load_data():
            # Load the extracted data
            for sample_name, sample_data in extracted_data.items():
                self.samples[sample_name] = sample_data
                
            # Update UI
            self.update_sample_combo()
            self.update_sample_checkboxes()
            
            # Select first sample
            if self.samples:
                first_sample = list(self.samples.keys())[0]
                self.sample_var.set(first_sample)
                self.load_sample_data(first_sample)
                
            self.update_plot()
            preview_window.destroy()
            
            messagebox.showinfo("Success", f"Loaded {len(extracted_data)} samples from image!")
            
        def manual_adjust():
            # Load data but keep preview open for manual adjustment
            load_data()
            # Keep preview window open so user can see what was extracted
            
        ttk.Button(button_frame, text="Load Data", command=load_data).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Load & Adjust", command=manual_adjust).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=preview_window.destroy).pack(side='right', padx=5)

class FormImageProcessor:
    """Process scanned sensory evaluation forms to extract data."""
    
    def __init__(self):
        # Define the expected form structure
        self.sample_positions = [
            # Top row (samples 1-3)
            {"row": 0, "col": 0, "name": "Sample 1"},
            {"row": 0, "col": 1, "name": "Sample 2"}, 
            {"row": 0, "col": 2, "name": "Sample 3"},
            # Bottom row (samples 4-6)
            {"row": 1, "col": 0, "name": "Sample 4"},
            {"row": 1, "col": 1, "name": "Sample 5"},
            {"row": 1, "col": 2, "name": "Sample 6"}
        ]
        
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        # Only lazy load cv2
        self.cv2 = None
        
    def _ensure_cv2_loaded(self):
        """Ensure cv2 is loaded."""
        if self.cv2 is None:
            self.cv2 = _lazy_import_cv2()
        return self.cv2 is not None
        
    def preprocess_image(self, image_path):
        """Preprocess the scanned form image."""
        # Only need to load cv2
        if not self._ensure_cv2_loaded():
            raise ImportError("OpenCV is required for image processing")
            
        cv2 = self.cv2
        
        print(f"DEBUG: Processing image: {image_path}")
        print(f"DEBUG: cv2 loaded: {cv2 is not None}")
        
        # Read image
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Could not load image: {image_path}")
            
        print(f"DEBUG: Image loaded successfully, shape: {img.shape}")
            
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Enhance contrast
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # Remove noise
        denoised = cv2.bilateralFilter(enhanced, 9, 75, 75)
        
        # Detect and correct skew
        corrected = self.correct_skew(denoised)
        
        return corrected
        
    def correct_skew(self, image):
        """Detect and correct image skew."""
        cv2 = self.cv2
    
        if cv2 is None:
            print("WARNING: OpenCV not loaded, skipping skew correction")
            return image
    
        try:
            # Find edges
            edges = cv2.Canny(image, 50, 150, apertureSize=3)
        
            # Find lines using Hough transform
            lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=100)
        
            if lines is not None and len(lines) > 0:
                # Calculate skew angle
                angles = []
                for line in lines[:20]:  # Use first 20 lines
                    # Fix the unpacking issue
                    if len(line) > 0 and len(line[0]) >= 2:
                        rho, theta = line[0]  # Unpack from line[0] not line directly
                        angle = theta * 180 / np.pi
                        # Normalize to [-45, 45] range
                        if angle > 45:
                            angle = angle - 90
                        elif angle < -45:
                            angle = angle + 90
                        angles.append(angle)
            
                # Use median angle to avoid outliers
                if angles:
                    skew_angle = np.median(angles)
                    print(f"DEBUG: Detected skew angle: {skew_angle:.2f} degrees")
                
                    # Rotate image if skew is significant
                    if abs(skew_angle) > 0.5:
                        return self.rotate_image(image, skew_angle)
            else:
                print("DEBUG: No lines detected for skew correction")
            
        except Exception as e:
            print(f"WARNING: Skew correction failed: {e}")
        
        return image
        
    def rotate_image(self, image, angle):
        """Rotate image by specified angle."""
        cv2 = self.cv2
        
        if cv2 is None:
            return image
            
        try:
            (h, w) = image.shape[:2]
            center = (w // 2, h // 2)
            
            # Calculate rotation matrix
            M = cv2.getRotationMatrix2D(center, angle, 1.0)
            
            # Perform rotation
            rotated = cv2.warpAffine(image, M, (w, h), 
                                   flags=cv2.INTER_CUBIC, 
                                   borderMode=cv2.BORDER_REPLICATE)
            return rotated
        except Exception as e:
            print(f"WARNING: Image rotation failed: {e}")
            return image
        
    def find_form_regions(self, image):
        """Find the sample grid regions in the form."""
        height, width = image.shape
    
        print(f"DEBUG: Image dimensions: {width}x{height}")
    
        # Define regions for 2x2 grid (4 samples total)
        # Top row (samples 1-2)
        top_region = {
            'y_start': int(height * 0.25),   # After header (adjusted)
            'y_end': int(height * 0.60),     # Middle of form
            'samples': [
                {'x_start': int(width * 0.20), 'x_end': int(width * 0.60)},  # Sample 1 (left)
                {'x_start': int(width * 0.60), 'x_end': int(width * 0.95)}   # Sample 2 (right)
            ]
        }
    
        # Bottom row (samples 3-4)  
        bottom_region = {
            'y_start': int(height * 0.60),   # Middle of form
            'y_end': int(height * 0.90),     # Before sample codes (adjusted)
            'samples': [
                {'x_start': int(width * 0.20), 'x_end': int(width * 0.60)},  # Sample 3 (left)
                {'x_start': int(width * 0.60), 'x_end': int(width * 0.95)}   # Sample 4 (right)
            ]
        }
    
        return [top_region, bottom_region]
        
    def detect_marked_circles(self, image, region):
        """Detect marked circles in a specific region."""
        cv2 = self.cv2
    
        if cv2 is None:
            return []
    
        try:
            # Extract region
            roi = image[region['y_start']:region['y_end'], 
                       region['x_start']:region['x_end']]
        
            if roi.size == 0:
                return []
        
            print(f"DEBUG: ROI size: {roi.shape}")
        
            # Apply threshold to highlight dark marks
            _, binary = cv2.threshold(roi, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        
            # Find contours
            contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
            marked_positions = []
        
            print(f"DEBUG: Found {len(contours)} contours")
        
            for i, contour in enumerate(contours):
                # Calculate contour properties
                area = cv2.contourArea(contour)
                perimeter = cv2.arcLength(contour, True)
            
                # Adjusted size filters for better detection
                if area > 100 and area < 5000:  # Increased minimum area
                    # Calculate circularity
                    if perimeter > 0:
                        circularity = 4 * np.pi * area / (perimeter * perimeter)
                    
                        if circularity > 0.2:  # More lenient circularity
                            # Get bounding box
                            x, y, w, h = cv2.boundingRect(contour)
                            center_x = x + w // 2
                            center_y = y + h // 2
                        
                            # Check aspect ratio (should be roughly square for circles)
                            aspect_ratio = float(w) / h if h > 0 else 0
                            if 0.5 <= aspect_ratio <= 2.0:  # Allow some variation
                                marked_positions.append({
                                    'x': center_x,
                                    'y': center_y,
                                    'area': area,
                                    'circularity': circularity,
                                    'aspect_ratio': aspect_ratio
                                })
                                print(f"DEBUG: Found mark at ({center_x}, {center_y}), area={area:.0f}, circularity={circularity:.2f}")
        
            return marked_positions
        
        except Exception as e:
            print(f"WARNING: Circle detection failed: {e}")
            return []
        
    def extract_ratings_from_sample(self, image, sample_region):
        """Extract ratings for all attributes from a single sample region."""
        roi_height = sample_region['y_end'] - sample_region['y_start']
        attribute_height = roi_height // len(self.attributes)
        
        ratings = {}
        
        for i, attribute in enumerate(self.attributes):
            # Define attribute region
            attr_region = {
                'y_start': sample_region['y_start'] + i * attribute_height,
                'y_end': sample_region['y_start'] + (i + 1) * attribute_height,
                'x_start': sample_region['x_start'],
                'x_end': sample_region['x_end']
            }
            
            # Find marked circles in this attribute row
            marked_circles = self.detect_marked_circles(image, attr_region)
            
            # Convert circle positions to ratings (1-9)
            if marked_circles:
                # Sort by x position (left to right = 1 to 9)
                marked_circles.sort(key=lambda c: c['x'])
                
                # Take the most prominent mark (largest area)
                best_mark = max(marked_circles, key=lambda c: c['area'])
                
                # Calculate rating based on x position
                region_width = attr_region['x_end'] - attr_region['x_start']
                if region_width > 0:
                    relative_x = best_mark['x'] / region_width
                    rating = min(9, max(1, int(relative_x * 9) + 1))
                else:
                    rating = 5
                
                ratings[attribute] = rating
            else:
                ratings[attribute] = 5  # Default if no mark detected
                
        return ratings
        
    def process_form_image(self, image_path):
        """Process a complete form image and extract all data."""
        try:
            print(f"DEBUG: Starting to process image: {image_path}")
        
            # Preprocess image
            processed_img = self.preprocess_image(image_path)
            print("DEBUG: Image preprocessing completed")
        
            # Find form regions
            regions = self.find_form_regions(processed_img)
            print(f"DEBUG: Found {len(regions)} regions")
        
            extracted_data = {}
            sample_counter = 1
        
            # Process each region (top and bottom) - should give us exactly 4 samples
            for region_idx, region in enumerate(regions):
                print(f"DEBUG: Processing region {region_idx + 1}")
                for sample_idx, sample_region in enumerate(region['samples']):
                    if sample_counter > 4:  # Limit to 4 samples
                        break
                    
                    # Define full sample region
                    full_sample_region = {
                        'y_start': region['y_start'],
                        'y_end': region['y_end'],
                        'x_start': sample_region['x_start'],
                        'x_end': sample_region['x_end']
                    }
                
                    print(f"DEBUG: Processing sample {sample_counter}")
                    print(f"DEBUG: Sample region: y={full_sample_region['y_start']}-{full_sample_region['y_end']}, x={full_sample_region['x_start']}-{full_sample_region['x_end']}")
                
                    # Extract ratings for this sample
                    ratings = self.extract_ratings_from_sample(processed_img, full_sample_region)
                
                    sample_name = f"Sample {sample_counter}"
                    extracted_data[sample_name] = ratings
                    extracted_data[sample_name]['comments'] = ''  # OCR can be added later
                
                    print(f"DEBUG: Sample {sample_counter} ratings: {ratings}")
                
                    sample_counter += 1
                
            print(f"DEBUG: Processing completed. Extracted {len(extracted_data)} samples")
            return extracted_data, processed_img
        
        except Exception as e:
            print(f"ERROR: Image processing failed: {str(e)}")
            import traceback
            traceback.print_exc()
            raise Exception(f"Error processing form image: {str(e)}")
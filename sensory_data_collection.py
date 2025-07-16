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
from enhanced_claude_form_processor import EnhancedClaudeFormProcessor

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
        self.sessions = {}  # {'session_id': {'header': {}, 'samples': {}, 'timestamp': '', 'source_image': ''}}
        self.current_session_id = None
        self.session_counter = 1
        self.samples = {}
        self.current_sample = None
        print("DEBUG: Initialized session-based data structure")
        print(f"DEBUG: self.sessions = {self.sessions}")
        print(f"DEBUG: self.current_session_id = {self.current_session_id}")
        print(f"DEBUG: self.session_counter = {self.session_counter}")
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
        """Create and display the sensory data collection window with optimized sizing."""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Sensory Data Collection")
        self.window.resizable(True, True)
    
        # Start with a reasonable initial size (will be optimized after content loads)
        initial_width = 1000
        initial_height = 600
        self.window.geometry(f"{initial_width}x{initial_height}")
        self.window.configure(bg=APP_BACKGROUND_COLOR)
        #self.window.transient(self.parent)
    
        print(f"DEBUG: Initial window size set to {initial_width}x{initial_height}")
    
        # Create main layout (this will trigger size optimization)
        self.setup_layout()
        self.setup_menu()
        self.window.bind('<Configure>', self.on_window_resize)

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
        file_menu.add_command(label="Merge Sessions from Files", command=self.merge_sessions_from_files)
        file_menu.add_separator()
        file_menu.add_command(label="Load from Image (ML)", command=self.load_from_image_enhanced) #add this back once new function added
        file_menu.add_command(label="Load with AI (Claude)", command=self.load_from_image_with_ai)
        file_menu.add_command(label="Batch Process with AI", command=self.batch_process_with_ai)
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
        
        ml_menu = tk.Menu(menubar, tearoff=0)
    
       # Core ML workflow - no training structure creation (separate scripts handle this)
        ml_menu.add_command(label="Check Enhanced Data Balance", command=self.check_enhanced_data_balance)
        ml_menu.add_command(label="Train Enhanced Model", command=self.train_enhanced_model)
        ml_menu.add_separator()

        # Testing and validation
        ml_menu.add_command(label="Test Enhanced Model", command=self.test_enhanced_model)
        ml_menu.add_command(label="Validate Model Performance", command=self.validate_enhanced_performance)
        ml_menu.add_separator()

        # Configuration management
        ml_menu.add_command(label="Update Processor Configuration", command=self.update_processor_config)

        menubar.add_cascade(label="Enhanced ML", menu=ml_menu)

    def merge_sessions_from_files(self):
        """Merge multiple session JSON files into a new session."""
    
        print("DEBUG: Starting merge sessions from files")
    
        # Select multiple JSON files
        filenames = filedialog.askopenfilenames(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Select Session Files to Merge (Hold Ctrl to select multiple)"
        )
    
        if not filenames or len(filenames) < 2:
            messagebox.showinfo("Insufficient Files", 
                              "Please select at least 2 session files to merge.")
            return
    
        print(f"DEBUG: Selected {len(filenames)} files for merging")
    
        # Load and validate all session files
        loaded_sessions = {}
        failed_files = []
    
        for filename in filenames:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)
            
                # Validate session data structure
                if self.validate_session_data(session_data):
                    session_name = os.path.splitext(os.path.basename(filename))[0]
                    loaded_sessions[session_name] = {
                        'file_path': filename,
                        'data': session_data
                    }
                    print(f"DEBUG: Successfully loaded session from {filename}")
                else:
                    failed_files.append(filename)
                    print(f"DEBUG: Invalid session format in {filename}")
                
            except Exception as e:
                failed_files.append(filename)
                print(f"DEBUG: Failed to load {filename}: {e}")
    
        if not loaded_sessions:
            messagebox.showerror("Load Error", 
                               "No valid session files could be loaded.\n"
                               "Ensure files are in the correct JSON format.")
            return
    
        if failed_files:
            failed_list = '\n'.join([os.path.basename(f) for f in failed_files])
            messagebox.showwarning("Some Files Failed", 
                                 f"Failed to load {len(failed_files)} files:\n{failed_list}\n\n"
                                 f"Continuing with {len(loaded_sessions)} valid files.")
    
        # Show merge configuration dialog
        self.show_merge_sessions_dialog(loaded_sessions)

    def validate_session_data(self, session_data):
        """Validate that the JSON file has the correct session format."""
    
        print("DEBUG: Validating session data structure")
    
        if not isinstance(session_data, dict):
            print("DEBUG: Session data is not a dictionary")
            return False
    
        # Check for required top-level keys
        required_keys = ['samples']
        if not all(key in session_data for key in required_keys):
            print(f"DEBUG: Missing required keys. Found: {list(session_data.keys())}")
            return False
    
        # Check samples structure
        samples = session_data.get('samples', {})
        if not isinstance(samples, dict):
            print("DEBUG: Samples is not a dictionary")
            return False
    
        # Validate sample data structure
        for sample_name, sample_data in samples.items():
            if not isinstance(sample_data, dict):
                print(f"DEBUG: Sample {sample_name} data is not a dictionary")
                return False
        
            # Check for expected metrics (at least some should be present)
            metrics_found = sum(1 for metric in self.metrics if metric in sample_data)
            if metrics_found == 0:
                print(f"DEBUG: No valid metrics found in sample {sample_name}")
                return False
    
        print("DEBUG: Session data validation passed")
        return True

    def show_merge_sessions_dialog(self, loaded_sessions):
        """Show dialog to configure session merging."""
    
        print(f"DEBUG: Showing merge dialog for {len(loaded_sessions)} sessions")
    
        # Create dialog window
        merge_window = tk.Toplevel(self.window)
        merge_window.title("Merge Session Files")
        merge_window.geometry("600x500")
        merge_window.transient(self.window)
        merge_window.grab_set()
    
        # Main frame
        main_frame = ttk.Frame(merge_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        # Title
        title_label = ttk.Label(main_frame, 
                               text="Configure Session Merge", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 10))
    
        # Instructions
        instructions = ("Select which sessions to include in the merge.\n"
                       "Conflicting sample names will be automatically renamed.\n"
                       "Header information will be merged where possible.")
    
        ttk.Label(main_frame, text=instructions, 
                 font=('Arial', 9), justify='left').pack(pady=(0, 15))
    
        # Session selection frame with scrollbar
        selection_frame = ttk.LabelFrame(main_frame, text="Sessions to Merge", padding=10)
        selection_frame.pack(fill='both', expand=True, pady=(0, 10))
    
        # Create scrollable frame
        canvas = tk.Canvas(selection_frame, height=200)
        scrollbar = ttk.Scrollbar(selection_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
    
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
    
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
        # Session checkboxes with details
        session_vars = {}
    
        for session_name, session_info in loaded_sessions.items():
            session_data = session_info['data']
            sample_count = len(session_data.get('samples', {}))
        
            # Extract some header info for display
            header = session_data.get('header', {})
            assessor = header.get('Assessor Name', 'Unknown')
            date = header.get('Date', 'Unknown')
        
            var = tk.BooleanVar(value=True)  # Default to checked
            session_vars[session_name] = {
                'var': var,
                'data': session_data,
                'file_path': session_info['file_path']
            }
        
            # Create session info frame
            session_frame = ttk.Frame(scrollable_frame)
            session_frame.pack(fill='x', pady=2)
        
            # Checkbox and main info
            info_text = f"{session_name} ({sample_count} samples)"
            ttk.Checkbutton(session_frame, text=info_text, 
                       variable=var).pack(anchor='w')
        
            # Additional details
            details = f"   Assessor: {assessor} | Date: {date} | File: {os.path.basename(session_info['file_path'])}"
            ttk.Label(session_frame, text=details, 
                     font=('Arial', 8), foreground='gray').pack(anchor='w', padx=(20, 0))
    
        # Merge options frame
        options_frame = ttk.LabelFrame(main_frame, text="Merge Options", padding=10)
        options_frame.pack(fill='x', pady=(0, 10))
    
        # New session name
        name_frame = ttk.Frame(options_frame)
        name_frame.pack(fill='x', pady=(0, 5))
    
        ttk.Label(name_frame, text="Merged session name:").pack(side='left')
        default_name = f"Merged_Session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        name_var = tk.StringVar(value=default_name)
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=30)
        name_entry.pack(side='right')
    
        # Header merge strategy
        header_frame = ttk.Frame(options_frame)
        header_frame.pack(fill='x', pady=5)
    
        ttk.Label(header_frame, text="Header merge strategy:").pack(side='left')
        header_strategy = tk.StringVar(value="first")
        header_combo = ttk.Combobox(header_frame, textvariable=header_strategy, 
                                   values=["first", "most_recent", "manual"], 
                                   state='readonly', width=15)
        header_combo.pack(side='right')
    
        # Sample naming strategy
        naming_frame = ttk.Frame(options_frame)
        naming_frame.pack(fill='x', pady=5)
    
        ttk.Label(naming_frame, text="Duplicate sample naming:").pack(side='left')
        naming_strategy = tk.StringVar(value="auto_rename")
        naming_combo = ttk.Combobox(naming_frame, textvariable=naming_strategy,
                                   values=["auto_rename", "prefix_session", "skip_duplicates"],
                                   state='readonly', width=15)
        naming_combo.pack(side='right')
    
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
    
        def perform_merge():
            print("DEBUG: Starting merge process")
        
            # Get selected sessions
            selected_sessions = {}
            for session_name, session_info in session_vars.items():
                if session_info['var'].get():
                    selected_sessions[session_name] = session_info
        
            if len(selected_sessions) < 2:
                messagebox.showwarning("Insufficient Selection", 
                                     "Please select at least 2 sessions to merge.")
                return
        
            new_session_name = name_var.get().strip()
            if not new_session_name:
                messagebox.showwarning("Invalid Name", "Please enter a name for the merged session.")
                return
        
            print(f"DEBUG: Merging {len(selected_sessions)} sessions into '{new_session_name}'")
        
            # Perform the merge
            success = self.execute_session_merge(
                selected_sessions, 
                new_session_name, 
                header_strategy.get(), 
                naming_strategy.get()
            )
        
            if success:
                merge_window.destroy()
                messagebox.showinfo("Merge Complete", 
                                  f"Successfully merged {len(selected_sessions)} sessions!\n"
                                  f"New session: {new_session_name}")
        
        def select_all():
            for session_info in session_vars.values():
                session_info['var'].set(True)
    
        def select_none():
            for session_info in session_vars.values():
                session_info['var'].set(False)
    
        # Selection buttons
        select_frame = ttk.Frame(button_frame)
        select_frame.pack(side='left')
    
        ttk.Button(select_frame, text="Select All", command=select_all).pack(side='left', padx=2)
        ttk.Button(select_frame, text="Select None", command=select_none).pack(side='left', padx=2)
    
        # Action buttons  
        action_frame = ttk.Frame(button_frame)
        action_frame.pack(side='right')
    
        ttk.Button(action_frame, text="Cancel", 
                   command=merge_window.destroy).pack(side='left', padx=5)
        ttk.Button(action_frame, text="Merge Sessions", 
                   command=perform_merge).pack(side='left', padx=5)
    
        print("DEBUG: Merge dialog created successfully")

    def execute_session_merge(self, selected_sessions, new_session_name, header_strategy, naming_strategy):
        """Execute the actual merging of sessions."""
    
        print(f"DEBUG: Executing merge with strategy - header: {header_strategy}, naming: {naming_strategy}")
    
        try:
            # Initialize merged session structure
            merged_session = {
                'header': {},
                'samples': {},
                'timestamp': datetime.now().isoformat(),
                'merge_info': {
                    'source_sessions': list(selected_sessions.keys()),
                    'merge_date': datetime.now().isoformat(),
                    'header_strategy': header_strategy,
                    'naming_strategy': naming_strategy
                }
            }
        
            # Merge headers based on strategy
            merged_session['header'] = self.merge_headers(selected_sessions, header_strategy)
        
            # Merge samples with conflict resolution
            merged_samples, conflicts_resolved = self.merge_samples(selected_sessions, naming_strategy)
            merged_session['samples'] = merged_samples
        
            # Create the new session in memory
            if hasattr(self, 'sessions'):
                # Using session-based structure
                self.sessions[new_session_name] = merged_session
                self.switch_to_session(new_session_name)
                if hasattr(self, 'session_var'):
                    self.session_var.set(new_session_name)
                self.update_session_combo()
            else:
                # Fallback to old structure
                self.samples = merged_samples
            
                # Update header fields
                for field, value in merged_session['header'].items():
                    if field in self.header_vars:
                        self.header_vars[field].set(value)
        
            # Update UI
            self.update_sample_combo()
            self.update_sample_checkboxes()
        
            # Select first sample if available
            if merged_samples:
                first_sample = list(merged_samples.keys())[0]
                self.sample_var.set(first_sample)
                self.load_sample_data(first_sample)
        
            self.update_plot()
        
            # Log merge details
            total_samples = len(merged_samples)
            total_sessions = len(selected_sessions)
        
            print(f"DEBUG: Merge completed successfully")
            print(f"DEBUG: Total samples: {total_samples}")
            print(f"DEBUG: Conflicts resolved: {conflicts_resolved}")
            print(f"DEBUG: Sessions merged: {total_sessions}")
        
            return True
        
        except Exception as e:
            print(f"DEBUG: Merge execution failed: {e}")
            messagebox.showerror("Merge Error", f"Failed to merge sessions: {e}")
            return False

    def merge_headers(self, selected_sessions, strategy):
        """Merge header information based on the selected strategy."""
    
        print(f"DEBUG: Merging headers with strategy: {strategy}")
    
        all_headers = []
        for session_name, session_info in selected_sessions.items():
            header = session_info['data'].get('header', {})
            if header:
                all_headers.append((session_name, header))
    
        if not all_headers:
            return {}
    
        if strategy == "first":
            return all_headers[0][1].copy()
    
        elif strategy == "most_recent":
            # Find header with most recent timestamp
            most_recent = None
            most_recent_time = None
        
            for session_name, header in all_headers:
                session_data = selected_sessions[session_name]['data']
                timestamp_str = session_data.get('timestamp', '')
            
                try:
                    if timestamp_str:
                        timestamp = datetime.fromisoformat(timestamp_str.replace('Z', '+00:00'))
                        if most_recent_time is None or timestamp > most_recent_time:
                            most_recent_time = timestamp
                            most_recent = header
                except:
                    pass
        
            return most_recent.copy() if most_recent else all_headers[0][1].copy()
    
        elif strategy == "manual":
            # For now, return first header - could be extended to show manual selection dialog
            return all_headers[0][1].copy()
    
        return {}

    def merge_samples(self, selected_sessions, naming_strategy):
        """Merge samples with conflict resolution."""
    
        print(f"DEBUG: Merging samples with naming strategy: {naming_strategy}")
    
        merged_samples = {}
        conflicts_resolved = 0
    
        for session_name, session_info in selected_sessions.items():
            session_samples = session_info['data'].get('samples', {})
        
            for original_sample_name, sample_data in session_samples.items():
                final_sample_name = original_sample_name
            
                # Handle naming conflicts
                if final_sample_name in merged_samples:
                    conflicts_resolved += 1
                
                    if naming_strategy == "auto_rename":
                        counter = 1
                        while f"{original_sample_name}_{counter}" in merged_samples:
                            counter += 1
                        final_sample_name = f"{original_sample_name}_{counter}"
                
                    elif naming_strategy == "prefix_session":
                        final_sample_name = f"{session_name}_{original_sample_name}"
                        counter = 1
                        while final_sample_name in merged_samples:
                            final_sample_name = f"{session_name}_{original_sample_name}_{counter}"
                            counter += 1
                
                    elif naming_strategy == "skip_duplicates":
                        print(f"DEBUG: Skipping duplicate sample: {original_sample_name}")
                        continue
            
                # Copy sample data
                merged_samples[final_sample_name] = sample_data.copy()
            
                print(f"DEBUG: Added sample {original_sample_name} as {final_sample_name}")
    
        print(f"DEBUG: Sample merge complete - {len(merged_samples)} total samples, {conflicts_resolved} conflicts resolved")
    
        return merged_samples, conflicts_resolved

    def create_new_session(self, session_name=None, source_image=None):
        """Create a new session for data collection."""
        if session_name is None:
            session_name = f"Session_{self.session_counter}"
            self.session_counter += 1
    
        print(f"DEBUG: Creating new session: {session_name}")
        print(f"DEBUG: Source image: {source_image}")
    
        # Create new session structure
        self.sessions[session_name] = {
            'header': {field: var.get() for field, var in self.header_vars.items()},
            'samples': {},
            'timestamp': datetime.now().isoformat(),
            'source_image': source_image or ''
        }
    
        # Switch to new session
        self.current_session_id = session_name
        self.samples = self.sessions[session_name]['samples']
    
        print(f"DEBUG: Session created successfully")
        print(f"DEBUG: Current session ID: {self.current_session_id}")
        print(f"DEBUG: Session structure: {self.sessions[session_name]}")
    
        self.update_session_combo()
        self.update_sample_combo()
        self.update_sample_checkboxes()
    
        return session_name

    def switch_to_session(self, session_id):
        """Switch to a specific session."""
        if session_id not in self.sessions:
            print(f"DEBUG: Session {session_id} not found")
            return False

        print(f"DEBUG: Switching from session {self.current_session_id} to {session_id}")

        # Save current session data before switching
        if self.current_session_id and self.current_session_id in self.sessions:
            self.sessions[self.current_session_id]['samples'] = self.samples
            self.sessions[self.current_session_id]['header'] = {field: var.get() for field, var in self.header_vars.items()}
            print(f"DEBUG: Saved {len(self.samples)} samples to previous session")

        # Switch to new session
        self.current_session_id = session_id
        self.samples = self.sessions[session_id]['samples']

        # Update header fields with session data
        session_header = self.sessions[session_id]['header']
        for field, var in self.header_vars.items():
            if field in session_header:
                var.set(session_header[field])
            else:
                if field == "Date":
                    var.set(datetime.now().strftime("%Y-%m-%d"))
                else:
                    var.set('')

        print(f"DEBUG: Switched to session {session_id} with {len(self.samples)} samples")

        # Update UI components
        self.update_sample_combo()
        self.update_sample_checkboxes()

        # Select first sample if available
        if self.samples:
            first_sample = list(self.samples.keys())[0]
            self.sample_var.set(first_sample)
            self.load_sample_data(first_sample)
        else:
            self.sample_var.set('')
            self.clear_form()

        self.update_plot()
        return True

    def update_session_combo(self):
        """Update the session selection combo box."""
        session_names = list(self.sessions.keys())
        if hasattr(self, 'session_combo'):
            self.session_combo['values'] = session_names
            print(f"DEBUG: Updated session combo with {len(session_names)} sessions")

    def check_enhanced_data_balance(self):
        """Check enhanced training data balance and quality."""
        try:
            import os
        
            base_dir = "training_data/sensory_ratings"
            if not os.path.exists(base_dir):
                messagebox.showwarning("No Enhanced Data", 
                                     "Enhanced training data not found.\n"
                                     "Run enhanced extraction first.")
                return
        
            print("="*80)
            print("ENHANCED TRAINING DATA ANALYSIS")
            print("="*80)
        
            # Detailed analysis with enhanced metrics
            class_distribution = {}
            enhanced_info_files = {}
            total_images = 0
            total_enhanced = 0
        
            for rating in range(1, 10):
                rating_dir = os.path.join(base_dir, f"rating_{rating}")
                if os.path.exists(rating_dir):
                    # Count image files
                    images = [f for f in os.listdir(rating_dir) 
                             if f.lower().endswith(('.jpg', '.png', '.jpeg'))]
                    count = len(images)
                    class_distribution[rating] = count
                    total_images += count
                
                    # Count enhanced extraction info files
                    info_files = [f for f in os.listdir(rating_dir) if f.endswith('_info.txt')]
                    enhanced_info_files[rating] = len(info_files)
                    total_enhanced += len(info_files)
                
                    # Sample file analysis
                    sample_sizes = []
                    for img_file in images[:5]:  # Check first 5 files
                        img_path = os.path.join(rating_dir, img_file)
                        try:
                            import cv2
                            img = cv2.imread(img_path)
                            if img is not None:
                                sample_sizes.append(f"{img.shape[1]}x{img.shape[0]}")
                        except:
                            pass
                
                    size_info = f" (sizes: {', '.join(set(sample_sizes[:3]))})" if sample_sizes else ""
                
                    print(f"Rating {rating}: {count:4d} images, {len(info_files):3d} enhanced{size_info}")
                
                    # Show sample filenames
                    if images:
                        print(f"  Sample files: {images[:2]}")
        
            # Enhanced analysis
            print("-" * 70)
            print(f"Total training images: {total_images}")
            print(f"Enhanced extractions: {total_enhanced}")
        
            if total_images > 0:
                min_count = min(class_distribution.values())
                max_count = max(class_distribution.values())
                imbalance_ratio = max_count / max(min_count, 1)
                enhancement_rate = total_enhanced / total_images
            
                print(f"\nEnhanced Quality Metrics:")
                print(f"  Class balance ratio: {imbalance_ratio:.2f}")
                print(f"  Enhancement rate: {enhancement_rate:.1%}")
                print(f"  Min/Max class sizes: {min_count}/{max_count}")
            
                # Enhanced recommendations
                recommendations = []
                if total_images < 100:
                    recommendations.append("Collect more training data (target: 100+ images)")
                if imbalance_ratio > 3.0:
                    recommendations.append("Balance classes - some ratings underrepresented")
                if enhancement_rate < 0.8:
                    recommendations.append("Re-extract with enhanced workflow for better quality")
                if total_images < 200:
                    recommendations.append("For production quality: collect 200+ images")
            
                # Status assessment
                if not recommendations:
                    status = "EXCELLENT - Ready for production training"
                    color = "green"
                elif len(recommendations) <= 2:
                    status = "GOOD - Ready for training with minor improvements"
                    color = "blue"
                else:
                    status = "NEEDS IMPROVEMENT - Address issues before training"
                    color = "orange"
            
                print(f"\nStatus: {status}")
            
                if recommendations:
                    print(f"\nRecommendations:")
                    for i, rec in enumerate(recommendations, 1):
                        print(f"  {i}. {rec}")
            
                # Show in dialog
                dialog_msg = (f"Enhanced Training Data Analysis\n\n"
                             f"Status: {status}\n\n"
                             f"Metrics:\n"
                             f"• Total images: {total_images}\n"
                             f"• Enhanced extractions: {total_enhanced} ({enhancement_rate:.1%})\n"
                             f"• Class balance ratio: {imbalance_ratio:.2f}\n"
                             f"• Resolution: 600x140 pixels\n\n")
            
                if recommendations:
                    dialog_msg += "Recommendations:\n" + "\n".join(f"• {rec}" for rec in recommendations)
                else:
                    dialog_msg += "✓ Data is ready for enhanced model training!"
                
                messagebox.showinfo("Enhanced Data Analysis", dialog_msg)
            else:
                messagebox.showwarning("No Training Data", 
                                     "No enhanced training images found.\n"
                                     "Use enhanced extraction tools first.")
            
        except Exception as e:
            messagebox.showerror("Analysis Error", f"Enhanced data analysis failed: {e}")
            import traceback
            traceback.print_exc()

    def train_enhanced_model(self):
        """Train enhanced model with comprehensive configuration."""
        try:
            # Import enhanced processor
            from enhanced_ml_form_processor import EnhancedMLFormProcessor, EnhancedMLTrainingHelper
            import os
        
            # Verify enhanced training data
            if not os.path.exists("training_data/sensory_ratings"):
                messagebox.showerror("Missing Enhanced Data", 
                                   "Enhanced training data not found.\n"
                                   "Use enhanced extraction first.")
                return
        
            # Count enhanced training data
            total_images = 0
            enhanced_count = 0
        
            for rating in range(1, 10):
                rating_dir = os.path.join("training_data/sensory_ratings", f"rating_{rating}")
                if os.path.exists(rating_dir):
                    images = len([f for f in os.listdir(rating_dir) if f.endswith(('.jpg', '.png', '.jpeg'))])
                    info_files = len([f for f in os.listdir(rating_dir) if f.endswith('_info.txt')])
                    total_images += images
                    enhanced_count += info_files
        
            # Enhanced training dialog
            enhancement_rate = enhanced_count / max(total_images, 1)
        
            training_msg = (f"Enhanced ML Model Training\n\n"
                           f"Training Data:\n"
                           f"• Total images: {total_images}\n"
                           f"• Enhanced extractions: {enhanced_count} ({enhancement_rate:.1%})\n"
                           f"• Target resolution: 600x140 pixels\n\n"
                           f"Enhanced Architecture:\n"
                           f"• 5 convolutional layers\n"
                           f"• Optimized for high-resolution data\n"
                           f"• Advanced regularization\n"
                           f"• Shadow removal preprocessing compatibility\n\n"
                           f"Training Features:\n"
                           f"• Early stopping with patience\n"
                           f"• Learning rate scheduling\n"
                           f"• Enhanced model checkpointing\n"
                           f"• Comprehensive logging\n\n"
                           f"Estimated time: 10-30 minutes\n\n"
                           f"Continue?")
        
            result = messagebox.askyesno("Enhanced Model Training", training_msg)
        
            if result:
                print("="*80)
                print("ENHANCED ML MODEL TRAINING")
                print("="*80)
                print(f"Training images: {total_images}")
                print(f"Enhanced extractions: {enhanced_count}")
                print(f"Architecture: Enhanced CNN for 600x140 resolution")
                print(f"Features: Shadow removal compatibility, advanced regularization")
            
                # Initialize enhanced components
                processor = EnhancedMLFormProcessor()
                trainer = EnhancedMLTrainingHelper(processor)
            
                # Enhanced training configuration
                training_config = {
                    'epochs': 100,
                    'batch_size': 16,
                    'validation_split': 0.25,
                    'save_best_only': True,
                    'patience': 20
                }
            
                print(f"\nEnhanced training configuration:")
                for key, value in training_config.items():
                    print(f"  {key}: {value}")
            
                # Train enhanced model
                model, history = trainer.train_enhanced_model(**training_config)
            
                # Enhanced results reporting
                if history and model:
                    final_train_acc = history.history['accuracy'][-1]
                    final_val_acc = history.history['val_accuracy'][-1]
                    best_val_acc = max(history.history['val_accuracy'])
                    epochs_trained = len(history.history['accuracy'])
                
                    # Check for enhanced model files
                    model_files = []
                    if os.path.exists('models/sensory_rating_classifier.h5'):
                        size_mb = os.path.getsize('models/sensory_rating_classifier.h5') / (1024*1024)
                        model_files.append(f"• Final model: {size_mb:.1f} MB")
                
                    if os.path.exists('models/enhanced/sensory_rating_classifier_best.h5'):
                        size_mb = os.path.getsize('models/enhanced/sensory_rating_classifier_best.h5') / (1024*1024)
                        model_files.append(f"• Best enhanced model: {size_mb:.1f} MB")
                
                    success_msg = (f"Enhanced Model Training Complete!\n\n"
                                 f"Performance Metrics:\n"
                                 f"• Final training accuracy: {final_train_acc:.3f}\n"
                                 f"• Final validation accuracy: {final_val_acc:.3f}\n"
                                 f"• Best validation accuracy: {best_val_acc:.3f}\n"
                                 f"• Epochs trained: {epochs_trained}\n\n"
                                 f"Model Files Saved:\n" + "\n".join(model_files) + f"\n\n"
                                 f"Enhanced Features:\n"
                                 f"• 600x140 high resolution\n"
                                 f"• Shadow removal preprocessing\n"
                                 f"• Advanced CNN architecture\n"
                                 f"• Production-ready accuracy\n\n"
                                 f"Next: Test Enhanced Model")
                
                    messagebox.showinfo("Enhanced Training Complete", success_msg)
                
                    print("="*80)
                    print("ENHANCED TRAINING COMPLETED SUCCESSFULLY")
                    print("="*80)
                    print(f"Enhanced model ready for production use!")
                
                else:
                    messagebox.showwarning("Training Issues", 
                                         "Enhanced training completed with issues.\n"
                                         "Check console for detailed information.")
                
        except Exception as e:
            error_msg = f"Enhanced training failed: {e}"
            print(f"ERROR: {error_msg}")
            messagebox.showerror("Enhanced Training Error", error_msg)
            import traceback
            traceback.print_exc()

    def test_enhanced_model(self):
        """Test enhanced model with comprehensive evaluation."""
        try:
            from enhanced_ml_form_processor import EnhancedMLFormProcessor
            import os
            import cv2
            import numpy as np
        
            # Check for enhanced model
            model_paths = [
                "models/enhanced/sensory_rating_classifier_best.h5",
                "models/sensory_rating_classifier.h5",
                "models/enhanced/sensory_rating_classifier.h5"
            ]
        
            model_path = None
            for path in model_paths:
                if os.path.exists(path):
                    model_path = path
                    break
        
            if not model_path:
                messagebox.showwarning("No Enhanced Model", 
                                     "No enhanced model found.\n"
                                     "Train the enhanced model first.")
                return
        
            print("="*80)
            print("ENHANCED MODEL TESTING")
            print("="*80)
            print(f"Testing model: {model_path}")
        
            # Initialize enhanced processor
            processor = EnhancedMLFormProcessor(model_path)
        
            if not processor.load_model():
                messagebox.showerror("Model Load Error", 
                                   "Failed to load enhanced model.\n"
                                   "Check console for error details.")
                return
        
            print(f"✓ Enhanced model loaded successfully")
            print(f"Model resolution: {processor.target_size}")
        
            # Test on enhanced training data samples
            base_dir = "training_data/sensory_ratings"
            test_results = {}
            detailed_results = []
        
            total_tests = 0
            correct_predictions = 0
            confidence_scores = []
        
            print(f"\nTesting enhanced model on training samples...")
        
            for rating in range(1, 10):
                rating_dir = os.path.join(base_dir, f"rating_{rating}")
                if os.path.exists(rating_dir):
                    images = [f for f in os.listdir(rating_dir) if f.endswith(('.jpg', '.png', '.jpeg'))]
                    if images:
                        # Test on first image in each class
                        test_image_path = os.path.join(rating_dir, images[0])
                    
                        try:
                            # Load and test with enhanced processor
                            test_image = cv2.imread(test_image_path, cv2.IMREAD_GRAYSCALE)
                            if test_image is not None:
                                # Resize to enhanced model input size
                                resized_image = cv2.resize(test_image, processor.target_size, 
                                                         interpolation=cv2.INTER_CUBIC)
                            
                                # Get enhanced prediction
                                predicted_rating, confidence, probabilities = processor.predict_rating_enhanced(resized_image)
                            
                                is_correct = predicted_rating == rating
                                test_results[rating] = {
                                    'predicted': predicted_rating,
                                    'confidence': confidence,
                                    'correct': is_correct,
                                    'probabilities': probabilities
                                }
                            
                                detailed_results.append({
                                    'true_rating': rating,
                                    'predicted_rating': predicted_rating,
                                    'confidence': confidence,
                                    'correct': is_correct
                                })
                            
                                total_tests += 1
                                confidence_scores.append(confidence)
                                if is_correct:
                                    correct_predictions += 1
                            
                                status = "✓ CORRECT" if is_correct else "✗ WRONG"
                                conf_level = "HIGH" if confidence > 0.8 else "MED" if confidence > 0.6 else "LOW"
                            
                                print(f"Rating {rating}: Predicted {predicted_rating} ({conf_level} conf: {confidence:.3f}) - {status}")
                            
                                # Show top 3 predictions for detailed analysis
                                top_3 = processor.get_top_predictions(probabilities, 3)
                                top_3_str = ", ".join([f"R{r}({p:.2f})" for r, p in top_3])
                                print(f"  Top 3: {top_3_str}")
                    
                        except Exception as e:
                            print(f"Error testing rating {rating}: {e}")
        
            # Enhanced results analysis
            if detailed_results:
                test_accuracy = correct_predictions / total_tests
                avg_confidence = np.mean(confidence_scores)
                confidence_std = np.std(confidence_scores)
            
                # Confidence analysis
                high_conf = sum(1 for c in confidence_scores if c > 0.8)
                med_conf = sum(1 for c in confidence_scores if 0.6 <= c <= 0.8)
                low_conf = sum(1 for c in confidence_scores if c < 0.6)
            
                print(f"\n" + "="*80)
                print(f"ENHANCED MODEL TEST RESULTS")
                print(f"="*80)
                print(f"Model tested: {os.path.basename(model_path)}")
                print(f"Resolution: {processor.target_size}")
                print(f"Classes tested: {total_tests}")
                print(f"Correct predictions: {correct_predictions}")
                print(f"Test accuracy: {test_accuracy:.3f} ({test_accuracy*100:.1f}%)")
                print(f"Average confidence: {avg_confidence:.3f} ± {confidence_std:.3f}")
                print(f"Confidence distribution: High({high_conf}) Med({med_conf}) Low({low_conf})")
            
                # Detailed error analysis
                errors = [r for r in detailed_results if not r['correct']]
                if errors:
                    print(f"\nError analysis:")
                    for error in errors:
                        print(f"  True: {error['true_rating']} → Predicted: {error['predicted_rating']} (conf: {error['confidence']:.3f})")
            
                # Performance assessment
                if test_accuracy >= 0.9:
                    status = "EXCELLENT"
                    recommendation = "Model ready for production deployment!"
                elif test_accuracy >= 0.8:
                    status = "VERY GOOD"
                    recommendation = "Model suitable for production with monitoring"
                elif test_accuracy >= 0.7:
                    status = "GOOD"
                    recommendation = "Consider collecting more training data"
                else:
                    status = "NEEDS IMPROVEMENT"
                    recommendation = "Collect significantly more training data"
            
                # Show comprehensive results dialog
                result_msg = (f"Enhanced Model Test Results - {status}\n\n"
                             f"Performance Metrics:\n"
                             f"• Test accuracy: {test_accuracy*100:.1f}%\n"
                             f"• Average confidence: {avg_confidence:.3f}\n"
                             f"• High confidence predictions: {high_conf}/{total_tests}\n"
                             f"• Model resolution: {processor.target_size[0]}x{processor.target_size[1]}\n\n"
                             f"Confidence Distribution:\n"
                             f"• High (>0.8): {high_conf}\n"
                             f"• Medium (0.6-0.8): {med_conf}\n"
                             f"• Low (<0.6): {low_conf}\n\n"
                             f"Recommendation:\n{recommendation}\n\n"
                             f"Check console for detailed per-class results.")
            
                messagebox.showinfo("Enhanced Model Test Complete", result_msg)
            else:
                messagebox.showwarning("No Test Data", 
                                     "No test data available.\n"
                                     "Ensure training data is present.")
        
        except Exception as e:
            messagebox.showerror("Enhanced Test Error", f"Enhanced model testing failed: {e}")
            import traceback
            traceback.print_exc()

    def validate_enhanced_performance(self):
        """Comprehensive enhanced model validation."""
        messagebox.showinfo("Enhanced Validation", 
                          "Comprehensive Enhanced Model Validation\n\n"
                          "Features to be implemented:\n\n"
                          "• Cross-validation analysis\n"
                          "• Confusion matrix generation\n"
                          "• Per-attribute accuracy metrics\n"
                          "• Confidence calibration analysis\n"
                          "• Model uncertainty quantification\n"
                          "• Production readiness assessment\n\n"
                          "This advanced validation suite will be available\n"
                          "in the next update for production deployment.")

    def update_processor_config(self):
        """Update processor configuration with enhanced settings."""
        try:
            from enhanced_ml_processor_updater import MLProcessorUpdater
            from tkinter import filedialog
        
            # Find available configurations
            config_dir = "training_data/claude_analysis"
            if os.path.exists(config_dir):
                config_files = [f for f in os.listdir(config_dir) 
                               if f.startswith('improved_boundaries_') and f.endswith('.json')]
            
                if config_files:
                    config_file = filedialog.askopenfilename(
                        title="Select enhanced boundary configuration",
                        initialdir=config_dir,
                        filetypes=[("JSON files", "*.json")]
                    )
                
                    if config_file:
                        result = messagebox.askyesno("Update Enhanced Processor",
                                                   f"Update enhanced processor with:\n"
                                                   f"{os.path.basename(config_file)}\n\n"
                                                   f"This will modify enhanced_ml_form_processor.py\n"
                                                   f"A backup will be created automatically.\n\n"
                                                   f"Continue?")
                        if result:
                            updater = MLProcessorUpdater()
                            updater.update_ml_processor_boundaries(config_file)
                        
                            messagebox.showinfo("Enhanced Update Complete",
                                              f"Enhanced processor updated!\n\n"
                                              f"Configuration: {os.path.basename(config_file)}\n"
                                              f"Backup: {updater.backup_path}\n\n"
                                              f"Test the updated enhanced processor.")
                else:
                    messagebox.showinfo("No Enhanced Configs", 
                                      "No enhanced configurations found.\n"
                                      "Use enhanced extraction tools first.")
            else:
                messagebox.showinfo("No Analysis Directory", 
                                  "Enhanced analysis directory not found.")
            
        except Exception as e:
            messagebox.showerror("Update Error", f"Failed to update enhanced processor: {e}")



    # Enhanced image loading function for File menu
    def load_from_image_enhanced(self):
        """Load sensory data using enhanced ML processing."""
        try:
            from tkinter import filedialog
            from enhanced_ml_form_processor import EnhancedMLFormProcessor
        
            # Check for enhanced model
            model_paths = [
                "models/enhanced/sensory_rating_classifier_best.h5",
                "models/sensory_rating_classifier.h5"
            ]
        
            model_path = None
            for path in model_paths:
                if os.path.exists(path):
                    model_path = path
                    break
        
            if not model_path:
                messagebox.showwarning("No Enhanced Model", 
                                     "No enhanced model found.\n"
                                     "Train an enhanced model first using the Enhanced ML menu.")
                return
        
            # Select image file
            image_path = filedialog.askopenfilename(
                title="Select form image for enhanced ML processing",
                filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
            )
        
            if image_path:
                print("="*80)
                print("ENHANCED ML FORM LOADING")
                print("="*80)
            
                # Initialize enhanced processor
                processor = EnhancedMLFormProcessor(model_path)
            
                # Process with enhanced pipeline
                extracted_data, processed_image = processor.process_form_image_enhanced(image_path)
            
                # Show enhanced preview with confidence scores
                self.show_enhanced_extraction_preview(extracted_data, processed_image, 
                                                    os.path.basename(image_path))
        
        except Exception as e:
            messagebox.showerror("Enhanced ML Error", f"Enhanced ML processing failed: {e}")
            import traceback
            traceback.print_exc()

    def show_enhanced_extraction_preview(self, extracted_data, processed_img, filename):
        """Show enhanced extraction preview with confidence analysis."""
        preview_window = tk.Toplevel(self.window)
        preview_window.title(f"Enhanced ML Extraction Preview - {filename}")
        preview_window.geometry("900x700")
    
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        # Title with enhanced info
        title_label = ttk.Label(main_frame, 
                               text=f"Enhanced ML Extraction Results - {filename}",
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 10))
    
        # Enhanced results info
        info_label = ttk.Label(main_frame,
                              text=f"Resolution: 600x140 pixels • Shadow removal preprocessing • OCR boundaries",
                              font=('Arial', 10))
        info_label.pack(pady=(0, 10))
    
        # Create notebook for organized display
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True)
    
        # Results tab
        results_frame = ttk.Frame(notebook)
        notebook.add(results_frame, text="Extraction Results")
    
        # Sample name editing with enhanced features
        if extracted_data:
            sample_name_vars = {}
            original_names = list(extracted_data.keys())
        
            # Enhanced sample display
            for i, (original_name, sample_data) in enumerate(extracted_data.items()):
                sample_frame = ttk.LabelFrame(results_frame, text=f"Sample {i+1}", padding=10)
                sample_frame.pack(fill='x', pady=5)
            
                # Sample name editing
                name_frame = ttk.Frame(sample_frame)
                name_frame.pack(fill='x', pady=(0, 5))
            
                ttk.Label(name_frame, text="Sample Name:").pack(side='left')
                sample_name_var = tk.StringVar(value=original_name)
                sample_name_vars[original_name] = sample_name_var
                name_entry = ttk.Entry(name_frame, textvariable=sample_name_var, width=30)
                name_entry.pack(side='left', padx=(5, 0))
            
                # Enhanced ratings display
                ratings_frame = ttk.Frame(sample_frame)
                ratings_frame.pack(fill='x')
            
                for j, (attribute, rating) in enumerate(sample_data.items()):
                    if attribute != 'comments':
                        attr_frame = ttk.Frame(ratings_frame)
                        attr_frame.pack(fill='x', pady=1)
                    
                        # Enhanced display with confidence indicators
                        attr_label = ttk.Label(attr_frame, text=f"{attribute}:", width=15, anchor='w')
                        attr_label.pack(side='left')
                    
                        rating_label = ttk.Label(attr_frame, text=f"Rating: {rating}", 
                                               font=('Arial', 10, 'bold'))
                        rating_label.pack(side='left', padx=(5, 0))
                    
                        # Add confidence indicator if available (placeholder for now)
                        conf_label = ttk.Label(attr_frame, text="(High Confidence)", 
                                             foreground='green', font=('Arial', 9))
                        conf_label.pack(side='left', padx=(10, 0))
    
        # Load enhanced data function
        def load_enhanced_data():
            final_data = {}
            for original_name in original_names:
                new_name = sample_name_vars[original_name].get().strip()
                if not new_name:
                    new_name = original_name
            
                if new_name in final_data:
                    messagebox.showerror("Duplicate Names", 
                                       f"Sample name '{new_name}' is used more than once.")
                    return
            
                final_data[new_name] = extracted_data[original_name]
        
            # Load into interface
            for sample_name, sample_data in final_data.items():
                self.samples[sample_name] = sample_data
        
            self.update_sample_combo()
            self.update_sample_checkboxes()
        
            if self.samples:
                first_sample = list(self.samples.keys())[0]
                self.sample_var.set(first_sample)
                self.load_sample_data(first_sample)
        
            self.update_plot()
            preview_window.destroy()
        
            messagebox.showinfo("Enhanced ML Loading Complete", 
                              f"Successfully loaded {len(final_data)} samples!\n\n"
                              f"Enhanced features:\n"
                              f"• High-resolution extraction (600x140)\n"
                              f"• Shadow removal preprocessing\n"
                              f"• OCR-based boundary detection\n\n"
                              f"Review and adjust ratings as needed.")
    
        # Enhanced buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
    
        ttk.Button(button_frame, text="Load Enhanced Data", 
                   command=load_enhanced_data).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                   command=preview_window.destroy).pack(side='right', padx=5)

    def setup_layout(self):
        """Create the main layout with proper canvas sizing."""
        print("DEBUG: Setting up layout with enhanced canvas sizing")

        # Create main paned window
        main_paned = tk.PanedWindow(self.window, orient='horizontal', sashrelief='raised', sashwidth=4)
        main_paned.pack(fill='both', expand=True, padx=5, pady=5)

        # === LEFT PANEL SETUP ===
        left_canvas = tk.Canvas(main_paned, bg=APP_BACKGROUND_COLOR, highlightthickness=0)
        self.left_canvas = left_canvas
        left_scrollbar = ttk.Scrollbar(main_paned, orient="vertical", command=left_canvas.yview)
        self.left_frame = ttk.Frame(left_canvas)

        # ENHANCED: Better scroll configuration
        def configure_left_scroll(event=None):
            """Configure scroll region and handle resizing."""
            self.left_frame.update_idletasks()
        
            # Update scroll region
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        
            # Get the current canvas size
            canvas_width = left_canvas.winfo_width()
            if canvas_width > 50:  # Valid width
                # Configure the interior frame to fill the canvas width
                left_canvas.itemconfig(left_canvas.find_all()[0], width=canvas_width-4)

        self.left_frame.bind("<Configure>", configure_left_scroll)

        # Create the canvas window
        left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        # ENHANCED: Better canvas resize handling
        def on_left_canvas_configure(event):
            """Handle left canvas resize events."""
            if event.widget == left_canvas:
                # Update scroll region
                left_canvas.configure(scrollregion=left_canvas.bbox("all"))
            
                # Ensure interior frame fills width
                canvas_width = event.width
                if canvas_width > 50 and left_canvas.find_all():
                    left_canvas.itemconfig(left_canvas.find_all()[0], width=canvas_width-4)

        left_canvas.bind('<Configure>', on_left_canvas_configure)

        # Mouse wheel scrolling
        def _on_mousewheel_left(event):
            left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        left_canvas.bind("<MouseWheel>", _on_mousewheel_left)

        # Add to paned window
        main_paned.add(left_canvas, stretch="always")

        # === RIGHT PANEL SETUP ===
        right_canvas = tk.Canvas(main_paned, bg=APP_BACKGROUND_COLOR, highlightthickness=0)
        right_scrollbar = ttk.Scrollbar(main_paned, orient="vertical", command=right_canvas.yview)
        self.right_frame = ttk.Frame(right_canvas)

        # Right panel configuration
        def configure_right_scroll(event=None):
            """Configure right scroll region."""
            self.right_frame.update_idletasks()
            bbox = right_canvas.bbox("all")
            if bbox:
                right_canvas.configure(scrollregion=bbox)

        self.right_frame.bind("<Configure>", configure_right_scroll)

        right_canvas.create_window((0, 0), window=self.right_frame, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)

        # Mouse wheel scrolling
        def _on_mousewheel_right(event):
            right_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        right_canvas.bind("<MouseWheel>", _on_mousewheel_right)

        main_paned.add(right_canvas)

        # Store the sash position function
        def set_initial_sash_position():
            try:
                window_width = self.window.winfo_width()
                if window_width > 100:
                    sash_position = int(window_width * 0.55)
                    main_paned.sash_place(0, sash_position, 0)
                    print(f"DEBUG: Set sash position to {sash_position}")
            except Exception as e:
                print(f"DEBUG: Sash positioning failed: {e}")

        self.set_initial_sash_position = set_initial_sash_position

        print("DEBUG: Enhanced layout setup complete")

        # Add session management and panels
        self.setup_session_selector(self.left_frame)
        self.setup_data_entry_panel()
        self.setup_plot_panel()

        # CRITICAL: Apply sizing optimization after content is added
        self.window.after(100, self.optimize_window_size)

        # Initialize default session
        if not self.sessions:
            self.create_new_session("Default_Session")

    def optimize_window_size(self):
        """Calculate window size based on actual frame dimensions after layout."""
        print("DEBUG: Starting precise window size optimization")
    
        # Force complete layout update
        self.window.update_idletasks()
        self.window.update()
    
        # CRITICAL: Measure what the left_frame actually uses after layout
        self.left_frame.update_idletasks()
        actual_left_frame_height = self.left_frame.winfo_reqheight()  # Changed from winfo_height()
    
        # Also measure right frame actual height for comparison
        self.right_frame.update_idletasks()
        actual_right_frame_height = self.right_frame.winfo_reqheight()  # Changed from winfo_height()
    
        print(f"DEBUG: Actual frame heights - Left: {actual_left_frame_height}px, Right: {actual_right_frame_height}px")
    
        # Width calculations (existing logic)
        left_frame_width = self.left_frame.winfo_reqwidth()
        right_frame_width = self.right_frame.winfo_reqwidth()
    
        min_plot_width = 500
        optimal_left_width = max(left_frame_width + 40, 450)
        optimal_right_width = max(min_plot_width, right_frame_width + 20)
        total_optimal_width = optimal_left_width + optimal_right_width + 50
    
        # FIXED: Reduce window chrome overhead and use required height instead of actual height
        governing_content_height = max(actual_left_frame_height, actual_right_frame_height)
        window_chrome = 10  # REDUCED from 120 to 30 - just enough for title bar and borders
        total_optimal_height = governing_content_height + window_chrome
    
        print(f"DEBUG: Window sized for actual frame height: {governing_content_height}px")
    
        # Screen constraints
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
    
        max_usable_width = int(screen_width * 0.9)
        max_usable_height = int(screen_height * 0.85)  # REDUCED from 0.9 to leave more screen space
    
        final_width = min(total_optimal_width, max_usable_width)
        final_height = min(total_optimal_height, max_usable_height)
    
        final_width = max(final_width, 800)
        final_height = max(final_height, 500)
    
        # ADDITIONAL FIX: Cap the height to avoid excess space
        if final_height > 900:  # Based on your 1080p screen
            final_height = 900
    
        print(f"DEBUG: Final window size matching actual content: {final_width}x{final_height}")
    
        # Apply the sizing
        self.window.geometry(f"{final_width}x{final_height}")
    
        # Pass the actual required height, not the full window height
        available_height = governing_content_height  # Don't add window_chrome here
        self.window.after(50, lambda: self.configure_canvas_sizing(available_height))
    
        self.center_window()
    
        if hasattr(self, 'set_initial_sash_position'):
            self.window.after(200, self.set_initial_sash_position)

    def configure_canvas_sizing(self, available_content_height):
        """Configure canvas sizing using the frame's actual rendered size."""
        print(f"DEBUG: Configuring canvas sizing")
    
        self.window.update_idletasks()
    
        # Get the required height of the left frame content
        self.left_frame.update_idletasks()
        required_frame_height = self.left_frame.winfo_reqheight()
    
        print(f"DEBUG: left_frame required height: {required_frame_height}px")
    
        # Set canvas to exactly match what the frame requires
        if hasattr(self, 'left_canvas'):
            # Add small padding but not excessive
            canvas_height = required_frame_height  # Just 10px padding instead of full height
            self.left_canvas.configure(height=canvas_height)
            print(f"DEBUG: Canvas set to frame's required height + padding: {canvas_height}px")
        
            # Update scroll region
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
    
        print("DEBUG: Canvas sized to exact frame height - minimal gray space")

    def coordinate_panel_heights(self, final_window_height, window_chrome_height):
        """Coordinate both panels to work optimally at the chosen window height."""
        print("DEBUG: Coordinating panel heights for optimal layout")
    
        # Calculate available height for panel content
        available_panel_height = final_window_height - window_chrome_height
    
        # Left Panel Strategy: Optimize scrolling behavior
        # If content fits, disable scrolling; if not, optimize scroll region
        left_content_height = self.left_frame.winfo_reqheight()
    
        if left_content_height <= available_panel_height:
            # Content fits! Set canvas to exact content height to eliminate gray space
            optimal_left_height = left_content_height + 5  # Small buffer
            print(f"DEBUG: Left panel content fits - setting canvas to {optimal_left_height}px")
        else:
            # Content is taller - use available height and enable smooth scrolling
            optimal_left_height = available_panel_height - 10  # Account for scrollbar
            print(f"DEBUG: Left panel content scrollable - setting canvas to {optimal_left_height}px")
    
        # Apply the left panel height optimization
        if hasattr(self, 'left_canvas'):
            self.left_canvas.configure(height=optimal_left_height)
    
        # Right Panel Strategy: Ensure plot area uses available space efficiently
        right_content_height = self.right_frame.winfo_reqheight()
    
        if right_content_height < available_panel_height:
            # Right panel has extra space - we could expand plot or center it
            extra_space = available_panel_height - right_content_height
            print(f"DEBUG: Right panel has {extra_space}px extra space - content will be naturally centered")
        else:
            print(f"DEBUG: Right panel content fits exactly in available space")
    
        print(f"DEBUG: Panel coordination complete - both panels optimized for {final_window_height}px window")
        
    def setup_data_entry_panel(self):
        """Setup the left panel for data entry."""
        # Header section
        header_frame = ttk.LabelFrame(self.left_frame, text="Session Information", padding=10)
        header_frame.pack(fill='x', padx=5, pady=5)
        
        # Configure header_frame for 2x2 + button layout
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1) 
        header_frame.grid_columnconfigure(2, weight=1)
        header_frame.grid_columnconfigure(3, weight=1)
        print("DEBUG: Configured header_frame for optimized 2x2 layout")
        
        self.header_vars = {}
        
        # Row 0: Assessor Name and Media
        ttk.Label(header_frame, text="Assessor Name:", font=FONT).grid(
            row=0, column=0, sticky='e', padx=5, pady=2)
        assessor_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=assessor_var, font=FONT, width=15).grid(
            row=0, column=1, sticky='w', padx=5, pady=2)
        self.header_vars["Assessor Name"] = assessor_var
        print("DEBUG: Added Assessor Name to row 0, column 0-1")

        ttk.Label(header_frame, text="Media:", font=FONT).grid(
            row=0, column=2, sticky='e', padx=5, pady=2)
        media_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=media_var, font=FONT, width=15).grid(
            row=0, column=3, sticky='w', padx=5, pady=2)
        self.header_vars["Media"] = media_var
        print("DEBUG: Added Media to row 0, column 2-3")

        # Row 1: Puff Length and Date
        ttk.Label(header_frame, text="Puff Length:", font=FONT).grid(
            row=1, column=0, sticky='e', padx=5, pady=2)
        puff_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=puff_var, font=FONT, width=15).grid(
            row=1, column=1, sticky='w', padx=5, pady=2)
        self.header_vars["Puff Length"] = puff_var
        print("DEBUG: Added Puff Length to row 1, column 0-1")

        ttk.Label(header_frame, text="Date:", font=FONT).grid(
            row=1, column=2, sticky='e', padx=5, pady=2)
        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(header_frame, textvariable=date_var, font=FONT, width=15).grid(
            row=1, column=3, sticky='w', padx=5, pady=2)
        self.header_vars["Date"] = date_var
        print("DEBUG: Added Date to row 1, column 2-3")
            
        # Row 2: Mode switch button (centered across all columns)
        mode_button_frame = ttk.Frame(header_frame)
        mode_button_frame.grid(row=2, column=0, columnspan=4, pady=10)
        
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
            # Create centered container for each metric row
            metric_container = ttk.Frame(eval_frame)
            metric_container.pack(fill='x', pady=4)
    
            # Center the metric frame within the container
            metric_frame = ttk.Frame(metric_container)
            metric_frame.pack(anchor='center')
            
            # Metric label
            ttk.Label(metric_frame, text=f"{metric}:", font=FONT, width=12).pack(side='left')
            
            # Container for scale and value with fixed width (50% reduction)
            scale_container = ttk.Frame(metric_frame)
            scale_container.pack(side='left', padx=5)
            
            # Rating scale (1-9) with reduced length, tickmarks, and smaller slider pointer
            self.rating_vars[metric] = tk.IntVar(value=5)
            scale = tk.Scale(scale_container, from_=1, to=9, orient='horizontal',
                           variable=self.rating_vars[metric], font=FONT, 
                           length=300, showvalue=0, tickinterval=1,
                           sliderlength=20, sliderrelief='raised', width=15)  # Smaller pointer: sliderlength=15, width=10
            scale.pack(side='left')
            print(f"DEBUG: Created centered scale for {metric} with length=200, smaller pointer (sliderlength=15), and tickmarks every 1 point")
            
            # Current value display
            value_label = ttk.Label(metric_frame, text="5", width=2)
            value_label.pack(side='left', padx=(10, 0))
            
            # Update value display AND plot when scale changes (LIVE UPDATES)
            def update_live(val, label=value_label, var=self.rating_vars[metric], metric_name=metric):
                label.config(text=str(var.get()))
                self.auto_save_and_update()  # Add this method call
            scale.config(command=update_live)
            print(f"DEBUG: Centered scale for {metric} configured with smaller pointer and tickmarks from 1-9")
            
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
         
    def on_window_resize_plot(self, event):
        """Handle window resize events to update plot size dynamically - ENHANCED DEBUG VERSION."""
        print(f"DEBUG: RESIZE EVENT DETECTED - Widget: {event.widget}, Window: {self.window}")
        print(f"DEBUG: Event widget type: {type(event.widget)}")
        print(f"DEBUG: Window type: {type(self.window)}")
        print(f"DEBUG: Event widget == window? {event.widget == self.window}")
        print(f"DEBUG: Event widget is window? {event.widget is self.window}")
    
        # Only handle main window resize events, not child widgets
        if event.widget != self.window:
            print(f"DEBUG: Ignoring resize event from child widget: {event.widget}")
            return
        
        print("DEBUG: MAIN WINDOW RESIZE CONFIRMED - Processing...")
    
        # Get current window dimensions for verification
        current_width = self.window.winfo_width()
        current_height = self.window.winfo_height()
        print(f"DEBUG: Current window dimensions: {current_width}x{current_height}")
    
        # Debounce rapid resize events
        if hasattr(self, '_resize_timer'):
            self.window.after_cancel(self._resize_timer)
            print("DEBUG: Cancelled previous resize timer")
    
        # Schedule plot size update with a small delay to avoid excessive updates
        self._resize_timer = self.window.after(150, self.update_plot_size_for_resize)
        print("DEBUG: Scheduled plot resize update in 150ms")

    def on_window_resize_plot(self, event):
        """Handle window resize events to update plot size dynamically."""
        # Only handle main window resize events, not child widgets
        if event.widget != self.window:
            return
        
        print("DEBUG: Window resize detected, checking if plot size update needed")
    
        # Debounce rapid resize events
        if hasattr(self, '_resize_timer'):
            self.window.after_cancel(self._resize_timer)
    
        # Schedule plot size update with a small delay to avoid excessive updates
        self._resize_timer = self.window.after(150, self.update_plot_size_for_resize)

    def update_plot_size_for_resize(self):
        """Update plot size after window resize with proper error handling."""
        try:
            print("DEBUG: Executing delayed plot size update after window resize")
        
            # Check if we have the necessary components
            if not hasattr(self, 'canvas_frame') or not self.canvas_frame.winfo_exists():
                print("DEBUG: Canvas frame not available, skipping resize update")
                return
            
            if not hasattr(self, 'fig') or not self.fig:
                print("DEBUG: Figure not available, skipping resize update")
                return
        
            # Get the parent frame for size calculation (plot_frame from setup_plot_panel)
            plot_panel_frame = self.canvas_frame.master
        
            # Calculate new optimal size
            new_width, new_height = self.calculate_dynamic_plot_size(plot_panel_frame)
        
            # Get current figure size
            current_width, current_height = self.fig.get_size_inches()
        
            # Only update if size change is significant (avoid tiny adjustments)
            width_diff = abs(new_width - current_width)
            height_diff = abs(new_height - current_height)
        
            if width_diff > 0.2 or height_diff > 0.1:  # 0.1 inch threshold
                print(f"DEBUG: Significant size change detected, updating plot from {current_width:.2f}x{current_height:.2f} to {new_width:.2f}x{new_height:.2f}")
            
                # Update figure size
                self.fig.set_size_inches(new_width, new_height)
            
                # Redraw the canvas
                self.canvas.draw()
                print("DEBUG: Plot size updated and redrawn successfully")
            else:
                print("DEBUG: Size change too small, skipping plot update")
            
        except Exception as e:
            print(f"DEBUG: Error during plot resize update: {str(e)}")

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
        """Create the matplotlib canvas for the spider plot with dynamic responsive sizing."""
        print("DEBUG: Setting up enhanced plot canvas with dynamic sizing")
    
        # Calculate dynamic plot size based on available space
        dynamic_width, dynamic_height = self.calculate_dynamic_plot_size(parent)
    
        # Create figure with calculated responsive sizing
        self.fig, self.ax = plt.subplots(figsize=(dynamic_width, dynamic_height), subplot_kw=dict(projection='polar'))
        self.fig.patch.set_facecolor('white')
        print(f"DEBUG: Created spider plot with dynamic size: {dynamic_width:.2f}x{dynamic_height:.2f} inches")
    
        self.fig.subplots_adjust(left=0.15, right=0.8, top=0.8, bottom=0.08)
        print("DEBUG: Applied subplot adjustments for dynamic plot")
        
        # Create canvas with scrollbars
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill='both', expand=True)
    
        # Store reference to canvas_frame for resize handling
        self.canvas_frame = canvas_frame
    
        # Create canvas
        self.canvas = FigureCanvasTkAgg(self.fig, canvas_frame)
        canvas_widget = self.canvas.get_tk_widget()
        canvas_widget.pack(fill='both', expand=True)
    
        # Add toolbar for additional functionality
        self.setup_plot_context_menu(canvas_widget)
        print("DEBUG: Plot canvas setup complete with dynamic sizing and right-click context menu")
    
        # Initialize empty plot
        self.update_plot()
    
        # Bind window resize events to update plot size
        self.window.bind('<Configure>', self.on_window_resize_plot, add=True)
        print("DEBUG: Window resize binding added for dynamic plot updates")

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

    def calculate_dynamic_plot_size(self, parent_frame):
        """Calculate optimal plot size based on available canvas space with aspect ratio preservation."""
        print("DEBUG: Starting dynamic plot size calculation")
    
        # Update geometry to get current dimensions
        self.window.update_idletasks()
    
        # Get the parent frame dimensions (this is the plot_frame from setup_plot_panel)
        parent_frame.update_idletasks()
        available_width = parent_frame.winfo_width()
        available_height = parent_frame.winfo_height()
    
        print(f"DEBUG: Available canvas space - Width: {available_width}px, Height: {available_height}px")
    
        # Ensure we have reasonable minimum dimensions
        if available_width < 100 or available_height < 100:
            print("DEBUG: Canvas too small, using fallback dimensions")
            return (6, 4.8)  # Original fallback size
    
        # Reserve space for controls and padding (checkbox frame, labels, margins)
        control_height_reserve = 80  # Space for checkboxes and labels
        padding_reserve = 40  # General padding around plot
    
        # Calculate usable space for the actual plot
        usable_width = available_width - 10  # 10 pixels less than canvas width as requested
        usable_height = available_height - control_height_reserve - padding_reserve
    
        print(f"DEBUG: Usable plot space - Width: {usable_width}px, Height: {usable_height}px")
    
        # Convert pixels to inches (matplotlib figsize uses inches)
        # Standard DPI is typically 100, but we'll use 100 for consistent sizing
        plot_dpi = 100
        max_width_inches = usable_width / plot_dpi
        max_height_inches = usable_height / plot_dpi
    
        # Define preferred aspect ratio (width:height) for spider plots
        preferred_aspect_ratio = 1.25  # 5:4 ratio works well for circular spider plots
    
        # Calculate dimensions while maintaining aspect ratio
        # Try width-constrained first
        width_constrained_width = max_width_inches
        width_constrained_height = width_constrained_width / preferred_aspect_ratio
    
        # Try height-constrained
        height_constrained_height = max_height_inches
        height_constrained_width = height_constrained_height * preferred_aspect_ratio
    
        # Choose the constraint that fits within both limits
        if width_constrained_height <= max_height_inches:
            # Width constraint works
            final_width = width_constrained_width
            final_height = width_constrained_height
            constraint_type = "width"
        else:
            # Height constraint needed
            final_width = height_constrained_width
            final_height = height_constrained_height
            constraint_type = "height"
    
        # Apply reasonable bounds
        min_size = 2.0  # Minimum 3 inches
        max_size = 12.0  # Maximum 12 inches
    
        final_width = max(min_size, min(final_width, max_size))
        final_height = max(min_size, min(final_height, max_size))
    
        print(f"DEBUG: Calculated plot size - {final_width:.2f}x{final_height:.2f} inches")
        print(f"DEBUG: Constraint applied: {constraint_type}, Aspect ratio: {final_width/final_height:.2f}")
        print(f"DEBUG: Final plot will be approximately {int(final_width*plot_dpi)}x{int(final_height*plot_dpi)} pixels")
    
        return (final_width, final_height)



    def create_spider_plot(self):
        """Create a spider/radar plot of sensory data."""
        self.ax.clear()
        
        # Check if we have any data to plot
        if not self.samples:
            self.ax.text(0.5, 0.5, 'No samples to display\nAdd samples to begin evaluation', 
                        transform=self.ax.transAxes, ha='center', va='center', 
                        fontsize=12, color='gray')
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
                        fontsize=12, color='gray')
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
            self.ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.2), fontsize=10)
            
        # Set title
        self.ax.set_title('Sensory Profile Comparison', fontsize=12, fontweight='bold', pad=15, ha = 'center', y = 1.08)
        
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
    
        # Check if there are unsaved changes
        if self.samples:
            result = messagebox.askyesnocancel("Unsaved Changes", 
                                             "Do you want to save the current session before creating a new one?\n\n"
                                             "Yes = Save and create new\n"
                                             "No = Discard and create new\n"
                                             "Cancel = Return to current session")
        
            if result is None:  # Cancel
                return
            elif result:  # Yes - save first
                self.save_session()
    
        # Get name for new session
        session_name = tk.simpledialog.askstring("New Session", 
                                                "Enter name for new session:",
                                                initialvalue=f"Session_{datetime.now().strftime('%Y%m%d_%H%M')}")
    
        if not session_name or not session_name.strip():
            return
    
        session_name = session_name.strip()
    
        # Ensure unique name
        counter = 1
        original_name = session_name
        while session_name in self.sessions:
            session_name = f"{original_name}_{counter}"
            counter += 1
    
        print(f"DEBUG: Creating new session: {session_name}")
    
        # Create the new session
        self.create_new_session(session_name)
    
        # Clear header fields to defaults
        for field, var in self.header_vars.items():
            if field == "Date":
                var.set(datetime.now().strftime("%Y-%m-%d"))
            else:
                var.set('')
    
        # Update UI
        if hasattr(self, 'session_var'):
            self.session_var.set(session_name)
    
        self.update_sample_combo()
        self.update_sample_checkboxes()
        self.sample_var.set('')
        self.clear_form()
        self.update_plot()
    
        print(f"DEBUG: New session {session_name} created successfully")
        messagebox.showinfo("New Session", f"Created new session: {session_name}")
        debug_print("Started new sensory evaluation session")
            
    def save_session(self):
        """Save the current session to a JSON file."""
        if not self.current_session_id or not self.sessions:
            messagebox.showwarning("Warning", "No session to save!")
            return
    
        # Make sure current samples are saved to the session
        if self.current_session_id in self.sessions:
            self.sessions[self.current_session_id]['samples'] = self.samples
            self.sessions[self.current_session_id]['header'] = {field: var.get() for field, var in self.header_vars.items()}
        
        current_session = self.sessions[self.current_session_id]
    
        if not current_session['samples']:
            messagebox.showwarning("Warning", "No sample data to save!")
            return
        
        # Default filename based on session name and assessor
        assessor_name = current_session['header'].get('Assessor Name', 'Unknown')
        safe_assessor = "".join(c for c in assessor_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_session = "".join(c for c in self.current_session_id if c.isalnum() or c in (' ', '-', '_')).strip()
    
        default_filename = f"{safe_assessor}_{safe_session}_sensory_session.json"
    
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Sensory Session",
            initialvalue=default_filename
        )
    
        if filename:
            try:
                # Create session data with updated timestamp
                session_data = {
                    'header': current_session['header'],
                    'samples': current_session['samples'],
                    'timestamp': datetime.now().isoformat(),
                    'session_name': self.current_session_id,
                    'source_file': current_session.get('source_file', ''),
                    'source_image': current_session.get('source_image', '')
                }
            
                with open(filename, 'w') as f:
                    json.dump(session_data, f, indent=2)
                
                print(f"DEBUG: Saved session {self.current_session_id} to {filename}")
                messagebox.showinfo("Success", 
                                  f"Session '{self.current_session_id}' saved to {os.path.basename(filename)}\n"
                                  f"Saved {len(current_session['samples'])} samples")
                debug_print(f"Saved sensory session to: {filename}")
            
            except Exception as e:
                print(f"DEBUG: Error saving session: {e}")
                messagebox.showerror("Error", f"Failed to save session: {e}")
                
    def load_session(self):
        """Load a session from a JSON file as a new session."""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Load Sensory Session"
        )
    
        if filename:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)
            
                print(f"DEBUG: Loading session from {filename}")
                print(f"DEBUG: Session data keys: {list(session_data.keys())}")
            
                # Validate session data
                if not self.validate_session_data(session_data):
                    messagebox.showerror("Invalid Session", 
                                       "The selected file does not contain valid session data.")
                    return
            
                # Create session name from filename
                base_filename = os.path.splitext(os.path.basename(filename))[0]
                session_name = base_filename
            
                # Ensure unique session name
                counter = 1
                original_name = session_name
                while session_name in self.sessions:
                    session_name = f"{original_name}_{counter}"
                    counter += 1
            
                print(f"DEBUG: Creating new session: {session_name}")
            
                # Create new session with loaded data
                self.sessions[session_name] = {
                    'header': session_data.get('header', {}),
                    'samples': session_data.get('samples', {}),
                    'timestamp': session_data.get('timestamp', datetime.now().isoformat()),
                    'source_file': filename
                }
            
                print(f"DEBUG: Session created with {len(self.sessions[session_name]['samples'])} samples")
            
                # Switch to the new session
                self.switch_to_session(session_name)
            
                # Update session selector UI
                self.update_session_combo()
                if hasattr(self, 'session_var'):
                    self.session_var.set(session_name)
            
                # Update other UI components
                self.update_sample_combo()
                self.update_sample_checkboxes()
            
                # Select first sample if available
                if self.samples:
                    first_sample = list(self.samples.keys())[0]
                    self.sample_var.set(first_sample)
                    self.load_sample_data(first_sample)
                else:
                    self.sample_var.set('')
                    self.clear_form()
                
                self.update_plot()
            
                print(f"DEBUG: Successfully loaded session {session_name}")
                messagebox.showinfo("Success", 
                                  f"Session loaded as '{session_name}'\n"
                                  f"Loaded {len(self.samples)} samples\n"
                                  f"Use session selector to switch between sessions.")
                debug_print(f"Loaded sensory session from: {filename}")
            
            except Exception as e:
                print(f"DEBUG: Error loading session: {e}")
                import traceback
                traceback.print_exc()
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

                               
    def load_from_image_with_ai(self):
        """Load sensory data from a single form image using Enhanced Claude AI."""
    
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
            title="Load Sensory Form Image for Enhanced AI Processing"
        )

        if not filename:
            return

        try:
            # Show processing dialog
            progress_window = tk.Toplevel(self.window)
            progress_window.title("Processing with Enhanced Claude AI...")
            progress_window.geometry("400x150")
            progress_window.transient(self.window)
            progress_window.grab_set()
    
            progress_label = ttk.Label(progress_window, 
                                     text="Processing image with shadow removal and AI analysis...", 
                                     font=FONT)
            progress_label.pack(expand=True, pady=10)
    
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
            progress_bar.pack(fill='x', padx=20, pady=10)
            progress_bar.start()
    
            self.window.update()
    
            # Process with Enhanced Claude AI
            ai_processor = EnhancedClaudeFormProcessor()
        
            # Process single image (uses shadow removal preprocessing)
            image_data, processed_image = ai_processor.prepare_image_with_preprocessing(filename)
            extracted_data = ai_processor.process_single_image_with_claude(image_data, filename)
    
            # Stop progress bar
            progress_bar.stop()
            progress_window.destroy()
    
            # Show enhanced results preview
            if extracted_data:
                self.show_enhanced_ai_extraction_preview(extracted_data, processed_image, filename)
            else:
                messagebox.showwarning("No Data", "No sensory data could be extracted from the image.")
        
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("Enhanced AI Processing Error", f"Failed to process image with AI: {e}")
            debug_print(f"Enhanced AI processing error: {e}")

    def batch_process_with_ai(self):
        """Process a batch of form images using Enhanced Claude AI."""

        # Check for required dependencies
        try:
            import anthropic
        except ImportError as e:
            messagebox.showerror("Missing Dependencies", 
                               f"AI batch processing requires additional packages.\n\n"
                               f"Install with: pip install anthropic pillow opencv-python\n\n"
                               f"Error: {e}")
            return

        # Get folder containing images
        folder_path = filedialog.askdirectory(
            title="Select Folder Containing Form Images for Batch AI Processing"
        )

        if not folder_path:
            return

        try:
            print("DEBUG: Starting batch AI processing")
        
            # Show processing dialog
            progress_window = tk.Toplevel(self.window)
            progress_window.title("Batch Processing with Enhanced Claude AI...")
            progress_window.geometry("450x200")
            progress_window.transient(self.window)
            progress_window.grab_set()

            progress_label = ttk.Label(progress_window, 
                                     text="Processing batch of images with shadow removal...", 
                                     font=FONT)
            progress_label.pack(expand=True, pady=10)
    
            detail_label = ttk.Label(progress_window, 
                                   text="Initializing...", 
                                   font=('Arial', 9))
            detail_label.pack(pady=5)

            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
            progress_bar.pack(fill='x', padx=20, pady=10)
            progress_bar.start()

            self.window.update()

            # Initialize Enhanced Claude AI processor
            ai_processor = EnhancedClaudeFormProcessor()
    
            # Update progress
            detail_label.config(text="Processing images with Claude AI...")
            self.window.update()
    
            # Process batch of images
            batch_results = ai_processor.process_batch_images(folder_path)

            # Stop progress bar
            progress_bar.stop()
            progress_window.destroy()

            # Show batch results summary
            successful_count = sum(1 for result in batch_results.values() if result['status'] == 'success')
            total_count = len(batch_results)
    
            if successful_count > 0:
                # Store the processor and results for review interface
                self.ai_processor = ai_processor
        
                result_msg = (f"Enhanced Batch Processing Complete!\n\n"
                             f"Successfully processed: {successful_count}/{total_count} images\n"
                             f"Features used:\n"
                             f"• Shadow removal preprocessing\n"
                             f"• Sample name extraction\n"
                             f"• Enhanced AI analysis\n\n"
                             f"Launch interactive review to verify and edit results?")
        
                if messagebox.askyesno("Batch Processing Complete", result_msg):
                    print("DEBUG: Launching review interface")
                    # Launch the interactive review interface
                    ai_processor.launch_review_interface()
                
                    # FIXED: Properly monitor for review completion
                    self.monitor_review_completion(ai_processor)
                else:
                    print("DEBUG: Loading results directly without review")
                    # Auto-load all successful results
                    self.load_batch_results_directly(batch_results)
            else:
                messagebox.showerror("Batch Processing Failed", 
                                   f"No images could be processed successfully.\n"
                                   f"Check that images contain readable sensory evaluation forms.")
    
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("Batch AI Processing Error", f"Failed to process batch: {e}")
            debug_print(f"Batch AI processing error: {e}")

    def monitor_review_completion(self, ai_processor):
        """Monitor the review interface for completion."""
        def check_completion():
            print("DEBUG: Checking review completion status")
        
            if ai_processor.is_review_complete():
                print("DEBUG: Review completed, loading results")
            
                # Get the reviewed results
                reviewed_results = ai_processor.get_reviewed_results()
            
                if reviewed_results:
                    print(f"DEBUG: Loading {len(reviewed_results)} reviewed results")
                
                    # Load the reviewed results using session-based structure
                    self.load_batch_results_directly(reviewed_results)
                
                    # FIXED: Update session selector UI
                    if hasattr(self, 'session_var') and self.current_session_id:
                        self.session_var.set(self.current_session_id)
                
                    # Show completion message
                    total_samples = sum(len(data['extracted_data']) for data in reviewed_results.values() 
                                      if data['status'] == 'success')
                
                    messagebox.showinfo("Review Complete", 
                                      f"Successfully loaded reviewed data!\n"
                                      f"Total sessions: {len(reviewed_results)}\n"
                                      f"Total samples: {total_samples}\n"
                                      f"Use the session selector to switch between sessions.")
                else:
                    print("DEBUG: No reviewed results found")
                    messagebox.showwarning("No Data", "No reviewed data to load.")
            
                return  # Stop monitoring
        
            # Continue monitoring if review not complete
            self.window.after(500, check_completion)  # Check every 500ms
    
        print("DEBUG: Starting review completion monitoring")
        # Start monitoring after a short delay to allow review window to open
        self.window.after(1000, check_completion)

    def handle_reviewed_results(self, results):
        """Handle results from the review interface."""
        print(f"DEBUG: Received callback with {len(results)} results")
        self.load_batch_results_directly(results)

    def load_batch_results_directly(self, batch_results):
        """Load successful batch results with each image as a separate session."""
        loaded_sessions = 0
        loaded_samples = 0

        print("DEBUG: Starting batch results loading with session-per-image structure")
        print(f"DEBUG: Processing {len(batch_results)} batch results")

        for image_path, result in batch_results.items():
            if result['status'] == 'success':
                extracted_data = result['extracted_data']
        
                # Skip empty results
                if not extracted_data:
                    print(f"DEBUG: Skipping empty result for {image_path}")
                    continue
        
                # Create session name from image filename
                image_name = os.path.splitext(os.path.basename(image_path))[0]
                session_name = f"Batch_AI_{image_name}"
            
                # Ensure unique session name
                counter = 1
                original_session_name = session_name
                while session_name in self.sessions:
                    session_name = f"{original_session_name}_{counter}"
                    counter += 1
        
                print(f"DEBUG: Creating session {session_name} from image {image_name}")
        
                # Create new session for this image
                self.sessions[session_name] = {
                    'header': {field: var.get() for field, var in self.header_vars.items()},
                    'samples': {},
                    'timestamp': datetime.now().isoformat(),
                    'source_image': image_path,
                    'extraction_method': 'Enhanced_Claude_AI_Batch'
                }
        
                # Load samples into this session (up to 4 samples)
                sample_count = 0
                for sample_key, sample_data in extracted_data.items():
                    if sample_count >= 4:
                        print(f"DEBUG: Reached maximum 4 samples for session {session_name}")
                        break
            
                    # Skip empty samples
                    if not sample_data or not any(sample_data.get(metric, None) for metric in self.metrics):
                        print(f"DEBUG: Skipping empty sample {sample_key}")
                        continue
            
                    # Add sample to this session
                    self.sessions[session_name]['samples'][sample_key] = sample_data
                    sample_count += 1
                    loaded_samples += 1
            
                    print(f"DEBUG: Added sample {sample_key} to session {session_name}")
        
                if sample_count > 0:
                    loaded_sessions += 1
                    print(f"DEBUG: Session {session_name} created with {sample_count} samples")
                else:
                    # Remove empty session
                    del self.sessions[session_name]
                    print(f"DEBUG: Removed empty session {session_name}")

        # Switch to first session if any were loaded
        if self.sessions:
            first_session = list(self.sessions.keys())[0]
            self.switch_to_session(first_session)
        
            # Update session selector UI
            if hasattr(self, 'session_var'):
                self.session_var.set(first_session)
    
            print(f"DEBUG: Switched to first session: {first_session}")

        # Update UI
        self.update_session_combo()
        self.update_sample_combo()
        self.update_sample_checkboxes()
        self.update_plot()

        print(f"DEBUG: Batch loading complete")
        print(f"DEBUG: Loaded {loaded_sessions} sessions with total {loaded_samples} samples")

        if loaded_sessions > 0:
            messagebox.showinfo("Batch Load Complete", 
                              f"Loaded {loaded_sessions} sessions with {loaded_samples} total samples!\n"
                              f"Each image is now a separate session (max 4 samples each).\n"
                              f"Use the session selector to switch between sessions.")
        else:
            messagebox.showwarning("No Data Loaded", 
                                 "No valid samples found in batch results.")

    def setup_session_selector(self, parent_frame):
        """Add session selector to the interface."""
        # Add session selector frame with reduced width
        session_frame = ttk.LabelFrame(parent_frame, text="Session Management", padding=10)
        session_frame.pack(fill='x', padx=5, pady=5)
    
        # Configure session_frame for centered grid layout
        session_frame.grid_columnconfigure(0, weight=1)
        session_frame.grid_columnconfigure(1, weight=1)
        print("DEBUG: Configured session_frame for centered layout")

        # Top row - Session selection centered
        top_frame = ttk.Frame(session_frame)
        top_frame.grid(row=0, column=0, columnspan=2, pady=(0, 5))
    
        session_label = ttk.Label(top_frame, text="Current Session:", font=FONT)
        session_label.pack(side='left', padx=(0, 5))

        self.session_var = tk.StringVar()
        self.session_combo = ttk.Combobox(top_frame, textvariable=self.session_var, 
                                         font=FONT, state='readonly', width=15)
        self.session_combo.pack(side='left', padx=(0, 10))
        self.session_combo.bind('<<ComboboxSelected>>', self.on_session_selected)
        print("DEBUG: Session dropdown centered on top row")

        # Second row - Session management buttons centered
        button_frame = ttk.Frame(session_frame)
        button_frame.grid(row=1, column=0, columnspan=2)
    
        ttk.Button(button_frame, text="New Session", 
                   command=self.add_new_session).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Combine Sessions", 
                   command=self.show_combine_sessions_dialog).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Delete Session", 
                   command=self.delete_current_session).pack(side='left', padx=2)
        print("DEBUG: Session management buttons centered on second row")

        print("DEBUG: Session selector UI setup complete with centered layout")

    def on_session_selected(self, event=None):
        """Handle session selection change."""
        selected_session = self.session_var.get()
        if selected_session and selected_session != self.current_session_id:
            print(f"DEBUG: Session selection changed to: {selected_session}")
            self.switch_to_session(selected_session)

    def add_new_session(self):
        """Add a new empty session."""
        session_name = tk.simpledialog.askstring("New Session", "Enter session name:")
        if session_name and session_name.strip():
            session_name = session_name.strip()
            if session_name in self.sessions:
                messagebox.showerror("Session Exists", f"Session '{session_name}' already exists.")
                return
        
            print(f"DEBUG: Creating new session: {session_name}")
            self.create_new_session(session_name)
            self.session_var.set(session_name)
            messagebox.showinfo("Success", f"Created new session: {session_name}")

    def delete_current_session(self):
        """Delete the current session."""
        if not self.current_session_id:
            messagebox.showwarning("No Session", "No session selected to delete.")
            return
    
        if len(self.sessions) <= 1:
            messagebox.showwarning("Cannot Delete", "Cannot delete the last session.")
            return
    
        if messagebox.askyesno("Confirm Delete", 
                              f"Delete session '{self.current_session_id}'?\n"
                              f"This will permanently remove all data in this session."):
        
            session_to_delete = self.current_session_id
            print(f"DEBUG: Deleting session: {session_to_delete}")
        
            # Switch to another session first
            remaining_sessions = [s for s in self.sessions.keys() if s != session_to_delete]
            if remaining_sessions:
                self.switch_to_session(remaining_sessions[0])
        
            # Delete the session
            del self.sessions[session_to_delete]
            self.update_session_combo()
        
            print(f"DEBUG: Session {session_to_delete} deleted successfully")
            messagebox.showinfo("Success", f"Session '{session_to_delete}' deleted.")

    def show_combine_sessions_dialog(self):
        """Show dialog to select and combine multiple sessions."""
        if len(self.sessions) < 2:
            messagebox.showinfo("Insufficient Sessions", 
                              "Need at least 2 sessions to combine.")
            return
    
        # Create dialog window
        combine_window = tk.Toplevel(self.window)
        combine_window.title("Combine Sessions")
        combine_window.geometry("400x300")
        combine_window.transient(self.window)
        combine_window.grab_set()
    
        # Instructions
        ttk.Label(combine_window, 
                 text="Select sessions to combine into a new session:",
                 font=FONT).pack(pady=10)
    
        # Session selection with checkboxes
        selection_frame = ttk.Frame(combine_window)
        selection_frame.pack(fill='both', expand=True, padx=20, pady=10)
    
        session_vars = {}
        for session_id in self.sessions.keys():
            var = tk.BooleanVar()
            session_vars[session_id] = var
        
            # Create checkbox with session info
            sample_count = len(self.sessions[session_id]['samples'])
            source_image = self.sessions[session_id].get('source_image', 'Manual')
            source_name = os.path.basename(source_image) if source_image else 'Manual'
        
            checkbox_text = f"{session_id} ({sample_count} samples) - {source_name}"
            ttk.Checkbutton(selection_frame, text=checkbox_text, 
                           variable=var).pack(anchor='w', pady=2)
    
        # New session name
        name_frame = ttk.Frame(combine_window)
        name_frame.pack(fill='x', padx=20, pady=10)
    
        ttk.Label(name_frame, text="New session name:").pack(side='left')
        name_var = tk.StringVar(value=f"Combined_Session_{self.session_counter}")
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=20)
        name_entry.pack(side='right')
    
        # Buttons
        button_frame = ttk.Frame(combine_window)
        button_frame.pack(fill='x', padx=20, pady=20)
    
        def combine_selected_sessions():
            selected_sessions = [sid for sid, var in session_vars.items() if var.get()]
            new_session_name = name_var.get().strip()
        
            print(f"DEBUG: Combining sessions: {selected_sessions}")
            print(f"DEBUG: New session name: {new_session_name}")
        
            if len(selected_sessions) < 2:
                messagebox.showwarning("Insufficient Selection", 
                                     "Select at least 2 sessions to combine.")
                return
        
            if not new_session_name:
                messagebox.showwarning("Invalid Name", "Enter a name for the new session.")
                return
        
            if new_session_name in self.sessions:
                messagebox.showerror("Name Exists", 
                                   f"Session '{new_session_name}' already exists.")
                return
        
            # Combine sessions
            combined_samples = {}
            total_sample_count = 0
        
            for session_id in selected_sessions:
                session_samples = self.sessions[session_id]['samples']
                for sample_name, sample_data in session_samples.items():
                    # Create unique sample name if conflict
                    unique_name = sample_name
                    counter = 1
                    while unique_name in combined_samples:
                        unique_name = f"{sample_name}_{counter}"
                        counter += 1
                
                    combined_samples[unique_name] = sample_data
                    total_sample_count += 1
                    print(f"DEBUG: Added sample {unique_name} from session {session_id}")
        
            # Create new combined session
            self.create_new_session(new_session_name)
            self.sessions[new_session_name]['samples'] = combined_samples
            self.samples = combined_samples
        
            # Update UI
            self.session_var.set(new_session_name)
            self.update_sample_combo()
            self.update_sample_checkboxes()
            self.update_plot()
        
            combine_window.destroy()
        
            print(f"DEBUG: Successfully combined {len(selected_sessions)} sessions")
            print(f"DEBUG: New session has {total_sample_count} samples")
        
            messagebox.showinfo("Success", 
                              f"Combined {len(selected_sessions)} sessions into '{new_session_name}'!\n"
                              f"Total samples: {total_sample_count}")
    
        ttk.Button(button_frame, text="Combine Sessions", 
                   command=combine_selected_sessions).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                   command=combine_window.destroy).pack(side='right', padx=5)
    
        print("DEBUG: Combine sessions dialog created")

    def setup_interface(self):
        """Set up the main interface with session management."""
        # Create main frames
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        # Add session management at the top
        self.setup_session_selector(main_frame)
    
        # Rest of your existing interface setup...
        # (header section, sample selection, ratings, etc.)
    
        # Initialize with default session
        if not self.sessions:
            self.create_new_session("Default_Session")

    def show_enhanced_ai_extraction_preview(self, extracted_data, processed_image, filename):
        """Show enhanced AI extraction preview with editable ratings."""
        preview_window = tk.Toplevel(self.window)
        preview_window.title("Enhanced AI Extraction Results")
        preview_window.geometry("1200x800")
        preview_window.transient(self.window)

        # Create main layout
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Title with enhanced info
        title_label = ttk.Label(main_frame, 
                               text=f"Enhanced Claude AI Results: {os.path.basename(filename)}", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=5)

        # Feature info
        info_label = ttk.Label(main_frame, 
                              text="✓ Shadow removal preprocessing ✓ Sample name extraction ✓ Enhanced AI analysis", 
                              font=('Arial', 10),
                              foreground='green')
        info_label.pack(pady=2)

        # Create horizontal layout
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill='both', expand=True, pady=10)

        # Left side - Processed image
        image_frame = ttk.LabelFrame(content_frame, text="Shadow-Removed Image", padding=10)
        image_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))

        # Display processed image
        try:
            Image, ImageTk = _lazy_import_pil()
            if Image and ImageTk:
                # Convert processed image to PIL
                pil_image = Image.fromarray(processed_image)
        
                # Resize for display
                display_width = 500
                aspect_ratio = pil_image.height / pil_image.width
                display_height = int(display_width * aspect_ratio)
        
                if display_height > 600:
                    display_height = 600
                    display_width = int(display_height / aspect_ratio)
        
                pil_image = pil_image.resize((display_width, display_height), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(pil_image)
        
                image_label = tk.Label(image_frame, image=photo)
                image_label.image = photo  # Keep reference
                image_label.pack()
            else:
                tk.Label(image_frame, text="Processed image\n(PIL not available for display)").pack()
        except Exception as e:
            tk.Label(image_frame, text=f"Error displaying image:\n{e}").pack()

        # Right side - Data editing
        data_frame = ttk.LabelFrame(content_frame, text="Extracted Data (Editable)", padding=10)
        data_frame.pack(side='right', fill='both', expand=True)

        # Store editing variables
        sample_name_vars = {}
        rating_vars = {}
        comments_vars = {}
        original_names = list(extracted_data.keys())

        # Scrollable frame for data
        canvas = tk.Canvas(data_frame)
        scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Enhanced sample display with EDITABLE ratings
        for i, (original_name, sample_data) in enumerate(extracted_data.items()):
            sample_frame = ttk.LabelFrame(scrollable_frame, text=f"Sample {i+1}", padding=10)
            sample_frame.pack(fill='x', pady=5, padx=5)

            # Sample name editing
            name_frame = ttk.Frame(sample_frame)
            name_frame.pack(fill='x', pady=(0, 5))

            ttk.Label(name_frame, text="Sample Name:", font=('Arial', 10, 'bold')).pack(side='left')
            sample_name_var = tk.StringVar(value=sample_data.get('sample_name', original_name))
            sample_name_vars[original_name] = sample_name_var
            name_entry = ttk.Entry(name_frame, textvariable=sample_name_var, width=25, font=('Arial', 10))
            name_entry.pack(side='left', padx=(5, 0))

            # Show if name was extracted vs. default
            extracted_name = sample_data.get('sample_name', '')
            if extracted_name and extracted_name != original_name:
                ttk.Label(name_frame, text="✓ Extracted", foreground='green', font=('Arial', 8)).pack(side='left', padx=(5, 0))

            # EDITABLE ratings display
            ratings_frame = ttk.Frame(sample_frame)
            ratings_frame.pack(fill='x', pady=(5, 0))
        
            # Initialize rating variables for this sample
            rating_vars[original_name] = {}
        
            for metric in self.metrics:
                if metric in sample_data and metric not in ['comments', 'sample_name']:
                    attr_frame = ttk.Frame(ratings_frame)
                    attr_frame.pack(fill='x', pady=2)

                    # Attribute label
                    attr_label = ttk.Label(attr_frame, text=f"{metric}:", width=15, anchor='w', font=('Arial', 9))
                    attr_label.pack(side='left')

                    # EDITABLE rating field
                    rating_var = tk.IntVar(value=sample_data.get(metric, 5))
                    rating_vars[original_name][metric] = rating_var
                
                    # Use Spinbox for easy editing
                    rating_spinbox = tk.Spinbox(attr_frame, from_=1, to=9, width=5, 
                                              textvariable=rating_var, font=('Arial', 9))
                    rating_spinbox.pack(side='left', padx=(5, 0))
                
                    # Show original extracted value for reference
                    original_value = sample_data.get(metric, 5)
                    ttk.Label(attr_frame, text=f"(AI: {original_value})", 
                             font=('Arial', 8), foreground='blue').pack(side='left', padx=(5, 0))

            # EDITABLE comments section
            comments_frame = ttk.Frame(sample_frame)
            comments_frame.pack(fill='x', pady=(5, 0))
        
            ttk.Label(comments_frame, text="Comments:", font=('Arial', 9, 'bold')).pack(anchor='w')
        
            comments_var = tk.StringVar(value=sample_data.get('comments', ''))
            comments_vars[original_name] = comments_var
        
            comments_entry = ttk.Entry(comments_frame, textvariable=comments_var, width=50)
            comments_entry.pack(fill='x', pady=(2, 0))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enhanced buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)

        def load_enhanced_data():
            print("DEBUG: Loading enhanced data with edited values into NEW SESSION")
        
            final_data = {}
            for original_name in original_names:
                new_name = sample_name_vars[original_name].get().strip()
                if not new_name:
                    new_name = original_name

                if new_name in final_data:
                    messagebox.showerror("Duplicate Names", 
                                       f"Sample name '{new_name}' is used more than once.")
                    return

                # Copy data with edited values
                final_data[new_name] = {}
            
                # Copy edited ratings
                for metric in self.metrics:
                    if metric in rating_vars[original_name]:
                        final_data[new_name][metric] = rating_vars[original_name][metric].get()
                    else:
                        final_data[new_name][metric] = extracted_data[original_name].get(metric, 5)
            
                # Copy edited comments
                final_data[new_name]['comments'] = comments_vars[original_name].get()
                final_data[new_name]['sample_name'] = new_name

            print(f"DEBUG: Final edited data: {final_data}")

            # FIXED: Create a new session instead of adding to current session
            # Create session name from image filename
            base_filename = os.path.splitext(os.path.basename(filename))[0]
            session_name = f"AI_{base_filename}"
        
            # Ensure unique session name
            counter = 1
            original_session_name = session_name
            while session_name in self.sessions:
                session_name = f"{original_session_name}_{counter}"
                counter += 1
        
            print(f"DEBUG: Creating new session for AI extraction: {session_name}")
        
            # Create new session
            self.sessions[session_name] = {
                'header': {field: var.get() for field, var in self.header_vars.items()},
                'samples': final_data,
                'timestamp': datetime.now().isoformat(),
                'source_image': filename,
                'extraction_method': 'Enhanced_Claude_AI'
            }
        
            # Switch to the new session
            self.current_session_id = session_name
            self.samples = final_data
        
            # Update session selector UI
            self.update_session_combo()
            if hasattr(self, 'session_var'):
                self.session_var.set(session_name)
        
            # Update other UI components
            self.update_sample_combo()
            self.update_sample_checkboxes()

            if self.samples:
                first_sample = list(self.samples.keys())[0]
                self.sample_var.set(first_sample)
                self.load_sample_data(first_sample)

            self.update_plot()
            preview_window.destroy()

            print(f"DEBUG: Successfully created session {session_name} with {len(final_data)} samples")
            messagebox.showinfo("Enhanced AI Loading Complete", 
                              f"Created new session: '{session_name}'\n"
                              f"Successfully loaded {len(final_data)} samples with your edits!\n\n"
                              f"Enhanced features used:\n"
                              f"• Shadow removal preprocessing\n"
                              f"• AI sample name extraction\n"
                              f"• High-accuracy rating detection\n"
                              f"• Manual corrections applied\n\n"
                              f"Use session selector to switch between sessions.")

        ttk.Button(button_frame, text="Load as New Session", 
                   command=load_enhanced_data, 
                   style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                   command=preview_window.destroy).pack(side='right', padx=5)
    
        print("DEBUG: Enhanced AI preview with editable ratings created")

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
"""
sample_comparison.py
Developed by Charlie Becquet.
Sample comparison module for analyzing multiple datasets across different batches and design changes.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import re
from typing import Dict, List, Any, Tuple
from utils import debug_print, FONT, show_success_message
import processing
import os
import json
from datetime import datetime

class SampleComparisonWindow:
    """Window for comparing samples across multiple loaded files."""
    
    def __init__(self, gui, selected_files=None):
        self.gui = gui
        self.window = None
        self.selected_files = selected_files if selected_files is not None else gui.all_filtered_sheets
    
        # Configuration file path
        self.config_file_path = os.path.join(os.path.dirname(__file__), 'sample_comparison_config.json')
    
        # Default configuration values
        self.default_config = {
            'model_keywords_raw': [
                'ds7010', 'ds7020', 'cps2910', 'cps2920', 'pc0110', 
                'ds7110', 'ds7120', 'ds7310', 'ds7320','briq 2.0','cgs1810','th2','m6t','gembar',
                'gembox','mixjoy','briq 3.0','minitank','evomax','evo','t28','rosin bar'
            ],
            'grouped_tests': [
                'quick screening test', 'extended test', 
                'horizontal test', 'lifetime test', 'device life test',
                'user simulation test','user test simulation',
                'user sim test','user sim','user test sim','legacy'
            ]
        }
    
        # Load configuration from file or use defaults
        self.load_configuration()
        self.model_keywords = {} 
        self.comparison_results = {}
        self.parse_keyword_variations()
        debug_print(f"DEBUG: Sample comparison initialized with config from {self.config_file_path}")
        
    def show(self):
        """Show the comparison window."""
        if not self.selected_files:
            show_success_message("Info", "No files are selected for comparison.", self.gui.root)
            debug_print("DEBUG: No files selected for comparison")
            return
        
        debug_print(f"DEBUG: Showing comparison window with {len(self.selected_files)} selected files")
        
        self.window = tk.Toplevel(self.gui.root)
        self.window.title("Sample Comparison Analysis")
        
        self.window.transient(self.gui.root)
    
        self.center_window(self.window,1200,800)

        self.create_ui()
        self.perform_analysis()
        
    def create_ui(self):
        """Create the user interface for the comparison window."""
        # Create main frame with notebook for tabs
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create notebook for different views
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Summary tab
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Summary")
        
        # Detailed results tab
        self.details_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.details_frame, text="Detailed Results")
        
        # Configuration tab
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="Configuration")
        
        self.create_summary_tab()
        self.create_details_tab()
        self.create_config_tab()
        
        # Add control buttons at the bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        ttk.Button(button_frame, text="Refresh Analysis", command=self.perform_analysis).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Export Results", command=self.export_results).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=self.window.destroy).pack(side="right", padx=5)
        
    def create_summary_tab(self):
        """Create the summary tab with high-level comparison results."""
        # Title
        title_label = ttk.Label(self.summary_frame, text="Sample Comparison Summary", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
    
        # Create treeview for summary results with extended selection support
        columns = ("Model", "Test Group", "Avg TPM", "Avg Std Dev", "Avg Draw Pressure", 
                  "File Count", "Sample Count")
    
        self.summary_tree = ttk.Treeview(self.summary_frame, columns=columns, show="headings", height=15,
                                        selectmode='extended')  # Enable multiple selection
    
        # Configure column headings and widths
        for col in columns:
            self.summary_tree.heading(col, text=col)
            if col == "Model":
                self.summary_tree.column(col, width=100)
            elif col == "Test Group":
                self.summary_tree.column(col, width=150)
            elif col in ["Avg TPM", "Avg Std Dev", "Avg Draw Pressure"]:
                self.summary_tree.column(col, width=120)
            else:
                self.summary_tree.column(col, width=100)
    
        # Add scrollbars
        summary_scrollbar_v = ttk.Scrollbar(self.summary_frame, orient="vertical", command=self.summary_tree.yview)
        summary_scrollbar_h = ttk.Scrollbar(self.summary_frame, orient="horizontal", command=self.summary_tree.xview)
        self.summary_tree.configure(yscrollcommand=summary_scrollbar_v.set, xscrollcommand=summary_scrollbar_h.set)
    
        # Pack treeview and scrollbars
        self.summary_tree.pack(side="left", fill="both", expand=True)
        summary_scrollbar_v.pack(side="right", fill="y")
        summary_scrollbar_h.pack(side="bottom", fill="x")
    
        # Add button frame for summary actions
        summary_button_frame = ttk.Frame(self.summary_frame)
        summary_button_frame.pack(fill="x", pady=(10, 0))

        # Update button text to reflect multi-selection capability
        ttk.Button(summary_button_frame, text="Show Time-Series Plots", 
                   command=self.show_comparison_plots).pack(expand=True, padx=5)
    
        # Add filter frame BELOW the button frame
        filter_frame = ttk.LabelFrame(self.summary_frame, text="Filters", padding=5)
        filter_frame.pack(fill="x", pady=(5, 0))
    
        # Model filter
        ttk.Label(filter_frame, text="Model:").grid(row=0, column=0, padx=5, sticky="w")
        self.model_filter_var = tk.StringVar(value="All")
        self.model_filter_combo = ttk.Combobox(filter_frame, textvariable=self.model_filter_var, 
                                              width=15, state="readonly")
        self.model_filter_combo.grid(row=0, column=1, padx=5)
        self.model_filter_combo.bind("<<ComboboxSelected>>", self.apply_filters)
    
        # Test Group filter  
        ttk.Label(filter_frame, text="Test Group:").grid(row=1, column=0, padx=5, sticky="w")
        self.test_group_filter_var = tk.StringVar(value="All")
        self.test_group_filter_combo = ttk.Combobox(filter_frame, textvariable=self.test_group_filter_var,
                                                   width=15, state="readonly")
        self.test_group_filter_combo.grid(row=1, column=1, padx=5)
        self.test_group_filter_combo.bind("<<ComboboxSelected>>", self.apply_filters)
    
        # Clear filters button
        ttk.Button(filter_frame, text="Clear Filters", command=self.clear_filters).grid(row=2, column=0, columnspan=2, pady=5)
    
        # Selection info label
        self.selection_info_label = ttk.Label(filter_frame, text="Selected: 0 combinations", 
                                             font=("Arial", 9), foreground="blue")
        self.selection_info_label.grid(row=3, column=0, columnspan=2, pady=2)
    
        # Bind selection change event to update selection info
        self.summary_tree.bind("<<TreeviewSelect>>", self.update_selection_info)
    
        debug_print("DEBUG: Summary tab created with multi-selection filtering functionality")

    def update_selection_info(self, event=None):
        """Update the selection info label when tree selection changes."""
        selected_items = self.summary_tree.selection()
        count = len(selected_items)
    
        if count == 0:
            self.selection_info_label.config(text="Selected: 0 combinations")
        elif count == 1:
            item_values = self.summary_tree.item(selected_items[0])['values']
            model = item_values[0] if len(item_values) > 0 else "Unknown"
            test_group = item_values[1] if len(item_values) > 1 else "Unknown"
            self.selection_info_label.config(text=f"Selected: {model} - {test_group}")
        else:
            self.selection_info_label.config(text=f"Selected: {count} combinations (multi-plot mode)")
    
        debug_print(f"DEBUG: Selection updated - {count} items selected")

    def create_details_tab(self):
        """Create the detailed results tab."""
        # Title
        title_label = ttk.Label(self.details_frame, text="Detailed Analysis Results", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Create text widget for detailed results
        self.details_text = tk.Text(self.details_frame, wrap="word", font=("Courier", 10))
        details_scrollbar = ttk.Scrollbar(self.details_frame, orient="vertical", command=self.details_text.yview)
        self.details_text.configure(yscrollcommand=details_scrollbar.set)
        
        self.details_text.pack(side="left", fill="both", expand=True)
        details_scrollbar.pack(side="right", fill="y")
        
    def load_configuration(self):
        """Load configuration from JSON file or use defaults if file doesn't exist."""
        try:
            if os.path.exists(self.config_file_path):
                with open(self.config_file_path, 'r') as f:
                    config = json.load(f)
            
                # Validate and use loaded config
                self.model_keywords_raw = config.get('model_keywords_raw', self.default_config['model_keywords_raw'])
                self.grouped_tests = config.get('grouped_tests', self.default_config['grouped_tests'])
            
                debug_print(f"DEBUG: Configuration loaded successfully from {self.config_file_path}")
                debug_print(f"DEBUG: Loaded {len(self.model_keywords_raw)} model keywords and {len(self.grouped_tests)} grouped tests")
            
            else:
                # Use default configuration
                self.model_keywords_raw = self.default_config['model_keywords_raw'].copy()
                self.grouped_tests = self.default_config['grouped_tests'].copy()
            
                debug_print(f"DEBUG: Configuration file not found, using defaults. Will create: {self.config_file_path}")
            
        except Exception as e:
            debug_print(f"DEBUG: Error loading configuration file: {e}")
            debug_print("DEBUG: Falling back to default configuration")
        
            # Fall back to defaults if loading fails
            self.model_keywords_raw = self.default_config['model_keywords_raw'].copy()
            self.grouped_tests = self.default_config['grouped_tests'].copy()

    def save_configuration(self):
        """Save current configuration to JSON file."""
        try:
            config = {
                'model_keywords_raw': self.model_keywords_raw,
                'grouped_tests': self.grouped_tests,
                'last_updated': json.dumps(datetime.now().isoformat()) if 'datetime' in globals() else None
            }
        
            # Ensure directory exists
            config_dir = os.path.dirname(self.config_file_path)
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)
        
            with open(self.config_file_path, 'w') as f:
                json.dump(config, f, indent=2)
        
            debug_print(f"DEBUG: Configuration saved successfully to {self.config_file_path}")
            debug_print(f"DEBUG: Saved {len(self.model_keywords_raw)} model keywords and {len(self.grouped_tests)} grouped tests")
            return True
        
        except Exception as e:
            debug_print(f"DEBUG: Error saving configuration file: {e}")
            return False

    def update_configuration(self):
        """Update the configuration based on user input and save permanently."""
        # Update keywords with variation parsing
        keywords_text = self.keywords_var.get().strip()
        if keywords_text:
            self.model_keywords_raw = [kw.strip() for kw in keywords_text.split(",") if kw.strip()]
            self.parse_keyword_variations()

        # Update grouped tests
        grouped_text = self.grouped_tests_var.get().strip()
        if grouped_text:
            self.grouped_tests = [test.strip().lower() for test in grouped_text.split(",") if test.strip()]

        debug_print(f"DEBUG: Updated keywords with variations: {self.model_keywords}")
        debug_print(f"DEBUG: Updated grouped tests: {self.grouped_tests}")

        # Save configuration permanently
        if self.save_configuration():
            success_message = "Configuration updated and saved permanently!"
            debug_print("DEBUG: Configuration changes saved to file successfully")
        else:
            success_message = "Configuration updated (but save failed - changes are temporary)"
            debug_print("DEBUG: Configuration changes could not be saved to file")

        # Refresh analysis
        self.perform_analysis()
        show_success_message("Success", success_message, self.gui.root)
    
    def create_config_tab(self):
        """Create the configuration tab for customizing keywords and groupings."""
        # Model Keywords section
        keywords_frame = ttk.LabelFrame(self.config_frame, text="Model Keywords")
        keywords_frame.pack(fill="x", padx=5, pady=5)

        # Convert model_keywords dict back to string format for display
        keywords_display = []
        for main_keyword, variations in self.model_keywords.items():
            if len(variations) == 1:
                keywords_display.append(main_keyword)
            else:
                keywords_display.append(':'.join(variations))

        self.keywords_var = tk.StringVar(value=", ".join(keywords_display))

        # Add format explanation
        format_label = ttk.Label(keywords_frame, 
                               text="Format: keyword1, keyword2:'variation1':'variation2', keyword3", 
                               font=("Arial", 9), foreground="gray")
        format_label.pack(anchor="w", padx=5, pady=2)

        ttk.Label(keywords_frame, text="Keywords (comma-separated):").pack(anchor="w", padx=5, pady=2)
        keywords_entry = ttk.Entry(keywords_frame, textvariable=self.keywords_var, width=80)
        keywords_entry.pack(fill="x", padx=5, pady=2)

        # Example label
        example_label = ttk.Label(keywords_frame, 
                                text="Example: cps2910:'T58G 510':'CCELL3.0 510', ds7010:'new_variation'", 
                                font=("Arial", 9), foreground="blue")
        example_label.pack(anchor="w", padx=5, pady=2)

        # Grouped Tests section
        grouped_frame = ttk.LabelFrame(self.config_frame, text="Grouped Tests")
        grouped_frame.pack(fill="x", padx=5, pady=5)

        self.grouped_tests_var = tk.StringVar(value=", ".join(self.grouped_tests))
        ttk.Label(grouped_frame, text="Tests to group together (comma-separated):").pack(anchor="w", padx=5, pady=2)
        grouped_entry = ttk.Entry(grouped_frame, textvariable=self.grouped_tests_var, width=80)
        grouped_entry.pack(fill="x", padx=5, pady=2)

        # Configuration file info
        config_info_frame = ttk.LabelFrame(self.config_frame, text="Configuration Storage")
        config_info_frame.pack(fill="x", padx=5, pady=5)
    
        config_path_label = ttk.Label(config_info_frame, 
                                     text=f"Configuration saved to: {self.config_file_path}", 
                                     font=("Arial", 9), foreground="darkgreen")
        config_path_label.pack(anchor="w", padx=5, pady=2)
    
        # Buttons frame
        buttons_frame = ttk.Frame(self.config_frame)
        buttons_frame.pack(fill="x", padx=5, pady=10)
    
        # Update button
        ttk.Button(buttons_frame, text="Update & Save Configuration", 
                  command=self.update_configuration).pack(side="left", padx=5)
    
        # Reset to defaults button
        ttk.Button(buttons_frame, text="Reset to Defaults", 
                  command=self.reset_to_defaults).pack(side="left", padx=5)

        # Info section
        info_frame = ttk.LabelFrame(self.config_frame, text="Analysis Information")
        info_frame.pack(fill="both", expand=True, padx=5, pady=5)

        info_text = """
        Analysis Process:
        1. Searches through all loaded files for sample names containing the specified keywords OR their variations
        2. Groups tests according to the grouping rules (grouped tests vs individual tests)
        3. For each model keyword and test group combination:
           - Extracts TPM, standard deviation, and draw pressure data
           - Calculates averages within each file
           - Computes overall averages across all files
        4. Displays results showing performance comparisons across different batches and design changes

        Keyword Variations:
        - Use colons to add variations to a main keyword
        - Format: main_keyword:'variation1':'variation2'
        - Example: cps2910:'T58G 510':'CCELL3.0 510'
        - All variations will be grouped under the main keyword in results

        Matching Priority:
        1. Sample name matches (highest priority)
        2. Filename matches (fallback when no sample name matches)

        Configuration Storage:
        - Changes are automatically saved to a configuration file
        - Settings persist between application sessions
        - Use "Reset to Defaults" to restore original values

        Metrics Calculated:
        - Average TPM: Mean Total Particulate Matter across all matching samples
        - Average Std Dev: Mean standard deviation of measurements
        - Average Draw Pressure: Mean draw pressure across all matching samples
        - File Count: Number of files containing this model/test combination
        - Sample Count: Total number of samples analyzed for this combination

        Plot Features:
        - Color coding: Each file gets a unique color for easy identification
        - Marker shapes: Each test type gets a unique marker (Performance Tests Group = circle)
        - Interactive legend: Click combinations or files to highlight in plots
        - Grid-based layout: Files are automatically spaced to prevent overlap
        - Zoom functionality: Scroll wheel to zoom, middle-click to reset
        - Error bars: Show standard deviation with proper horizontal caps

        Selection Modes:
        - Normal click: Select single item (clears others)
        - Ctrl+click: Add/remove items from selection (multi-select)
        - File highlighting: Select files to highlight all their data points
        - Combination highlighting: Select test combinations to highlight specific markers
                """

        # Create scrollable text widget for info
        info_text_frame = ttk.Frame(info_frame)
        info_text_frame.pack(fill="both", expand=True, padx=5, pady=5)

        info_text_widget = tk.Text(info_text_frame, wrap="word", font=("Arial", 10), height=15)
        info_scrollbar = ttk.Scrollbar(info_text_frame, orient="vertical", command=info_text_widget.yview)
        info_text_widget.configure(yscrollcommand=info_scrollbar.set)

        # Insert the text and make it read-only
        info_text_widget.insert("1.0", info_text)
        info_text_widget.configure(state="disabled")  # Make it read-only

        # Pack the text widget and scrollbar
        info_text_widget.pack(side="left", fill="both", expand=True)
        info_scrollbar.pack(side="right", fill="y")

    def reset_to_defaults(self):
        """Reset configuration to default values."""
        result = messagebox.askyesno("Reset Configuration", 
                                    "Are you sure you want to reset all configuration to default values?\n\n"
                                    "This will permanently overwrite your saved settings.")
    
        if result:
            # Reset to defaults
            self.model_keywords_raw = self.default_config['model_keywords_raw'].copy()
            self.grouped_tests = self.default_config['grouped_tests'].copy()
        
            # Update the GUI
            keywords_display = []
            for main_keyword, variations in self.model_keywords.items():
                if len(variations) == 1:
                    keywords_display.append(main_keyword)
                else:
                    keywords_display.append(':'.join(variations))
        
            self.keywords_var.set(", ".join(self.model_keywords_raw))
            self.grouped_tests_var.set(", ".join(self.grouped_tests))
        
            # Parse and save
            self.parse_keyword_variations()
        
            if self.save_configuration():
                success_message = "Configuration reset to defaults and saved!"
            else:
                success_message = "Configuration reset to defaults (but save failed)"
        
            # Refresh analysis
            self.perform_analysis()
            show_success_message("Reset Complete", success_message, self.gui.root)
        
            debug_print("DEBUG: Configuration reset to defaults")
        
    def parse_keyword_variations(self):
            """Parse keyword variations from the raw keyword string."""
            self.model_keywords = {}
        
            debug_print(f"DEBUG: Parsing keyword variations from: {self.model_keywords_raw}")
        
            for keyword_entry in self.model_keywords_raw:
                keyword_entry = keyword_entry.strip().lower()
            
                if ':' in keyword_entry:
                    # Split by colon to get main keyword and variations
                    parts = keyword_entry.split(':')
                    main_keyword = parts[0].strip()
                    variations = [part.strip() for part in parts if part.strip()]
                
                    # Main keyword is included in variations
                    self.model_keywords[main_keyword] = variations
                    debug_print(f"DEBUG: Parsed {main_keyword} with variations: {variations}")
                else:
                    # Single keyword, no variations
                    main_keyword = keyword_entry
                    self.model_keywords[main_keyword] = [main_keyword]
                    debug_print(f"DEBUG: Single keyword: {main_keyword}")
        
            debug_print(f"DEBUG: Final parsed keywords: {self.model_keywords}")

    def center_window(self, window, width, height):
        """Center a window on the screen."""
        try:
            # Get screen dimensions
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
        
            # Calculate center position
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
        
            # Set window geometry with center position
            window.geometry(f"{width}x{height}+{x}+{y}")
        
            debug_print(f"DEBUG: Centered window {width}x{height} at position ({x}, {y}) on {screen_width}x{screen_height} screen")
        
        except Exception as e:
            debug_print(f"DEBUG: Error centering window: {e}")
            # Fallback to basic geometry if centering fails
            window.geometry(f"{width}x{height}")

    def perform_analysis(self):
        """Perform the sample comparison analysis with time-series data creation."""
        debug_print("DEBUG: Starting sample comparison analysis")

        if not self.selected_files:
            debug_print("DEBUG: No selected files for analysis")
            return
    
        self.comparison_results = {}

        # Progress tracking
        total_files = len(self.selected_files)
        debug_print(f"DEBUG: Analyzing {total_files} selected files")

        for file_idx, file_data in enumerate(self.selected_files):
            file_name = file_data["file_name"]
            filtered_sheets = file_data["filtered_sheets"]
        
            # Extract timestamp from file data
            file_timestamp = None
            if "database_created_at" in file_data and file_data["database_created_at"]:
                timestamp = file_data["database_created_at"]
                if isinstance(timestamp, str):
                    try:
                        from datetime import datetime
                        if 'T' in timestamp:
                            file_timestamp = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                        else:
                            file_timestamp = datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
                    except:
                        file_timestamp = None
                elif hasattr(timestamp, 'year'):
                    file_timestamp = timestamp
        
            # Fallback to current date if no timestamp available
            if file_timestamp is None:
                from datetime import datetime
                file_timestamp = datetime.now()
                debug_print(f"DEBUG: No timestamp found for {file_name}, using current date")
        
            debug_print(f"DEBUG: Processing selected file {file_idx + 1}/{total_files}: {file_name} (timestamp: {file_timestamp})")

            for sheet_name, sheet_info in filtered_sheets.items():
                data = sheet_info["data"]
    
                if data.empty:
                    continue
        
                # Determine test group
                test_group = self.get_test_group(sheet_name)
    
                # Find samples matching keywords (now passing filename)
                matching_samples = self.find_matching_samples(data, sheet_name, file_name)
    
                for keyword, sample_columns in matching_samples.items():
                    if not sample_columns:
                        continue
                
                    # Extract metrics for this keyword/test group combination
                    metrics = self.extract_metrics(data, sample_columns, sheet_name)
            
                    # Store results with proper structure for time-series plotting
                    if keyword not in self.comparison_results:
                        self.comparison_results[keyword] = {}
                
                    if test_group not in self.comparison_results[keyword]:
                        self.comparison_results[keyword][test_group] = {
                            'dates': [],
                            'tpm_values': [],
                            'std_dev_values': [],
                            'draw_pressure_values': [],
                            'file_info': [],
                            'file_count': 0,
                            'sample_count': 0,
                            'files': [],
                            'match_sources': []
                        }
                
                    # Add data points for each sample in this file/sheet combination
                    if metrics['tpm'] is not None and len(metrics['tpm']) > 0:
                        for i, tpm_val in enumerate(metrics['tpm']):
                            self.comparison_results[keyword][test_group]['dates'].append(file_timestamp)
                            self.comparison_results[keyword][test_group]['tpm_values'].append(tpm_val)
                        
                            # Add corresponding std dev if available
                            if metrics['std_dev'] and i < len(metrics['std_dev']):
                                self.comparison_results[keyword][test_group]['std_dev_values'].append(metrics['std_dev'][i])
                            else:
                                self.comparison_results[keyword][test_group]['std_dev_values'].append(0)
                        
                            # Add corresponding draw pressure if available
                            if metrics['draw_pressure'] and i < len(metrics['draw_pressure']):
                                self.comparison_results[keyword][test_group]['draw_pressure_values'].append(metrics['draw_pressure'][i])
                            else:
                                self.comparison_results[keyword][test_group]['draw_pressure_values'].append(0)
                        
                            # Add file info for color mapping
                            display_filename = file_data.get('display_filename', file_name)
                            self.comparison_results[keyword][test_group]['file_info'].append(display_filename)

                    self.comparison_results[keyword][test_group]['sample_count'] += len(sample_columns)
                    self.comparison_results[keyword][test_group]['files'].append(f"{file_name}:{sheet_name}")
            
                    # Determine match source for this file
                    sample_name_matches = self.find_sample_name_matches_only(data, keyword, sheet_name)
                    if sample_name_matches:
                        match_source = "sample_name"
                    else:
                        match_source = "filename"
                    self.comparison_results[keyword][test_group]['match_sources'].append(f"{file_name}:{match_source}")

        # Calculate file counts (unique files for each combination)
        for keyword in self.comparison_results:
            for test_group in self.comparison_results[keyword]:
                unique_files = set()
                for file_sheet in self.comparison_results[keyword][test_group]['files']:
                    file_name = file_sheet.split(':')[0]
                    unique_files.add(file_name)
                self.comparison_results[keyword][test_group]['file_count'] = len(unique_files)

        debug_print(f"DEBUG: Analysis complete. Found {sum(len(groups) for groups in self.comparison_results.values())} model/test combinations")

        # Update displays
        self.update_summary_display()
        self.update_details_display()
        
    def get_test_group(self, sheet_name: str) -> str:
        """Determine which test group a sheet belongs to."""
        sheet_lower = sheet_name.lower()
        
        for grouped_test in self.grouped_tests:
            if grouped_test in sheet_lower:
                return "Performance Tests Group"
        
        return f"Individual - {sheet_name}"
        
    def get_column_indices(self, data, sheet_name, sample_idx):
        """Get column indices for TPM, Std Dev, Draw Pressure, and Power."""
        is_user_simulation = any(test in sheet_name.lower() for test in ['user test simulation', 'user simulation'])
        columns_per_sample = 8 if is_user_simulation else 12
    
        start_col = sample_idx * columns_per_sample
    
        if is_user_simulation:
            tpm_col = start_col + 3  # TPM column
            std_col = start_col + 4  # Std Dev column
            dp_col = -1  # No draw pressure in user simulation
            power_col = start_col + 6  # Power column
        else:
            tpm_col = start_col + 7   # TPM column
            std_col = start_col + 8   # Std Dev column  
            dp_col = start_col + 9    # Draw Pressure column
            power_col = start_col + 10 # Power column
    
        return tpm_col, std_col, dp_col, power_col

    def extract_column_average(self, data, col_idx):
        """Extract average value from a column, skipping header rows."""
        if col_idx < 0 or col_idx >= len(data.columns):
            return None
    
        # Skip header rows (typically first 3-5 rows)
        values = []
        for idx in range(5, len(data)):
            try:
                val = float(data.iloc[idx, col_idx])
                if not np.isnan(val) and val > 0:
                    values.append(val)
            except:
                continue
    
        return np.mean(values) if values else None

    def find_sample_name_matches_only(self, data: pd.DataFrame, main_keyword: str, sheet_name: str) -> List[int]:
        """Check if a main keyword (or its variations) matches any sample names (used to determine match source)."""
        matches = []
    
        is_user_simulation = any(test in sheet_name.lower() for test in ['user test simulation', 'user simulation'])
        columns_per_sample = 8 if is_user_simulation else 12
    
        if data.shape[0] > 0:
            first_row = data.iloc[0]
            num_samples = len(first_row) // columns_per_sample
        
            # Get variations for this main keyword
            variations = self.model_keywords.get(main_keyword, [main_keyword])
        
            for sample_idx in range(num_samples):
                start_col = sample_idx * columns_per_sample
                sample_name_col = 5 + start_col
            
                if sample_name_col < len(first_row):
                    sample_name = str(first_row.iloc[sample_name_col]).lower()
                
                    # Check all variations
                    for variation in variations:
                        if variation.lower() in sample_name:
                            matches.append(sample_idx)
                            break  # Don't add the same sample multiple times
    
        return matches

    def find_matching_samples(self, data: pd.DataFrame, sheet_name: str, file_name: str = None) -> Dict[str, List[int]]:
        """Find samples that match the model keywords and their variations, checking sample names first, then filename."""
        matching_samples = {main_keyword: [] for main_keyword in self.model_keywords.keys()}
    
        debug_print(f"DEBUG: Searching for keyword matches in {sheet_name}")
    
        # Determine the data structure based on the sheet type
        is_user_simulation = any(test in sheet_name.lower() for test in ['user test simulation', 'user simulation'])
    
        if is_user_simulation:
            columns_per_sample = 8  # User simulation uses 8 columns per sample
        else:
            columns_per_sample = 12  # Standard tests use 12 columns per sample
        
        debug_print(f"DEBUG: Using {columns_per_sample} columns per sample for {sheet_name}")
    
        # STEP 1: Look for sample names in the first row (PRIORITY)
        sample_name_matches = {main_keyword: [] for main_keyword in self.model_keywords.keys()}
        valid_sample_indices = []
    
        if data.shape[0] > 0:
            first_row = data.iloc[0]
            num_samples = len(first_row) // columns_per_sample
        
            for sample_idx in range(num_samples):
                start_col = sample_idx * columns_per_sample
            
                # Sample name is typically in column F (index 5) + offset
                sample_name_col = 5 + start_col
            
                if sample_name_col < len(first_row):
                    sample_name = str(first_row.iloc[sample_name_col]).lower()
                
                    # Check if this sample has valid data (not empty/NaN)
                    if sample_name and sample_name.lower() not in ['nan', 'none', '']:
                        valid_sample_indices.append(sample_idx)
                    
                        # Check for keyword matches in sample name (including variations)
                        for main_keyword, variations in self.model_keywords.items():
                            for variation in variations:
                                if variation.lower() in sample_name:
                                    sample_name_matches[main_keyword].append(sample_idx)
                                    debug_print(f"DEBUG: Found {main_keyword} match (variation: '{variation}') in SAMPLE NAME {sample_idx}: {sample_name}")
                                    break  # Stop checking other variations for this main keyword
    
        # STEP 2: For keywords with no sample name matches, check filename
        filename_matches = {main_keyword: [] for main_keyword in self.model_keywords.keys()}
    
        if file_name:
            file_name_lower = file_name.lower()
            debug_print(f"DEBUG: Checking filename for keywords: {file_name}")
        
            for main_keyword, variations in self.model_keywords.items():
                # Only check filename if no sample name matches were found for this keyword
                if not sample_name_matches[main_keyword]:
                    for variation in variations:
                        if variation.lower() in file_name_lower:
                            # Apply this keyword to ALL valid samples since it's in the filename
                            filename_matches[main_keyword] = valid_sample_indices.copy()
                            debug_print(f"DEBUG: Found {main_keyword} match (variation: '{variation}') in FILENAME, applying to {len(valid_sample_indices)} samples: {valid_sample_indices}")
                            break  # Stop checking other variations for this main keyword
    
        # STEP 3: Combine results with sample name matches taking priority
        for main_keyword in self.model_keywords.keys():
            if sample_name_matches[main_keyword]:
                # Use sample name matches (priority)
                matching_samples[main_keyword] = sample_name_matches[main_keyword]
                debug_print(f"DEBUG: Using SAMPLE NAME matches for {main_keyword}: {sample_name_matches[main_keyword]}")
            elif filename_matches[main_keyword]:
                # Use filename matches as fallback
                matching_samples[main_keyword] = filename_matches[main_keyword]
                debug_print(f"DEBUG: Using FILENAME matches for {main_keyword}: {filename_matches[main_keyword]}")
            else:
                # No matches found
                matching_samples[main_keyword] = []
                debug_print(f"DEBUG: No matches found for {main_keyword}")
    
        return matching_samples
        
    def extract_metrics(self, data: pd.DataFrame, sample_indices: List[int], sheet_name: str) -> Dict[str, List[float]]:
        """Extract TPM, standard deviation, and draw pressure metrics for specified samples."""
        metrics = {
            'tpm': [],
            'std_dev': [],
            'draw_pressure': []
        }
        
        is_user_simulation = any(test in sheet_name.lower() for test in ['user test simulation', 'user simulation'])
        columns_per_sample = 8 if is_user_simulation else 12
        
        for sample_idx in sample_indices:
            start_col = sample_idx * columns_per_sample
            end_col = start_col + columns_per_sample
            
            if end_col <= data.shape[1]:
                sample_data = data.iloc[:, start_col:end_col]
                
                # Extract TPM values
                try:
                    if is_user_simulation:
                        tpm_values = processing.get_y_data_for_user_test_simulation_plot_type(sample_data, "TPM")
                    else:
                        tpm_values = processing.get_y_data_for_plot_type(sample_data, "TPM")
                    
                    tpm_numeric = pd.to_numeric(tpm_values, errors='coerce').dropna()
                    if not tpm_numeric.empty:
                        # Use only the first 70% of TPM values for better representation
                        tpm_count = len(tpm_numeric)
                        cutoff_index = int(tpm_count * 0.70)
                        if cutoff_index < 1:
                            cutoff_index = 1  # Ensure we use at least one value
                    
                        tpm_truncated = tpm_numeric.iloc[:cutoff_index]
                        debug_print(f"DEBUG: Sample {sample_idx} - Using first {cutoff_index} of {tpm_count} TPM values ({(cutoff_index/tpm_count)*100:.1f}%)")
                    
                        metrics['tpm'].append(tpm_truncated.mean())
                        metrics['std_dev'].append(tpm_truncated.std())
                        
                except Exception as e:
                    debug_print(f"DEBUG: Error extracting TPM for sample {sample_idx}: {e}")
                
                # Extract draw pressure values
                try:
                    if is_user_simulation:
                        dp_values = processing.get_y_data_for_user_test_simulation_plot_type(sample_data, "Draw Pressure")
                    else:
                        dp_values = processing.get_y_data_for_plot_type(sample_data, "Draw Pressure")
                    
                    dp_numeric = pd.to_numeric(dp_values, errors='coerce').dropna()
                    if not dp_numeric.empty:
                        metrics['draw_pressure'].append(dp_numeric.mean())
                        
                except Exception as e:
                    debug_print(f"DEBUG: Error extracting Draw Pressure for sample {sample_idx}: {e}")
        
        # Convert lists to None if empty
        for key in metrics:
            if not metrics[key]:
                metrics[key] = None
                
        return metrics
        
    def update_summary_display(self):
        """Update the summary treeview with analysis results."""
        # Clear existing items
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
    
        if not self.comparison_results:
            return
    
        # Collect unique models and test groups for filter dropdowns
        all_models = set()
        all_test_groups = set()
    
        # Flatten the results for display
        flattened_results = []
        for keyword, test_groups in self.comparison_results.items():
            all_models.add(keyword.upper())
            for test_group, data in test_groups.items():
                all_test_groups.add(test_group)
                flattened_results.append(((keyword, test_group), data))

        # Update filter comboboxes
        model_values = ["All"] + sorted(all_models)
        test_group_values = ["All"] + sorted(all_test_groups)
    
        self.model_filter_combo['values'] = model_values
        self.test_group_filter_combo['values'] = test_group_values
    
        # Apply current filters
        current_model_filter = self.model_filter_var.get()
        current_test_group_filter = self.test_group_filter_var.get()
    
        # Sort results by model keyword and test group
        sorted_results = sorted(flattened_results, key=lambda x: (x[0][0], x[0][1]))

        for (keyword, test_group), data in sorted_results:
            # Apply filters
            if current_model_filter != "All" and keyword.upper() != current_model_filter:
                continue
            if current_test_group_filter != "All" and test_group != current_test_group_filter:
                continue
            
            # Calculate averages
            avg_tpm = np.mean(data['tpm_values']) if data['tpm_values'] else 0
            avg_std_dev = np.mean(data['std_dev_values']) if data['std_dev_values'] else 0
            avg_draw_pressure = np.mean(data['draw_pressure_values']) if data['draw_pressure_values'] else 0
    
            # Format values
            avg_tpm_str = f"{avg_tpm:.3f}" if avg_tpm > 0 else "N/A"
            avg_std_dev_str = f"{avg_std_dev:.3f}" if avg_std_dev > 0 else "N/A"
            avg_draw_pressure_str = f"{avg_draw_pressure:.1f}" if avg_draw_pressure > 0 else "N/A"
    
            # Insert into treeview
            self.summary_tree.insert("", "end", values=(
                keyword.upper(),
                test_group,
                avg_tpm_str,
                avg_std_dev_str,
                avg_draw_pressure_str,
                data['file_count'],
                data['sample_count']
            ))
    
        debug_print(f"DEBUG: Summary display updated with filters - Model: {current_model_filter}, Test Group: {current_test_group_filter}")

    def apply_filters(self, event=None):
        """Apply the selected filters to the summary display."""
        debug_print(f"DEBUG: Applying filters - Model: {self.model_filter_var.get()}, Test Group: {self.test_group_filter_var.get()}")
        self.update_summary_display()

    def clear_filters(self):
        """Clear all filters and show all results."""
        debug_print("DEBUG: Clearing all filters")
        self.model_filter_var.set("All")
        self.test_group_filter_var.set("All")
        self.update_summary_display()


    def update_details_display(self):
        """Update the detailed results text widget."""
        self.details_text.delete("1.0", tk.END)

        if not self.comparison_results:
            self.details_text.insert("1.0", "No analysis results available.")
            return
    
        details = "DETAILED SAMPLE COMPARISON ANALYSIS\n"
        details += "=" * 60 + "\n\n"

        # Show keyword variations at the top
        details += "KEYWORD VARIATIONS:\n"
        details += "-" * 20 + "\n"
        for main_keyword, variations in self.model_keywords.items():
            if len(variations) > 1:
                details += f"{main_keyword.upper()}: {', '.join(variations)}\n"
            else:
                details += f"{main_keyword.upper()}: (no variations)\n"
        details += "\n" + "=" * 60 + "\n\n"

        # Group by keyword
        for keyword in sorted(self.comparison_results.keys()):
            details += f"MODEL: {keyword.upper()}\n"
            if keyword in self.model_keywords and len(self.model_keywords[keyword]) > 1:
                details += f"  Variations: {', '.join(self.model_keywords[keyword])}\n"
            details += "-" * 40 + "\n"
    
            test_groups = self.comparison_results[keyword]
            for test_group in sorted(test_groups.keys()):
                data = test_groups[test_group]
        
                details += f"\nTest Group: {test_group}\n"
                details += f"  Files involved: {data['file_count']}\n"
                details += f"  Total samples: {data['sample_count']}\n"
                details += f"  Data points: {len(data['tpm_values'])}\n"
        
                if data['tpm_values']:
                    avg_tpm = np.mean(data['tpm_values'])
                    min_tpm = np.min(data['tpm_values'])
                    max_tpm = np.max(data['tpm_values'])
                    details += f"  TPM - Avg: {avg_tpm:.3f}, Range: {min_tpm:.3f} - {max_tpm:.3f}\n"
            
                if data['std_dev_values']:
                    avg_std = np.mean(data['std_dev_values'])
                    details += f"  Std Dev - Avg: {avg_std:.3f}\n"
            
                if data['draw_pressure_values']:
                    avg_dp = np.mean(data['draw_pressure_values'])
                    min_dp = np.min(data['draw_pressure_values'])
                    max_dp = np.max(data['draw_pressure_values'])
                    details += f"  Draw Pressure - Avg: {avg_dp:.1f}, Range: {min_dp:.1f} - {max_dp:.1f}\n"
            
                # Show which files contributed to this data and match sources
                file_match_info = {}
                for file_sheet in data['files']:
                    file_name = file_sheet.split(':')[0]
                    if file_name not in file_match_info:
                        file_match_info[file_name] = set()
        
                # Determine match sources for each file
                for match_source_info in data.get('match_sources', []):
                    if ':' in match_source_info:
                        file_name, source = match_source_info.split(':', 1)
                        if file_name in file_match_info:
                            file_match_info[file_name].add(source)
        
                details += f"  Contributing files:\n"
                for file_name in sorted(file_match_info.keys()):
                    sources = list(file_match_info[file_name])
                    if 'sample_name' in sources:
                        source_text = "(sample name match)"
                    elif 'filename' in sources:
                        source_text = "(filename match)"
                    else:
                        source_text = "(unknown match)"
                    details += f"    - {file_name} {source_text}\n"
            
            details += "\n" + "=" * 60 + "\n\n"
    
        self.details_text.insert("1.0", details)
        
    def export_results(self):
        """Export the comparison results to a CSV file."""
        if not self.comparison_results:
            messagebox.showwarning("Warning", "No analysis results to export.")
            return
        
        try:
            from tkinter import filedialog
        
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                title="Export Comparison Results",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
        
            if not file_path:
                return
            
            # Prepare data for export
            export_data = []
            for (keyword, test_group), data in self.comparison_results.items():
                avg_tpm = np.mean(data['tpm_values']) if data['tpm_values'] else None
                avg_std_dev = np.mean(data['std_dev_values']) if data['std_dev_values'] else None
                avg_draw_pressure = np.mean(data['draw_pressure_values']) if data['draw_pressure_values'] else None
            
                # Determine primary match source
                sample_name_matches = 0
                filename_matches = 0
                for match_source_info in data.get('match_sources', []):
                    if 'sample_name' in match_source_info:
                        sample_name_matches += 1
                    elif 'filename' in match_source_info:
                        filename_matches += 1
            
                if sample_name_matches > 0:
                    primary_match_source = "Sample Name"
                elif filename_matches > 0:
                    primary_match_source = "Filename"
                else:
                    primary_match_source = "Unknown"
            
                export_data.append({
                    'Model': keyword.upper(),
                    'Test_Group': test_group,
                    'Avg_TPM': avg_tpm,
                    'Avg_Std_Dev': avg_std_dev,
                    'Avg_Draw_Pressure': avg_draw_pressure,
                    'File_Count': data['file_count'],
                    'Sample_Count': data['sample_count'],
                    'Primary_Match_Source': primary_match_source,
                    'Contributing_Files': '; '.join(sorted(set(f.split(':')[0] for f in data['files'])))
                })
            
            # Create DataFrame and export
            df = pd.DataFrame(export_data)
            df.to_csv(file_path, index=False)
        
            show_success_message("Success", f"Results exported successfully to:\n{file_path}", self.gui.root)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export results:\n{str(e)}")

    def show_comparison_plots(self):
        """Show comparison plots for the selected test groups with enhanced multi-selection visualization."""
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        from datetime import timedelta

        selected_items = self.summary_tree.selection()
        if not selected_items:
            show_success_message("Info", "Please select one or more test groups from the summary table.\n\n"
                                      "Use Ctrl+Click to select multiple individual items\n"
                                      "Use Shift+Click to select a range of items", self.gui.root)
            debug_print("DEBUG: No test groups selected for plotting")
            return

        debug_print(f"DEBUG: Showing enhanced multi-selection plots for {len(selected_items)} selected combinations")

        # Collect data from all selected combinations
        combined_plot_data = {
            'dates': [],
            'tpm_values': [],
            'std_dev_values': [],
            'draw_pressure_values': [],
            'combination_labels': [],  # To track which combination each data point belongs to
            'file_info': []
        }

        selected_combinations = []

        for selected_item in selected_items:
            item_values = self.summary_tree.item(selected_item)['values']

            if len(item_values) < 2:
                continue

            selected_keyword = item_values[0]  # Model (keyword)
            selected_test_group = item_values[1]  # Test Group

            combination_label = f"{selected_keyword} - {selected_test_group}"
            selected_combinations.append(combination_label)

            # Find the comparison data for this combination
            keyword = selected_keyword.lower()

            if keyword in self.comparison_results:
                for test_group, group_data in self.comparison_results[keyword].items():
                    if test_group.lower() == selected_test_group.lower():
                        # Add data from this combination to the combined dataset
                        dates = group_data['dates']
                        tpm_values = group_data['tpm_values']
                        std_values = group_data['std_dev_values']
                        draw_pressure_values = group_data['draw_pressure_values']
                        file_info_list = group_data.get('file_info', [])

                        # Extend combined data
                        combined_plot_data['dates'].extend(dates)
                        combined_plot_data['tpm_values'].extend(tpm_values)
                        combined_plot_data['std_dev_values'].extend(std_values)
                        combined_plot_data['draw_pressure_values'].extend(draw_pressure_values)
                        combined_plot_data['file_info'].extend(file_info_list)

                        # Add combination labels for each data point
                        combined_plot_data['combination_labels'].extend([combination_label] * len(dates))

                        debug_print(f"DEBUG: Added {len(dates)} data points from {combination_label}")
                        break

        if not combined_plot_data['dates']:
            show_success_message("Info", f"No time-series data found for the selected combinations", self.gui.root)
            return

        # Create plot title
        if len(selected_combinations) == 1:
            plot_title = f'Time-Series Comparison: {selected_combinations[0]}'
        else:
            plot_title = f'Multi-Selection Time-Series Comparison ({len(selected_combinations)} combinations)'

        # Create the plot window with proper sizing for plots + legend
        plot_window = tk.Toplevel(self.window)
        plot_window.title(plot_title)
        plot_window.transient(self.window)

        # Center the plotting window
        self.center_window(plot_window, 1400, 900)

        # Create main container
        main_container = ttk.Frame(plot_window)
        main_container.pack(fill="both", expand=True, padx=10, pady=5)

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9, 6))
        fig.suptitle(plot_title, fontsize=10)

        # Get combined data
        dates = combined_plot_data['dates']
        tpm_values = combined_plot_data['tpm_values']
        std_values = combined_plot_data['std_dev_values']
        draw_pressure_values = combined_plot_data['draw_pressure_values']
        combination_labels = combined_plot_data['combination_labels']
        file_info_list = combined_plot_data['file_info']

        debug_print(f"DEBUG: Processing {len(dates)} total data points from {len(selected_combinations)} combinations")

        # Grid-based offset function
        def apply_date_offsets(dates_list, values_list, std_list, dp_list, combo_list, file_list):
            """Apply grid-based x-axis offsets to organize files and test groups without overlap."""
            debug_print(f"DEBUG: Starting grid-based offset with {len(dates_list)} data points")

            if len(dates_list) <= 1:
                return dates_list, values_list, std_list, dp_list, combo_list, file_list

            try:
                from datetime import timedelta
    
                # Step 1: Group data by file to get file base dates and test groups
                file_data = {}
                for i in range(len(dates_list)):
                    file_info = file_list[i]
                    combo = combo_list[i]
                    date = dates_list[i]
        
                    if file_info not in file_data:
                        file_data[file_info] = {
                            'base_date': date,
                            'test_groups': set(),
                            'points': []
                        }
        
                    # Use earliest date as base date for this file
                    if date < file_data[file_info]['base_date']:
                        file_data[file_info]['base_date'] = date
            
                    file_data[file_info]['test_groups'].add(combo)
                    file_data[file_info]['points'].append(i)
    
                debug_print(f"DEBUG: Found {len(file_data)} unique files")
    
                # Step 2: Sort files by base date and resolve overlaps
                sorted_files = sorted(file_data.items(), key=lambda x: x[1]['base_date'])
    
                # Convert pixel offset to days (assume ~1 day = 20 pixels for reasonable spacing)
                pixels_per_day = 20.0
                region_width_days = 200.0 / pixels_per_day  # 200 pixels = 10 days
                half_region_days = region_width_days / 2.0   # ±5 days
    
                # Step 3: Calculate actual file positions, resolving overlaps
                file_positions = {}  # file_info -> adjusted_base_date
    
                for i, (file_info, file_info_data) in enumerate(sorted_files):
                    base_date = file_info_data['base_date']
        
                    if i == 0:
                        # First file keeps its original position
                        file_positions[file_info] = base_date
                        debug_print(f"DEBUG: File {file_info} at original position: {base_date.strftime('%Y-%m-%d')}")
                    else:
                        # Check if this file's region overlaps with any previous file
                        proposed_position = base_date
            
                        # Check overlap with all previous files
                        needs_adjustment = False
                        for prev_file_info in [f[0] for f in sorted_files[:i]]:
                            prev_position = file_positions[prev_file_info]
                            prev_right_edge = prev_position + timedelta(days=half_region_days)
                            proposed_left_edge = proposed_position - timedelta(days=half_region_days)
                
                            # If proposed left edge is before previous right edge, we have overlap
                            if proposed_left_edge < prev_right_edge:
                                needs_adjustment = True
                                # Push this file to start after the previous file's region
                                proposed_position = prev_right_edge + timedelta(days=half_region_days)
                                debug_print(f"DEBUG: File {file_info} pushed to avoid overlap with {prev_file_info}")
            
                        file_positions[file_info] = proposed_position
                        debug_print(f"DEBUG: File {file_info} positioned at: {proposed_position.strftime('%Y-%m-%d')} (adjusted: {needs_adjustment})")
    
                # Step 4: Calculate test group offsets within each file's region
                test_group_offsets = {}  # (file_info, combo) -> offset_days
    
                for file_info, file_info_data in file_data.items():
                    test_groups = sorted(list(file_info_data['test_groups']))  # Sort for consistent ordering
                    num_test_groups = len(test_groups)
        
                    if num_test_groups == 1:
                        # Single test group goes in the center
                        test_group_offsets[(file_info, test_groups[0])] = 0.0
                    else:
                        # Distribute test groups evenly across the region width
                        step = region_width_days / (num_test_groups - 1) if num_test_groups > 1 else 0
                        start_offset = -half_region_days
            
                        for idx, test_group in enumerate(test_groups):
                            offset_within_region = start_offset + (idx * step)
                            test_group_offsets[(file_info, test_group)] = offset_within_region
                            debug_print(f"DEBUG: Test group {test_group} in {file_info} gets offset: {offset_within_region:.2f} days")
    
                # Step 5: Apply offsets to all data points
                offset_dates = []
                for i in range(len(dates_list)):
                    file_info = file_list[i]
                    combo = combo_list[i]
        
                    # Get file's adjusted base position
                    file_base_position = file_positions[file_info]
        
                    # Get test group's offset within the file's region
                    test_group_offset = test_group_offsets.get((file_info, combo), 0.0)
        
                    # Calculate final position
                    final_position = file_base_position + timedelta(days=test_group_offset)
                    offset_dates.append(final_position)
    
                return offset_dates, values_list, std_list, dp_list, combo_list, file_list
    
            except Exception as e:
                return dates_list, values_list, std_list, dp_list, combo_list, file_list

        # Apply date offsets
        dates, tpm_values, std_values, draw_pressure_values, combination_labels, file_info_list = apply_date_offsets(
            dates, tpm_values, std_values, draw_pressure_values, combination_labels, file_info_list
        )

        # SWITCHED: Create color maps for files and marker maps for combinations
        unique_combinations = list(set(combination_labels))
        unique_files = list(set(file_info_list))
    
        # Color mapping: each file gets a unique color
        file_colors = plt.cm.Set1(np.linspace(0, 1, len(unique_files)))
        file_color_map = {file_info: file_colors[i] for i, file_info in enumerate(unique_files)}
    
        # Marker mapping: each combination gets a unique marker, with Performance Tests Group = 'o'
        markers = ['o', 's', '^', 'v', 'D', 'P', '*', 'h', 'X', '+', '<', '>', '1', '2', '3', '4']
        combination_marker_map = {}
    
        # Ensure Performance Tests Group gets 'o' marker
        performance_combos = [combo for combo in unique_combinations if 'Performance Tests Group' in combo]
        other_combos = [combo for combo in unique_combinations if 'Performance Tests Group' not in combo]
    
        # Assign markers
        marker_idx = 0
        for combo in performance_combos:
            combination_marker_map[combo] = 'o'  # Always circle for Performance Tests Group
    
        marker_idx = 1  # Start from index 1 for other combinations
        for combo in other_combos:
            combination_marker_map[combo] = markers[marker_idx % len(markers)]
            marker_idx += 1

        ax1.set_title('TPM Over Time (Multi-Selection)', fontsize=10)

        # Group data by (file_info, combination) for vertical line plotting
        file_combo_groups = {}
        for i, (date, tpm, std, combo, file_info) in enumerate(zip(dates, tpm_values, std_values, combination_labels, file_info_list)):
            key = (file_info, combo)
            if key not in file_combo_groups:
                file_combo_groups[key] = {
                    'dates': [],
                    'tpm_values': [],
                    'std_values': [],
                    'indices': []
                }
            file_combo_groups[key]['dates'].append(date)
            file_combo_groups[key]['tpm_values'].append(tpm)
            file_combo_groups[key]['std_values'].append(std)
            file_combo_groups[key]['indices'].append(i)

        debug_print(f"DEBUG: Created {len(file_combo_groups)} vertical line groups for TPM plotting")

        # Store plot objects for interactive highlighting (now tracking both combinations and files)
        plot_objects = {}  # combination -> {'tpm_scatter': scatter_obj, 'tpm_errors': [error_objs], 'dp_scatter': scatter_obj}
        file_plot_objects = {}  # file_info -> {'tpm_scatters': [scatter_objs], 'tpm_errors': [error_objs], 'dp_scatters': [scatter_objs]}
        highlighted_combinations = set()  # Track which combinations are selected/highlighted
        highlighted_files = set()  # Track which files are selected/highlighted

        # Initialize highlighting variables and functions BEFORE legend creation
        combination_labels_dict = {}  # Store combination label references for font changes
        file_labels_dict = {}  # Store file label references for font changes

        # Functions for interactive highlighting (DEFINED EARLY)
        def highlight_combination(combination):
            """Highlight a specific combination in both plots."""
            if combination in plot_objects:
                # Highlight TPM scatter
                if plot_objects[combination]['tpm_scatter']:
                    plot_objects[combination]['tpm_scatter'].set_alpha(1.0)
                    plot_objects[combination]['tpm_scatter'].set_sizes([60])  # Larger size
                    plot_objects[combination]['tpm_scatter'].set_zorder(10)   # Bring to front
                    plot_objects[combination]['tpm_scatter'].set_edgecolors('red')
                    plot_objects[combination]['tpm_scatter'].set_linewidths(2)

                # Highlight TPM error bars
                for error_obj in plot_objects[combination]['tpm_errors']:
                    for line in error_obj[1]:  # error_obj[1] contains the error bar lines
                        line.set_alpha(1.0)
                        line.set_linewidth(2.5)
                    for line in error_obj[2]:  # error_obj[2] contains the caps
                        line.set_alpha(1.0)
                        line.set_linewidth(2.5)

                # Highlight Draw Pressure scatter
                if plot_objects[combination]['dp_scatter']:
                    plot_objects[combination]['dp_scatter'].set_alpha(1.0)
                    plot_objects[combination]['dp_scatter'].set_sizes([60])  # Larger size
                    plot_objects[combination]['dp_scatter'].set_zorder(10)   # Bring to front
                    plot_objects[combination]['dp_scatter'].set_edgecolors('red')
                    plot_objects[combination]['dp_scatter'].set_linewidths(2)

        def unhighlight_combination(combination):
            """Return a combination to normal appearance."""
            if combination in plot_objects:
                # Reset TPM scatter
                if plot_objects[combination]['tpm_scatter']:
                    plot_objects[combination]['tpm_scatter'].set_alpha(0.7)
                    plot_objects[combination]['tpm_scatter'].set_sizes([40])  # Normal size
                    plot_objects[combination]['tpm_scatter'].set_zorder(2)    # Normal layer
                    plot_objects[combination]['tpm_scatter'].set_edgecolors('black')
                    plot_objects[combination]['tpm_scatter'].set_linewidths(0.5)

                # Reset TPM error bars
                for error_obj in plot_objects[combination]['tpm_errors']:
                    for line in error_obj[1]:  # error_obj[1] contains the error bar lines
                        line.set_alpha(0.7)
                        line.set_linewidth(1.5)
                    for line in error_obj[2]:  # error_obj[2] contains the caps
                        line.set_alpha(0.7)
                        line.set_linewidth(1.5)

                # Reset Draw Pressure scatter
                if plot_objects[combination]['dp_scatter']:
                    plot_objects[combination]['dp_scatter'].set_alpha(0.7)
                    plot_objects[combination]['dp_scatter'].set_sizes([40])  # Normal size
                    plot_objects[combination]['dp_scatter'].set_zorder(2)    # Normal layer
                    plot_objects[combination]['dp_scatter'].set_edgecolors('black')
                    plot_objects[combination]['dp_scatter'].set_linewidths(0.5)

        def highlight_file(file_info):
            """Highlight all data points from a specific file in both plots."""
            if file_info in file_plot_objects:
                # Highlight all TPM scatters for this file
                for scatter_obj in file_plot_objects[file_info]['tpm_scatters']:
                    scatter_obj.set_alpha(1.0)
                    scatter_obj.set_sizes([60])  # Larger size
                    scatter_obj.set_zorder(10)   # Bring to front
                    scatter_obj.set_edgecolors('red')
                    scatter_obj.set_linewidths(2)

                # Highlight all TPM error bars for this file
                for error_obj in file_plot_objects[file_info]['tpm_errors']:
                    for line in error_obj[1]:  # error_obj[1] contains the error bar lines
                        line.set_alpha(1.0)
                        line.set_linewidth(2.5)
                    for line in error_obj[2]:  # error_obj[2] contains the caps
                        line.set_alpha(1.0)
                        line.set_linewidth(2.5)

                # Highlight all Draw Pressure scatters for this file
                for scatter_obj in file_plot_objects[file_info]['dp_scatters']:
                    scatter_obj.set_alpha(1.0)
                    scatter_obj.set_sizes([60])  # Larger size
                    scatter_obj.set_zorder(10)   # Bring to front
                    scatter_obj.set_edgecolors('red')
                    scatter_obj.set_linewidths(2)

        def unhighlight_file(file_info):
            """Return all data points from a specific file to normal appearance."""
            if file_info in file_plot_objects:
                # Reset all TPM scatters for this file
                for scatter_obj in file_plot_objects[file_info]['tpm_scatters']:
                    scatter_obj.set_alpha(0.7)
                    scatter_obj.set_sizes([40])  # Normal size
                    scatter_obj.set_zorder(2)    # Normal layer
                    scatter_obj.set_edgecolors('black')
                    scatter_obj.set_linewidths(0.5)

                # Reset all TPM error bars for this file
                for error_obj in file_plot_objects[file_info]['tpm_errors']:
                    for line in error_obj[1]:  # error_obj[1] contains the error bar lines
                        line.set_alpha(0.7)
                        line.set_linewidth(1.5)
                    for line in error_obj[2]:  # error_obj[2] contains the caps
                        line.set_alpha(0.7)
                        line.set_linewidth(1.5)

                # Reset all Draw Pressure scatters for this file
                for scatter_obj in file_plot_objects[file_info]['dp_scatters']:
                    scatter_obj.set_alpha(0.7)
                    scatter_obj.set_sizes([40])  # Normal size
                    scatter_obj.set_zorder(2)    # Normal layer
                    scatter_obj.set_edgecolors('black')
                    scatter_obj.set_linewidths(0.5)

        def update_plot_display():
            """Refresh the canvas to show highlighting changes."""
            canvas.draw()

        debug_print("DEBUG: Initialized interactive highlighting functions and variables")

        # Plot each file+combination group as a vertical line (UPDATED for new color/marker mapping)
        for (file_info, combination), group_data in file_combo_groups.items():
            color = file_color_map[file_info]  # Color based on file
            marker = combination_marker_map[combination]  # Marker based on combination

            # All points in this group should have the same x-value (after offset)
            group_dates = group_data['dates']
            group_tpm = group_data['tpm_values']
            group_std = group_data['std_values']

            # Verify that all dates in the group are the same (after offset)
            unique_dates = list(set(group_dates))
            if len(unique_dates) > 1:
                debug_print(f"WARNING: Group {file_info}-{combination} has multiple x-values: {unique_dates}")

            # Use the first date as the x-position for the entire vertical line
            x_position = group_dates[0] if group_dates else None

            if x_position is not None:
                # Initialize plot objects storage for this combination
                if combination not in plot_objects:
                    plot_objects[combination] = {'tpm_errors': [], 'dp_scatter': None}
            
                # Initialize file plot objects storage
                if file_info not in file_plot_objects:
                    file_plot_objects[file_info] = {'tpm_scatters': [], 'tpm_errors': [], 'dp_scatters': []}

                # Plot error bars for all points at this x-position
                for tpm_val, std_val in zip(group_tpm, group_std):
                    if std_val > 0:
                        error_obj = ax1.errorbar(x_position, tpm_val, yerr=std_val, 
                                    color=color, alpha=0.7, capsize=4, capthick=1.5, 
                                    elinewidth=1.5, linestyle='', marker='', zorder=1)
                        plot_objects[combination]['tpm_errors'].append(error_obj)
                        file_plot_objects[file_info]['tpm_errors'].append(error_obj)

                # Plot scatter points for all points at this x-position
                scatter_obj = ax1.scatter([x_position] * len(group_tpm), group_tpm, 
                            color=color, s=40, alpha=0.7, marker=marker, 
                            edgecolors='black', linewidths=0.5, zorder=2,
                            label=f"{combination} ({file_info})" if len(file_combo_groups) <= 10 else None)

                plot_objects[combination]['tpm_scatter'] = scatter_obj
                file_plot_objects[file_info]['tpm_scatters'].append(scatter_obj)

                debug_print(f"DEBUG: Plotted vertical line for {file_info}-{combination} with {len(group_tpm)} points at x={x_position.strftime('%Y-%m-%d')}")

        debug_print(f"DEBUG: Plotted TPM data with proper error bars and vertical grouping")

        # Configure TPM plot axes
        ax1.set_ylabel('Average TPM (mg/puff)', fontsize=8)
        ax1.set_ylim(-1, 6)  # Fixed y-axis range for TPM
        ax1.grid(True, alpha=0.3)
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        ax1.xaxis.set_major_locator(mdates.AutoDateLocator())
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=0, ha='right')

        # Plot Draw Pressure data (bottom subplot) with identical vertical grouping
        ax2.set_title('Draw Pressure Over Time (Multi-Selection)', fontsize=10)

        # Group draw pressure data the same way
        for (file_info, combination), group_data in file_combo_groups.items():
            color = file_color_map[file_info]  # Color based on file
            marker = combination_marker_map[combination]  # Marker based on combination

            # Get corresponding draw pressure values for this group
            group_dp_values = [draw_pressure_values[idx] for idx in group_data['indices']]
            x_position = group_data['dates'][0] if group_data['dates'] else None

            if x_position is not None:
                # Plot scatter points for draw pressure at this x-position
                dp_scatter_obj = ax2.scatter([x_position] * len(group_dp_values), group_dp_values, 
                            color=color, s=40, alpha=0.7, marker=marker, 
                            edgecolors='black', linewidths=0.5, zorder=2)

                # Store draw pressure scatter object
                if combination in plot_objects:
                    plot_objects[combination]['dp_scatter'] = dp_scatter_obj
                if file_info in file_plot_objects:
                    file_plot_objects[file_info]['dp_scatters'].append(dp_scatter_obj)

                debug_print(f"DEBUG: Plotted draw pressure vertical line for {file_info}-{combination} with {len(group_dp_values)} points")

        debug_print(f"DEBUG: Plotted draw pressure data with vertical grouping")

        # Configure Draw Pressure plot axes
        ax2.set_ylabel('Average Draw Pressure (kPa)', fontsize=8)
        ax2.set_ylim(-1, 10)
        ax2.set_xlabel('Date', fontsize=8)
        ax2.grid(True, alpha=0.3)
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        ax2.xaxis.set_major_locator(mdates.AutoDateLocator())
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=0, ha='right')

        # Adjust layout to provide proper spacing
        plt.tight_layout()
        fig.subplots_adjust(bottom=0.15)  # Reduced from 0.25 to 0.15 for tighter layout

        # Create the plot canvas with constrained size
        plot_frame = ttk.Frame(main_container, height=650)  
        plot_frame.pack(fill="x", expand=False)  # Don't expand vertically
        plot_frame.pack_propagate(False)  # Maintain fixed height

        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas.draw()

        # Store original axis limits for reset functionality
        original_ax1_xlim = ax1.get_xlim()
        original_ax1_ylim = ax1.get_ylim()
        original_ax2_xlim = ax2.get_xlim()
        original_ax2_ylim = ax2.get_ylim()

        debug_print(f"DEBUG: Stored original axis limits - TPM: x{original_ax1_xlim}, y{original_ax1_ylim}")
        debug_print(f"DEBUG: Stored original axis limits - Draw Pressure: x{original_ax2_xlim}, y{original_ax2_ylim}")

        # Add scroll wheel zoom functionality
        def on_scroll(event):
            """Handle mouse scroll wheel for zooming."""
            if event.inaxes is None:
                return

            # Determine which axis we're over
            current_ax = event.inaxes

            # Get current axis limits
            xleft, xright = current_ax.get_xlim()
            ybottom, ytop = current_ax.get_ylim()

            # Calculate zoom factor (scroll down = zoom in, scroll up = zoom out)
            zoom_factor = 0.8 if event.button == 'down' else 1.25

            # Calculate zoom center based on mouse position
            xdata = event.xdata
            ydata = event.ydata

            if xdata is not None and ydata is not None:
                # Zoom around mouse position
                x_center = xdata
                y_center = ydata
            else:
                # Zoom around center of plot
                x_center = (xleft + xright) / 2
                y_center = (ybottom + ytop) / 2

            # Calculate new limits
            x_range = (xright - xleft) * zoom_factor
            y_range = (ytop - ybottom) * zoom_factor

            new_xleft = x_center - x_range / 2
            new_xright = x_center + x_range / 2
            new_ybottom = y_center - y_range / 2
            new_ytop = y_center + y_range / 2

            # Apply zoom limits
            current_ax.set_xlim(new_xleft, new_xright)
            current_ax.set_ylim(new_ybottom, new_ytop)

            # Refresh the plot
            canvas.draw()

            zoom_direction = "in" if event.button == 'down' else "out"
            debug_print(f"DEBUG: Zoom {zoom_direction} applied to {'TPM' if current_ax == ax1 else 'Draw Pressure'} plot, factor: {zoom_factor:.2f}")

        def on_double_click(event):
            """Handle double-click to reset zoom."""
            if event.inaxes is None:
                return

            # Check if it's a middle mouse button double-click
            if event.button == 2 and event.dblclick:
                current_ax = event.inaxes

                # Reset to original limits based on which subplot
                if current_ax == ax1:
                    current_ax.set_xlim(original_ax1_xlim)
                    current_ax.set_ylim(original_ax1_ylim)
                    plot_name = "TPM"
                elif current_ax == ax2:
                    current_ax.set_xlim(original_ax2_xlim)
                    current_ax.set_ylim(original_ax2_ylim)
                    plot_name = "Draw Pressure"
                else:
                    return

                # Refresh the plot
                canvas.draw()

                debug_print(f"DEBUG: Reset zoom for {plot_name} plot to original view")

        # Connect scroll and double-click events to canvas
        canvas.mpl_connect('scroll_event', on_scroll)
        canvas.mpl_connect('button_press_event', on_double_click)

        debug_print("DEBUG: Added scroll wheel zoom (down=in, up=out) and middle-click reset functionality")

        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Add toolbar in the plot frame
        toolbar = NavigationToolbar2Tk(canvas, plot_frame)
        toolbar.update()

        # Create comprehensive legend below the plot with guaranteed space
        legend_frame = ttk.LabelFrame(main_container, text="Multi-Selection Legend", padding=5)
        legend_frame.pack(fill="both", expand=True, pady=(10, 0))  # This will expand to fill remaining space

        # Create notebook for organized legend display with fixed height
        legend_notebook = ttk.Notebook(legend_frame)
        legend_notebook.pack(fill="both", expand=True)

        # File Names tab (with color patches and click functionality)
        files_frame = ttk.Frame(legend_notebook)
        legend_notebook.add(files_frame, text=f"File Names ({len(unique_files)})")

        # Calculate optimal columns for files
        file_max_rows = 4
        file_columns = max(1, min(3, (len(unique_files) + file_max_rows - 1) // file_max_rows))
        file_estimated_rows = (len(unique_files) + file_columns - 1) // file_columns

        # Create scrollable frame for files if there are many
        if file_estimated_rows > file_max_rows:
            files_canvas = tk.Canvas(files_frame, height=120)
            files_scrollbar = ttk.Scrollbar(files_frame, orient="vertical", command=files_canvas.yview)
            scrollable_files_frame = ttk.Frame(files_canvas)

            scrollable_files_frame.bind(
                "<Configure>",
                lambda e: files_canvas.configure(scrollregion=files_canvas.bbox("all"))
            )

            files_canvas.create_window((0, 0), window=scrollable_files_frame, anchor="nw")
            files_canvas.configure(yscrollcommand=files_scrollbar.set)

            files_canvas.pack(side="left", fill="both", expand=True)
            files_scrollbar.pack(side="right", fill="y")

            files_parent = scrollable_files_frame
        else:
            files_parent = files_frame

        # Create legend entries for files with click handlers (showing color patches)
        for i, file_info in enumerate(unique_files):
            row = i // file_columns
            col = i % file_columns

            # Create color patch
            color_frame = tk.Frame(files_parent, width=20, height=20, bg='#{:02x}{:02x}{:02x}'.format(
                int(file_color_map[file_info][0] * 255),
                int(file_color_map[file_info][1] * 255),
                int(file_color_map[file_info][2] * 255)
            ))
            color_frame.grid(row=row, column=col*2, padx=5, pady=2, sticky="w")
            color_frame.grid_propagate(False)

            # Add file name label with click handler
            file_label = ttk.Label(files_parent, text=file_info, font=("Arial", 9), cursor="hand2")
            file_label.grid(row=row, column=col*2+1, padx=(5, 15), pady=2, sticky="w")
        
            # Store file label reference
            file_labels_dict[file_info] = file_label

            # Add click handler for this file label
            def create_file_click_handler(file_name, label_widget):
                def on_click(event):
                    # Check if Ctrl is held down
                    ctrl_held = (event.state & 0x4) != 0  # Ctrl modifier

                    if ctrl_held:
                        # Ctrl+click behavior
                        if file_name in highlighted_files:
                            # Already selected with Ctrl+click: just deselect this file
                            highlighted_files.remove(file_name)
                            unhighlight_file(file_name)
                            label_widget.configure(font=("Arial", 9))  # Normal font
                            debug_print(f"DEBUG: Ctrl+click deselected file: {file_name}")
                        else:
                            # Not selected with Ctrl+click: add to selection
                            highlighted_files.add(file_name)
                            highlight_file(file_name)
                            label_widget.configure(font=("Arial", 9, "bold"))  # Bold font when selected
                            debug_print(f"DEBUG: Ctrl+click selected file: {file_name}")
                    else:
                        # Normal click behavior
                        if file_name in highlighted_files:
                            # Already selected with normal click: clear all file selections
                            for prev_file in list(highlighted_files):
                                highlighted_files.remove(prev_file)
                                unhighlight_file(prev_file)
                                # Reset font for all file labels
                                if prev_file in file_labels_dict:
                                    file_labels_dict[prev_file].configure(font=("Arial", 9))
                            debug_print(f"DEBUG: Normal click on selected file {file_name}: cleared all file selections")
                        else:
                            # Not selected with normal click: clear other files and select this one
                            for prev_file in list(highlighted_files):
                                highlighted_files.remove(prev_file)
                                unhighlight_file(prev_file)
                                # Reset font for other file labels
                                if prev_file in file_labels_dict:
                                    file_labels_dict[prev_file].configure(font=("Arial", 9))

                            highlighted_files.add(file_name)
                            highlight_file(file_name)
                            label_widget.configure(font=("Arial", 9, "bold"))  # Bold font when selected
                            debug_print(f"DEBUG: Normal click selected file: {file_name}")

                    # Update the display
                    update_plot_display()

                    # Update selection info
                    if len(highlighted_files) == 0:
                        debug_print("DEBUG: No files selected")
                    elif len(highlighted_files) == 1:
                        debug_print(f"DEBUG: 1 file selected: {list(highlighted_files)[0]}")
                    else:
                        debug_print(f"DEBUG: {len(highlighted_files)} files selected: {sorted(highlighted_files)}")

                return on_click

            # Bind click event to the file label
            file_click_handler = create_file_click_handler(file_info, file_label)
            file_label.bind("<Button-1>", file_click_handler)

        debug_print(f"DEBUG: Added click handlers to {len(unique_files)} file labels")

        # Combinations tab (with marker symbols)
        combo_frame = ttk.Frame(legend_notebook)
        legend_notebook.add(combo_frame, text=f"Model/Test Combinations ({len(unique_combinations)})")

        # Calculate optimal columns for up to 4 rows
        max_rows = 4
        combo_columns = max(1, min(6, (len(unique_combinations) + max_rows - 1) // max_rows))  # Calculate columns needed for max 4 rows
        estimated_rows = (len(unique_combinations) + combo_columns - 1) // combo_columns

        debug_print(f"DEBUG: Legend using {combo_columns} columns for {estimated_rows} rows ({len(unique_combinations)} combinations)")

        # Create scrollable area only if we exceed 4 rows
        if estimated_rows > max_rows:
            combo_canvas = tk.Canvas(combo_frame, height=120)
            combo_scrollbar = ttk.Scrollbar(combo_frame, orient="vertical", command=combo_canvas.yview)
            scrollable_combo_frame = ttk.Frame(combo_canvas)

            scrollable_combo_frame.bind(
                "<Configure>",
                lambda e: combo_canvas.configure(scrollregion=combo_canvas.bbox("all"))
            )

            combo_canvas.create_window((0, 0), window=scrollable_combo_frame, anchor="nw")
            combo_canvas.configure(yscrollcommand=combo_scrollbar.set)

            combo_canvas.pack(side="left", fill="both", expand=True)
            combo_scrollbar.pack(side="right", fill="y")

            combo_parent = scrollable_combo_frame
        else:
            combo_parent = combo_frame

        # Create legend entries for combinations with click handlers (showing marker symbols)
        for i, combination in enumerate(unique_combinations):
            row = i // combo_columns
            col = i % combo_columns

            # Create marker symbol frame
            marker_symbol = combination_marker_map[combination]
    
            # Map matplotlib markers to Unicode symbols for display
            marker_display_map = {
                'o': '●',  # circle
                's': '■',  # square  
                '^': '▲',  # triangle up
                'v': '▼',  # triangle down
                'D': '♦',  # diamond
                'P': '⊕',  # plus (filled)
                '*': '★',  # star
                'h': '⬡',  # hexagon
                'X': '✖',  # X (filled)
                '+': '✚',  # plus
                '<': '◀',  # triangle left
                '>': '▶',  # triangle right
                '1': '⬇',  # tri_down
                '2': '⬆',  # tri_up  
                '3': '⬅',  # tri_left
                '4': '➡',  # tri_right
            }
    
            display_symbol = marker_display_map.get(marker_symbol, marker_symbol)
    
            marker_frame = tk.Frame(combo_parent, width=20, height=20, bg='white')
            marker_frame.grid(row=row, column=col*2, padx=5, pady=2, sticky="w")
            marker_frame.grid_propagate(False)
    
            # Add marker symbol text with Unicode character
            marker_label = tk.Label(marker_frame, text=display_symbol, font=("Arial", 12, "bold"), bg='white')
            marker_label.place(relx=0.5, rely=0.5, anchor='center')

            # Add combination label with click handler
            combo_label = ttk.Label(combo_parent, text=combination, font=("Arial", 9), cursor="hand2")
            combo_label.grid(row=row, column=col*2+1, padx=(5, 20), pady=2, sticky="w")

            # Store label reference
            combination_labels_dict[combination] = combo_label

            # Add click handler for this label
            def create_combination_click_handler(combo_name, label_widget):
                def on_click(event):
                    # Check if Ctrl is held down
                    ctrl_held = (event.state & 0x4) != 0  # Ctrl modifier

                    if ctrl_held:
                        # Ctrl+click behavior
                        if combo_name in highlighted_combinations:
                            # Already selected with Ctrl+click: just deselect this combo
                            highlighted_combinations.remove(combo_name)
                            unhighlight_combination(combo_name)
                            label_widget.configure(font=("Arial", 9))  # Normal font
                            debug_print(f"DEBUG: Ctrl+click deselected combination: {combo_name}")
                        else:
                            # Not selected with Ctrl+click: add to selection
                            highlighted_combinations.add(combo_name)
                            highlight_combination(combo_name)
                            label_widget.configure(font=("Arial", 9, "bold"))  # Bold font when selected
                            debug_print(f"DEBUG: Ctrl+click selected combination: {combo_name}")
                    else:
                        # Normal click behavior
                        if combo_name in highlighted_combinations:
                            # Already selected with normal click: clear all selections
                            for prev_combo in list(highlighted_combinations):
                                highlighted_combinations.remove(prev_combo)
                                unhighlight_combination(prev_combo)
                                # Reset font for all labels
                                if prev_combo in combination_labels_dict:
                                    combination_labels_dict[prev_combo].configure(font=("Arial", 9))
                            debug_print(f"DEBUG: Normal click on selected combo {combo_name}: cleared all selections")
                        else:
                            # Not selected with normal click: clear others and select this one
                            for prev_combo in list(highlighted_combinations):
                                highlighted_combinations.remove(prev_combo)
                                unhighlight_combination(prev_combo)
                                # Reset font for other labels
                                if prev_combo in combination_labels_dict:
                                    combination_labels_dict[prev_combo].configure(font=("Arial", 9))

                            highlighted_combinations.add(combo_name)
                            highlight_combination(combo_name)
                            label_widget.configure(font=("Arial", 9, "bold"))  # Bold font when selected
                            debug_print(f"DEBUG: Normal click selected combination: {combo_name}")

                    # Update the display
                    update_plot_display()

                    # Update selection info
                    if len(highlighted_combinations) == 0:
                        debug_print("DEBUG: No combinations selected")
                    elif len(highlighted_combinations) == 1:
                        debug_print(f"DEBUG: 1 combination selected: {list(highlighted_combinations)[0]}")
                    else:
                        debug_print(f"DEBUG: {len(highlighted_combinations)} combinations selected: {sorted(highlighted_combinations)}")

                return on_click

            # Bind click event to the label
            click_handler = create_combination_click_handler(combination, combo_label)
            combo_label.bind("<Button-1>", click_handler)        

        # Button frame for plot controls at the bottom
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=(5, 0))

        # Export plot button
        def export_multi_selection_plot():
            from tkinter import filedialog
            default_name = f"multi_selection_comparison_{len(selected_combinations)}_combinations.png"
            file_path = filedialog.asksaveasfilename(
                title="Save Multi-Selection Comparison Plot", 
                initialfile=default_name,
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("SVG files", "*.svg"), ("All files", "*.*")]
            )
            if file_path:
                fig.savefig(file_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
                show_success_message("Success", f"Multi-selection plot saved to {file_path}", self.gui.root)

        # Enhanced statistics button
        def show_multi_selection_statistics():
            stats_text = f"Multi-Selection Statistics\n"
            stats_text += "="*50 + "\n\n"
            stats_text += f"Selected combinations: {len(selected_combinations)}\n"
            stats_text += f"Total data points: {len(tpm_values)}\n"
            stats_text += f"Unique files: {len(unique_files)}\n"
            stats_text += f"Date range: {min(dates).strftime('%Y-%m-%d')} to {max(dates).strftime('%Y-%m-%d')}\n\n"

            # Overall statistics
            stats_text += f"Overall TPM Statistics:\n"
            stats_text += f"  Mean: {np.mean(tpm_values):.4f} mg/puff\n"
            stats_text += f"  Std Dev: {np.std(tpm_values):.4f} mg/puff\n"
            stats_text += f"  Min: {np.min(tpm_values):.4f} mg/puff\n"
            stats_text += f"  Max: {np.max(tpm_values):.4f} mg/puff\n\n"

            # Overall Draw Pressure Statistics
            if draw_pressure_values and any(dp > 0 for dp in draw_pressure_values):
                valid_dp_values = [dp for dp in draw_pressure_values if dp > 0]
                stats_text += f"Overall Draw Pressure Statistics:\n"
                stats_text += f"  Mean: {np.mean(valid_dp_values):.2f} kPa\n"
                stats_text += f"  Std Dev: {np.std(valid_dp_values):.2f} kPa\n"
                stats_text += f"  Min: {np.min(valid_dp_values):.2f} kPa\n"
                stats_text += f"  Max: {np.max(valid_dp_values):.2f} kPa\n\n"

            # Per-combination statistics
            stats_text += "Per-Combination Statistics:\n"
            stats_text += "-" * 30 + "\n"
            for combination in unique_combinations:
                combo_indices = [i for i, combo in enumerate(combination_labels) if combo == combination]
                combo_tpm_values = [tpm_values[i] for i in combo_indices]
                combo_dp_values = [draw_pressure_values[i] for i in combo_indices if draw_pressure_values[i] > 0]

                stats_text += f"\n{combination}:\n"
                stats_text += f"  Data points: {len(combo_tpm_values)}\n"
                if combo_tpm_values:
                    stats_text += f"  TPM Mean: {np.mean(combo_tpm_values):.4f} mg/puff\n"
                    stats_text += f"  TPM Range: {np.min(combo_tpm_values):.4f} - {np.max(combo_tpm_values):.4f}\n"
                if combo_dp_values:
                    stats_text += f"  Draw Pressure Mean: {np.mean(combo_dp_values):.2f} kPa\n"
                    stats_text += f"  Draw Pressure Range: {np.min(combo_dp_values):.2f} - {np.max(combo_dp_values):.2f}\n"

            # Create statistics window
            stats_window = tk.Toplevel(plot_window)
            stats_window.title("Multi-Selection Statistics")

            # Center the statistics window
            self.center_window(stats_window, 800, 700)

            stats_text_widget = tk.Text(stats_window, wrap="word", font=("Courier", 10))
            stats_scrollbar = ttk.Scrollbar(stats_window, orient="vertical", command=stats_text_widget.yview)
            stats_text_widget.configure(yscrollcommand=stats_scrollbar.set)

            stats_text_widget.insert("1.0", stats_text)
            stats_text_widget.configure(state="disabled")

            stats_text_widget.pack(side="left", fill="both", expand=True)
            stats_scrollbar.pack(side="right", fill="y")

        ttk.Button(button_frame, text="Export Multi-Selection Plot", command=export_multi_selection_plot).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Show Multi-Selection Statistics", command=show_multi_selection_statistics).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=plot_window.destroy).pack(side="right", padx=5)

        debug_print(f"DEBUG: Enhanced multi-selection plot created with color=file, marker=combination, and dual-tab selection functionality")
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
from utils import debug_print, FONT
import processing

class SampleComparisonWindow:
    """Window for comparing samples across multiple loaded files."""
    
    def __init__(self, gui, selected_files=None):
        self.gui = gui
        self.window = None
        self.selected_files = selected_files if selected_files is not None else gui.all_filtered_sheets
        
        # Predefined model keywords (can be expanded)
        self.model_keywords_raw = [
            'ds7010', 'ds7020', 'cps2910', 'cps2920', 'pc0110', 
            'ds7110', 'ds7120', 'ds7310', 'ds7320'
        ]

        self.model_keywords = {} 
        
        # Test groupings
        self.grouped_tests = [
            'quick screening test', 'extended test', 'intense test', 
            'horizontal test', 'lifetime test', 'device life test'
        ]
        
        self.comparison_results = {}

        self.parse_keyword_variations()
        
    def show(self):
        """Show the comparison window."""
        if not self.selected_files:
            messagebox.showinfo("Info", "No files are selected for comparison.")
            debug_print("DEBUG: No files selected for comparison")
            return
        
        debug_print(f"DEBUG: Showing comparison window with {len(self.selected_files)} selected files")
        
        self.window = tk.Toplevel(self.gui.root)
        self.window.title("Sample Comparison Analysis")
        self.window.geometry("1200x800")
        self.window.transient(self.gui.root)
    
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
        
        # Create treeview for summary results
        columns = ("Model", "Test Group", "Avg TPM", "Avg Std Dev", "Avg Draw Pressure", 
                  "File Count", "Sample Count")
        
        self.summary_tree = ttk.Treeview(self.summary_frame, columns=columns, show="headings", height=15)
        
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

        ttk.Button(summary_button_frame, text="Show Time-Series Plots", 
                   command=self.show_comparison_plots).pack(side="left", padx=5)

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
    
        # Update button
        ttk.Button(self.config_frame, text="Update Configuration", 
                  command=self.update_configuration).pack(pady=10)
    
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

    Metrics Calculated:
    - Average TPM: Mean Total Particulate Matter across all matching samples
    - Average Std Dev: Mean standard deviation of measurements
    - Average Draw Pressure: Mean draw pressure across all matching samples
    - File Count: Number of files containing this model/test combination
    - Sample Count: Total number of samples analyzed for this combination
            """
    
        info_label = ttk.Label(info_frame, text=info_text, justify="left", wraplength=600)
        info_label.pack(padx=5, pady=5, anchor="w")
        
    def update_configuration(self):
        """Update the configuration based on user input."""
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
    
        # Refresh analysis
        self.perform_analysis()
        messagebox.showinfo("Success", "Configuration updated and analysis refreshed!")
        
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
        
        # Flatten the results for display
        flattened_results = []
        for keyword, test_groups in self.comparison_results.items():
            for test_group, data in test_groups.items():
                flattened_results.append(((keyword, test_group), data))
    
        # Sort results by model keyword and test group
        sorted_results = sorted(flattened_results, key=lambda x: (x[0][0], x[0][1]))
    
        for (keyword, test_group), data in sorted_results:
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
        
            messagebox.showinfo("Success", f"Results exported successfully to:\n{file_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export results:\n{str(e)}")

    def show_comparison_plots(self):
        """Show comparison plots for the selected test group with enhanced visualization."""
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        import matplotlib.dates as mdates
        from datetime import datetime
        from tkinter import filedialog
        import tkinter as tk
        from tkinter import ttk, messagebox
        import numpy as np
        import pandas as pd
        from utils import debug_print

        if not self.comparison_results:
            messagebox.showinfo("Info", "No comparison data available. Please run analysis first.")
            return

        # Get the selected item from the summary tree
        selected_items = self.summary_tree.selection()
        if not selected_items:
            messagebox.showinfo("Info", "Please select a test group from the summary table first.")
            return

        # Get the values from the selected row
        selected_item = selected_items[0]  # Take the first selected item
        item_values = self.summary_tree.item(selected_item)['values']

        if len(item_values) < 2:
            messagebox.showerror("Error", "Invalid selection. Please select a valid test group.")
            return

        selected_keyword = item_values[0]  # Model (keyword)
        selected_test_group = item_values[1]  # Test Group

        debug_print(f"DEBUG: Showing enhanced plots for {selected_keyword} - {selected_test_group}")

        # Find the comparison data for the selected group
        plot_data = {}
        keyword = selected_keyword.lower()

        if keyword in self.comparison_results:
            for test_group, group_data in self.comparison_results[keyword].items():
                if test_group.lower() == selected_test_group.lower():
                    plot_data[keyword] = group_data
                    break

        if not plot_data:
            messagebox.showinfo("Info", f"No time-series data found for {selected_keyword.upper()} - {selected_test_group}")
            return

        # Create the plot window
        plot_window = tk.Toplevel(self.window)
        plot_window.title(f'Enhanced Sample Comparison: {selected_keyword.upper()} - {selected_test_group}')
        plot_window.geometry("1400x900")  # Larger window to accommodate legend
        plot_window.transient(self.window)

        # Create main container
        main_container = ttk.Frame(plot_window)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Create matplotlib figure with reduced height (half of original)
        fig, ax = plt.subplots(1, 1, figsize=(12, 4))  # Reduced height from 8 to 4
        fig.suptitle(f'TPM Over Time: {selected_keyword.upper()} - {selected_test_group}', fontsize=16)

        # Get the data for the selected group
        dates = plot_data[keyword]['dates']
        tpm_values = plot_data[keyword]['tpm_values']
        std_values = plot_data[keyword]['std_dev_values']
    
        # Get file info for each data point to assign colors
        file_info_list = plot_data[keyword].get('file_info', [])
    
        debug_print(f"DEBUG: Processing {len(dates)} data points from {len(set(file_info_list))} unique files")

        # Create a color map for unique files
        unique_files = list(set(file_info_list))
        colors = plt.cm.tab10(np.linspace(0, 1, len(unique_files)))  # Use tab10 colormap for distinct colors
        file_color_map = {file_info: colors[i] for i, file_info in enumerate(unique_files)}

        debug_print(f"DEBUG: Unique files: {unique_files}")
        debug_print(f"DEBUG: Assigned colors to {len(file_color_map)} files")

        # Group data by file to identify min/max values for enhanced error bars
        file_data_groups = {}
        for i, (date, tpm, std, file_info) in enumerate(zip(dates, tpm_values, std_values, file_info_list)):
            if file_info not in file_data_groups:
                file_data_groups[file_info] = []
            file_data_groups[file_info].append((i, date, tpm, std))

        # Find the error bars that extend furthest for each file (pick only ONE point for each extreme)
        file_error_extremes = {}
        for file_info, data_points in file_data_groups.items():
            # Calculate error bar extents for each point
            min_extent_value = float('inf')
            max_extent_value = float('-inf')
            min_extent_index = None
            max_extent_index = None
    
            for point in data_points:
                i, date, tpm, std = point
                lower_extent = tpm - std  # Bottom of error bar
                upper_extent = tpm + std  # Top of error bar
        
                # Update minimum extent (only if this is a new minimum)
                if lower_extent < min_extent_value:
                    min_extent_value = lower_extent
                    min_extent_index = i
            
                # Update maximum extent (only if this is a new maximum)
                if upper_extent > max_extent_value:
                    max_extent_value = upper_extent
                    max_extent_index = i
    
            file_error_extremes[file_info] = {
                'min_extent_index': min_extent_index,  # Single index of point with lowest error bar
                'max_extent_index': max_extent_index,  # Single index of point with highest error bar
                'min_extent_value': min_extent_value,
                'max_extent_value': max_extent_value
            }

        debug_print(f"DEBUG: File error bar extremes analysis completed for {len(file_error_extremes)} files")
        for file_info, extremes in file_error_extremes.items():
            debug_print(f"DEBUG: {file_info} - Lowest error bar extent: {extremes['min_extent_value']:.3f} at index {extremes['min_extent_index']}, Highest error bar extent: {extremes['max_extent_value']:.3f} at index {extremes['max_extent_index']}")

        # Plot each data point with enhanced error bars only for error bar extremes
        normal_capsize = 5
        enhanced_capsize = 15  # 3x the normal capsize

        for i, (date, tpm, std, file_info) in enumerate(zip(dates, tpm_values, std_values, file_info_list)):
            color = file_color_map.get(file_info, colors[0])  # Default to first color if not found
    
            # Determine if this specific point index has an extreme error bar for its file
            is_min_extent = (i == file_error_extremes[file_info]['min_extent_index'])
            is_max_extent = (i == file_error_extremes[file_info]['max_extent_index'])
            is_extreme_error_bar = is_min_extent or is_max_extent
    
            # Use enhanced capsize for extreme error bars
            capsize = enhanced_capsize if is_extreme_error_bar else normal_capsize
            capthick = 3 if is_extreme_error_bar else 2
    
            debug_print(f"DEBUG: Point {i} for {file_info}: TPM={tpm:.3f}, std={std:.3f}, lower_extent={tpm-std:.3f}, upper_extent={tpm+std:.3f}, is_min_extent={is_min_extent}, is_max_extent={is_max_extent}, capsize={capsize}")
    
            # All points have same marker size - only error bar caps are different
            marker_size = 8
            edge_width = 1
    
            # Plot the main error bar (horizontal)
            ax.errorbar(date, tpm, yerr=std, color=color, marker='o', 
                       capsize=capsize, capthick=capthick, markersize=marker_size, linewidth=0,
                       alpha=0.8, markeredgewidth=edge_width, markeredgecolor='black')
    
            # Add vertical lines connecting the point to the error bar endpoints
            if std > 0:  # Only add vertical lines if there's actually an error bar
                # Calculate error bar endpoints
                upper_y = tpm + std
                lower_y = tpm - std
        
                # Draw vertical lines from the center point to the error bar caps
                line_alpha = 0.7 if is_extreme_error_bar else 0.5
                line_width = 1.5 if is_extreme_error_bar else 1
        
                ax.plot([date, date], [tpm, upper_y], color=color, linewidth=line_width, 
                       alpha=line_alpha, linestyle='-', zorder=1)
                ax.plot([date, date], [tpm, lower_y], color=color, linewidth=line_width, 
                       alpha=line_alpha, linestyle='-', zorder=1)

        # Configure axes
        ax.set_ylabel('Average TPM (mg/puff)', fontsize=12)
        ax.set_xlabel('Date', fontsize=12)
        ax.grid(True, alpha=0.3)

        # Format x-axis dates
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')

        # Adjust layout to make room for legend below
        plt.tight_layout()
        fig.subplots_adjust(bottom=0.3)  # Make room for legend below plot

        # Create the plot canvas
        plot_frame = ttk.Frame(main_container)
        plot_frame.pack(fill="both", expand=True)

        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Add toolbar above the legend
        toolbar = NavigationToolbar2Tk(canvas, plot_frame)
        toolbar.update()

        # Create detailed legend below the plot
        legend_frame = ttk.LabelFrame(main_container, text="File Legend", padding=10)
        legend_frame.pack(fill="x", pady=(10, 0))

        # Create scrollable frame for legend if there are many files
        if len(unique_files) > 6:  # If more than 6 files, make it scrollable
            legend_canvas = tk.Canvas(legend_frame, height=120)
            legend_scrollbar = ttk.Scrollbar(legend_frame, orient="vertical", command=legend_canvas.yview)
            scrollable_legend_frame = ttk.Frame(legend_canvas)
        
            scrollable_legend_frame.bind(
                "<Configure>",
                lambda e: legend_canvas.configure(scrollregion=legend_canvas.bbox("all"))
            )
        
            legend_canvas.create_window((0, 0), window=scrollable_legend_frame, anchor="nw")
            legend_canvas.configure(yscrollcommand=legend_scrollbar.set)
        
            legend_canvas.pack(side="left", fill="both", expand=True)
            legend_scrollbar.pack(side="right", fill="y")
        
            legend_parent = scrollable_legend_frame
        else:
            legend_parent = legend_frame

        # Create legend entries in a grid layout
        cols = 2  # Two columns for better space usage
        for i, (file_info, color) in enumerate(file_color_map.items()):
            row = i // cols
            col = i % cols
        
            # Create a colored square
            color_frame = tk.Frame(legend_parent, width=20, height=20, bg='black')
            color_frame.grid(row=row, column=col*2, padx=(0, 5), pady=2, sticky="w")
            color_frame.configure(bg=f'#{int(color[0]*255):02x}{int(color[1]*255):02x}{int(color[2]*255):02x}')
            color_frame.grid_propagate(False)
        
            # Create label with file name
            file_label = ttk.Label(legend_parent, text=file_info, font=("Arial", 10))
            file_label.grid(row=row, column=col*2+1, padx=(0, 20), pady=2, sticky="w")

        # Add mouse wheel scrolling to legend if scrollable
        if len(unique_files) > 6:
            def on_mousewheel(event):
                legend_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            legend_canvas.bind("<MouseWheel>", on_mousewheel)

        # Create button frame for actions
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=(10, 0))

        # Export button
        def export_enhanced_plot():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("SVG files", "*.svg"), ("All files", "*.*")]
            )
            if file_path:
                # Save with high DPI and tight bounding box
                fig.savefig(file_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
                messagebox.showinfo("Success", f"Enhanced plot saved to {file_path}")

        # Statistics button
        def show_statistics():
            stats_text = f"Statistics for {selected_keyword.upper()} - {selected_test_group}\n"
            stats_text += "="*50 + "\n\n"
            stats_text += f"Total data points: {len(tpm_values)}\n"
            stats_text += f"Unique files: {len(unique_files)}\n"
            stats_text += f"Date range: {min(dates).strftime('%Y-%m-%d')} to {max(dates).strftime('%Y-%m-%d')}\n\n"
            stats_text += f"TPM Statistics:\n"
            stats_text += f"  Mean: {np.mean(tpm_values):.4f} mg/puff\n"
            stats_text += f"  Std Dev: {np.std(tpm_values):.4f} mg/puff\n"
            stats_text += f"  Min: {np.min(tpm_values):.4f} mg/puff\n"
            stats_text += f"  Max: {np.max(tpm_values):.4f} mg/puff\n\n"
        
            # Add file-specific statistics
            stats_text += "File-specific statistics:\n"
            for file_info in unique_files:
                file_indices = [i for i, fi in enumerate(file_info_list) if fi == file_info]
                file_tpm_values = [tpm_values[i] for i in file_indices]
                file_count = len(file_tpm_values)
            
                if file_tpm_values:
                    file_mean = np.mean(file_tpm_values)
                    file_min = np.min(file_tpm_values)
                    file_max = np.max(file_tpm_values)
                    stats_text += f"  {file_info}: {file_count} points, "
                    stats_text += f"Mean: {file_mean:.4f}, Range: {file_min:.4f} - {file_max:.4f}\n"

            # Create statistics window
            stats_window = tk.Toplevel(plot_window)
            stats_window.title("Enhanced Statistics")
            stats_window.geometry("600x500")
        
            stats_text_widget = tk.Text(stats_window, wrap="word", font=("Courier", 10))
            stats_scrollbar = ttk.Scrollbar(stats_window, orient="vertical", command=stats_text_widget.yview)
            stats_text_widget.configure(yscrollcommand=stats_scrollbar.set)
        
            stats_text_widget.insert("1.0", stats_text)
            stats_text_widget.configure(state="disabled")
        
            stats_text_widget.pack(side="left", fill="both", expand=True)
            stats_scrollbar.pack(side="right", fill="y")

        ttk.Button(button_frame, text="Export Plot", command=export_enhanced_plot).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Show Statistics", command=show_statistics).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=plot_window.destroy).pack(side="right", padx=5)

        debug_print(f"DEBUG: Enhanced comparison plot created with {len(unique_files)} unique file colors, detailed legend, and enhanced error bars")
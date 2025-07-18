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
    
    def __init__(self, gui):
        self.gui = gui
        self.window = None
        
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
        if not self.gui.all_filtered_sheets:
            messagebox.showinfo("Info", "No files are currently loaded. Please load some data files first.")
            return
            
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
        """Perform the sample comparison analysis."""
        debug_print("DEBUG: Starting sample comparison analysis")
    
        if not self.gui.all_filtered_sheets:
            return
        
        self.comparison_results = {}
    
        # Progress tracking
        total_files = len(self.gui.all_filtered_sheets)
        debug_print(f"DEBUG: Analyzing {total_files} loaded files")
    
        for file_idx, file_data in enumerate(self.gui.all_filtered_sheets):
            file_name = file_data["file_name"]
            filtered_sheets = file_data["filtered_sheets"]
        
            debug_print(f"DEBUG: Processing file {file_idx + 1}/{total_files}: {file_name}")
        
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
                
                    # Store results
                    key = (keyword, test_group)
                    if key not in self.comparison_results:
                        self.comparison_results[key] = {
                            'tpm_values': [],
                            'std_dev_values': [],
                            'draw_pressure_values': [],
                            'file_count': 0,
                            'sample_count': 0,
                            'files': [],
                            'match_sources': []  # NEW: Track whether matches came from sample names or filenames
                        }
                
                    if metrics['tpm'] is not None:
                        self.comparison_results[key]['tpm_values'].extend(metrics['tpm'])
                    if metrics['std_dev'] is not None:
                        self.comparison_results[key]['std_dev_values'].extend(metrics['std_dev'])
                    if metrics['draw_pressure'] is not None:
                        self.comparison_results[key]['draw_pressure_values'].extend(metrics['draw_pressure'])
                
                    self.comparison_results[key]['sample_count'] += len(sample_columns)
                    self.comparison_results[key]['files'].append(f"{file_name}:{sheet_name}")
                
                    # Determine match source for this file
                    sample_name_matches = self.find_sample_name_matches_only(data, keyword, sheet_name)
                    if sample_name_matches:
                        match_source = "sample_name"
                    else:
                        match_source = "filename"
                    self.comparison_results[key]['match_sources'].append(f"{file_name}:{match_source}")
    
        # Calculate file counts (unique files for each combination)
        for key in self.comparison_results:
            unique_files = set()
            for file_sheet in self.comparison_results[key]['files']:
                file_name = file_sheet.split(':')[0]
                unique_files.add(file_name)
            self.comparison_results[key]['file_count'] = len(unique_files)
    
        debug_print(f"DEBUG: Analysis complete. Found {len(self.comparison_results)} model/test combinations")
    
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
                        metrics['tpm'].append(tpm_numeric.mean())
                        metrics['std_dev'].append(tpm_numeric.std())
                        
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
            
        # Sort results by model keyword and test group
        sorted_results = sorted(self.comparison_results.items(), 
                               key=lambda x: (x[0][0], x[0][1]))
        
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
        keyword_groups = {}
        for (keyword, test_group), data in self.comparison_results.items():
            if keyword not in keyword_groups:
                keyword_groups[keyword] = {}
            keyword_groups[keyword][test_group] = data
        
        for keyword in sorted(keyword_groups.keys()):
            details += f"MODEL: {keyword.upper()}\n"
            if keyword in self.model_keywords and len(self.model_keywords[keyword]) > 1:
                details += f"  Variations: {', '.join(self.model_keywords[keyword])}\n"
            details += "-" * 40 + "\n"
        
            test_groups = keyword_groups[keyword]
            for test_group in sorted(test_groups.keys()):
                data = test_groups[test_group]
            
                details += f"\nTest Group: {test_group}\n"
                details += f"  Files involved: {data['file_count']}\n"
                details += f"  Total samples: {data['sample_count']}\n"
            
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
        """Create a new window with subplots comparing samples by date."""
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        import matplotlib.dates as mdates
        from datetime import datetime
    
        if not self.comparison_results:
            messagebox.showinfo("Info", "No comparison data available. Please run analysis first.")
            return
    
        # Create new window
        plot_window = tk.Toplevel(self.window)
        plot_window.title("Sample Comparison Plots")
        plot_window.geometry("1200x800")
        plot_window.transient(self.window)
    
        # Prepare data for plotting
        plot_data = {}
    
        # Extract data for each sample with timestamps
        for (keyword, test_group), data in self.comparison_results.items():
            if keyword not in plot_data:
                plot_data[keyword] = {
                    'dates': [],
                    'tpm_values': [],
                    'std_dev_values': [],
                    'draw_pressure_values': [],
                    'tpm_per_power': [],
                    'labels': []
                }
        
            # Process each file's data
            for file_info in data['files']:
                # file_info is the display name, we need to find the actual file data
                timestamp = None
                actual_filename = None
            
                # Search through all loaded files to find the matching one
                for file_data_item in self.gui.all_filtered_sheets:
                    # Check if this is the file we're looking for
                    if file_info in file_data_item.get("display_filename", "") or file_info in file_data_item["file_name"]:
                    
                        # First priority: Check for database filename (which includes timestamp)
                        if "database_filename" in file_data_item:
                            actual_filename = file_data_item["database_filename"]
                            debug_print(f"DEBUG: Using database filename: {actual_filename}")
                        elif "original_filename" in file_data_item:
                            actual_filename = file_data_item["original_filename"]
                            debug_print(f"DEBUG: Using original filename: {actual_filename}")
                        else:
                            actual_filename = file_data_item["file_name"]
                            debug_print(f"DEBUG: Using file_name: {actual_filename}")
                    
                        # Check for stored timestamp in metadata
                        if "database_created_at" in file_data_item:
                            timestamp = file_data_item["database_created_at"]
                            if isinstance(timestamp, str):
                                try:
                                    timestamp = datetime.fromisoformat(timestamp)
                                except:
                                    timestamp = None
                            debug_print(f"DEBUG: Found database_created_at timestamp: {timestamp}")
                        elif "created_at" in file_data_item:
                            timestamp = file_data_item["created_at"]
                            debug_print(f"DEBUG: Found created_at timestamp: {timestamp}")
                    
                        break
            
                # If we found an actual filename with potential timestamp and no timestamp yet, try to parse it
                if actual_filename and not timestamp:
                    import re
                
                    # Look for timestamp pattern AFTER the extension
                    # Pattern: .vap3 - YYYY-MM-DD HH:MM:SS
                    timestamp_pattern = r'\.vap3\s*-\s*(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})'
                    match = re.search(timestamp_pattern, actual_filename)
                
                    if match:
                        try:
                            date_part = match.group(1)
                            time_part = match.group(2)
                            timestamp_str = f"{date_part} {time_part}"
                            timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                            debug_print(f"DEBUG: Parsed timestamp {timestamp} from {actual_filename}")
                        except Exception as e:
                            debug_print(f"DEBUG: Failed to parse timestamp: {e}")
                
                    # If no standard timestamp found, check for other patterns after extension
                    if not timestamp:
                        # Alternative patterns that might appear after .vap3
                        alt_patterns = [
                            (r'\.vap3.*_(\d{8})_(\d{6})', '%Y%m%d_%H%M%S'),  # .vap3_20240315_143022
                            (r'\.vap3.*_(\d{4}-\d{2}-\d{2})_(\d{2}-\d{2}-\d{2})', '%Y-%m-%d_%H-%M-%S'),  # .vap3_2024-03-15_14-30-22
                            (r'\.vap3\s*-\s*(\d{4}-\d{2}-\d{2})', '%Y-%m-%d'),  # .vap3 - 2024-03-15 (date only)
                        ]
                    
                        for pattern, date_format in alt_patterns:
                            match = re.search(pattern, actual_filename)
                            if match:
                                try:
                                    if len(match.groups()) == 2:
                                        timestamp_str = f"{match.group(1)}_{match.group(2)}"
                                    else:
                                        timestamp_str = match.group(1)
                                    timestamp = datetime.strptime(timestamp_str, date_format)
                                    debug_print(f"DEBUG: Parsed alternative timestamp {timestamp} from {actual_filename}")
                                    break
                                except:
                                    continue
            
                # If no timestamp found, use current time as fallback
                if not timestamp:
                    timestamp = datetime.now()
                    debug_print(f"WARNING: No timestamp found in {actual_filename or file_info}, using current time")
            
                # Get the corresponding sample data
                for file_data_item in self.gui.all_filtered_sheets:
                    if file_info in file_data_item.get("display_filename", "") or file_info in file_data_item["file_name"]:
                        for sheet_name, sheet_info in file_data_item["filtered_sheets"].items():
                            if self.get_test_group(sheet_name) == test_group:
                                sheet_data = sheet_info["data"]
                            
                                # Calculate averages for this file
                                sample_indices = self.find_matching_samples(sheet_data, sheet_name, file_data_item["file_name"]).get(keyword, [])
                            
                                if sample_indices:
                                    # Extract values for matching samples
                                    tpm_vals = []
                                    std_vals = []
                                    dp_vals = []
                                    power_vals = []
                                
                                    for idx in sample_indices:
                                        tpm_col, std_col, dp_col, power_col = self.get_column_indices(sheet_data, sheet_name, idx)
                                    
                                        if tpm_col >= 0:
                                            tpm_val = self.extract_column_average(sheet_data, tpm_col)
                                            if tpm_val is not None:
                                                tpm_vals.append(tpm_val)
                                    
                                        if std_col >= 0:
                                            std_val = self.extract_column_average(sheet_data, std_col)
                                            if std_val is not None:
                                                std_vals.append(std_val)
                                    
                                        if dp_col >= 0:
                                            dp_val = self.extract_column_average(sheet_data, dp_col)
                                            if dp_val is not None:
                                                dp_vals.append(dp_val)
                                    
                                        if power_col >= 0:
                                            power_val = self.extract_column_average(sheet_data, power_col)
                                            if power_val is not None:
                                                power_vals.append(power_val)
                                
                                    # Calculate averages and add to plot data
                                    if tpm_vals:
                                        avg_tpm = np.mean(tpm_vals)
                                        avg_std = np.mean(std_vals) if std_vals else 0
                                        avg_dp = np.mean(dp_vals) if dp_vals else 0
                                        avg_power = np.mean(power_vals) if power_vals else 1  # Avoid division by zero
                                    
                                        plot_data[keyword]['dates'].append(timestamp)
                                        plot_data[keyword]['tpm_values'].append(avg_tpm)
                                        plot_data[keyword]['std_dev_values'].append(avg_std)
                                        plot_data[keyword]['draw_pressure_values'].append(avg_dp)
                                        plot_data[keyword]['tpm_per_power'].append(avg_tpm / avg_power if avg_power > 0 else 0)
                                        plot_data[keyword]['labels'].append(os.path.basename(file_info))
                            
                                break
                        break
    
        # Check if we have any data to plot
        has_data = False
        for model_data in plot_data.values():
            if model_data['dates']:
                has_data = True
                break
    
        if not has_data:
            messagebox.showwarning("No Data", "No data available for plotting. Make sure samples have been analyzed.")
            plot_window.destroy()
            return
    
        # Create matplotlib figure with 3 subplots
        fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 10), sharex=True)
        fig.suptitle('Sample Comparison Over Time', fontsize=16)
    
        # Color map for different models
        colors = plt.cm.tab10(np.linspace(0, 1, len(plot_data)))
    
        # Plot for each model
        for idx, (model, data) in enumerate(plot_data.items()):
            if not data['dates']:
                continue
        
            # Sort by date
            sorted_indices = np.argsort(data['dates'])
            dates = [data['dates'][i] for i in sorted_indices]
            tpm_values = [data['tpm_values'][i] for i in sorted_indices]
            std_values = [data['std_dev_values'][i] for i in sorted_indices]
            dp_values = [data['draw_pressure_values'][i] for i in sorted_indices]
            tpm_power_values = [data['tpm_per_power'][i] for i in sorted_indices]
        
            color = colors[idx]
        
            # Plot 1: Average TPM with error bars
            if tpm_values:  # Only plot if we have data
                ax1.errorbar(dates, tpm_values, yerr=std_values, 
                            label=model.upper(), color=color, marker='o', 
                            capsize=5, capthick=2, markersize=8)
        
            # Plot 2: Average Draw Pressure
            if dp_values and any(v > 0 for v in dp_values):  # Only plot if we have non-zero data
                ax2.plot(dates, dp_values, label=model.upper(), 
                        color=color, marker='s', markersize=8, linewidth=2)
        
            # Plot 3: Average TPM/Power
            if tpm_power_values and any(v > 0 for v in tpm_power_values):  # Only plot if we have data
                ax3.plot(dates, tpm_power_values, label=model.upper(), 
                        color=color, marker='^', markersize=8, linewidth=2)
    
        # Configure axes
        ax1.set_ylabel('Average TPM', fontsize=12)
        ax1.grid(True, alpha=0.3)
        if ax1.get_lines():  # Only add legend if there are lines
            ax1.legend(loc='upper left')
    
        ax2.set_ylabel('Average Draw Pressure', fontsize=12)
        ax2.grid(True, alpha=0.3)
    
        ax3.set_ylabel('Average TPM/Power', fontsize=12)
        ax3.set_xlabel('Date', fontsize=12)
        ax3.grid(True, alpha=0.3)
    
        # Format x-axis dates
        for ax in [ax1, ax2, ax3]:
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            ax.xaxis.set_major_locator(mdates.AutoDateLocator())
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
    
        # Adjust layout
        plt.tight_layout()
    
        # Create tkinter canvas
        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
        # Add toolbar
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, plot_window)
        toolbar.update()
    
        # Add export button
        def export_plot():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if file_path:
                fig.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Plot saved to {file_path}")
    
        button_frame = ttk.Frame(plot_window)
        button_frame.pack(fill="x", padx=10, pady=5)
    
        ttk.Button(button_frame, text="Export Plot", command=export_plot).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=plot_window.destroy).pack(side="right", padx=5)
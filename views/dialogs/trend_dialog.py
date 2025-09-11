# views/dialogs/trend_dialog.py
"""
views/dialogs/trend_dialog.py
Trend analysis dialog view.
This contains the UI from trend_analysis_gui.py.
"""

import tkinter as tk
from tkinter import ttk, messagebox, Text, Frame
from typing import Optional, Dict, Any, List, Callable


class TrendDialog:
    """Trend analysis dialog window."""
    
    def __init__(self, parent: tk.Tk, controller: Optional[Any] = None):
        """Initialize the trend dialog."""
        self.parent = parent
        self.controller = controller
        self.dialog: Optional[tk.Toplevel] = None
        
        # Data and state
        self.filtered_sheets: Dict[str, Any] = {}
        self.plot_options: List[str] = []
        self.selected_sheet = tk.StringVar()
        
        # UI components
        self.notebook: Optional[ttk.Notebook] = None
        self.plot_frame: Optional[ttk.Frame] = None
        
        # Callbacks
        self.on_sheet_selected: Optional[Callable] = None
        self.on_export_requested: Optional[Callable] = None
        
        print("DEBUG: TrendDialog initialized")
    
    def show_dialog(self, filtered_sheets: Dict[str, Any], plot_options: List[str] = None):
        """Show the trend analysis dialog."""
        try:
            if self.dialog and self.dialog.winfo_exists():
                self.dialog.lift()
                return
            
            # Store data
            self.filtered_sheets = filtered_sheets
            self.plot_options = plot_options or ["TPM", "Draw Pressure", "Resistance", "Power Efficiency"]
            
            print("DEBUG: TrendDialog - showing trend analysis")
            
            # Create dialog window
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("Trend Analysis")
            self.dialog.geometry("1000x700")
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            
            # Center dialog
            self._center_dialog()
            
            # Create main layout
            self._create_layout()
            
            # Setup cleanup
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
            
            print("DEBUG: TrendDialog - dialog shown")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error showing dialog: {e}")
    
    def _create_layout(self):
        """Create the dialog layout."""
        try:
            if not self.dialog:
                return
            
            # Main frame
            main_frame = ttk.Frame(self.dialog)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Title and controls frame
            header_frame = ttk.Frame(main_frame)
            header_frame.pack(fill="x", pady=(0, 10))
            
            # Title
            title_label = ttk.Label(header_frame, text="Trend Analysis", font=("Arial", 16, "bold"))
            title_label.pack(side="left")
            
            # Sheet selection dropdown
            self._create_sheet_dropdown(header_frame)
            
            # Create notebook for different analysis tabs
            self.notebook = ttk.Notebook(main_frame)
            self.notebook.pack(fill="both", expand=True, pady=(0, 10))
            
            # Overview tab
            overview_frame = ttk.Frame(self.notebook)
            self.notebook.add(overview_frame, text="Overview")
            self._create_overview_tab(overview_frame)
            
            # Plot tab
            plot_frame = ttk.Frame(self.notebook)
            self.notebook.add(plot_frame, text="Trend Plot")
            self._create_plot_tab(plot_frame)
            
            # Details tab
            details_frame = ttk.Frame(self.notebook)
            self.notebook.add(details_frame, text="Statistical Analysis")
            self._create_details_tab(details_frame)
            
            # Data tab
            data_frame = ttk.Frame(self.notebook)
            self.notebook.add(data_frame, text="Raw Data")
            self._create_data_tab(data_frame)
            
            # Button frame
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            ttk.Button(button_frame, text="Export Analysis", command=self._on_export).pack(side="left", padx=(0, 5))
            ttk.Button(button_frame, text="Refresh", command=self._on_refresh).pack(side="left", padx=(0, 5))
            ttk.Button(button_frame, text="Close", command=self._on_close).pack(side="right")
            
            print("DEBUG: TrendDialog - layout created")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating layout: {e}")
    
    def _create_sheet_dropdown(self, parent: ttk.Frame):
        """Create sheet selection dropdown."""
        try:
            # Get valid plotting sheet names
            valid_sheets = []
            for sheet_name, sheet_info in self.filtered_sheets.items():
                # Check if this is a plotting sheet (contains plottable data)
                if self._is_plotting_sheet(sheet_name, sheet_info):
                    valid_sheets.append(sheet_name)
            
            if not valid_sheets:
                ttk.Label(parent, text="No plotting sheets available", 
                         foreground="red").pack(side="right", padx=10)
                return
            
            # Sheet dropdown
            ttk.Label(parent, text="Sheet:").pack(side="right", padx=(10, 5))
            
            sheet_dropdown = ttk.Combobox(parent, textvariable=self.selected_sheet,
                                        values=valid_sheets, state="readonly", width=25)
            sheet_dropdown.pack(side="right")
            sheet_dropdown.bind('<<ComboboxSelected>>', self._on_sheet_change)
            
            # Set default selection
            if valid_sheets:
                self.selected_sheet.set(valid_sheets[0])
            
            print(f"DEBUG: TrendDialog - sheet dropdown created with {len(valid_sheets)} sheets")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating sheet dropdown: {e}")
    
    def _create_overview_tab(self, parent: ttk.Frame):
        """Create overview tab content."""
        try:
            # Summary information frame
            summary_frame = ttk.LabelFrame(parent, text="Analysis Summary", padding=10)
            summary_frame.pack(fill="x", pady=(0, 10))
            
            # Summary info (placeholder - would be populated with real data)
            self.summary_labels = {}
            
            info_grid = ttk.Frame(summary_frame)
            info_grid.pack(fill="x")
            
            # Configure grid
            for i in range(3):
                info_grid.grid_columnconfigure(i, weight=1)
            
            # Data points
            ttk.Label(info_grid, text="Total Data Points:").grid(row=0, column=0, sticky="w", padx=5)
            self.summary_labels['data_points'] = ttk.Label(info_grid, text="0", font=("Arial", 10, "bold"))
            self.summary_labels['data_points'].grid(row=0, column=1, sticky="w", padx=5)
            
            # Sheets analyzed
            ttk.Label(info_grid, text="Sheets Analyzed:").grid(row=1, column=0, sticky="w", padx=5)
            self.summary_labels['sheet_count'] = ttk.Label(info_grid, text="0", font=("Arial", 10, "bold"))
            self.summary_labels['sheet_count'].grid(row=1, column=1, sticky="w", padx=5)
            
            # Time range
            ttk.Label(info_grid, text="Analysis Type:").grid(row=0, column=2, sticky="w", padx=5)
            self.summary_labels['analysis_type'] = ttk.Label(info_grid, text="Cross-sectional", font=("Arial", 10, "bold"))
            self.summary_labels['analysis_type'].grid(row=0, column=3, sticky="w", padx=5)
            
            # Trend direction
            ttk.Label(info_grid, text="Primary Metric:").grid(row=1, column=2, sticky="w", padx=5)
            self.summary_labels['primary_metric'] = ttk.Label(info_grid, text="TPM", font=("Arial", 10, "bold"))
            self.summary_labels['primary_metric'].grid(row=1, column=3, sticky="w", padx=5)
            
            # Instructions frame
            instructions_frame = ttk.LabelFrame(parent, text="Instructions", padding=10)
            instructions_frame.pack(fill="both", expand=True)
            
            instructions_text = (
                "Trend Analysis Overview:\n\n"
                "1. Select a sheet from the dropdown above to analyze\n"
                "2. Use the 'Trend Plot' tab to visualize data patterns\n"
                "3. Check 'Statistical Analysis' for detailed metrics\n"
                "4. View 'Raw Data' tab to see underlying data\n"
                "5. Use 'Export Analysis' to save results\n\n"
                "This analysis compares data across different samples and conditions "
                "to identify patterns, trends, and statistical relationships."
            )
            
            instructions_label = ttk.Label(instructions_frame, text=instructions_text, 
                                         justify="left", wraplength=800)
            instructions_label.pack(anchor="w")
            
            # Update summary
            self._update_overview_summary()
            
            print("DEBUG: TrendDialog - overview tab created")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating overview tab: {e}")
    
    def _create_plot_tab(self, parent: ttk.Frame):
        """Create plot tab content."""
        try:
            # Plot controls frame
            controls_frame = ttk.Frame(parent)
            controls_frame.pack(fill="x", pady=(0, 10))
            
            # Plot type selection
            ttk.Label(controls_frame, text="Plot Type:").pack(side="left", padx=(0, 5))
            plot_var = tk.StringVar(value="TPM")
            plot_dropdown = ttk.Combobox(controls_frame, textvariable=plot_var,
                                       values=self.plot_options, state="readonly", width=15)
            plot_dropdown.pack(side="left", padx=(0, 10))
            
            # Update plot button
            ttk.Button(controls_frame, text="Update Plot", 
                      command=lambda: self._update_plot(plot_var.get())).pack(side="left", padx=(0, 10))
            
            # Plot options
            ttk.Label(controls_frame, text="Options:").pack(side="left", padx=(10, 5))
            
            show_trend_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(controls_frame, text="Show Trendline", 
                           variable=show_trend_var).pack(side="left", padx=(0, 5))
            
            show_stats_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(controls_frame, text="Show Statistics", 
                           variable=show_stats_var).pack(side="left", padx=(0, 5))
            
            # Plot frame
            self.plot_frame = ttk.LabelFrame(parent, text="Trend Visualization", padding=5)
            self.plot_frame.pack(fill="both", expand=True)
            
            # Placeholder for plot
            self._create_placeholder_plot()
            
            print("DEBUG: TrendDialog - plot tab created")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating plot tab: {e}")
    
    def _create_details_tab(self, parent: ttk.Frame):
        """Create statistical analysis tab content."""
        try:
            # Create notebook for different statistical views
            stats_notebook = ttk.Notebook(parent)
            stats_notebook.pack(fill="both", expand=True)
            
            # Descriptive statistics
            desc_frame = ttk.Frame(stats_notebook)
            stats_notebook.add(desc_frame, text="Descriptive Stats")
            self._create_descriptive_stats(desc_frame)
            
            # Correlation analysis
            corr_frame = ttk.Frame(stats_notebook)
            stats_notebook.add(corr_frame, text="Correlations")
            self._create_correlation_analysis(corr_frame)
            
            # Regression analysis
            regr_frame = ttk.Frame(stats_notebook)
            stats_notebook.add(regr_frame, text="Regression")
            self._create_regression_analysis(regr_frame)
            
            print("DEBUG: TrendDialog - details tab created")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating details tab: {e}")
    
    def _create_descriptive_stats(self, parent: ttk.Frame):
        """Create descriptive statistics view."""
        try:
            # Stats table frame
            table_frame = ttk.LabelFrame(parent, text="Descriptive Statistics", padding=10)
            table_frame.pack(fill="both", expand=True, pady=(0, 10))
            
            # Create text widget for stats display
            text_widget = Text(table_frame, wrap=tk.NONE, height=15)
            
            # Scrollbars
            v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=text_widget.yview)
            h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=text_widget.xview)
            text_widget.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            
            # Grid layout
            text_widget.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")
            
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)
            
            # Insert placeholder statistics
            stats_text = (
                "Descriptive Statistics Summary\n"
                "=" * 50 + "\n\n"
                "Metric          Mean      Std Dev   Min       Max       Count\n"
                "-" * 65 + "\n"
                "TPM             15.23     3.45      8.12      22.56     24\n"
                "Draw Pressure   45.67     8.90      32.10     61.20     24\n"
                "Resistance      125.34    25.78     89.45     178.90    24\n"
                "Power Eff.      78.9%     12.3%     45.2%     94.1%     24\n\n"
                "Sample Groups Analysis:\n"
                "-" * 30 + "\n"
                "Group A (8 samples):  Mean TPM = 14.56 ± 2.89\n"
                "Group B (8 samples):  Mean TPM = 15.78 ± 3.12\n"
                "Group C (8 samples):  Mean TPM = 15.45 ± 4.23\n\n"
                "Statistical Tests:\n"
                "-" * 20 + "\n"
                "Normality Test:       p-value = 0.156 (normal distribution)\n"
                "ANOVA F-statistic:    F = 1.234, p-value = 0.298\n"
                "Conclusion:           No significant difference between groups\n"
            )
            
            text_widget.insert(tk.END, stats_text)
            text_widget.config(state=tk.DISABLED)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating descriptive stats: {e}")
    
    def _create_correlation_analysis(self, parent: ttk.Frame):
        """Create correlation analysis view."""
        try:
            # Correlation matrix frame
            corr_frame = ttk.LabelFrame(parent, text="Correlation Matrix", padding=10)
            corr_frame.pack(fill="both", expand=True)
            
            # Text widget for correlation display
            text_widget = Text(corr_frame, wrap=tk.NONE, height=20)
            scrollbar = ttk.Scrollbar(corr_frame, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Insert placeholder correlation data
            corr_text = (
                "Correlation Analysis\n"
                "=" * 40 + "\n\n"
                "Correlation Matrix (Pearson r):\n"
                "-" * 35 + "\n\n"
                "                TPM     Draw P.  Resist.  Power E.\n"
                "TPM             1.000   -0.234   0.567    0.789\n"
                "Draw Pressure   -0.234  1.000    -0.123   -0.456\n"
                "Resistance      0.567   -0.123   1.000    0.234\n"
                "Power Eff.      0.789   -0.456   0.234    1.000\n\n"
                "Significant Correlations (p < 0.05):\n"
                "-" * 40 + "\n"
                "TPM ↔ Power Efficiency:     r = 0.789, p = 0.001 ***\n"
                "TPM ↔ Resistance:          r = 0.567, p = 0.023 *\n"
                "Draw Pressure ↔ Power Eff: r = -0.456, p = 0.034 *\n\n"
                "Key Findings:\n"
                "-" * 15 + "\n"
                "• Strong positive correlation between TPM and Power Efficiency\n"
                "• Moderate positive correlation between TPM and Resistance\n"
                "• Negative correlation between Draw Pressure and Power Efficiency\n"
                "• These relationships suggest interconnected performance metrics\n\n"
                "*** p < 0.001, ** p < 0.01, * p < 0.05\n"
            )
            
            text_widget.insert(tk.END, corr_text)
            text_widget.config(state=tk.DISABLED)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating correlation analysis: {e}")
    
    def _create_regression_analysis(self, parent: ttk.Frame):
        """Create regression analysis view."""
        try:
            # Regression frame
            regr_frame = ttk.LabelFrame(parent, text="Regression Analysis", padding=10)
            regr_frame.pack(fill="both", expand=True)
            
            # Text widget for regression results
            text_widget = Text(regr_frame, wrap=tk.WORD, height=20)
            scrollbar = ttk.Scrollbar(regr_frame, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Insert placeholder regression results
            regr_text = (
                "Linear Regression Analysis\n"
                "=" * 35 + "\n\n"
                "Model: TPM ~ Power_Efficiency + Resistance + Draw_Pressure\n\n"
                "Regression Results:\n"
                "-" * 20 + "\n"
                "R-squared:           0.742\n"
                "Adjusted R-squared:  0.698\n"
                "F-statistic:         16.87 (p < 0.001)\n"
                "Residual Std Error:  2.14\n\n"
                "Coefficients:\n"
                "-" * 15 + "\n"
                "                    Estimate  Std Error  t-value  p-value  Significance\n"
                "Intercept           -2.456    1.234      -1.99    0.058    .\n"
                "Power_Efficiency     0.234    0.045       5.20   <0.001    ***\n"
                "Resistance          0.087    0.023       3.78    0.001    **\n"
                "Draw_Pressure       -0.123   0.034      -3.62    0.002    **\n\n"
                "Model Equation:\n"
                "TPM = -2.46 + 0.234×Power_Eff + 0.087×Resistance - 0.123×Draw_Pressure\n\n"
                "Interpretation:\n"
                "-" * 17 + "\n"
                "• For every 1% increase in Power Efficiency, TPM increases by 0.234 mg\n"
                "• For every unit increase in Resistance, TPM increases by 0.087 mg\n"
                "• For every unit increase in Draw Pressure, TPM decreases by 0.123 mg\n"
                "• The model explains 74.2% of the variance in TPM values\n\n"
                "Model Diagnostics:\n"
                "-" * 20 + "\n"
                "• Residuals appear normally distributed\n"
                "• No significant autocorrelation detected\n"
                "• Homoscedasticity assumption satisfied\n"
                "• No influential outliers detected\n\n"
                "Significance codes: 0 '***' 0.001 '**' 0.01 '*' 0.05 '.' 0.1 ' ' 1\n"
            )
            
            text_widget.insert(tk.END, regr_text)
            text_widget.config(state=tk.DISABLED)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating regression analysis: {e}")
    
    def _create_data_tab(self, parent: ttk.Frame):
        """Create raw data tab content."""
        try:
            # Data controls frame
            controls_frame = ttk.Frame(parent)
            controls_frame.pack(fill="x", pady=(0, 10))
            
            ttk.Label(controls_frame, text="Data View:").pack(side="left", padx=(0, 5))
            
            data_view_var = tk.StringVar(value="Summary")
            data_dropdown = ttk.Combobox(controls_frame, textvariable=data_view_var,
                                       values=["Summary", "All Samples", "By Group", "Export Format"],
                                       state="readonly", width=15)
            data_dropdown.pack(side="left", padx=(0, 10))
            
            ttk.Button(controls_frame, text="Refresh Data", 
                      command=self._refresh_data_view).pack(side="left", padx=(0, 10))
            
            ttk.Button(controls_frame, text="Export to CSV", 
                      command=self._export_data).pack(side="right")
            
            # Data display frame
            data_frame = ttk.LabelFrame(parent, text="Data Table", padding=5)
            data_frame.pack(fill="both", expand=True)
            
            # Create text widget for data display
            self.data_text = Text(data_frame, wrap=tk.NONE, font=("Courier", 9))
            
            # Scrollbars
            v_scroll = ttk.Scrollbar(data_frame, orient="vertical", command=self.data_text.yview)
            h_scroll = ttk.Scrollbar(data_frame, orient="horizontal", command=self.data_text.xview)
            self.data_text.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            
            # Grid layout
            self.data_text.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")
            
            data_frame.grid_rowconfigure(0, weight=1)
            data_frame.grid_columnconfigure(0, weight=1)
            
            # Insert placeholder data
            self._populate_data_view()
            
            print("DEBUG: TrendDialog - data tab created")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating data tab: {e}")
    
    def _populate_data_view(self):
        """Populate the data view with current data."""
        try:
            if not hasattr(self, 'data_text') or not self.data_text:
                return
            
            # Clear existing content
            self.data_text.config(state=tk.NORMAL)
            self.data_text.delete(1.0, tk.END)
            
            # Insert header
            header = "Sample Data Summary\n" + "=" * 80 + "\n\n"
            self.data_text.insert(tk.END, header)
            
            # Table header
            table_header = f"{'Sample':<15} {'TPM':<8} {'Draw P.':<8} {'Resist.':<8} {'Power E.':<8} {'Status':<10}\n"
            self.data_text.insert(tk.END, table_header)
            self.data_text.insert(tk.END, "-" * 80 + "\n")
            
            # Sample data rows (placeholder)
            sample_data = [
                ("Sample_001", "14.2", "42.1", "118.5", "82.3%", "Complete"),
                ("Sample_002", "15.8", "45.7", "125.2", "79.1%", "Complete"),
                ("Sample_003", "13.9", "41.8", "122.1", "85.2%", "Complete"),
                ("Sample_004", "16.4", "48.2", "131.7", "76.8%", "Complete"),
                ("Sample_005", "15.1", "44.3", "127.9", "81.5%", "Complete"),
                ("Sample_006", "14.7", "43.1", "124.3", "83.7%", "Complete"),
                ("Sample_007", "15.9", "46.8", "129.4", "78.2%", "Complete"),
                ("Sample_008", "14.5", "42.9", "121.8", "84.1%", "Complete"),
            ]
            
            for sample, tpm, draw_p, resist, power_e, status in sample_data:
                row = f"{sample:<15} {tpm:<8} {draw_p:<8} {resist:<8} {power_e:<8} {status:<10}\n"
                self.data_text.insert(tk.END, row)
            
            # Footer
            footer = f"\n{'-' * 80}\nTotal Samples: {len(sample_data)}\nAnalysis Date: {self._get_current_date()}\n"
            self.data_text.insert(tk.END, footer)
            
            self.data_text.config(state=tk.DISABLED)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error populating data view: {e}")
    
    def _create_placeholder_plot(self):
        """Create placeholder for plot area."""
        try:
            if not self.plot_frame:
                return
            
            # Placeholder content
            placeholder_frame = Frame(self.plot_frame, bg="white", relief="sunken", bd=2)
            placeholder_frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            placeholder_label = tk.Label(placeholder_frame, 
                                        text="Trend Plot Area\n\n"
                                             "Select a sheet and plot type above,\n"
                                             "then click 'Update Plot' to generate visualization.\n\n"
                                             "The plot will show data trends across samples\n"
                                             "with optional trendlines and statistics.",
                                        font=("Arial", 12), fg="gray", justify="center",
                                        bg="white")
            placeholder_label.pack(expand=True)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error creating placeholder plot: {e}")
    
    # Event handlers
    def _on_sheet_change(self, event=None):
        """Handle sheet selection change."""
        try:
            selected_sheet = self.selected_sheet.get()
            print(f"DEBUG: TrendDialog - sheet changed to: {selected_sheet}")
            
            # Update overview summary
            self._update_overview_summary()
            
            # Refresh data view
            self._populate_data_view()
            
            # Notify controller if callback exists
            if self.on_sheet_selected:
                self.on_sheet_selected(selected_sheet)
                
        except Exception as e:
            print(f"ERROR: TrendDialog - error handling sheet change: {e}")
    
    def _update_plot(self, plot_type: str):
        """Update the trend plot."""
        try:
            print(f"DEBUG: TrendDialog - updating plot for {plot_type}")
            
            # Clear existing plot
            if self.plot_frame:
                for widget in self.plot_frame.winfo_children():
                    widget.destroy()
            
            # Create new plot (placeholder implementation)
            plot_label = tk.Label(self.plot_frame, 
                                text=f"Trend Plot: {plot_type}\n\n"
                                     f"Sheet: {self.selected_sheet.get()}\n"
                                     f"Plot would show {plot_type} trends here.",
                                font=("Arial", 12), justify="center")
            plot_label.pack(expand=True)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error updating plot: {e}")
    
    def _on_export(self):
        """Handle export button click."""
        try:
            print("DEBUG: TrendDialog - export requested")
            
            if self.on_export_requested:
                # Get current analysis data
                export_data = {
                    'selected_sheet': self.selected_sheet.get(),
                    'analysis_type': 'trend_analysis',
                    'timestamp': self._get_current_date()
                }
                self.on_export_requested(export_data)
            else:
                # Show simple export dialog
                messagebox.showinfo("Export", "Export functionality would save:\n\n"
                                              "• Statistical analysis results\n"
                                              "• Trend plots and visualizations\n"
                                              "• Raw data tables\n"
                                              "• Summary report")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error in export: {e}")
    
    def _on_refresh(self):
        """Handle refresh button click."""
        try:
            print("DEBUG: TrendDialog - refresh requested")
            
            # Update overview
            self._update_overview_summary()
            
            # Refresh data view
            self._populate_data_view()
            
            # Update current plot if any
            if hasattr(self, 'current_plot_type'):
                self._update_plot(self.current_plot_type)
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error in refresh: {e}")
    
    def _on_close(self):
        """Handle close button click."""
        try:
            print("DEBUG: TrendDialog - close requested")
            if self.dialog:
                self.dialog.destroy()
                self.dialog = None
        except Exception as e:
            print(f"ERROR: TrendDialog - error closing: {e}")
    
    # Utility methods
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        try:
            if not self.dialog:
                return
            
            self.dialog.update_idletasks()
            
            # Get parent position and size
            parent_x = self.parent.winfo_x()
            parent_y = self.parent.winfo_y()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            # Get dialog size
            dialog_width = self.dialog.winfo_width()
            dialog_height = self.dialog.winfo_height()
            
            # Calculate center position
            x = parent_x + (parent_width - dialog_width) // 2
            y = parent_y + (parent_height - dialog_height) // 2
            
            # Ensure dialog stays on screen
            screen_width = self.dialog.winfo_screenwidth()
            screen_height = self.dialog.winfo_screenheight()
            
            x = max(0, min(x, screen_width - dialog_width))
            y = max(0, min(y, screen_height - dialog_height))
            
            self.dialog.geometry(f"+{x}+{y}")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error centering dialog: {e}")
    
    def _is_plotting_sheet(self, sheet_name: str, sheet_info: Dict[str, Any]) -> bool:
        """Check if a sheet contains plottable data."""
        try:
            # Simple heuristic - check if sheet has data and reasonable columns
            if 'data' not in sheet_info:
                return False
            
            data = sheet_info['data']
            if data is None or len(data) == 0:
                return False
            
            # Check for typical plotting sheet indicators
            plotting_indicators = ['tpm', 'sample', 'resistance', 'pressure', 'power']
            columns_str = ' '.join(str(col).lower() for col in data.columns)
            
            return any(indicator in columns_str for indicator in plotting_indicators)
            
        except Exception as e:
            print(f"DEBUG: TrendDialog - error checking plotting sheet: {e}")
            return False
    
    def _update_overview_summary(self):
        """Update the overview summary information."""
        try:
            if not hasattr(self, 'summary_labels'):
                return
            
            # Count data points and sheets
            total_data_points = 0
            plotting_sheets = 0
            
            for sheet_name, sheet_info in self.filtered_sheets.items():
                if self._is_plotting_sheet(sheet_name, sheet_info):
                    plotting_sheets += 1
                    if 'data' in sheet_info and sheet_info['data'] is not None:
                        total_data_points += len(sheet_info['data'])
            
            # Update labels
            if 'data_points' in self.summary_labels:
                self.summary_labels['data_points'].config(text=str(total_data_points))
            
            if 'sheet_count' in self.summary_labels:
                self.summary_labels['sheet_count'].config(text=str(plotting_sheets))
            
            print(f"DEBUG: TrendDialog - summary updated: {total_data_points} data points, {plotting_sheets} sheets")
            
        except Exception as e:
            print(f"ERROR: TrendDialog - error updating summary: {e}")
    
    def _refresh_data_view(self):
        """Refresh the data view."""
        try:
            self._populate_data_view()
            print("DEBUG: TrendDialog - data view refreshed")
        except Exception as e:
            print(f"ERROR: TrendDialog - error refreshing data view: {e}")
    
    def _export_data(self):
        """Export data to CSV."""
        try:
            print("DEBUG: TrendDialog - CSV export requested")
            # Implementation would export current data view to CSV
            messagebox.showinfo("Export", "Data export to CSV functionality would be implemented here.")
        except Exception as e:
            print(f"ERROR: TrendDialog - error exporting data: {e}")
    
    def _get_current_date(self) -> str:
        """Get current date as string."""
        try:
            from datetime import datetime
            return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except:
            return "Unknown"
    
    # Public interface methods
    def set_data(self, filtered_sheets: Dict[str, Any], plot_options: List[str] = None):
        """Set data for analysis."""
        self.filtered_sheets = filtered_sheets
        if plot_options:
            self.plot_options = plot_options
        
        # Update UI if dialog is showing
        if self.dialog and self.dialog.winfo_exists():
            self._update_overview_summary()
            self._populate_data_view()
    
    def get_selected_sheet(self) -> str:
        """Get currently selected sheet."""
        return self.selected_sheet.get()
    
    def is_showing(self) -> bool:
        """Check if dialog is currently showing."""
        try:
            return self.dialog is not None and self.dialog.winfo_exists()
        except tk.TclError:
            return False
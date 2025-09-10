# views/calculators/viscosity_view.py
"""
views/calculators/viscosity_view.py
Viscosity calculator view.
This will contain the UI from viscosity_calculator.py.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, Any, Callable
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class ViscosityView:
    """View for viscosity calculator interface."""
    
    def __init__(self, parent: tk.Widget):
        """Initialize the viscosity view."""
        self.parent = parent
        self.frame = ttk.Frame(parent)
        
        # UI Components
        self.notebook: Optional[ttk.Notebook] = None
        self.data_tab: Optional[ttk.Frame] = None
        self.training_tab: Optional[ttk.Frame] = None
        self.prediction_tab: Optional[ttk.Frame] = None
        self.analysis_tab: Optional[ttk.Frame] = None
        
        # Matplotlib components
        self.figure: Optional[Figure] = None
        self.canvas: Optional[FigureCanvasTkAgg] = None
        self.axes: Optional[Any] = None
        
        # Input variables
        self.temperature_var = tk.DoubleVar(value=25.0)
        self.viscosity_var = tk.DoubleVar(value=1.0)
        self.sample_id_var = tk.StringVar(value="Sample_1")
        self.model_name_var = tk.StringVar(value="Model_1")
        self.prediction_temp_var = tk.DoubleVar(value=25.0)
        
        # Callbacks
        self.on_data_added: Optional[Callable] = None
        self.on_model_trained: Optional[Callable] = None
        self.on_prediction_requested: Optional[Callable] = None
        
        print("DEBUG: ViscosityView initialized")
    
    def setup_view(self):
        """Set up the viscosity calculator view."""
        print("DEBUG: ViscosityView setting up layout")
        
        # Create notebook for tabs
        self._create_notebook()
        
        # Create individual tabs
        self._create_data_tab()
        self._create_training_tab()
        self._create_prediction_tab()
        self._create_analysis_tab()
        
        print("DEBUG: ViscosityView layout complete")
    
    def _create_notebook(self):
        """Create the main notebook widget."""
        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        print("DEBUG: ViscosityView notebook created")
    
    def _create_data_tab(self):
        """Create the data input tab."""
        self.data_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.data_tab, text="Data Input")
        
        # Data input frame
        input_frame = ttk.LabelFrame(self.data_tab, text="Add Viscosity Data", padding=10)
        input_frame.pack(fill="x", padx=10, pady=10)
        
        # Temperature input
        ttk.Label(input_frame, text="Temperature (°C):").grid(row=0, column=0, sticky="w", padx=(0, 5))
        temp_entry = ttk.Entry(input_frame, textvariable=self.temperature_var, width=15)
        temp_entry.grid(row=0, column=1, padx=(0, 10))
        
        # Viscosity input
        ttk.Label(input_frame, text="Viscosity (Pa·s):").grid(row=0, column=2, sticky="w", padx=(0, 5))
        visc_entry = ttk.Entry(input_frame, textvariable=self.viscosity_var, width=15)
        visc_entry.grid(row=0, column=3, padx=(0, 10))
        
        # Sample ID input
        ttk.Label(input_frame, text="Sample ID:").grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(5, 0))
        sample_entry = ttk.Entry(input_frame, textvariable=self.sample_id_var, width=15)
        sample_entry.grid(row=1, column=1, padx=(0, 10), pady=(5, 0))
        
        # Add data button
        add_button = ttk.Button(input_frame, text="Add Data", command=self._on_add_data_clicked)
        add_button.grid(row=1, column=2, columnspan=2, pady=(5, 0))
        
        # Data display frame
        display_frame = ttk.LabelFrame(self.data_tab, text="Current Data", padding=10)
        display_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create data listbox
        self.data_listbox = tk.Listbox(display_frame, height=10)
        data_scrollbar = ttk.Scrollbar(display_frame, orient="vertical", command=self.data_listbox.yview)
        self.data_listbox.configure(yscrollcommand=data_scrollbar.set)
        
        self.data_listbox.pack(side="left", fill="both", expand=True)
        data_scrollbar.pack(side="right", fill="y")
        
        print("DEBUG: ViscosityView data tab created")
    
    def _create_training_tab(self):
        """Create the model training tab."""
        self.training_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.training_tab, text="Model Training")
        
        # Training controls frame
        controls_frame = ttk.LabelFrame(self.training_tab, text="Training Controls", padding=10)
        controls_frame.pack(fill="x", padx=10, pady=10)
        
        # Model name input
        ttk.Label(controls_frame, text="Model Name:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        model_entry = ttk.Entry(controls_frame, textvariable=self.model_name_var, width=20)
        model_entry.grid(row=0, column=1, padx=(0, 10))
        
        # Model type selection
        ttk.Label(controls_frame, text="Model Type:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.model_type_var = tk.StringVar(value="polynomial")
        model_type_combo = ttk.Combobox(
            controls_frame,
            textvariable=self.model_type_var,
            values=["polynomial", "arrhenius", "linear"],
            state="readonly",
            width=15
        )
        model_type_combo.grid(row=0, column=3, padx=(0, 10))
        
        # Training buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.grid(row=1, column=0, columnspan=4, pady=(10, 0))
        
        ttk.Button(button_frame, text="Train Model", command=self._on_train_model_clicked).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="Clear Models", command=self._on_clear_models_clicked).pack(side="left")
        
        # Model results frame
        results_frame = ttk.LabelFrame(self.training_tab, text="Training Results", padding=10)
        results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Results text widget
        self.results_text = tk.Text(results_frame, height=15, wrap=tk.WORD)
        results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=results_scrollbar.set)
        
        self.results_text.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")
        
        print("DEBUG: ViscosityView training tab created")
    
    def _create_prediction_tab(self):
        """Create the prediction tab."""
        self.prediction_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.prediction_tab, text="Prediction")
        
        # Prediction controls frame
        pred_controls_frame = ttk.LabelFrame(self.prediction_tab, text="Prediction Controls", padding=10)
        pred_controls_frame.pack(fill="x", padx=10, pady=10)
        
        # Temperature input for prediction
        ttk.Label(pred_controls_frame, text="Temperature (°C):").grid(row=0, column=0, sticky="w", padx=(0, 5))
        pred_temp_entry = ttk.Entry(pred_controls_frame, textvariable=self.prediction_temp_var, width=15)
        pred_temp_entry.grid(row=0, column=1, padx=(0, 10))
        
        # Model selection for prediction
        ttk.Label(pred_controls_frame, text="Model:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.prediction_model_var = tk.StringVar()
        pred_model_combo = ttk.Combobox(
            pred_controls_frame,
            textvariable=self.prediction_model_var,
            state="readonly",
            width=20
        )
        pred_model_combo.grid(row=0, column=3, padx=(0, 10))
        
        # Prediction button
        pred_button = ttk.Button(pred_controls_frame, text="Predict", command=self._on_predict_clicked)
        pred_button.grid(row=0, column=4, padx=(10, 0))
        
        # Prediction results frame
        pred_results_frame = ttk.LabelFrame(self.prediction_tab, text="Prediction Results", padding=10)
        pred_results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Results display
        self.prediction_result_var = tk.StringVar(value="No prediction yet")
        result_label = ttk.Label(pred_results_frame, textvariable=self.prediction_result_var, font=("Arial", 12))
        result_label.pack(pady=20)
        
        print("DEBUG: ViscosityView prediction tab created")
    
    def _create_analysis_tab(self):
        """Create the analysis and visualization tab."""
        self.analysis_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.analysis_tab, text="Analysis")
        
        # Analysis controls
        analysis_controls_frame = ttk.LabelFrame(self.analysis_tab, text="Analysis Controls", padding=10)
        analysis_controls_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        ttk.Button(analysis_controls_frame, text="Plot Data", command=self._on_plot_data_clicked).pack(side="left", padx=(0, 5))
        ttk.Button(analysis_controls_frame, text="Plot Model", command=self._on_plot_model_clicked).pack(side="left", padx=(0, 5))
        ttk.Button(analysis_controls_frame, text="Compare Models", command=self._on_compare_models_clicked).pack(side="left")
        
        # Create matplotlib plot area
        plot_frame = ttk.LabelFrame(self.analysis_tab, text="Visualization", padding=10)
        plot_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        
        # Create figure and canvas
        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.axes = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, plot_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Initialize plot
        self.axes.set_xlabel("Temperature (°C)")
        self.axes.set_ylabel("Viscosity (Pa·s)")
        self.axes.set_title("Viscosity vs Temperature")
        self.axes.grid(True, alpha=0.3)
        
        print("DEBUG: ViscosityView analysis tab created")
    
    # Event handlers
    def _on_add_data_clicked(self):
        """Handle add data button click."""
        print("DEBUG: ViscosityView - add data clicked")
        if self.on_data_added:
            data = {
                'temperature': self.temperature_var.get(),
                'viscosity': self.viscosity_var.get(),
                'sample_id': self.sample_id_var.get()
            }
            self.on_data_added(data)
    
    def _on_train_model_clicked(self):
        """Handle train model button click."""
        print("DEBUG: ViscosityView - train model clicked")
        if self.on_model_trained:
            config = {
                'model_name': self.model_name_var.get(),
                'model_type': self.model_type_var.get()
            }
            self.on_model_trained(config)
    
    def _on_clear_models_clicked(self):
        """Handle clear models button click."""
        print("DEBUG: ViscosityView - clear models clicked")
        self.clear_results()
    
    def _on_predict_clicked(self):
        """Handle predict button click."""
        print("DEBUG: ViscosityView - predict clicked")
        if self.on_prediction_requested:
            config = {
                'temperature': self.prediction_temp_var.get(),
                'model_name': self.prediction_model_var.get()
            }
            self.on_prediction_requested(config)
    
    def _on_plot_data_clicked(self):
        """Handle plot data button click."""
        print("DEBUG: ViscosityView - plot data clicked")
        # Placeholder for plotting raw data
    
    def _on_plot_model_clicked(self):
        """Handle plot model button click."""
        print("DEBUG: ViscosityView - plot model clicked")
        # Placeholder for plotting model predictions
    
    def _on_compare_models_clicked(self):
        """Handle compare models button click."""
        print("DEBUG: ViscosityView - compare models clicked")
        # Placeholder for model comparison
    
    # Public methods for updating the view
    def add_data_to_list(self, data_text: str):
        """Add data to the data listbox."""
        self.data_listbox.insert(tk.END, data_text)
        print(f"DEBUG: ViscosityView added data to list: {data_text}")
    
    def clear_data_list(self):
        """Clear the data listbox."""
        self.data_listbox.delete(0, tk.END)
        print("DEBUG: ViscosityView cleared data list")
    
    def update_model_list(self, models: list):
        """Update the available models for prediction."""
        if hasattr(self, 'prediction_model_var'):
            # Update combobox values
            for widget in self.prediction_tab.winfo_children():
                if isinstance(widget, ttk.LabelFrame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Combobox) and child['textvariable'] == str(self.prediction_model_var):
                            child['values'] = models
                            break
        
        print(f"DEBUG: ViscosityView updated model list with {len(models)} models")
    
    def display_training_results(self, results: str):
        """Display training results in the results text widget."""
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, results)
        print("DEBUG: ViscosityView displayed training results")
    
    def display_prediction_result(self, result: str):
        """Display prediction result."""
        self.prediction_result_var.set(result)
        print(f"DEBUG: ViscosityView displayed prediction: {result}")
    
    def clear_results(self):
        """Clear all results displays."""
        if hasattr(self, 'results_text'):
            self.results_text.delete(1.0, tk.END)
        self.prediction_result_var.set("No prediction yet")
        print("DEBUG: ViscosityView cleared results")
    
    def plot_data(self, temperatures: list, viscosities: list, labels: list = None):
        """Plot viscosity data."""
        if not self.axes:
            return
        
        self.axes.clear()
        self.axes.set_xlabel("Temperature (°C)")
        self.axes.set_ylabel("Viscosity (Pa·s)")
        self.axes.set_title("Viscosity vs Temperature")
        self.axes.grid(True, alpha=0.3)
        
        if labels:
            for i, label in enumerate(labels):
                if i < len(temperatures) and i < len(viscosities):
                    self.axes.scatter(temperatures[i], viscosities[i], label=label, s=50)
        else:
            self.axes.scatter(temperatures, viscosities, s=50)
        
        if labels:
            self.axes.legend()
        
        self.canvas.draw()
        print(f"DEBUG: ViscosityView plotted {len(temperatures)} data points")
    
    def set_callbacks(self, data_callback: Callable, training_callback: Callable, prediction_callback: Callable):
        """Set callback functions."""
        self.on_data_added = data_callback
        self.on_model_trained = training_callback
        self.on_prediction_requested = prediction_callback
        print("DEBUG: ViscosityView callbacks set")
    
    def get_widget(self) -> ttk.Frame:
        """Get the main widget frame."""
        return self.frame
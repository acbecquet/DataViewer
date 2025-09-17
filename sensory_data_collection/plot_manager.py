"""
Plot Manager
Handles all plotting operations including spider plots, sizing, and canvas management
"""

import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
import math
import numpy as np
from utils import debug_print


class PlotManager:
    """Manages sensory profile plotting and visualization."""
    
    def __init__(self, sensory_window):
        """Initialize the plot manager with reference to main window."""
        self.sensory_window = sensory_window
        
    def on_window_resize_plot(self, event):
        """Handle window resize events to update plot size dynamically - ENHANCED DEBUG VERSION."""
        debug_print(f"DEBUG: RESIZE EVENT DETECTED - Widget: {event.widget}, Window: {self.window}")
        debug_print(f"DEBUG: Event widget type: {type(event.widget)}")
        debug_print(f"DEBUG: Window type: {type(self.window)}")
        debug_print(f"DEBUG: Event widget == window? {event.widget == self.window}")
        debug_print(f"DEBUG: Event widget is window? {event.widget is self.window}")

        # Only handle main window resize events, not child widgets
        if event.widget != self.window:
            debug_print(f"DEBUG: Ignoring resize event from child widget: {event.widget}")
            return

        debug_print("DEBUG: MAIN WINDOW RESIZE CONFIRMED - Processing...")

        # Get current window dimensions for verification
        current_width = self.window.winfo_width()
        current_height = self.window.winfo_height()
        debug_print(f"DEBUG: Current window dimensions: {current_width}x{current_height}")

        # Debounce rapid resize events
        if hasattr(self, '_resize_timer'):
            self.window.after_cancel(self._resize_timer)
            debug_print("DEBUG: Cancelled previous resize timer")

        # Schedule plot size update with a small delay to avoid excessive updates
        self._resize_timer = self.window.after(1000, self.update_plot_size_for_resize)
        debug_print("DEBUG: Scheduled plot resize update in 150ms")

    def update_plot_size_for_resize(self):
        """Update plot size with artifact prevention and frame validation."""
        try:
            # Check if we have the necessary components
            if not hasattr(self, 'canvas_frame') or not self.canvas_frame.winfo_exists():
                return

            if not hasattr(self, 'fig') or not self.fig:
                return

            # Wait for frame geometry to stabilize
            self.window.update_idletasks()

            # Use the actual plot container for sizing
            if hasattr(self, 'plot_container') and self.plot_container.winfo_exists():
                parent_for_sizing = self.plot_container
            else:
                parent_for_sizing = self.canvas_frame.master

            parent_for_sizing.update_idletasks()

            # Validate that frames have reasonable dimensions before proceeding
            parent_width = parent_for_sizing.winfo_width()
            parent_height = parent_for_sizing.winfo_height()

            if parent_width < 200 or parent_height < 200:
                debug_print("DEBUG: Parent frame size too small, skipping this resize update")
                return  # Just return, don't schedule another call

            # Calculate new size based on validated frame dimensions
            new_width, new_height = self.calculate_dynamic_plot_size(parent_for_sizing)

            # Get current figure size for comparison
            current_width, current_height = self.fig.get_size_inches()

            # Only update if change is significant
            width_diff = abs(new_width - current_width)
            height_diff = abs(new_height - current_height)
            threshold = 1  # Threshold to reduce excessive updates

            if width_diff > threshold or height_diff > threshold:
                debug_print(f"DEBUG: Significant size change detected - updating plot")

                # Apply the new size
                self.fig.set_size_inches(new_width, new_height)

                self.canvas.draw_idle()
            else:
                debug_print("DEBUG: Size change below threshold, skipping update to prevent artifacts")

        except Exception as e:
            debug_print(f"DEBUG: Error during plot resize: {str(e)}")
            import traceback
            debug_print(f"DEBUG: Full traceback: {traceback.format_exc()}")

    def setup_plot_panel(self):
        """Setup the right panel for spider plot visualization with proper resizing."""

        # Create the main plot frame with proper expansion settings
        plot_frame = ttk.LabelFrame(self.right_frame, text="Sensory Profile Comparison", padding=10)
        plot_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Sample selection for plotting (fixed height)
        control_frame = ttk.Frame(plot_frame)
        control_frame.pack(side='top', fill='x', pady=(0, 5))

        ttk.Label(control_frame, text="Select Samples to Display:", font=FONT).pack(anchor='w')

        # Checkboxes frame (will be populated when samples are added)
        self.checkbox_frame = ttk.Frame(control_frame)
        self.checkbox_frame.pack(fill='x', pady=5)

        self.sample_checkboxes = {}

        # Canvas container frame (expandable)
        canvas_container = ttk.Frame(plot_frame)
        canvas_container.pack(side='top', fill='both', expand=True)

        # Store reference to the container for proper sizing
        self.plot_container = canvas_container

        # Plot canvas (pass the expandable container)
        self.setup_plot_canvas(canvas_container)

    def setup_plot_canvas(self, parent):
        """Create the matplotlib canvas for the spider plot with dynamic responsive sizing."""

        # Calculate dynamic plot size based on available space
        dynamic_width, dynamic_height = self.calculate_dynamic_plot_size(parent)

        # Create figure with calculated responsive sizing
        self.fig, self.ax = plt.subplots(figsize=(dynamic_width, dynamic_height), subplot_kw=dict(projection='polar'))
        self.fig.patch.set_facecolor('white')
        self.fig.subplots_adjust(left=0.1, right=0.9, top=0.85, bottom=0.1)

        # Create canvas with proper expansion configuration
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

        # Initialize empty plot
        self.update_plot()

        # Bind window resize events to update plot size
        self.window.bind('<Configure>', self.on_window_resize_plot, add=True)

    def ensure_canvas_expansion(self):
        """Ensure canvas frame expands to use full available height."""
        if hasattr(self, 'canvas_frame') and self.canvas_frame.winfo_exists():
            # Force the canvas frame to expand
            self.canvas_frame.update_idletasks()

            # Get parent dimensions
            parent = self.canvas_frame.master
            parent_height = parent.winfo_height()

            # Check if canvas is using full height
            canvas_height = self.canvas_frame.winfo_height()
            debug_print(f"DEBUG: Canvas frame height: {canvas_height}px, Parent height: {parent_height}px")

            if canvas_height < parent_height - 100:  # If significantly smaller
                debug_print("DEBUG: Canvas not using full height - forcing expansion")
                # Force redraw
                if hasattr(self, 'update_plot_size_for_resize'):
                    self.update_plot_size_for_resize()

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
        canvas_widget.bind("<Button-3>", show_context_menu)
        canvas_widget.bind("<Button-2>", show_context_menu)

    def calculate_dynamic_plot_size(self, parent_frame):
        """Calculate plot size that directly scales with window size - with enhanced debugging."""
        debug_print("DEBUG: Starting SIMPLE dynamic plot size calculation")

        # Force geometry update to get current dimensions
        self.window.update_idletasks()


        if hasattr(self, 'plot_container') and self.plot_container.winfo_exists():
            parent_frame = self.plot_container
            parent_frame.update_idletasks()

        # Get the actual dimensions
        available_width = parent_frame.winfo_width()
        available_height = parent_frame.winfo_height()

        aspect_ratio = available_height/available_width

        debug_print(f"DEBUG: Parent frame: {parent_frame.__class__.__name__}")
        debug_print(f"DEBUG: Available dimensions - Width: {available_width}px, Height: {available_height}px")

        # Simple fallback for initial sizing or very small windows
        if available_width < 200 or available_height < 200:
            debug_print("DEBUG: Using fallback size for small window")
            return (6, 4.8)

        # Reserve space for controls (checkboxes, labels, etc.)
        control_space = 100  # Space for labelframe title and checkbox row
        scrollbar_space = 20  # Account for potential scrollbar

        # Calculate available space more accurately
        plot_height_available = available_height - control_space - scrollbar_space
        plot_width_available = available_width - scrollbar_space

        # Use most of the available space for the plot
        plot_width_pixels = plot_width_available - 20  # 100px margin on each side for the legend
        plot_height_pixels = plot_height_available - 10  # 5px margin top/bottom

        debug_print(f"DEBUG: Plot space in pixels - Width: {plot_width_pixels}px, Height: {plot_height_pixels}px")

        # Convert to inches for matplotlib (using standard 100 DPI)
        plot_width_inches = plot_width_pixels / 100.0
        plot_height_inches = plot_width_pixels*aspect_ratio/100

        # Apply minimum sizes to prevent too small plots
        plot_width_inches = max(plot_width_inches, 4.0)
        plot_height_inches = max(plot_height_inches, 3.0)

        debug_print(f"DEBUG: FINAL plot size - Width: {plot_width_inches:.2f} inches, Height: {plot_height_inches:.2f} inches")

        return (plot_width_inches, plot_height_inches)

    def initialize_plot_size(self):
        """Initialize plot size after window is fully rendered to prevent startup artifacts."""
        debug_print("DEBUG: Initializing plot size after window render")

        # Give window time to fully render
        self.window.update_idletasks()

        # Force canvas updates
        if hasattr(self, 'right_canvas'):
            self.right_canvas.update_idletasks()

        # Update plot size
        if hasattr(self, 'update_plot_size_for_resize'):
            self.update_plot_size_for_resize()

        debug_print("DEBUG: Initial plot size set")

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
        self.ax.set_ylim(0, 9)
        self.ax.set_yticks(range(1, 10))
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
            self.ax.legend(loc='upper right', bbox_to_anchor=(1.1, 1.2), fontsize=8)

        # Set title
        self.ax.set_title('Sensory Profile Comparison', fontsize=12, fontweight='bold', pad=15, ha = 'center', y = 1.08)

        # Force canvas update
        self.canvas.draw_idle()

    def update_plot(self):
        """Update the spider plot."""
        self.create_spider_plot()
        self.bring_to_front()
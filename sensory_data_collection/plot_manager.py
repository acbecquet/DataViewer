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
from utils import debug_print, FONT


class PlotManager:
    """Manages sensory profile plotting and visualization."""
    
    def __init__(self, sensory_window, mode_manager=None):
        """Initialize the plot manager with reference to main window."""
        self.sensory_window = sensory_window
        self.mode_manager = mode_manager
        
    def on_window_resize_plot(self, event):
        """Handle window resize events to update plot size dynamically - ENHANCED DEBUG VERSION."""
        debug_print(f"DEBUG: RESIZE EVENT DETECTED - Widget: {event.widget}, Window: {self.sensory_window.window}")
        debug_print(f"DEBUG: Event widget type: {type(event.widget)}")
        debug_print(f"DEBUG: Window type: {type(self.sensory_window.window)}")
        debug_print(f"DEBUG: Event widget == window? {event.widget == self.sensory_window.window}")
        debug_print(f"DEBUG: Event widget is window? {event.widget is self.sensory_window.window}")

        # Only handle main window resize events, not child widgets
        if event.widget != self.sensory_window.window:
            debug_print(f"DEBUG: Ignoring resize event from child widget: {event.widget}")
            return

        debug_print("DEBUG: MAIN WINDOW RESIZE CONFIRMED - Processing...")

        # Get current window dimensions for verification
        current_width = self.sensory_window.window.winfo_width()
        current_height = self.sensory_window.window.winfo_height()
        debug_print(f"DEBUG: Current window dimensions: {current_width}x{current_height}")

        # Debounce rapid resize events
        if hasattr(self, '_resize_timer'):
            self.sensory_window.window.after_cancel(self._resize_timer)
            debug_print("DEBUG: Cancelled previous resize timer")

        # Schedule plot size update with a small delay to avoid excessive updates
        self._resize_timer = self.sensory_window.window.after(1000, self.update_plot_size_for_resize)
        debug_print("DEBUG: Scheduled plot resize update in 150ms")

    def update_plot_size_for_resize(self):
        """Update plot size with artifact prevention and frame validation."""
        try:
            # Check if we have the necessary components
            if not hasattr(self.sensory_window, 'canvas_frame') or not self.sensory_window.canvas_frame.winfo_exists():
                return

            if not hasattr(self.sensory_window, 'fig') or not self.sensory_window.fig:
                return

            # Wait for frame geometry to stabilize
            self.sensory_window.window.update_idletasks()

            # Use the actual plot container for sizing
            if hasattr(self, 'plot_container') and self.sensory_window.plot_container.winfo_exists():
                parent_for_sizing = self.sensory_window.plot_container
            else:
                parent_for_sizing = self.sensory_window.canvas_frame.master

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
            current_width, current_height = self.sensory_window.fig.get_size_inches()

            # Only update if change is significant
            width_diff = abs(new_width - current_width)
            height_diff = abs(new_height - current_height)
            threshold = 1  # Threshold to reduce excessive updates

            if width_diff > threshold or height_diff > threshold:
                debug_print(f"DEBUG: Significant size change detected - updating plot")

                # Apply the new size
                self.sensory_window.fig.set_size_inches(new_width, new_height)

                self.sensory_window.canvas.draw_idle()
            else:
                debug_print("DEBUG: Size change below threshold, skipping update to prevent artifacts")

        except Exception as e:
            debug_print(f"DEBUG: Error during plot resize: {str(e)}")
            import traceback
            debug_print(f"DEBUG: Full traceback: {traceback.format_exc()}")

    def setup_plot_panel(self):
        """Setup the right panel for spider plot visualization with proper resizing."""

        # Create the main plot frame with proper expansion settings
        plot_frame = ttk.LabelFrame(self.sensory_window.right_frame, text="Sensory Profile Comparison", padding=10)
        plot_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Sample selection for plotting (fixed height)
        control_frame = ttk.Frame(plot_frame)
        control_frame.pack(side='top', fill='x', pady=(0, 5))

        ttk.Label(control_frame, text="Select Samples to Display:", font=FONT).pack(anchor='w')

        # Checkboxes frame (will be populated when samples are added)
        self.sensory_window.checkbox_frame = ttk.Frame(control_frame)
        self.sensory_window.checkbox_frame.pack(fill='x', pady=5)

        # Canvas container frame (expandable)
        canvas_container = ttk.Frame(plot_frame)
        canvas_container.pack(side='top', fill='both', expand=True)

        # Store reference to the container for proper sizing
        self.sensory_window.plot_container = canvas_container

        # Plot canvas (pass the expandable container)
        self.setup_plot_canvas(canvas_container)

    def setup_plot_canvas(self, parent):
        """Create the matplotlib canvas for the spider plot with dynamic responsive sizing."""

        # Calculate dynamic plot size based on available space
        dynamic_width, dynamic_height = self.calculate_dynamic_plot_size(parent)

        # Create figure with calculated responsive sizing
        self.sensory_window.fig, self.sensory_window.ax = plt.subplots(figsize=(dynamic_width, dynamic_height), subplot_kw=dict(projection='polar'))
        self.sensory_window.fig.patch.set_facecolor('white')
        self.sensory_window.fig.subplots_adjust(left=0.1, right=0.9, top=0.85, bottom=0.1)

        # Create canvas with proper expansion configuration
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill='both', expand=True)

        # Store reference to canvas_frame for resize handling
        self.sensory_window.canvas_frame = canvas_frame

        # Create canvas
        self.sensory_window.canvas = FigureCanvasTkAgg(self.sensory_window.fig, canvas_frame)
        canvas_widget = self.sensory_window.canvas.get_tk_widget()
        canvas_widget.pack(fill='both', expand=True)

        # Add toolbar for additional functionality
        self.setup_plot_context_menu(canvas_widget)

        # Initialize empty plot
        self.update_plot()

        # Bind window resize events to update plot size
        self.sensory_window.window.bind('<Configure>', self.on_window_resize_plot, add=True)

    def ensure_canvas_expansion(self):
        """Ensure canvas frame expands to use full available height."""
        if hasattr(self, 'canvas_frame') and self.sensory_window.canvas_frame.winfo_exists():
            # Force the canvas frame to expand
            self.sensory_window.canvas_frame.update_idletasks()

            # Get parent dimensions
            parent = self.sensory_window.canvas_frame.master
            parent_height = parent.winfo_height()

            # Check if canvas is using full height
            canvas_height = self.sensory_window.canvas_frame.winfo_height()
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
        self.sensory_window.hidden_toolbar = NavigationToolbar2Tk(self.sensory_window.canvas, canvas_widget.master)
        self.sensory_window.hidden_toolbar.pack_forget()  # Hide the toolbar but keep functionality

        def show_context_menu(event):
            """Show simplified right-click context menu with essential plot options."""
            context_menu = tk.Menu(self.sensory_window.window, tearoff=0)

            # Essential options only
            context_menu.add_command(label="⚙️ Configure Plot",
                                   command=self.sensory_window.hidden_toolbar.configure_subplots)
            context_menu.add_separator()
            context_menu.add_command(label="💾 Save Plot...",
                                   command=self.sensory_window.hidden_toolbar.save_figure)

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
        self.sensory_window.window.update_idletasks()


        if hasattr(self.sensory_window, 'plot_container') and self.sensory_window.plot_container.winfo_exists():
            parent_frame = self.sensory_window.plot_container
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
        self.sensory_window.window.update_idletasks()

        # Force canvas updates
        if hasattr(self, 'right_canvas'):
            self.right_canvas.update_idletasks()

        # Update plot size
        if hasattr(self, 'update_plot_size_for_resize'):
            self.update_plot_size_for_resize()

        debug_print("DEBUG: Initial plot size set")

    def create_spider_plot(self):
        """Create a spider/radar plot of sensory data with comprehensive error handling."""
        debug_print("DEBUG: Starting create_spider_plot() with robust validation")
    
        try:
            # Validate essential components exist
            if not hasattr(self.sensory_window, 'ax') or self.sensory_window.ax is None:
                debug_print("ERROR: No axes object available for plotting")
                return
            
            if not hasattr(self.sensory_window, 'canvas') or self.sensory_window.canvas is None:
                debug_print("ERROR: No canvas object available for plotting")
                return
            
            # Clear the plot safely
            try:
                self.sensory_window.ax.clear()
                debug_print("DEBUG: Successfully cleared axes")
            except Exception as e:
                debug_print(f"ERROR: Failed to clear axes: {e}")
                return

            # Validate metrics exist and are properly configured
            if not hasattr(self.sensory_window, 'metrics') or not self.sensory_window.metrics:
                debug_print("ERROR: No metrics defined for plotting")
                self.sensory_window.ax.text(0.5, 0.5, 'Configuration Error\nNo sensory metrics defined',
                            transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                            fontsize=12, color='red')
                self.sensory_window.canvas.draw()
                return
            
            if len(self.sensory_window.metrics) < 3:
                debug_print("WARNING: Less than 3 metrics defined - spider plot may not display well")

            # Check if we have any data to plot
            if not hasattr(self.sensory_window, 'samples') or not self.sensory_window.samples:
                debug_print("DEBUG: No samples available for plotting")
                self.sensory_window.ax.text(0.5, 0.5, 'No samples to display\nAdd samples to begin evaluation',
                            transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                            fontsize=12, color='gray')
                self.sensory_window.canvas.draw()
                return

            debug_print(f"DEBUG: Found {len(self.sensory_window.samples)} total samples")

            # Get selected samples for plotting with validation
            selected_samples = []
            checkbox_system_available = (hasattr(self, 'sample_manager') and 
                                       hasattr(self.sample_manager, 'sample_checkboxes') and 
                                       self.sample_manager.sample_checkboxes)
        
            if checkbox_system_available:
                debug_print(f"DEBUG: Checkbox system available with {len(self.sample_manager.sample_checkboxes)} checkboxes")
            
                for sample_name, checkbox_var in self.sample_manager.sample_checkboxes.items():
                    try:
                        is_checked = checkbox_var.get()
                        is_in_samples = sample_name in self.sensory_window.samples
                        debug_print(f"DEBUG: Sample '{sample_name}' - Checked: {is_checked}, In samples: {is_in_samples}")
                    
                        if is_checked and is_in_samples:
                            # Validate sample data integrity
                            sample_data = self.sensory_window.samples[sample_name]
                            if isinstance(sample_data, dict):
                                selected_samples.append(sample_name)
                            else:
                                debug_print(f"WARNING: Sample '{sample_name}' has invalid data format")
                            
                    except Exception as e:
                        debug_print(f"ERROR: Failed to check sample '{sample_name}': {e}")
                        continue
            else:
                debug_print("DEBUG: Checkbox system not available - fallback to all samples")
                # Fallback: validate and select all samples with valid data
                for sample_name, sample_data in self.sensory_window.samples.items():
                    if isinstance(sample_data, dict):
                        selected_samples.append(sample_name)
                    else:
                        debug_print(f"WARNING: Skipping sample '{sample_name}' - invalid data format")

            debug_print(f"DEBUG: Selected {len(selected_samples)} samples for plotting: {selected_samples}")

            if not selected_samples:
                debug_print("DEBUG: No valid samples selected for plotting")
                self.sensory_window.ax.text(0.5, 0.5, 'No samples selected\nUse checkboxes to select samples',
                            transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                            fontsize=12, color='gray')
                self.sensory_window.canvas.draw()
                return

            # Validate sample data quality
            valid_samples = []
            for sample_name in selected_samples:
                sample_data = self.sensory_window.samples[sample_name]
            
                # Check if sample has data for at least some metrics
                metric_count = sum(1 for metric in self.sensory_window.metrics if metric in sample_data)
                if metric_count > 0:
                    valid_samples.append(sample_name)
                    debug_print(f"DEBUG: Sample '{sample_name}' has data for {metric_count}/{len(self.sensory_window.metrics)} metrics")
                else:
                    debug_print(f"WARNING: Sample '{sample_name}' has no metric data - skipping")

            if not valid_samples:
                debug_print("ERROR: No samples with valid metric data")
                self.sensory_window.ax.text(0.5, 0.5, 'No valid data to plot\nSamples exist but contain no metric data',
                            transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                            fontsize=12, color='orange')
                self.sensory_window.canvas.draw()
                return

            debug_print(f"DEBUG: Proceeding to plot {len(valid_samples)} samples with valid data")

            # Setup the spider plot with error handling
            try:
                num_metrics = len(self.sensory_window.metrics)
                angles = np.linspace(0, 2 * np.pi, num_metrics, endpoint=False).tolist()
                angles += angles[:1]  # Complete the circle

                # Set up the plot
                self.sensory_window.ax.set_theta_offset(np.pi / 2)
                self.sensory_window.ax.set_theta_direction(-1)
                self.sensory_window.ax.set_thetagrids(np.degrees(angles[:-1]), self.sensory_window.metrics, fontsize=10)
                self.sensory_window.ax.set_ylim(0, 9)
                self.sensory_window.ax.set_yticks(range(1, 10))
                self.sensory_window.ax.set_yticklabels(range(1, 10))
                self.sensory_window.ax.grid(True, alpha=0.3)
            
                debug_print("DEBUG: Spider plot axes configured successfully")

            except Exception as e:
                debug_print(f"ERROR: Failed to configure spider plot axes: {e}")
                self.sensory_window.ax.text(0.5, 0.5, 'Plot Configuration Error\nFailed to set up spider plot',
                            transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                            fontsize=12, color='red')
                self.sensory_window.canvas.draw()
                return

            # Colors and styles for different samples
            colors = ['red', 'green', 'blue', 'orange', 'purple', 'brown', 'pink', 'cyan', 'magenta', 'lime']
            line_styles = ['-', '--', '-.', ':']
        
            # Plot each valid sample with error handling
            plotted_count = 0
            for i, sample_name in enumerate(valid_samples):
                try:
                    sample_data = self.sensory_window.samples[sample_name]
                
                    # Build values array with validation
                    values = []
                    missing_metrics = []
                
                    for metric in self.sensory_window.metrics:
                        if metric in sample_data:
                            value = sample_data[metric]
                            # Validate numeric value
                            try:
                                numeric_value = float(value)
                                # Clamp to valid range
                                numeric_value = max(1, min(9, numeric_value))
                                values.append(numeric_value)
                            except (ValueError, TypeError):
                                debug_print(f"WARNING: Invalid value for {metric} in sample {sample_name}: {value}")
                                values.append(5)  # Default fallback
                                missing_metrics.append(metric)
                        else:
                            values.append(5)  # Default to neutral for missing data
                            missing_metrics.append(metric)
                
                    if missing_metrics:
                        debug_print(f"WARNING: Sample '{sample_name}' missing data for: {missing_metrics}")
                
                    values += values[:1]  # Complete the circle

                    # Get colors and styles with wraparound
                    color = colors[i % len(colors)]
                    line_style = line_styles[i % len(line_styles)]

                    # Plot the line and markers
                    self.sensory_window.ax.plot(angles, values, 'o', linewidth=2.5, label=sample_name,
                                color=color, linestyle=line_style, markersize=8, alpha=0.8)
                
                    # Fill the area
                    self.sensory_window.ax.fill(angles, values, alpha=0.1, color=color)
                
                    plotted_count += 1
                    debug_print(f"DEBUG: Successfully plotted sample '{sample_name}'")

                except Exception as e:
                    debug_print(f"ERROR: Failed to plot sample '{sample_name}': {e}")
                    continue

            if plotted_count == 0:
                debug_print("ERROR: Failed to plot any samples")
                self.sensory_window.ax.text(0.5, 0.5, 'Plotting Error\nFailed to render any sample data',
                            transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                            fontsize=12, color='red')
            else:
                debug_print(f"DEBUG: Successfully plotted {plotted_count} samples")
            
                # Add legend with error handling
                try:
                    if plotted_count > 0:
                        self.sensory_window.ax.legend(loc='upper right', bbox_to_anchor=(1.1, 1.2), fontsize=8)
                        debug_print("DEBUG: Legend added successfully")
                except Exception as e:
                    debug_print(f"WARNING: Failed to add legend: {e}")

                # Set title with error handling
                try:
                    self.sensory_window.ax.set_title('Sensory Profile Comparison', fontsize=12, fontweight='bold', 
                                                   pad=15, ha='center', y=1.08)
                    debug_print("DEBUG: Title set successfully")
                except Exception as e:
                    debug_print(f"WARNING: Failed to set title: {e}")

            # Force canvas update with error handling
            try:
                self.sensory_window.canvas.draw_idle()
                debug_print("DEBUG: Canvas update completed successfully")
            except Exception as e:
                debug_print(f"ERROR: Failed to update canvas: {e}")
                # Try immediate draw as fallback
                try:
                    self.sensory_window.canvas.draw()
                    debug_print("DEBUG: Fallback canvas draw successful")
                except Exception as e2:
                    debug_print(f"ERROR: Fallback canvas draw also failed: {e2}")

            debug_print("DEBUG: create_spider_plot() completed")

        except Exception as e:
            debug_print(f"CRITICAL ERROR in create_spider_plot(): {e}")
            import traceback
            debug_print(f"DEBUG: Full traceback: {traceback.format_exc()}")
        
            # Emergency fallback - try to show error message
            try:
                if hasattr(self.sensory_window, 'ax') and self.sensory_window.ax:
                    self.sensory_window.ax.clear()
                    self.sensory_window.ax.text(0.5, 0.5, f'Critical Plot Error\n{str(e)}',
                                transform=self.sensory_window.ax.transAxes, ha='center', va='center',
                                fontsize=12, color='red')
                    if hasattr(self.sensory_window, 'canvas') and self.sensory_window.canvas:
                        self.sensory_window.canvas.draw()
            except:
                debug_print("ERROR: Even emergency error display failed")

    def update_plot(self):
        """Update the spider plot."""
        debug_print("DEBUG: Starting plot update")
        try:
            self.create_spider_plot()
            if self.mode_manager:
                self.mode_manager.bring_to_front()
            debug_print("DEBUG: Plot update completed successfully")
        except Exception as e:
            debug_print(f"DEBUG: Error updating plot: {str(e)}")
            import traceback
            debug_print(f"DEBUG: Full traceback: {traceback.format_exc()}")
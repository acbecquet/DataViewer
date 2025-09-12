# plot_manager.py
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, Label, Button
from utils import wrap_text,APP_BACKGROUND_COLOR,BUTTON_COLOR, PLOT_CHECKBOX_TITLE,FONT, debug_print
import processing


def lazy_import_matplotlib_components():
    """Lazy import matplotlib components."""
    try:
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        from matplotlib.widgets import CheckButtons
        print("TIMING: Lazy loaded matplotlib components in PlotManager")
        return plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons
    except ImportError as e:
        print(f"Error importing matplotlib: {e}")
        return None, None, None, None

def lazy_import_pandas():
    """Lazy import pandas."""
    try:
        import pandas as pd
        print("TIMING: Lazy loaded pandas in PlotManager")
        return pd
    except ImportError as e:
        print(f"Error importing pandas: {e}")
        return None

class PlotManager:
    def __init__(self, parent):
        """
        The PlotManager is initialized with a reference to the DataViewer (parent)
        so that it can access shared variables such as selected_plot_type, selected_sheet,
        filtered_sheets, and plot_options.
        """

        # Add lazy loading infrastructure - THIS IS WHAT'S MISSING
        self._matplotlib_loaded = False
        self._pandas_loaded = False

        self.parent = parent
        self.figure = None
        self.canvas = None
        self.axes = None
        self.lines = None
        self.plot_dropdown = None
        # For convenience, reference shared variables from parent:
        self.selected_plot_type = self.parent.selected_plot_type
        self.selected_sheet = self.parent.selected_sheet
        self.plot_options = self.parent.plot_options
        self.check_buttons = None

    def get_matplotlib_components(self):
        """Get matplotlib components with lazy loading."""
        if not self._matplotlib_loaded:
            self.plt, self.FigureCanvasTkAgg, self.NavigationToolbar2Tk, self.CheckButtons = lazy_import_matplotlib_components()
            self._matplotlib_loaded = True
        return self.plt, self.FigureCanvasTkAgg, self.NavigationToolbar2Tk, self.CheckButtons

    def get_pandas(self):
        """Get pandas with lazy loading."""
        if not self._pandas_loaded:
            self.pd = lazy_import_pandas()
            self._pandas_loaded = True
        return self.pd

    @property
    def line_labels(self):
        """Get line labels, creating them if necessary."""
        if not hasattr(self.parent, 'line_labels'):
            self.parent.line_labels = []
        return self.parent.line_labels

    @line_labels.setter
    def line_labels(self, value):
        """Set line labels."""
        self.parent.line_labels = value

    def embed_plot_in_frame(self, fig, frame: ttk.Frame):
        """
        Embed a Matplotlib figure into a Tkinter frame with proper layout control.
        Enhanced to handle User Test Simulation split plots.
        """
        plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()

        # Error handling
        if not plt or not FigureCanvasTkAgg or not NavigationToolbar2Tk:
            print("ERROR: Could not load matplotlib components in embed_plot_in_frame")
            return None

        # Clear existing widgets
        for widget in frame.winfo_children():
            widget.destroy()

            # Create a separate frame for the dropdown - ensure it's visible
        self.dropdown_frame = ttk.Frame(frame)
        self.dropdown_frame.pack(side='bottom', fill='x', pady=10)

        # Create a container frame for the plot
        plot_container = ttk.Frame(frame)
        plot_container.pack(fill='both', expand=True, pady=(0, 0))

        if self.figure:
            plt.close(self.figure)

        # Check if this is a split plot and adjust margins accordingly
        is_split_plot = hasattr(fig, 'is_split_plot') and fig.is_split_plot

        if is_split_plot:
            # For split plots, adjust margins to accommodate checkboxes for both plots
            fig.subplots_adjust(right=0.80)
            debug_print("DEBUG: Split plot detected, adjusted margins")
        else:
            # Standard right margin - don't make it too small
            fig.subplots_adjust(right=0.82)

        # Embed figure in the plot container
        self.canvas = FigureCanvasTkAgg(fig, master=plot_container)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

        # Add toolbar to the plot container
        toolbar = NavigationToolbar2Tk(self.canvas, plot_container)
        toolbar.update()

        # Save references to the figure and its axes
        self.figure = fig

        if is_split_plot:
            # For split plots, we have multiple axes
            self.axes = fig.axes  # List of axes
            self.lines = []  # Will be handled by the split plot logic
            debug_print(f"DEBUG: Split plot - found {len(self.axes)} axes")
        else:
            # Standard single plot
            self.axes = fig.gca()
            self.lines = self.axes.lines

        # Bind scroll event for zooming
        self.canvas.mpl_connect("scroll_event", lambda event: self.zoom(event))

        # Add checkboxes for toggling plot elements
        self.add_checkboxes()

        return self.canvas

    def on_user_test_simulation_checkbox_click(self, wrapped_label):
        """
        Handle checkbox clicks for User Test Simulation split plots.
        Controls both Phase 1 and Phase 2 plots simultaneously.
        """
        debug_print(f"DEBUG: User Test Simulation checkbox clicked: {wrapped_label}")

        if not hasattr(self.figure, 'phase1_lines') or not hasattr(self.figure, 'phase2_lines'):
            debug_print("DEBUG: No phase lines found on figure")
            return

        # Get the original label from the wrapped label
        original_label = self.label_mapping.get(wrapped_label)
        if original_label is None:
            debug_print(f"DEBUG: Could not find original label for: {wrapped_label}")
            return

        # Find the index of the clicked sample
        try:
            index = self.parent.line_labels.index(original_label)
            debug_print(f"DEBUG: Found sample index: {index} for label: {original_label}")
        except ValueError:
            debug_print(f"DEBUG: Could not find index for label: {original_label}")
            return

        # Get the corresponding lines from both plots
        if index < len(self.figure.phase1_lines) and index < len(self.figure.phase2_lines):
            phase1_line = self.figure.phase1_lines[index]
            phase2_line = self.figure.phase2_lines[index]

            # Toggle visibility for both lines
            is_visible = phase1_line.get_visible()
            new_visibility = not is_visible

            phase1_line.set_visible(new_visibility)
            phase2_line.set_visible(new_visibility)

            debug_print(f"DEBUG: Toggled sample '{original_label}' visibility to {new_visibility} in both plots")

            # Redraw the canvas
            if self.canvas:
                self.canvas.draw_idle()
        else:
            debug_print(f"DEBUG: Index {index} out of range for phase lines")

    def on_user_test_simulation_bar_checkbox_click(self, wrapped_label):
        """
        Handle checkbox clicks for User Test Simulation split bar charts.
        Controls both Phase 1 and Phase 2 bar charts simultaneously.
        """
        debug_print(f"DEBUG: User Test Simulation bar chart checkbox clicked: {wrapped_label}")

        if not hasattr(self.figure, 'phase1_bars') or not hasattr(self.figure, 'phase2_bars'):
            debug_print("DEBUG: No phase bars found on figure")
            return

        # Get the original label from the wrapped label
        original_label = self.label_mapping.get(wrapped_label)
        if original_label is None:
            debug_print(f"DEBUG: Could not find original label for: {wrapped_label}")
            return

        # Find the index of the clicked sample
        try:
            index = self.parent.line_labels.index(original_label)
            debug_print(f"DEBUG: Found sample index: {index} for label: {original_label}")
        except ValueError:
            debug_print(f"DEBUG: Could not find index for label: {original_label}")
            return

        # Get the corresponding bars from both plots
        if index < len(self.figure.phase1_bars) and index < len(self.figure.phase2_bars):
            phase1_bar = self.figure.phase1_bars[index]
            phase2_bar = self.figure.phase2_bars[index]

            # Toggle visibility for both bars
            is_visible = phase1_bar.get_visible()
            new_visibility = not is_visible

            phase1_bar.set_visible(new_visibility)
            phase2_bar.set_visible(new_visibility)

            debug_print(f"DEBUG: Toggled sample '{original_label}' bar visibility to {new_visibility} in both plots")

            # Redraw the canvas
            if self.canvas:
                self.canvas.draw_idle()
        else:
            debug_print(f"DEBUG: Index {index} out of range for phase bars")

    def plot_all_samples(self, frame: ttk.Frame, full_sample_data, num_columns_per_sample: int) -> None:
        """
        Plot the provided sample data in the given frame.
        Enhanced to handle empty data for data collection.
        Uses processing.plot_all_samples to generate the figure (and sample names, if any),
        then embeds the figure into the frame.
        """
        pd = self.get_pandas()
        plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()
        # Error handling
        if not pd:
            print("ERROR: Could not load pandas in plot_all_samples")
            return None

        if not plt:
            print("ERROR: Could not load matplotlib in plot_all_samples")
            return None

        # Validate that full_sample_data is actually a DataFrame (optional safety check)
        if not hasattr(full_sample_data, 'empty') or not hasattr(full_sample_data, 'shape'):
            print("ERROR: full_sample_data is not a valid DataFrame-like object")
            return None
        debug_print(f"DEBUG: plot_all_samples called with data shape: {full_sample_data.shape}")

        # Clear frame contents
        for widget in frame.winfo_children():
            widget.pack_forget()

        # Check if data is empty or invalid
        if full_sample_data.empty:
            debug_print("DEBUG: Data is completely empty - showing placeholder for data collection")
            self.show_empty_plot_placeholder(frame, "No data loaded yet.\nUse 'Collect Data' to add measurements.")
            return

        # Check if data contains only NaN values
        if full_sample_data.isna().all().all():
            debug_print("DEBUG: Data contains only NaN values - showing placeholder for data collection")
            self.show_empty_plot_placeholder(frame, "No measurement data available yet.\nUse 'Collect Data' to add measurements.")
            return

        # Check if there's any numeric data for plotting
        numeric_data = full_sample_data.apply(pd.to_numeric, errors='coerce')
        if numeric_data.isna().all().all():
            debug_print("DEBUG: No numeric data available for plotting - showing placeholder")
            self.show_empty_plot_placeholder(frame, "No numeric data available for plotting.\nUse 'Collect Data' to add measurement values.")
            return

        debug_print("DEBUG: Data appears valid for plotting, proceeding with plot generation")

        # Extract sample names from the processed data if available
        sample_names = None
        try:
            # Try to get sample names from the parent GUI if it has processed data
            if hasattr(self.parent, 'current_sheet_data') and hasattr(self.parent.current_sheet_data, 'iloc'):
                # Extract sample names from the processed data table if available
                if 'Sample Name' in self.parent.current_sheet_data.columns:
                    sample_names = self.parent.current_sheet_data['Sample Name'].tolist()
                    debug_print(f"DEBUG: Extracted sample names from processed data: {sample_names}")
        except Exception as e:
            debug_print(f"DEBUG: Could not extract sample names from processed data: {e}")

        try:
            # Correct argument order - plot_type comes second, then num_columns_per_sample, then sample_names
            result = processing.plot_all_samples(
                full_sample_data,
                self.selected_plot_type.get(),
                num_columns_per_sample,
                sample_names
            )

            if isinstance(result, tuple):
                fig, extracted_sample_names = result
            else:
                fig, extracted_sample_names = result, None

            debug_print("DEBUG: Plot generated successfully, embedding in frame")

            # Embed the figure
            self.canvas = self.embed_plot_in_frame(fig, frame)

            # Add checkboxes using the extracted sample names
            self.add_checkboxes(sample_names=extracted_sample_names)

            debug_print("DEBUG: Plot embedded and checkboxes added successfully")

        except Exception as e:
            debug_print(f"ERROR: Failed to generate or embed plot: {e}")
            import traceback
            traceback.print_exc()
            # Show error message instead of plot
            self.show_empty_plot_placeholder(frame, f"Error generating plot: {str(e)}\nTry using 'Collect Data' to add valid measurements.")

    def show_empty_plot_placeholder(self, frame: ttk.Frame, message: str) -> None:
        """
        Show a placeholder message in the plot frame when no data is available.

        Args:
            frame (ttk.Frame): The frame to display the placeholder in
            message (str): The message to display
        """
        print(f"DEBUG: Showing empty plot placeholder with message: {message}")

        # Clear any existing widgets
        for widget in frame.winfo_children():
            widget.destroy()

        # Create a frame for the placeholder content
        placeholder_frame = ttk.Frame(frame)
        placeholder_frame.pack(fill="both", expand=True)

        # Add the placeholder message
        placeholder_label = tk.Label(
            placeholder_frame,
            text=message,
            font=("Arial", 12),
            fg="gray",
            justify="center",
            wraplength=300  # Wrap text for better readability
        )
        placeholder_label.pack(expand=True)

        # Add a small instruction
        instruction_label = tk.Label(
            placeholder_frame,
            text="Data collection will enable real-time plotting.",
            font=("Arial", 10),
            fg="lightgray",
            justify="center"
        )
        instruction_label.pack(pady=(10, 0))

        debug_print("DEBUG: Empty plot placeholder displayed successfully")

    def update_plot(self, full_sample_data, num_columns_per_sample, frame=None):
        """
        Update the plot dynamically. The frame must be provided.
        """
        if frame is None:
            raise ValueError("Plot frame must be provided.")
        for widget in frame.winfo_children():
            widget.pack_forget()
        self.plot_all_samples(frame, full_sample_data, num_columns_per_sample)

    def add_plot_dropdown(self, frame: ttk.Frame) -> None:
        """
        Create or update the plot type dropdown, centered with no background color.
        """
        # Use existing dropdown frame or create a new one
        if not hasattr(self, 'dropdown_frame') or not self.dropdown_frame.winfo_exists():
            self.dropdown_frame = ttk.Frame(frame)
            self.dropdown_frame.pack(side='bottom', fill='x', pady=5)

        # Clear existing widgets
        for widget in self.dropdown_frame.winfo_children():
            widget.destroy()

        # Create a container frame to center the dropdown components
        center_frame = ttk.Frame(self.dropdown_frame)
        center_frame.pack(side="top", fill="x")

        # Use pack with expand=True to center the content
        label_frame = ttk.Frame(center_frame)
        label_frame.pack(side="top", fill="none", expand=True)

        # Create the dropdown label
        dropdown_label = ttk.Label(
            label_frame,
            text="Plot Type:",
            font=('Arial', 10), foreground = 'white'
        )
        dropdown_label.pack(side="left", padx=(0, 5))

        # Create the dropdown with no background styling
        self.plot_dropdown = ttk.Combobox(
            label_frame,
            textvariable=self.selected_plot_type,
            values=self.plot_options,
            state="readonly",
            font=('Arial', 10),
            width=15
        )
        self.plot_dropdown.pack(side="left")

        # Remove ALL background styling
        style = ttk.Style()
        style.map('TCombobox', fieldbackground=[('readonly', 'white')])
        style.map('TCombobox', selectbackground=[('readonly', 'white')])
        style.configure('TCombobox', background='white')

        # Additional styling to ensure no blue background
        self.plot_dropdown.configure(background='white')

        # Bind event handler
        self.plot_dropdown.bind("<<ComboboxSelected>>", self.update_plot_from_dropdown)

        # Set default value if needed
        if self.selected_plot_type.get() not in self.plot_options:
            self.selected_plot_type.set(self.plot_options[0])

    def update_plot_from_dropdown(self, event) -> None:
        """
        Callback for when the plot type dropdown changes.
        Retrieves the current sheet and data from the parent and updates the plot.
        """
        selected_plot = self.selected_plot_type.get()
        if not selected_plot:
            messagebox.showerror("Error", "No plot type selected.")
            return
        try:
            current_sheet_name = self.parent.selected_sheet.get()
            if current_sheet_name not in self.parent.filtered_sheets:
                messagebox.showerror("Error", f"Sheet '{current_sheet_name}' not found.")
                return
            sheet_info = self.parent.filtered_sheets[current_sheet_name]
            sheet_data = sheet_info["data"]
            if sheet_data.empty:
                messagebox.showwarning("Warning", "The selected sheet has no data.")
                return
            # Process data for the selected plot option.
            process_function = processing.get_processing_function(current_sheet_name)
            processed_data, _, full_sample_data = process_function(sheet_data)

            # Use the correct number of columns per sample
            num_columns = getattr(self.parent, 'num_columns_per_sample', 12)
            debug_print(f"DEBUG: update_plot_from_dropdown using {num_columns} columns per sample")

            # Update the plot
            self.plot_all_samples(self.parent.plot_frame, full_sample_data, num_columns)

            # Refresh the dropdown in case it was recreated.
            if not self.plot_dropdown or not self.plot_dropdown.winfo_exists():
                self.add_plot_dropdown(self.parent.plot_frame)
            else:
                self.plot_dropdown["values"] = self.parent.plot_options  # Use parent's plot options
                self.plot_dropdown.set(selected_plot)

        except Exception as e:
            print(f"ERROR: Error in update_plot_from_dropdown: {e}")
            messagebox.showerror("Error", f"An error occurred while updating the plot: {e}")

    def clear_plot_area(self):
        """
        Clear the plot area and release Matplotlib resources.
        """
        plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()

        if not plt:
            print("ERROR: Could not load matplotlib")
            return None

        if hasattr(self.parent, 'plot_frame') and self.parent.plot_frame.winfo_exists():
            for widget in self.parent.plot_frame.winfo_children():
                widget.destroy()
        if self.figure:
            plt.close(self.figure)
            self.figure = None
        if self.canvas:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None
        self.axes = None
        self.lines = None
        self.parent.line_labels = []
        self.parent.original_lines_data = []
        self.check_buttons = None
        self.checkbox_cid = None

    def zoom(self, event):
        """
        Zoom in/out on the plot with the cursor as the zoom origin.
        """
        if event.inaxes is not self.axes:
            return
        x_min, x_max = self.axes.get_xlim()
        y_min, y_max = self.axes.get_ylim()
        cursor_x, cursor_y = event.xdata, event.ydata
        zoom_scale = 0.9 if event.button == 'up' else 1.1
        if cursor_x is None or cursor_y is None:
            return
        new_x_min = cursor_x - (cursor_x - x_min) * zoom_scale
        new_x_max = cursor_x + (x_max - cursor_x) * zoom_scale
        new_y_min = cursor_y - (cursor_y - y_min) * zoom_scale
        new_y_max = cursor_y + (y_max - cursor_y) * zoom_scale
        self.axes.set_xlim(new_x_min, new_x_max)
        self.axes.set_ylim(new_y_min, new_y_max)
        if self.canvas:
            self.canvas.draw_idle()

    def add_checkboxes(self, sample_names=None):
        """
        Add checkboxes to toggle visibility of plot elements with properly balanced spacing.
        Enhanced to handle User Test Simulation split plots and split bar charts.
        """
        plt, FigureCanvasTkAgg, NavigationToolbar2Tk, CheckButtons = self.get_matplotlib_components()

        if not plt:
            print("ERROR: Could not load matplotlib")
            return None

        self.parent.line_labels = []
        self.parent.original_lines_data = []
        self.label_mapping = {}
        wrapped_labels = []

        if self.check_buttons:
            self.check_buttons.ax.clear()
            self.check_buttons = None

        is_bar_chart = self.selected_plot_type.get() == "TPM (Bar)"

        # Check if this is a User Test Simulation split plot
        is_split_plot = hasattr(self.figure, 'is_split_plot') and self.figure.is_split_plot
        is_split_bar_chart = is_split_plot and hasattr(self.figure, 'is_bar_chart') and self.figure.is_bar_chart
        debug_print(f"DEBUG: add_checkboxes - is_split_plot: {is_split_plot}, is_bar_chart: {is_bar_chart}, is_split_bar_chart: {is_split_bar_chart}")

        if sample_names is None:
            sample_names = self.parent.line_labels

        if sample_names:
            self.parent.line_labels = sample_names
            debug_print(f"DEBUG: Using provided sample_names: {sample_names}")

        # Handle different plot types in priority order
        if is_split_bar_chart and hasattr(self.figure, 'phase1_bars'):
            # For split bar charts, store bar data from both phases
            if not self.parent.line_labels:
                self.parent.line_labels = sample_names or [f"Sample {i+1}" for i in range(len(self.figure.phase1_bars))]

            # Store original data for both phases (bar charts use different data structure)
            phase1_data = [(bar.get_x(), bar.get_height()) for bar in self.figure.phase1_bars]
            phase2_data = [(bar.get_x(), bar.get_height()) for bar in self.figure.phase2_bars]
            self.parent.original_lines_data = list(zip(phase1_data, phase2_data))
            debug_print(f"DEBUG: Split bar chart - stored data for {len(self.parent.line_labels)} samples")
        elif is_split_plot and hasattr(self.figure, 'phase1_lines'):
            # For split line plots, use the sample names and store line data from both phases
            if not self.parent.line_labels:
                self.parent.line_labels = sample_names or [f"Sample {i+1}" for i in range(len(self.figure.phase1_lines))]

            # Store original data for both phases
            phase1_data = [(line.get_xdata(), line.get_ydata()) for line in self.figure.phase1_lines]
            phase2_data = [(line.get_xdata(), line.get_ydata()) for line in self.figure.phase2_lines]
            self.parent.original_lines_data = list(zip(phase1_data, phase2_data))
            debug_print(f"DEBUG: Split plot - stored data for {len(self.parent.line_labels)} samples")
        elif is_bar_chart and sample_names:
            # For regular bar charts (not split plots)
            self.parent.line_labels = sample_names
            self.parent.original_lines_data = [(patch.get_x(), patch.get_height()) for patch in self.axes.patches]
        else:
            if not self.parent.line_labels:
                self.parent.line_labels = [line.get_label() for line in self.lines]
                self.parent.original_lines_data = [(line.get_xdata(), line.get_ydata()) for line in self.lines]

        for label in self.parent.line_labels:
            wrapped_label = wrap_text(text=label, max_width=12)
            self.label_mapping[wrapped_label] = label
            wrapped_labels.append(wrapped_label)

        # CREATE THE CHECKBOXES - This was missing!
        # Adjusted positioning - increase top margin, reduce width on right
        # Format: [left, bottom, width, height]
        checkbox_ax = self.figure.add_axes([
            0.835,      # Left position
            0.33,       # Bottom position - lowered to give more space on top
            0.125,      # Width - reduced to remove empty space on right
            0.55        # Height - increased slightly to ensure all items fit
        ])

        self.check_buttons = CheckButtons(checkbox_ax, wrapped_labels, [True]*len(self.parent.line_labels))
        checkbox_ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)
        checkbox_ax.grid(False)

        for spine in checkbox_ax.spines.values():
            spine.set_visible(False)

        # Make font slightly smaller to ensure all labels fit
        for label in self.check_buttons.labels:
            label.set_fontsize(7)

        # Adjust the rectangle border to have proper margins around checkboxes
        rect = plt.Rectangle(
            (0.05, 0.02),     # Slight adjustment to left and bottom insets
            0.95,              # Width relative to axis
            0.96,             # Height relative to axis - increased to fit content
            transform=checkbox_ax.transAxes,
            facecolor='none',
            edgecolor='black',
            linewidth=1.5,
            zorder=10
        )
        checkbox_ax.add_patch(rect)

        # Adjust title position and border
        title_text = PLOT_CHECKBOX_TITLE
        title_x = 0.835 + 0.125/2  # Center in checkbox area
        title_y = 0.33 + 0.55 + 0.025  # Just above the checkbox area

        self.figure.text(title_x, title_y, title_text, fontsize=8, ha='center', va='center', wrap=True)

        # Adjust title border
        title_border_x = 0.825
        title_border_y = title_y - 0.03
        border_width = 0.14
        border_height = 0.065

        self.figure.add_artist(plt.Rectangle(
            (title_border_x, title_border_y),
            border_width,
            border_height,
            edgecolor='black',
            facecolor='white',
            lw=1,
            zorder=2
        ))

        # Bind callbacks based on plot type
        if is_split_bar_chart:
            debug_print("DEBUG: Using User Test Simulation split bar chart checkbox callback")
            self.checkbox_cid = self.check_buttons.on_clicked(self.on_user_test_simulation_bar_checkbox_click)
        elif is_split_plot:
            debug_print("DEBUG: Using User Test Simulation checkbox callback")
            self.checkbox_cid = self.check_buttons.on_clicked(self.on_user_test_simulation_checkbox_click)
        elif is_bar_chart:
            debug_print("DEBUG: Using bar chart checkbox callback")
            self.checkbox_cid = self.check_buttons.on_clicked(self.on_bar_checkbox_click)
        else:
            debug_print("DEBUG: Using standard line plot checkbox callback")
            self.checkbox_cid = self.check_buttons.on_clicked(self.on_checkbox_click)

        if self.canvas:
            self.canvas.draw_idle()

    def on_bar_checkbox_click(self, wrapped_label):
        """
        Toggle visibility of a bar when its checkbox is clicked.
        """
        original_label = self.label_mapping.get(wrapped_label)
        if original_label is None:
            return
        index = self.parent.line_labels.index(original_label)
        bar = self.axes.patches[index]
        bar.set_visible(not bar.get_visible())
        if self.canvas:
            self.canvas.draw_idle()

    def on_checkbox_click(self, wrapped_label):
        """
        Toggle visibility of a line when its checkbox is clicked.
        """
        original_label = self.label_mapping.get(wrapped_label)
        if original_label is None:
            return
        index = list(self.label_mapping.values()).index(original_label)
        if self.selected_plot_type.get() == "TPM (Bar)":
            bar = self.lines[index]
            bar.set_visible(not bar.get_visible())
        else:
            line = self.lines[index]
            line.set_visible(not line.get_visible())
        if self.canvas:
            self.canvas.draw_idle()

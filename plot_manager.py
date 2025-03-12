# plot_manager.py
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.widgets import CheckButtons
from tkinter import ttk, messagebox, Toplevel, Label, Button
from utils import wrap_text,APP_BACKGROUND_COLOR,BUTTON_COLOR, PLOT_CHECKBOX_TITLE,FONT
import processing
import pandas as pd

class PlotManager:
    def __init__(self, parent):
        """
        The PlotManager is initialized with a reference to the TestingGUI (parent)
        so that it can access shared variables such as selected_plot_type, selected_sheet,
        filtered_sheets, and plot_options.
        """
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

    def embed_plot_in_frame(self, fig: plt.Figure, frame: ttk.Frame) -> FigureCanvasTkAgg:
        """
        Embed a Matplotlib figure into a Tkinter frame with proper layout control.
        """
        # Clear existing widgets
        for widget in frame.winfo_children():
            widget.destroy()
    
        # Create a container frame for the plot with fixed height
        plot_container = ttk.Frame(frame)
        plot_container.pack(fill='both', expand=True, pady=(0, 5))  # Add padding at bottom for dropdown
    
        # Create a separate frame for the dropdown
        self.dropdown_frame = ttk.Frame(frame)
        self.dropdown_frame.pack(side='bottom', fill='x', pady=2, before = plot_container)
    
        if self.figure:
            plt.close(self.figure)
    
        fig.subplots_adjust(right=0.82)
    
        # Embed figure in the plot container (not the whole frame)
        self.canvas = FigureCanvasTkAgg(fig, master=plot_container)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
    
        # Add toolbar to the plot container
        toolbar = NavigationToolbar2Tk(self.canvas, plot_container)
        toolbar.update()
    
        # Save references to the figure and its axes
        self.figure = fig
        self.axes = fig.gca()
        self.lines = self.axes.lines
    
        # Bind scroll event for zooming
        self.canvas.mpl_connect("scroll_event", lambda event: self.zoom(event))
    
        # Add checkboxes for toggling plot elements
        self.add_checkboxes()
    
        return self.canvas

    def plot_all_samples(self, frame: ttk.Frame, full_sample_data: pd.DataFrame, num_columns_per_sample: int) -> None:
        """
        Plot the provided sample data in the given frame.
        Uses processing.plot_all_samples to generate the figure (and sample names, if any),
        then embeds the figure into the frame.
        """
        # Clear frame contents
        # Clear frame contents
        for widget in frame.winfo_children():
            widget.pack_forget()
        if full_sample_data.empty or full_sample_data.isna().all().all():
            messagebox.showwarning("Warning", "The data is empty or invalid for plotting.")
            return
        # Generate figure (and optionally sample names) via processing module.
        result = processing.plot_all_samples(full_sample_data, num_columns_per_sample, self.selected_plot_type.get())
        if isinstance(result, tuple):
            fig, sample_names = result
        else:
            fig, sample_names = result, None
        # Embed the figure
        self.canvas = self.embed_plot_in_frame(fig, frame)
        # Add checkboxes using the (optional) sample names
        self.add_checkboxes(sample_names=sample_names)

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
            font=('Arial', 10)
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
            # Update the plot (assuming parent.plot_frame exists)
            self.plot_all_samples(self.parent.plot_frame, full_sample_data, 12)
            # Refresh the dropdown in case it was recreated.
            if not self.plot_dropdown or not self.plot_dropdown.winfo_exists():
                self.add_plot_dropdown(self.parent.plot_frame)
            else:
                self.plot_dropdown["values"] = self.plot_options
                self.plot_dropdown.set(selected_plot)
            # Optionally, update the displayed sheet in the parent.
            self.parent.update_displayed_sheet(current_sheet_name)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while updating the plot: {e}")

    def clear_plot_area(self):
        """
        Clear the plot area and release Matplotlib resources.
        """
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
        """
        self.parent.line_labels = []
        self.parent.original_lines_data = []
        self.label_mapping = {}
        wrapped_labels = []
    
        if self.check_buttons:
            self.check_buttons.ax.clear()
            self.check_buttons = None
        
        is_bar_chart = self.selected_plot_type.get() == "TPM (Bar)"
    
        if sample_names is None:
            sample_names = self.parent.line_labels
        
        if is_bar_chart and sample_names:
            self.parent.line_labels = sample_names
            self.parent.original_lines_data = [(patch.get_x(), patch.get_height()) for patch in self.axes.patches]
        else:
            if not self.parent.line_labels:
                self.parent.line_labels = [line.get_label() for line in self.lines]
                self.parent.original_lines_data = [(line.get_xdata(), line.get_ydata()) for line in self.lines]
            
        for label in self.parent.line_labels:
            wrapped_label = wrap_text(text=label, max_width=10)
            self.label_mapping[wrapped_label] = label
            wrapped_labels.append(wrapped_label)
    
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
            label.set_fontsize(7.5)
    
        # Adjust the rectangle border to have proper margins around checkboxes
        rect = plt.Rectangle(
            (0.05, 0.02),     # Slight adjustment to left and bottom insets
            0.9,              # Width relative to axis
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
        title_border_x = 0.835
        title_border_y = title_y - 0.03
        border_width = 0.125
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
    
        # Bind callbacks
        if is_bar_chart:
            self.checkbox_cid = self.check_buttons.on_clicked(self.on_bar_checkbox_click)
        else:
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

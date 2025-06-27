# trend_analysis_gui.py
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, Label, Button
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import pandas as pd
import processing
from utils import get_resource_path, debug_print

# Global constants
APP_BACKGROUND_COLOR = '#0504AA'
BUTTON_COLOR = '#4169E1'
FONT = ('Arial', 12)
PLOT_CHECKBOX_TITLE = "Click Checkbox to \nAdd/Remove Item \nFrom Plot"

class TrendAnalysisGUI:
    def __init__(self, root, filtered_sheets: dict, plot_options: list):
        """
        Initialize the TrendAnalysisGUI with a reference to the main Tk window,
        the filtered sheets from TestingGUI, and the list of plot options.
        """
        self.root = root
        self.filtered_sheets = filtered_sheets
        self.plot_options = plot_options
        self.selected_agg_plot = None  # Will be set when opening the window
        self.aggregate_plots = {}  # Cache for aggregate plots if desired

    def center_window(self, window: tk.Toplevel, width: int = None, height: int = None) -> None:
        """
        Center the given Toplevel window on the screen.
        """
        window.update_idletasks()
        w = width or window.winfo_width()
        h = height or window.winfo_height()
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        window.geometry(f"{w}x{h}+{x}+{y}")

    def build_aggregate_plots_dict(self) -> dict:
        """
        For each plotting sheet in filtered_sheets, compute aggregate metrics,
        generate an aggregate plot, and store it in a dictionary.
        Returns a dictionary mapping sheet names to their Matplotlib figure.
        """
        aggregate_plots = {}
        plot_sheet_names = processing.get_plot_sheet_names()
        for sheet_name, sheet_info in self.filtered_sheets.items():
            if sheet_name not in plot_sheet_names:
                continue
            full_sample_data = sheet_info["data"]
            if full_sample_data.empty:
                continue
            agg_df = processing.aggregate_sheet_metrics(full_sample_data, num_columns_per_sample=12)
            agg_df["Sheet"] = sheet_name
            fig = processing.plot_aggregate_trends(agg_df)
            aggregate_plots[sheet_name] = fig
        return aggregate_plots

    def collect_trend_data_aggregate(self, is_plotting_sheet: bool = True) -> pd.DataFrame:
        """
        For each sheet in filtered_sheets that qualifies as a plotting sheet,
        compute aggregate metrics and add a 'Sheet' column.
        Returns a DataFrame with columns: "Sheet", "Sample Name", "Average TPM", "Total Puffs".
        """
        aggregate_data = []
        plot_sheet_names = processing.get_plot_sheet_names()
        for sheet_name, sheet_info in self.filtered_sheets.items():
            if is_plotting_sheet and sheet_name not in plot_sheet_names:
                continue
            full_sample_data = sheet_info["data"]
            if not full_sample_data.empty:
                agg_df = processing.aggregate_sheet_metrics(full_sample_data, num_columns_per_sample=12)
                agg_df["Sheet"] = sheet_name
                aggregate_data.append(agg_df)
        if aggregate_data:
            return pd.concat(aggregate_data, ignore_index=True)
        else:
            return pd.DataFrame()

    def open_trend_analysis_window(self) -> None:
        """
        Open a new window for trend analysis. The window contains a dropdown of
        valid plotting sheets and a frame to display the aggregate plot.
        """
        trend_window = tk.Toplevel(self.root)
        trend_window.title("Trend Analysis")
        trend_window.geometry("800x600")
        self.center_window(trend_window)

        # Frame for dropdown
        dropdown_frame = ttk.Frame(trend_window)
        dropdown_frame.pack(pady=10)

        # Get valid plotting sheet names from filtered_sheets
        valid_plot_sheet_names = [
            sheet_name for sheet_name in self.filtered_sheets.keys()
            if sheet_name in processing.get_plot_sheet_names()
        ]
        if not valid_plot_sheet_names:
            messagebox.showerror("Error", "No valid plotting sheets found for trend analysis.")
            trend_window.destroy()
            return

        # Set up a StringVar and create the dropdown
        self.selected_agg_plot = tk.StringVar(value=valid_plot_sheet_names[0])
        dropdown = ttk.Combobox(
            dropdown_frame,
            textvariable=self.selected_agg_plot,
            values=valid_plot_sheet_names,
            state="readonly",
            font=FONT
        )
        dropdown.pack()
        # Create a frame for the aggregate plot
        plot_frame = ttk.Frame(trend_window)
        plot_frame.pack(fill="both", expand=True)
        # Bind dropdown selection to update the plot
        dropdown.bind("<<ComboboxSelected>>", lambda event: self.update_aggregate_plot(plot_frame))
        # Initially generate and display the aggregate plot for the first valid sheet
        self.update_aggregate_plot(plot_frame)

    def display_aggregate_plot(self, parent, fig: plt.Figure) -> None:
        """
        Embed the given Matplotlib figure into the parent widget.
        """
        # Clear existing widgets (except for dropdowns)
        for widget in parent.winfo_children():
            if isinstance(widget, ttk.Combobox):
                continue
            widget.destroy()
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar = NavigationToolbar2Tk(canvas, parent)
        toolbar.update()
        toolbar.pack(side="bottom", fill="x")

    def update_aggregate_plot(self, parent) -> None:
        """
        Regenerate and display the aggregate plot for the sheet selected in the dropdown.
        """
        selected_sheet = self.selected_agg_plot.get()
        sheet_info = self.filtered_sheets.get(selected_sheet)
        if not sheet_info:
            messagebox.showerror("Error", f"Sheet '{selected_sheet}' not found.")
            return
        full_sample_data = sheet_info["data"]
        debug_print(f"Selected sheet: {selected_sheet}, shape: {full_sample_data.shape}")
        if full_sample_data.empty:
            messagebox.showerror("Error", f"No data available for sheet '{selected_sheet}'.")
            return
        agg_df = processing.aggregate_sheet_metrics(full_sample_data, num_columns_per_sample=12)
        agg_df["Sheet"] = selected_sheet
        new_fig = processing.plot_aggregate_trends(agg_df)
        # Optionally, cache the figure
        if not hasattr(self, "aggregate_plots"):
            self.aggregate_plots = {}
        self.aggregate_plots[selected_sheet] = new_fig
        self.display_aggregate_plot(parent, new_fig)

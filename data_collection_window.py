"""
data_collection_window.py
Developed by Charlie Becquet.
Interface for rapid test data collection.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import numpy as np
import os
import openpyxl
from openpyxl.styles import PatternFill
from utils import FONT

class DataCollectionWindow:
    def __init__(self, parent, file_path, test_name, header_data):
        """
        Initialize the data collection window.
        
        Args:
            parent (tk.Tk): The parent window.
            file_path (str): Path to the Excel file.
            test_name (str): Name of the test being conducted.
            header_data (dict): Dictionary containing header data.
        """
        self.parent = parent
        self.file_path = file_path
        self.test_name = test_name
        self.header_data = header_data
        self.num_samples = header_data["num_samples"]
        self.result = None
        
        # Create the window
        self.window = tk.Toplevel(parent)
        self.window.title(f"Data Collection - {test_name}")
        self.window.geometry("1100x600")  # Wider window to accommodate TPM panel
        self.window.minsize(1000, 500)
        
        # Default puff interval
        self.puff_interval = 10  # Default to 10
        
        # Set up keyboard shortcut flags
        self.hotkeys_enabled = True
        self.hotkey_bindings = {}
        self.navigation_keys_enable = True
        
        # Create the style for ttk widgets
        self.style = ttk.Style()
        self.setup_styles()
        
        # Data storage
        self.data = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            self.data[sample_id] = {
                "puffs": [],           # Will be populated dynamically
                "before_weight": [],   # Will be populated dynamically
                "after_weight": [],    # Will be populated dynamically
                "draw_pressure": [],   # Will be populated dynamically
                "smell": [],           # Will be populated dynamically
                "notes": [],           # Will be populated dynamically
                "tpm": [],             # Will store calculated TPM values
                "current_row_index": 0, # Track the current editable row
                "avg_tpm": 0.0         # Track average TPM
            }
            
            # Add first row with initial puff interval
            self.data[sample_id]["puffs"].append(self.puff_interval)
            self.data[sample_id]["before_weight"].append("")
            self.data[sample_id]["after_weight"].append("")
            self.data[sample_id]["draw_pressure"].append("")
            self.data[sample_id]["smell"].append("")
            self.data[sample_id]["notes"].append("")
            self.data[sample_id]["tpm"].append(None)
        
        # Create the UI
        self.create_widgets()
        
        # Center the window
        self.center_window()
        
        # Set up event handlers
        self.setup_event_handlers()
    
    def setup_styles(self):
        """Set up styles for ttk widgets to ensure no blue backgrounds."""
        # Configure ttk styles to use system defaults
        self.style.configure('TFrame', background='SystemButtonFace')
        self.style.configure('TLabel', background='SystemButtonFace')
        self.style.configure('TLabelframe', background='SystemButtonFace')
        self.style.configure('TLabelframe.Label', background='SystemButtonFace')
        self.style.configure('TNotebook', background='SystemButtonFace')
        self.style.configure('TNotebook.Tab', background='SystemButtonFace')
        
        # Create special styles for headers
        self.style.configure('Header.TLabel', 
                             font=("Arial", 14, "bold"), 
                             background='SystemButtonFace')
        
        self.style.configure('SubHeader.TLabel', 
                             font=("Arial", 12), 
                             background='SystemButtonFace')
        
        # Style for stats panel
        self.style.configure('Stats.TLabel', 
                             font=("Arial", 14, "bold"), 
                             background='SystemButtonFace')
        
        # Style for sample info
        self.style.configure('SampleInfo.TLabel',
                             font=("Arial", 11),
                             background='SystemButtonFace')
    
    def center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """Create the data collection UI."""
        # Set window background to system default
        self.window.configure(background='SystemButtonFace')
        
        # Create a horizontal split layout
        main_frame = ttk.Frame(self.window, padding=10, style='TFrame')
        main_frame.pack(fill="both", expand=True)
        
        # Header with test information
        header_frame = ttk.Frame(main_frame, style='TFrame')
        header_frame.pack(fill="x", pady=(0, 10))
        
        # Use styled labels
        ttk.Label(header_frame, text=f"Test: {self.test_name}", style='Header.TLabel').pack(side="left")
        ttk.Label(header_frame, text=f"Tester: {self.header_data['common']['tester']}", style='SubHeader.TLabel').pack(side="right")
        
        # Create a horizontal paned window to split the main area
        paned_window = ttk.PanedWindow(main_frame, orient="horizontal")
        paned_window.pack(fill="both", expand=True)
        
        # Left side - Data entry area
        data_frame = ttk.Frame(paned_window, style='TFrame')
        paned_window.add(data_frame, weight=3)  # 75% of the width
        
        # Right side - TPM stats panel
        self.stats_frame = ttk.Frame(paned_window, style='TFrame')
        paned_window.add(self.stats_frame, weight=1)  # 25% of the width
        
        # Setup data entry area with notebook
        self.notebook = ttk.Notebook(data_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Create a tab for each sample
        self.sample_frames = []
        self.sample_trees = []
        
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            sample_frame = ttk.Frame(self.notebook, padding=10, style='TFrame')
            self.notebook.add(sample_frame, text=f"Sample {i+1} - {self.header_data['samples'][i]['id']}")
            self.sample_frames.append(sample_frame)
            
            # Create the sample tab content
            tree = self.create_sample_tab(sample_frame, sample_id, i)
            self.sample_trees.append(tree)
        
        # Create the TPM stats panel
        self.create_tpm_stats_panel()
        
        # Control buttons at the bottom
        button_frame = ttk.Frame(main_frame, style='TFrame')
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Left side controls
        left_controls = ttk.Frame(button_frame, style='TFrame')
        left_controls.pack(side="left", fill="x")
        
        ttk.Label(left_controls, text="Puff Interval:", style='TLabel').pack(side="left")
        self.puff_interval_var = tk.IntVar(value=self.puff_interval)
        puff_spinbox = ttk.Spinbox(
            left_controls, 
            from_=1, 
            to=100, 
            textvariable=self.puff_interval_var, 
            width=5,
            command=self.update_puff_interval
        )
        puff_spinbox.pack(side="left", padx=5)
        
        # Sample navigation
        nav_frame = ttk.Frame(button_frame, style='TFrame')
        nav_frame.pack(side="left", padx=20)
        
        ttk.Button(nav_frame, text="← Prev Sample", command=self.go_to_previous_sample).pack(side="left")
        ttk.Button(nav_frame, text="Next Sample →", command=self.go_to_next_sample).pack(side="left", padx=5)
        
        # Right side controls
        ttk.Button(button_frame, text="Save Data", command=self.save_data).pack(side="right", padx=(10, 0))
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side="right")
        
        # Bind notebook tab change to update stats panel
        self.notebook.bind("<<NotebookTabChanged>>", self.update_stats_panel)
    
    def create_sample_tab(self, parent_frame, sample_id, sample_index):
        """Create a tab for a single sample with fast data entry."""
        # Sample metadata display
        info_frame = ttk.Frame(parent_frame, style='TFrame')
        info_frame.pack(fill="x", pady=(0, 10))
        
        # Display sample metadata without background color
        ttk.Label(info_frame, 
                 text=f"Sample ID: {self.header_data['samples'][sample_index]['id']}", 
                 style='SampleInfo.TLabel').pack(side="left", padx=(0, 20))
                 
        ttk.Label(info_frame, 
                 text=f"Resistance: {self.header_data['samples'][sample_index]['resistance']} Ω", 
                 style='SampleInfo.TLabel').pack(side="left", padx=(0, 20))
                 
        ttk.Label(info_frame, 
                 text=f"Voltage: {self.header_data['common']['voltage']} V", 
                 style='SampleInfo.TLabel').pack(side="left", padx=(0, 20))
        
        # Create a frame for the data table
        table_frame = ttk.Frame(parent_frame, style='TFrame')
        table_frame.pack(fill="both", expand=True)
        
        # Create the treeview (table)
        columns = ("puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", 
                           selectmode="browse", height=20)
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        y_scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=y_scrollbar.set)
        
        # Define column headings
        tree.heading("puffs", text="Puffs")
        tree.heading("before_weight", text="Before weight/g")
        tree.heading("after_weight", text="After weight/g")
        tree.heading("draw_pressure", text="Draw Pressure (kPa)")
        tree.heading("smell", text="Smell")
        tree.heading("notes", text="Notes")
        
        # Define column widths
        tree.column("puffs", width=80, anchor="center")
        tree.column("before_weight", width=120, anchor="center")
        tree.column("after_weight", width=120, anchor="center")
        tree.column("draw_pressure", width=120, anchor="center")
        tree.column("smell", width=80, anchor="center")
        tree.column("notes", width=150, anchor="w")
        
        tree.pack(fill="both", expand=True)
        
        # Add the initial row
        self.update_treeview(tree, sample_id)
        
        # Enable in-place editing with single click
        tree.bind("<ButtonRelease-1>", lambda event, tree=tree, sample_id=sample_id: 
                 self.on_cell_click(event, tree, sample_id))
        
        # Bind Tab key for quick navigation between cells
        tree.bind("<Tab>", lambda event, tree=tree, sample_id=sample_id: 
                 self.tab_to_next_cell(event, tree, sample_id))
        # Bind left/right arrow keys for cell navigation during data entry
        tree.bind("<Right>", lambda event, tree=tree, sample_id=sample_id: 
                 self.move_to_next_cell(event, tree, sample_id, "right"))
        tree.bind("<Left>", lambda event, tree=tree, sample_id=sample_id: 
        tree.bind("<Left>", lambda event, tree=tree, sample_id=sample_id: 
                 self.move_to_next_cell(event, tree, sample_id, "left"))
        
        # Store reference to the tree
        self.data[sample_id]["tree"] = tree
        
        return tree
    
    def create_tpm_stats_panel(self):
        """Create the TPM statistics panel on the right side."""
        # Clear existing widgets
        for widget in self.stats_frame.winfo_children():
            widget.destroy()
            
        # Create a title without background color
        title_label = ttk.Label(self.stats_frame, text="TPM Statistics", style='Stats.TLabel')
        title_label.pack(pady=(0, 20))
        
        # Create frame for statistics
        stats_container = ttk.Frame(self.stats_frame, style='TFrame')
        stats_container.pack(fill="both", expand=True)
        
        # Create individual stat frames for each sample
        self.tpm_labels = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            sample_name = self.header_data["samples"][i]["id"]
            
            # Create a frame for this sample - using standard ttk frames instead of LabelFrame
            sample_frame = ttk.Frame(stats_container, style='TFrame')
            sample_frame.pack(fill="x", pady=5, padx=10)
            
            # Add a header label
            ttk.Label(sample_frame, 
                     text=f"Sample {i+1}: {sample_name}",
                     style='SampleInfo.TLabel').pack(anchor="w", pady=(0, 5))
            
            # Add a separator
            ttk.Separator(sample_frame, orient="horizontal").pack(fill="x", pady=2)
            
            # Add TPM statistics
            stat_grid = ttk.Frame(sample_frame, style='TFrame')
            stat_grid.pack(fill="x", pady=5, padx=10)
            
            # Row 1: Average TPM
            ttk.Label(stat_grid, text="Average TPM:", style='TLabel').grid(row=0, column=0, sticky="w", pady=2)
            avg_tpm_label = ttk.Label(stat_grid, text="N/A", font=("Arial", 10, "bold"), style='TLabel')
            avg_tpm_label.grid(row=0, column=1, sticky="e", pady=2)
            
            # Row 2: Latest TPM
            ttk.Label(stat_grid, text="Latest TPM:", style='TLabel').grid(row=1, column=0, sticky="w", pady=2)
            latest_tpm_label = ttk.Label(stat_grid, text="N/A", font=("Arial", 10), style='TLabel')
            latest_tpm_label.grid(row=1, column=1, sticky="e", pady=2)
            
            # Row 3: Puff Count
            ttk.Label(stat_grid, text="Puff Count:", style='TLabel').grid(row=2, column=0, sticky="w", pady=2)
            puff_count_label = ttk.Label(stat_grid, text="0", style='TLabel')
            puff_count_label.grid(row=2, column=1, sticky="e", pady=2)
            
            # Store references to labels
            self.tpm_labels[sample_id] = {
                "avg_tpm": avg_tpm_label,
                "latest_tpm": latest_tpm_label,
                "puff_count": puff_count_label
            }
            
        # Update the statistics for the current sample
        self.update_stats_panel()
    
    def update_treeview(self, tree, sample_id):
        """Update the treeview with current data."""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Add rows from data
        for i in range(len(self.data[sample_id]["puffs"])):
            tree.insert("", "end", values=(
                self.data[sample_id]["puffs"][i],
                self.data[sample_id]["before_weight"][i],
                self.data[sample_id]["after_weight"][i],
                self.data[sample_id]["draw_pressure"][i],
                self.data[sample_id]["smell"][i],
                self.data[sample_id]["notes"][i]
            ))
    
    def on_cell_click(self, event, tree, sample_id):
        """Handle cell click for in-place editing."""
        # Get clicked region
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        # Get the item and column
        item = tree.identify("item", event.x, event.y)
        column = tree.identify("column", event.x, event.y)
        
        if not item or not column:
            return
        
        # Select the item
        tree.selection_set(item)
        
        # Get the column index
        column_idx = int(column[1:]) - 1
        
        # Don't allow editing the puffs column
        if column_idx == 0:  # Puffs column
            return
            
        # Get column name
        columns = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"]
        column_name = columns[column_idx]
        
        # Get the current value
        row_idx = tree.index(item)
        current_value = self.data[sample_id][column_name][row_idx] if row_idx < len(self.data[sample_id][column_name]) else ""
        
        # Create entry widget for editing
        self.edit_cell(tree, item, column, column_name, row_idx, sample_id, current_value)
    
    def edit_cell(self, tree, item, column, column_name, row_idx, sample_id, current_value):
        """Create an entry widget for editing a cell."""
        # Disable hotkeys during editing
        self.hotkeys_enabled = False
        self.navigation_keys_enabled = True
        # Get the bbox of the cell
        x, y, width, height = tree.bbox(item, column)
        
        # Create a frame for the entry
        frame = tk.Frame(tree, borderwidth=0, highlightthickness=1, highlightbackground="black")
        frame.place(x=x, y=y, width=width, height=height)
        
        # Create the entry widget
        entry = tk.Entry(frame, borderwidth=0)
        entry.pack(fill="both", expand=True)
        
        # Set current value
        if current_value:
            entry.insert(0, current_value)
        
        # Select all text
        entry.select_range(0, tk.END)
        
        # Focus the entry
        entry.focus_set()
        
        # Callback for when editing is done
        def on_entry_done(event=None):
            value = entry.get()
            
            # Update the data
            if row_idx < len(self.data[sample_id][column_name]):
                self.data[sample_id][column_name][row_idx] = value
            
            # Update the tree item
            values = list(tree.item(item, "values"))
            values[int(column[1:]) - 1] = value
            tree.item(item, values=values)
            
            # Calculate TPM if both weights are present
            if column_name in ["before_weight", "after_weight"]:
                self.calculate_tpm(sample_id)
                self.update_stats_panel()
            
            # Check if both before and after weight are filled
            # If so, add a new row if this is the last row
            if (row_idx == len(self.data[sample_id]["puffs"]) - 1 and
                column_name == "after_weight" and
                self.data[sample_id]["before_weight"][row_idx] and 
                self.data[sample_id]["after_weight"][row_idx]):
                
                # Add new row with next puff interval
                next_puff = self.data[sample_id]["puffs"][row_idx] + self.puff_interval
                self.data[sample_id]["puffs"].append(next_puff)
                
                # Set the before weight of the new row to the after weight of the current row
                after_weight = self.data[sample_id]["after_weight"][row_idx]
                self.data[sample_id]["before_weight"].append(after_weight)  # Auto-populate!
                
                # Add empty values for other columns
                self.data[sample_id]["after_weight"].append("")
                self.data[sample_id]["draw_pressure"].append("")
                self.data[sample_id]["smell"].append("")
                self.data[sample_id]["notes"].append("")
                self.data[sample_id]["tpm"].append(None)
                
                # Update the treeview
                self.update_treeview(tree, sample_id)
                
                # Select the after_weight cell of the new row
                new_item = tree.get_children()[-1]
                tree.selection_set(new_item)
                tree.focus(new_item)
                tree.see(new_item)
                
                # Re-enable hotkeys before simulating the click
                self.hotkeys_enabled = True
                
                # Simulate clicking on the "after_weight" cell
                col = "#3"  # after_weight column
                tree.event_generate("<Button-1>", x=tree.bbox(new_item, col)[0] + 5, 
                                  y=tree.bbox(new_item, col)[1] + 5)
                
                # Update the statistics panel
                self.update_stats_panel()
            else:
                # Re-enable hotkeys
                self.hotkeys_enabled = True
            
            # Destroy the entry widget
            frame.destroy()

            # Handle special keys for navigation if this was triggered by an arrow key
            if event and event.keysym in ("Right", "Left"):
                # Re-enable navigation after closing the entry
                self.navigation_keys_enabled = True
            
                # Move to the next/prev cell
                direction = "right" if event.keysym == "Right" else "left"
                frame.destroy()  # First destroy the frame
                self.move_to_next_cell(event, tree, sample_id, direction)
                return
        
        # Handle escape key to cancel editing but restore hotkeys
        def on_escape(event=None):
            self.hotkeys_enabled = True
            frame.destroy()
        
        # Bind events for entry completion
        entry.bind("<Return>", on_entry_done)
        entry.bind("<Tab>", lambda e: (on_entry_done(), self.tab_to_next_cell(e, tree, sample_id)))
        entry.bind("<Escape>", on_escape)
        entry.bind("<Right>", on_entry_done)
        entry.bind("<Left>", on_entry_done)
        
        # For focus out, we need to re-enable hotkeys
        entry.bind("<FocusOut>", on_entry_done)
    
    def move_to_next_cell(self, event, tree, sample_id, direction):
        """
        Move to the next/previous cell in the current row.
    
        Args:
            event: The key event
            tree: The treeview
            sample_id: The current sample ID
            direction: "right" or "left"
        """
        # Only respond if navigation keys are enabled
        if not self.navigation_keys_enabled:
            return "break"
        
        # Get current item and column
        item = tree.focus()
        if not item:
            return "break"
        
        column = tree.identify_column(tree.winfo_pointerx() - tree.winfo_rootx())
        if not column:
            return "break"
        
        # Get column index
        col_idx = int(column[1:]) - 1
    
        # Calculate new column index
        if direction == "right" and col_idx < 5:  # Not the last column
            new_col_idx = col_idx + 1
            new_col = f"#{new_col_idx + 1}"
        elif direction == "left" and col_idx > 0:  # Not the first column
            new_col_idx = col_idx - 1
            new_col = f"#{new_col_idx + 1}"
        else:
            return "break"  # At edge columns, do nothing
        
        # Select the new cell
        x, y, width, height = tree.bbox(item, new_col)
        tree.event_generate("<Button-1>", x=x+5, y=y+5)
    
        return "break"  # Prevent default behavior

    def tab_to_next_cell(self, event, tree, sample_id):
        """Handle Tab key to move to the next editable cell."""
        item = tree.focus()
        if not item:
            return "break"  # Prevent default tab behavior
            
        # Get current column
        column = tree.identify_column(tree.winfo_pointerx() - tree.winfo_rootx())
        if not column:
            return "break"
            
        # Get column index
        col_idx = int(column[1:]) - 1
        
        # Move to next column or first column of next row
        if col_idx < 5:  # Not the last column
            # Move to next column
            new_col_idx = col_idx + 1
            tree.event_generate("<Right>")
        else:
            # Get next item (row)
            items = tree.get_children()
            current_idx = items.index(item)
            
            if current_idx < len(items) - 1:
                # Move to first editable column of next row
                new_item = items[current_idx + 1]
                tree.selection_set(new_item)
                tree.focus(new_item)
                tree.see(new_item)
                
                # Simulate clicking on the appropriate cell (before_weight or after_weight)
                # If before_weight is empty, focus there, otherwise focus on after_weight
                row_idx = current_idx + 1
                if not self.data[sample_id]["before_weight"][row_idx]:
                    col = "#2"  # before_weight column
                else:
                    col = "#3"  # after_weight column
                
                tree.event_generate("<Button-1>", x=tree.bbox(new_item, col)[0] + 5, 
                                   y=tree.bbox(new_item, col)[1] + 5)
            else:
                # Last row, last column - go to next tab
                self.go_to_next_sample()
        
        return "break"  # Prevent default tab behavior
    
    def update_puff_interval(self):
        """Update the puff interval for future rows."""
        try:
            self.puff_interval = self.puff_interval_var.get()
        except:
            messagebox.showerror("Error", "Invalid puff interval. Please enter a positive number.")
            self.puff_interval_var.set(self.puff_interval)
    
    def go_to_previous_sample(self):
        """Navigate to the previous sample tab."""
        if not self.hotkeys_enabled:
            return
            
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab > 0:
            self.notebook.select(current_tab - 1)
        else:
            # Wrap around to last tab
            self.notebook.select(len(self.sample_frames) - 1)
    
    def go_to_next_sample(self):
        """Navigate to the next sample tab."""
        if not self.hotkeys_enabled:
            return
            
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab < len(self.sample_frames) - 1:
            self.notebook.select(current_tab + 1)
        else:
            # Wrap around to first tab
            self.notebook.select(0)
    
    def calculate_tpm(self, sample_id):
        """Calculate TPM for all rows with before and after weights."""
        for i in range(len(self.data[sample_id]["puffs"])):
            try:
                before_weight_str = self.data[sample_id]["before_weight"][i]
                after_weight_str = self.data[sample_id]["after_weight"][i]
                
                # Skip if either weight is missing
                if not before_weight_str or not after_weight_str:
                    continue
                    
                before_weight = float(before_weight_str)
                after_weight = float(after_weight_str)
                
                puff_interval = self.data[sample_id]["puffs"][i]
                
                # Calculate TPM
                if i > 0:
                    prev_puff = self.data[sample_id]["puffs"][i - 1]
                    puffs_in_interval = puff_interval - prev_puff
                else:
                    puffs_in_interval = puff_interval
                
                if puffs_in_interval > 0 and before_weight > after_weight:
                    tpm = (before_weight - after_weight) / puffs_in_interval
                    
                    # Ensure tpm list is long enough
                    while len(self.data[sample_id]["tpm"]) <= i:
                        self.data[sample_id]["tpm"].append(None)
                        
                    self.data[sample_id]["tpm"][i] = round(tpm, 6)
                    
            except (ValueError, TypeError, ZeroDivisionError) as e:
                print(f"Error calculating TPM: {e}")
    
    def update_stats_panel(self, event=None):
        """Update the TPM statistics panel based on current data."""
        # Update stats for all samples
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"
            
            # Calculate TPM values if needed
            self.calculate_tpm(sample_id)
            
            # Get TPM values (filtering out None values)
            tpm_values = [v for v in self.data[sample_id]["tpm"] if v is not None]
            
            # Update labels
            if tpm_values:
                avg_tpm = sum(tpm_values) / len(tpm_values)
                self.data[sample_id]["avg_tpm"] = avg_tpm
                self.tpm_labels[sample_id]["avg_tpm"].config(text=f"{avg_tpm:.6f}")
                self.tpm_labels[sample_id]["latest_tpm"].config(text=f"{tpm_values[-1]:.6f}")
                puff_count = self.data[sample_id]["puffs"][-1] if self.data[sample_id]["puffs"] else 0
                self.tpm_labels[sample_id]["puff_count"].config(text=str(puff_count))
            else:
                self.tpm_labels[sample_id]["avg_tpm"].config(text="N/A")
                self.tpm_labels[sample_id]["latest_tpm"].config(text="N/A")
                self.tpm_labels[sample_id]["puff_count"].config(text="0")
    
    def setup_event_handlers(self):
        """Set up event handlers for the window."""
        # Handle window close
        self.window.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # Set up hotkeys
        self.setup_hotkeys()

    def setup_hotkeys(self):
        """Set up keyboard shortcuts for navigation."""
        if not hasattr(self, 'hotkey_bindings'):
            self.hotkey_bindings = {}
        
        # Clear any existing bindings
        for key, binding_id in self.hotkey_bindings.items():
            self.window.unbind(key, binding_id)
        
        self.hotkey_bindings.clear()
        
        # Bind Enter key to go to next sample
        binding_id = self.window.bind("<Return>", lambda e: self.go_to_next_sample() if self.hotkeys_enabled else None)
        self.hotkey_bindings["<Return>"] = binding_id
        
        # Bind number keys to select tabs
        for i in range(1, 10):  # Bind keys 1-9
            if i <= self.num_samples:
                binding_id = self.window.bind(str(i), 
                                             lambda e, tab=i-1: self.notebook.select(tab) if self.hotkeys_enabled else None)
                self.hotkey_bindings[str(i)] = binding_id
    
    def save_data(self):
        """Save the collected data to the Excel file."""
        try:
            # Confirm save
            if not messagebox.askyesno("Confirm Save", "Save the collected data to the file?"):
                return
                
            # Ensure TPM values are calculated for all samples
            for i in range(self.num_samples):
                sample_id = f"Sample {i+1}"
                self.calculate_tpm(sample_id)
            
            # Load the workbook
            wb = openpyxl.load_workbook(self.file_path)
            
            # Get the sheet for this test
            if self.test_name not in wb.sheetnames:
                messagebox.showerror("Error", f"Sheet '{self.test_name}' not found in the file.")
                return
                
            ws = wb[self.test_name]
            
            # Define green fill for TPM cells
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            
            # For each sample, write the data
            for sample_idx in range(self.num_samples):
                sample_id = f"Sample {sample_idx+1}"
                
                # Calculate column offset (12 columns per sample)
                col_offset = sample_idx * 12
                
                # Write the puffs data starting at row 5
                for i, puff in enumerate(self.data[sample_id]["puffs"]):
                    row = i + 5  # Row 5 is the first data row
                    
                    # Puffs column (A + offset)
                    ws.cell(row=row, column=1 + col_offset, value=puff)
                    
                    # Before weight column (B + offset)
                    if self.data[sample_id]["before_weight"][i]:
                        try:
                            ws.cell(row=row, column=2 + col_offset, value=float(self.data[sample_id]["before_weight"][i]))
                        except:
                            ws.cell(row=row, column=2 + col_offset, value=self.data[sample_id]["before_weight"][i])
                    
                    # After weight column (C + offset)
                    if self.data[sample_id]["after_weight"][i]:
                        try:
                            ws.cell(row=row, column=3 + col_offset, value=float(self.data[sample_id]["after_weight"][i]))
                        except:
                            ws.cell(row=row, column=3 + col_offset, value=self.data[sample_id]["after_weight"][i])
                    
                    # Draw pressure column (D + offset)
                    if self.data[sample_id]["draw_pressure"][i]:
                        try:
                            ws.cell(row=row, column=4 + col_offset, value=float(self.data[sample_id]["draw_pressure"][i]))
                        except:
                            ws.cell(row=row, column=4 + col_offset, value=self.data[sample_id]["draw_pressure"][i])
                    
                    # Smell column (F + offset)
                    if self.data[sample_id]["smell"][i]:
                        try:
                            ws.cell(row=row, column=6 + col_offset, value=float(self.data[sample_id]["smell"][i]))
                        except:
                            ws.cell(row=row, column=6 + col_offset, value=self.data[sample_id]["smell"][i])
                    
                    # Notes column (H + offset)
                    if self.data[sample_id]["notes"][i]:
                        ws.cell(row=row, column=8 + col_offset, value=str(self.data[sample_id]["notes"][i]))
                    
                    # TPM column (I + offset) - if calculated
                    if i < len(self.data[sample_id]["tpm"]) and self.data[sample_id]["tpm"][i] is not None:
                        tpm_cell = ws.cell(row=row, column=9 + col_offset, value=float(self.data[sample_id]["tpm"][i]))
                        tpm_cell.fill = green_fill
            
            # Save the workbook
            wb.save(self.file_path)
            
            messagebox.showinfo("Success", "Data saved successfully.")
            
            # Close the window
            self.window.destroy()
            
            # Signal to load the file for viewing
            self.result = "load_file"
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the data: {str(e)}")
    
    def on_cancel(self):
        """Handle cancel button click or window close."""
        if messagebox.askyesno("Confirm", "Are you sure you want to cancel? All data will be lost."):
            self.result = "cancel"
            self.window.destroy()
    
    def show(self):
        """
        Show the window and wait for user input.
        
        Returns:
            str: "load_file" if data was saved and file should be loaded for viewing,
                 "cancel" if the user cancelled.
        """
        self.window.wait_window()
        return self.result
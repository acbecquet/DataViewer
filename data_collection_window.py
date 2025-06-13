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
        
        # Tracking variables for cell editing
        self.editing = False
        self.current_edit_widget = None
        
        # Set up keyboard shortcut flags
        self.hotkeys_enabled = True
        self.hotkey_bindings = {}
        
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
                             
        # Style for Treeview with gridlines
        self.style.configure('Treeview', 
                            showlines=True,
                            background='white',
                            fieldbackground='white')
        
        # Add gridlines to the Treeview cells
        self.style.configure('Treeview', rowheight=25)
        self.style.map('Treeview', background=[('selected', '#CCCCCC')])
        
        # Configure tag for focused items
        self.style.map('Treeview', 
                      background=[('selected', '#CCCCCC')],
                      foreground=[('selected', 'black')])
    
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
                           selectmode="browse", height=20, style='Treeview')
        
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
        
        # Enable cell editing with click
        tree.bind("<ButtonRelease-1>", lambda e: self.on_cell_click(e, tree, sample_id))
        
        # Bind keyboard navigation
        tree.bind("<Tab>", lambda e: self.handle_tab_key(e, tree, sample_id))
        tree.bind("<Left>", lambda e: self.handle_arrow_key(e, tree, sample_id, "left"))
        tree.bind("<Right>", lambda e: self.handle_arrow_key(e, tree, sample_id, "right"))
        
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
        """Handle clicking on a cell to edit it."""
        # Get the item and column clicked
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        item = tree.identify("item", event.x, event.y)
        column = tree.identify("column", event.x, event.y)
        
        if not item or not column:
            return
            
        # Get column index (1-based)
        col_idx = int(column[1:])
        
        # Skip puffs column (not editable)
        if col_idx == 1:
            return
            
        # Get row index
        row_idx = tree.index(item)
        
        # Get columns names for lookup
        columns = ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"]
        column_name = columns[col_idx - 1]  # Convert 1-based to 0-based for array index
        
        # Start editing this cell
        self.edit_cell(tree, item, column, row_idx, sample_id, column_name)
    
    def edit_cell(self, tree, item, column, row_idx, sample_id, column_name):
        """Create an entry widget for editing a cell."""
        # Cancel any existing edit
        self.end_editing()
        
        # Mark that we're editing
        self.editing = True
        self.hotkeys_enabled = False
        
        # Get the current value
        current_value = ""
        if row_idx < len(self.data[sample_id][column_name]):
            current_value = self.data[sample_id][column_name][row_idx]
        
        # Get the cell coordinates
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
        
        # Save references
        self.current_edit = {
            "frame": frame,
            "entry": entry,
            "tree": tree,
            "item": item,
            "column": column,
            "column_name": column_name,
            "row_idx": row_idx,
            "sample_id": sample_id
        }
        
        # Focus the entry
        entry.focus_set()
        
        # Bind events for the entry widget
        entry.bind("<Return>", self.finish_edit)
        entry.bind("<Tab>", self.handle_tab_in_edit)
        entry.bind("<Escape>", self.cancel_edit)
        entry.bind("<FocusOut>", self.finish_edit)
        entry.bind("<Left>", self.handle_arrow_in_edit)
        entry.bind("<Right>", self.handle_arrow_in_edit)
    
    def finish_edit(self, event=None):
        """Save the current edit and move to the next cell if needed."""
        if not self.editing or not hasattr(self, 'current_edit'):
            return
        
        value = self.current_edit["entry"].get()
        tree = self.current_edit["tree"]
        item = self.current_edit["item"]
        column = self.current_edit["column"]
        column_name = self.current_edit["column_name"]
        row_idx = self.current_edit["row_idx"]
        sample_id = self.current_edit["sample_id"]
    
        # Update data storage
        if row_idx < len(self.data[sample_id][column_name]):
            self.data[sample_id][column_name][row_idx] = value
    
        # Update the tree
        col_idx = int(column[1:]) - 1
        values = list(tree.item(item, "values"))
        values[col_idx] = value
        tree.item(item, values=values)
    
        # If we're in the after_weight column and we're at the last row, check if we need to add a new row
        if column_name == "after_weight" and row_idx == len(self.data[sample_id]["puffs"]) - 1:
            if self.data[sample_id]["before_weight"][row_idx] and self.data[sample_id]["after_weight"][row_idx]:
                # Add a new row with the next puff interval
                next_puff = self.data[sample_id]["puffs"][row_idx] + self.puff_interval
                self.data[sample_id]["puffs"].append(next_puff)
            
                # Set before_weight to the previous after_weight
                after_weight = self.data[sample_id]["after_weight"][row_idx]
                self.data[sample_id]["before_weight"].append(after_weight)
            
                # Add empty values for other columns
                self.data[sample_id]["after_weight"].append("")
                self.data[sample_id]["draw_pressure"].append("")
                self.data[sample_id]["smell"].append("")
                self.data[sample_id]["notes"].append("")
                self.data[sample_id]["tpm"].append(None)
            
                # Update the tree without changing focus
                self.update_treeview(tree, sample_id)
            
                # Make sure we re-select the current item to maintain focus
                tree.selection_set(item)
                tree.focus(item)
                tree.see(item)
            
                # Calculate TPM and update stats
                self.calculate_tpm(sample_id)
                self.update_stats_panel()
    
        # Calculate TPM if weight was changed
        if column_name in ["before_weight", "after_weight"]:
            self.calculate_tpm(sample_id)
            self.update_stats_panel()
    
        # Check if this was triggered by Tab or an arrow key
        move_to_next = False
        if event and event.keysym in ["Tab", "Right", "Left"]:
            move_to_next = True
    
        # End the current edit
        self.end_editing()
    
        # If triggered by navigation key, move focus
        if move_to_next:
            if event.keysym == "Tab":
                self.handle_tab_key(event, tree, sample_id)
            elif event.keysym == "Right":
                self.handle_arrow_key(event, tree, sample_id, "right")
            elif event.keysym == "Left":
                self.handle_arrow_key(event, tree, sample_id, "left")
    
    def handle_tab_in_edit(self, event):
        """Handle tab key pressed while editing a cell."""
        # Save the current edit
        self.finish_edit(event)
        return "break"  # Stop event propagation
    
    def handle_arrow_in_edit(self, event):
        """Handle arrow keys pressed while editing a cell."""
        # Only handle left/right arrows
        if event.keysym not in ["Left", "Right"]:
            return
            
        # Check if at beginning or end of text
        entry = self.current_edit["entry"]
        cursor_pos = entry.index(tk.INSERT)
        
        # If at beginning and pressing left, or at end and pressing right, navigate to next cell
        if (cursor_pos == 0 and event.keysym == "Left") or \
           (cursor_pos == len(entry.get()) and event.keysym == "Right"):
            self.finish_edit(event)
            return "break"  # Stop event propagation
    
    def cancel_edit(self, event=None):
        """Cancel the current edit without saving."""
        self.end_editing()
        return "break"  # Stop event propagation
    
    def end_editing(self):
        """Clean up editing widgets and state."""
        if not self.editing:
            return
            
        if hasattr(self, 'current_edit') and self.current_edit:
            if "frame" in self.current_edit:
                self.current_edit["frame"].destroy()
            self.current_edit = None
            
        self.editing = False
        self.hotkeys_enabled = True
    
    def handle_tab_key(self, event, tree, sample_id):
        """Handle Tab key press in the treeview."""
        if self.editing:
            return "break"  # Already handled in edit mode
            
        # Get the current selection
        item = tree.focus()
        if not item:
            # Select the first item and the first editable column
            if tree.get_children():
                item = tree.get_children()[0]
                tree.selection_set(item)
                tree.focus(item)
                self.edit_cell(tree, item, "#2", 0, sample_id, "before_weight")
            return "break"
            
        # Get the current column
        column = tree.identify_column(event.x) if event else None
        if not column:
            # Start at the first editable column
            column = "#2"
            
        # Get column index (1-based)
        col_idx = int(column[1:])
        
        # Get next column
        next_col_idx = col_idx + 1
        
        # If at the last column, move to the first editable column of the next row
        if next_col_idx > 6:  # We have 6 columns
            # Get the next item
            items = tree.get_children()
            idx = items.index(item)
            
            if idx < len(items) - 1:
                # Move to next row
                next_item = items[idx + 1]
                tree.selection_set(next_item)
                tree.focus(next_item)
                tree.see(next_item)
                
                # Edit the first editable column
                self.edit_cell(tree, next_item, "#2", idx + 1, sample_id, "before_weight")
            else:
                # At the last row, move to next sample
                self.go_to_next_sample()
        else:
            # Skip the puffs column
            if next_col_idx == 1:
                next_col_idx = 2
                
            # Edit the next column in the same row
            row_idx = tree.index(item)
            self.edit_cell(tree, item, f"#{next_col_idx}", row_idx, sample_id, 
                          ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][next_col_idx - 1])
            
        return "break"  # Stop event propagation
    
    def handle_arrow_key(self, event, tree, sample_id, direction):
        """Handle arrow key press in the treeview."""
        if self.editing or not self.hotkeys_enabled:
            return "break"  # Already handled in edit mode
            
        # Get the current selection
        item = tree.focus()
        if not item:
            return "break"
            
        # Get the current column
        column = tree.identify_column(event.x) if event else None
        if not column:
            # Default to first editable column
            column = "#2"
            
        # Get column index (1-based)
        col_idx = int(column[1:])
        
        if direction == "right":
            # Move to the next column
            next_col_idx = col_idx + 1
            if next_col_idx > 6:  # Last column
                return "break"
                
            # Skip the puffs column
            if next_col_idx == 1:
                next_col_idx = 2
                
            # Edit the cell
            row_idx = tree.index(item)
            self.edit_cell(tree, item, f"#{next_col_idx}", row_idx, sample_id, 
                          ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][next_col_idx - 1])
                          
        elif direction == "left":
            # Move to the previous column
            prev_col_idx = col_idx - 1
            if prev_col_idx < 2:  # Before first editable column
                return "break"
                
            # Edit the cell
            row_idx = tree.index(item)
            self.edit_cell(tree, item, f"#{prev_col_idx}", row_idx, sample_id, 
                          ["puffs", "before_weight", "after_weight", "draw_pressure", "smell", "notes"][prev_col_idx - 1])
        
        return "break"  # Stop event propagation
    
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
            # End any active editing
            self.end_editing()
            
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
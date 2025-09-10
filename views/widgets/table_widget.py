# views/widgets/table_widget.py
"""
views/widgets/table_widget.py
Data table display widget.
This will contain the table display logic from main_gui.py.
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
from typing import Optional, Dict, Any


class TableWidget:
    """Widget for displaying data tables."""
    
    def __init__(self, parent: tk.Widget):
        """Initialize the table widget."""
        self.parent = parent
        self.frame = ttk.Frame(parent)
        
        # Table components
        self.tree: Optional[ttk.Treeview] = None
        self.v_scrollbar: Optional[ttk.Scrollbar] = None
        self.h_scrollbar: Optional[ttk.Scrollbar] = None
        
        # Data
        self.current_data: Optional[pd.DataFrame] = None
        
        print("DEBUG: TableWidget initialized")
    
    def setup_widget(self):
        """Set up the table widget layout."""
        print("DEBUG: TableWidget setting up layout")
        
        # Create treeview with scrollbars
        self._create_table()
        
        print("DEBUG: TableWidget layout complete")
    
    def _create_table(self):
        """Create the data table with scrollbars."""
        # Create frame for table and scrollbars
        table_frame = ttk.Frame(self.frame)
        table_frame.pack(fill="both", expand=True)
        
        # Create treeview
        self.tree = ttk.Treeview(table_frame, show="tree headings")
        
        # Create scrollbars
        self.v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        
        # Configure treeview scrolling
        self.tree.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        
        # Pack components
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        # Configure grid weights
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        print("DEBUG: TableWidget table created")
    
    def display_data(self, data: pd.DataFrame, sheet_name: str = ""):
        """Display data in the table."""
        print(f"DEBUG: TableWidget displaying data for sheet: {sheet_name}")
        
        if not self.tree or data.empty:
            print("WARNING: TableWidget - no tree or empty data")
            return
        
        try:
            # Store current data
            self.current_data = data.copy()
            
            # Clear existing data
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Configure columns
            columns = list(data.columns)
            self.tree["columns"] = columns
            self.tree["show"] = "headings"
            
            # Set column headings and widths
            for col in columns:
                self.tree.heading(col, text=str(col))
                self.tree.column(col, width=100, minwidth=50)
            
            # Insert data rows
            for index, row in data.iterrows():
                values = [str(val) if pd.notna(val) else "" for val in row]
                self.tree.insert("", "end", values=values)
            
            print(f"DEBUG: TableWidget displayed {len(data)} rows, {len(columns)} columns")
            
        except Exception as e:
            print(f"ERROR: TableWidget failed to display data: {e}")
            self._show_error_message(f"Error displaying data: {e}")
    
    def _show_error_message(self, message: str):
        """Show error message in the table area."""
        if not self.tree:
            return
        
        # Clear table
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Set single column for error
        self.tree["columns"] = ["Error"]
        self.tree["show"] = "headings"
        self.tree.heading("Error", text="Error")
        self.tree.column("Error", width=400)
        
        # Insert error message
        self.tree.insert("", "end", values=[message])
    
    def clear_table(self):
        """Clear the table."""
        if self.tree:
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            self.tree["columns"] = []
            self.current_data = None
            
            print("DEBUG: TableWidget cleared")
    
    def export_data(self, file_path: str):
        """Export current data to file."""
        if not self.current_data is None:
            return False
        
        try:
            self.current_data.to_excel(file_path, index=False)
            print(f"DEBUG: TableWidget exported data to {file_path}")
            return True
        except Exception as e:
            print(f"ERROR: TableWidget failed to export data: {e}")
            return False
    
    def get_selected_data(self) -> Optional[Dict[str, Any]]:
        """Get data from selected row."""
        if not self.tree or not self.current_data is not None:
            return None
        
        selection = self.tree.selection()
        if not selection:
            return None
        
        try:
            # Get selected item
            item = selection[0]
            item_index = self.tree.index(item)
            
            # Get row data
            row_data = self.current_data.iloc[item_index].to_dict()
            
            print(f"DEBUG: TableWidget selected row {item_index}")
            return row_data
            
        except Exception as e:
            print(f"ERROR: TableWidget failed to get selected data: {e}")
            return None
    
    def filter_data(self, filter_text: str):
        """Filter displayed data based on text."""
        if not self.current_data is not None or not filter_text:
            return
        
        try:
            # Simple text-based filtering
            mask = self.current_data.astype(str).apply(
                lambda x: x.str.contains(filter_text, case=False, na=False)
            ).any(axis=1)
            
            filtered_data = self.current_data[mask]
            self.display_data(filtered_data)
            
            print(f"DEBUG: TableWidget filtered to {len(filtered_data)} rows")
            
        except Exception as e:
            print(f"ERROR: TableWidget failed to filter data: {e}")
    
    def get_widget(self) -> ttk.Frame:
        """Get the main widget frame."""
        return self.frame
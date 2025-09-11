# views/widgets/table_widget.py
"""
views/widgets/table_widget.py
Table widget for displaying data tables.
This contains table display functionality from main_gui.py display_table method.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List, Callable
import threading


class TableWidget:
    """Widget for displaying data tables."""
    
    def __init__(self, parent: tk.Widget, controller: Optional[Any] = None):
        """Initialize the table widget."""
        self.parent = parent
        self.controller = controller
        
        # Table components
        self.table_frame: Optional[ttk.Frame] = None
        self.table_canvas: Optional[Any] = None
        self.sheet_widget: Optional[Any] = None
        self.current_sheet_widgets: Dict[str, Any] = {}
        
        # Table configuration
        self.use_tksheet = True  # Prefer tksheet over tkintertable
        self.default_cellwidth = 120
        self.default_rowheight = 25
        self.font_config = ('Arial', 10)
        
        # Callbacks
        self.on_cell_clicked: Optional[Callable] = None
        self.on_cell_edited: Optional[Callable] = None
        self.on_view_raw_data: Optional[Callable] = None
        
        print("DEBUG: TableWidget initialized")
    
    def get_table_libraries(self):
        """Lazy load table libraries."""
        tksheet = None
        tkintertable = None
        
        try:
            from tksheet import Sheet
            tksheet = Sheet
            print("DEBUG: TableWidget - tksheet loaded")
        except ImportError:
            print("DEBUG: TableWidget - tksheet not available")
        
        try:
            from tkintertable import TableCanvas, TableModel
            tkintertable = (TableCanvas, TableModel)
            print("DEBUG: TableWidget - tkintertable loaded")
        except ImportError:
            print("DEBUG: TableWidget - tkintertable not available")
        
        return tksheet, tkintertable
    
    def display_table(self, frame: tk.Widget, data: Any, sheet_name: str, 
                     is_plotting_sheet: bool = False) -> bool:
        """Display data table in the specified frame."""
        try:
            print(f"DEBUG: TableWidget - displaying table for sheet: {sheet_name}")
            
            if not frame or not frame.winfo_exists():
                print(f"ERROR: TableWidget - frame does not exist")
                return False
            
            # Clear existing widgets
            for widget in frame.winfo_children():
                widget.destroy()
            
            # Validate and prepare data
            cleaned_data = self._prepare_data(data, sheet_name)
            if cleaned_data is None:
                return False
            
            # Choose table implementation
            if self.use_tksheet:
                success = self._create_tksheet_table(frame, cleaned_data, sheet_name, is_plotting_sheet)
            else:
                success = self._create_tkintertable_table(frame, cleaned_data, sheet_name, is_plotting_sheet)
            
            if success:
                print(f"DEBUG: TableWidget - table displayed successfully for {sheet_name}")
            else:
                self._show_error_message(frame, sheet_name, "Failed to create table")
            
            return success
            
        except Exception as e:
            print(f"ERROR: TableWidget - error displaying table: {e}")
            import traceback
            traceback.print_exc()
            self._show_error_message(frame, sheet_name, str(e))
            return False
    
    def _prepare_data(self, data: Any, sheet_name: str) -> Optional[Any]:
        """Prepare and validate data for display."""
        try:
            pd = self._get_pandas()
            if not pd:
                print("ERROR: TableWidget - pandas not available")
                return None
            
            if not isinstance(data, pd.DataFrame):
                print(f"ERROR: TableWidget - data is not a DataFrame: {type(data)}")
                return None
            
            # Clean column names
            data = self._clean_columns(data)
            data.columns = data.columns.map(str)
            
            # Check if data is empty
            if data.empty or len(data) == 0 or data.shape[0] == 0:
                print(f"DEBUG: TableWidget - data is empty for sheet: {sheet_name}")
                return pd.DataFrame([{"Status": "No data available", 
                                    "Instructions": "Use data collection tools to add information"}])
            
            # Convert to string and handle NaN values
            data = data.astype(str)
            data = data.replace(['nan', 'None', 'NaN', '<NA>'], '', regex=False)
            
            print(f"DEBUG: TableWidget - data prepared: shape {data.shape}")
            return data
            
        except Exception as e:
            print(f"ERROR: TableWidget - error preparing data: {e}")
            return None
    
    def _create_tksheet_table(self, frame: tk.Widget, data: Any, sheet_name: str, 
                            is_plotting_sheet: bool) -> bool:
        """Create table using tksheet library."""
        try:
            Sheet, _ = self.get_table_libraries()
            if not Sheet:
                print("DEBUG: TableWidget - tksheet not available, falling back to tkintertable")
                return self._create_tkintertable_table(frame, data, sheet_name, is_plotting_sheet)
            
            # Create table frame
            self.table_frame = ttk.Frame(frame, padding=(2, 1))
            self.table_frame.pack(fill='both', expand=True)
            
            # Configure grid weights
            self.table_frame.grid_rowconfigure(0, weight=1)
            self.table_frame.grid_columnconfigure(0, weight=1)
            
            # Convert DataFrame to list format for tksheet
            headers = list(data.columns)
            table_data = data.values.tolist()
            
            # Create sheet widget
            sheet = Sheet(self.table_frame,
                         column_headers=headers,
                         startup_select=(0, 1, "rows"),
                         headers_bg="#4CC9F0",
                         headers_fg="black", 
                         headers_font=self.font_config,
                         font=self.font_config,
                         table_bg="white",
                         table_fg="black")
            
            # Set data
            sheet.set_sheet_data(table_data)
            
            # Configure columns
            self._configure_tksheet_columns(sheet, data, is_plotting_sheet)
            
            # Disable editing and configure bindings
            sheet.disable_bindings("edit_cell", "paste", "delete", "insert_columns", 
                                 "insert_rows", "delete_columns", "delete_rows")
            
            sheet.enable_bindings("single_select", "drag_select", "select_all",
                                "column_select", "row_select", "column_width_resize",
                                "double_click_column_resize", "row_height_resize",
                                "arrowkeys", "right_click_popup_menu", "rc_select")
            
            # Grid the sheet
            sheet.grid(row=0, column=0, sticky="nsew")
            
            # Refresh and store reference
            sheet.refresh()
            self.current_sheet_widgets[sheet_name] = sheet
            self.sheet_widget = sheet
            
            print(f"DEBUG: TableWidget - tksheet table created successfully")
            return True
            
        except Exception as e:
            print(f"ERROR: TableWidget - error creating tksheet table: {e}")
            return False
    
    def _configure_tksheet_columns(self, sheet: Any, data: Any, is_plotting_sheet: bool):
        """Configure tksheet column widths and properties."""
        try:
            # Calculate column widths
            col_widths = []
            max_width = 200
            min_width = 80
            
            for i, col in enumerate(data.columns):
                # Calculate width based on content
                header_width = len(str(col)) * 8 + 20
                
                # Sample content width (check first few rows)
                content_width = min_width
                for j in range(min(5, len(data))):
                    try:
                        cell_content = str(data.iloc[j, i])
                        cell_width = len(cell_content) * 7 + 10
                        content_width = max(content_width, cell_width)
                    except:
                        continue
                
                # Final width calculation
                final_width = max(header_width, content_width)
                final_width = min(final_width, max_width)
                final_width = max(final_width, min_width)
                
                col_widths.append(final_width)
                
                # Set column width
                sheet.column_width(column=i, width=int(final_width))
            
            # Configure row heights for plotting sheets
            if is_plotting_sheet:
                sheet.set_all_row_heights(height=35)
            else:
                sheet.set_all_row_heights(height=self.default_rowheight)
            
            print(f"DEBUG: TableWidget - configured {len(col_widths)} columns")
            
        except Exception as e:
            print(f"ERROR: TableWidget - error configuring columns: {e}")
    
    def _create_tkintertable_table(self, frame: tk.Widget, data: Any, sheet_name: str, 
                                 is_plotting_sheet: bool) -> bool:
        """Create table using tkintertable library."""
        try:
            _, tkintertable_libs = self.get_table_libraries()
            if not tkintertable_libs:
                print("ERROR: TableWidget - no table libraries available")
                return False
            
            TableCanvas, TableModel = tkintertable_libs
            
            # Create table frame
            self.table_frame = ttk.Frame(frame, padding=(2, 1))
            self.table_frame.pack(fill='both', expand=True)
            
            # Calculate dimensions
            frame.update_idletasks()
            available_width = self.table_frame.winfo_width()
            num_columns = len(data.columns)
            calculated_cellwidth = max(self.default_cellwidth, available_width // num_columns)
            
            # Create scrollbars
            v_scrollbar = ttk.Scrollbar(self.table_frame, orient='vertical')
            h_scrollbar = ttk.Scrollbar(self.table_frame, orient='horizontal')
            
            # Create model and import data
            model = TableModel()
            table_data_dict = data.to_dict(orient='index')
            model.importDict(table_data_dict)
            
            # Calculate row height for plotting sheets
            row_height = self._calculate_row_height(data, is_plotting_sheet)
            
            # Create table canvas
            self.table_canvas = TableCanvas(
                self.table_frame, 
                model=model,
                cellwidth=calculated_cellwidth,
                cellbackgr='#4CC9F0',
                thefont=self.font_config,
                rowheight=row_height,
                rowselectedcolor='#FFFFFF',
                editable=False,
                yscrollcommand=v_scrollbar.set,
                xscrollcommand=h_scrollbar.set,
                showGrid=True
            )
            
            # Grid components
            self.table_canvas.grid(row=0, column=0, sticky='nsew')
            v_scrollbar.grid(row=0, column=1, sticky='ns')
            h_scrollbar.grid(row=1, column=0, sticky='ew')
            
            # Configure grid weights
            self.table_frame.grid_rowconfigure(0, weight=1)
            self.table_frame.grid_columnconfigure(0, weight=1)
            
            # Show table
            self.table_canvas.show()
            
            print(f"DEBUG: TableWidget - tkintertable table created successfully")
            return True
            
        except Exception as e:
            print(f"ERROR: TableWidget - error creating tkintertable table: {e}")
            return False
    
    def _calculate_row_height(self, data: Any, is_plotting_sheet: bool) -> int:
        """Calculate appropriate row height based on content."""
        try:
            if is_plotting_sheet:
                # Calculate based on content length for plotting sheets
                font_height = 16
                char_per_line = 12
                
                row_heights = []
                for _, row in data.iterrows():
                    max_lines = 1
                    for cell in row:
                        if cell:
                            cell_length = len(str(cell))
                            lines = (cell_length // char_per_line) + 1
                            max_lines = max(max_lines, lines)
                    row_heights.append(max_lines * font_height)
                
                return max(row_heights) if row_heights else self.default_rowheight
            else:
                return self.default_rowheight
                
        except Exception as e:
            print(f"ERROR: TableWidget - error calculating row height: {e}")
            return self.default_rowheight
    
    def _show_error_message(self, frame: tk.Widget, sheet_name: str, error_msg: str):
        """Show error message when table creation fails."""
        try:
            for widget in frame.winfo_children():
                widget.destroy()
            
            error_label = tk.Label(
                frame,
                text=f"Error displaying table for '{sheet_name}'.\n{error_msg}\nPlease check the console for details.",
                font=("Arial", 12),
                fg="red",
                justify="center"
            )
            error_label.pack(expand=True, pady=50)
            
            print(f"DEBUG: TableWidget - error message shown for {sheet_name}")
            
        except Exception as e:
            print(f"ERROR: TableWidget - error showing error message: {e}")
    
    def _show_empty_message(self, frame: tk.Widget):
        """Show message when no data is available."""
        try:
            empty_label = tk.Label(
                frame,
                text="This sheet is empty.",
                font=("Arial", 14),
                fg="red"
            )
            empty_label.pack(anchor="center", pady=20)
            
            print("DEBUG: TableWidget - empty message shown")
            
        except Exception as e:
            print(f"ERROR: TableWidget - error showing empty message: {e}")
    
    # Utility methods
    def _get_pandas(self):
        """Get pandas with lazy loading."""
        try:
            import pandas as pd
            import numpy as np
            return pd
        except ImportError:
            return None
    
    def _clean_columns(self, data: Any):
        """Clean duplicate column names."""
        try:
            # Simple column cleaning - could be enhanced
            cols = data.columns.tolist()
            seen = {}
            for i, col in enumerate(cols):
                if col in seen:
                    seen[col] += 1
                    cols[i] = f"{col}_{seen[col]}"
                else:
                    seen[col] = 0
            
            data.columns = cols
            return data
            
        except Exception as e:
            print(f"ERROR: TableWidget - error cleaning columns: {e}")
            return data
    
    # Public interface methods
    def get_current_sheet_widget(self, sheet_name: str = None) -> Optional[Any]:
        """Get the current sheet widget."""
        if sheet_name and sheet_name in self.current_sheet_widgets:
            return self.current_sheet_widgets[sheet_name]
        return self.sheet_widget
    
    def clear_table(self, frame: tk.Widget = None):
        """Clear the current table."""
        try:
            target_frame = frame or self.parent
            if target_frame and target_frame.winfo_exists():
                for widget in target_frame.winfo_children():
                    widget.destroy()
            
            self.table_canvas = None
            self.sheet_widget = None
            
            print("DEBUG: TableWidget - table cleared")
            
        except Exception as e:
            print(f"ERROR: TableWidget - error clearing table: {e}")
    
    def export_table_data(self, sheet_name: str = None) -> Optional[Any]:
        """Export current table data."""
        try:
            widget = self.get_current_sheet_widget(sheet_name)
            if not widget:
                return None
            
            # Implementation depends on which library is being used
            if hasattr(widget, 'get_sheet_data'):  # tksheet
                return widget.get_sheet_data()
            elif hasattr(widget, 'model'):  # tkintertable
                return widget.model.getAllCells()
            
            return None
            
        except Exception as e:
            print(f"ERROR: TableWidget - error exporting table data: {e}")
            return None
    
    def refresh_table(self):
        """Refresh the current table display."""
        try:
            if self.sheet_widget and hasattr(self.sheet_widget, 'refresh'):
                self.sheet_widget.refresh()
                print("DEBUG: TableWidget - table refreshed")
            elif self.table_canvas and hasattr(self.table_canvas, 'redraw'):
                self.table_canvas.redraw()
                print("DEBUG: TableWidget - table redrawn")
        except Exception as e:
            print(f"ERROR: TableWidget - error refreshing table: {e}")
    
    def set_table_font(self, font_config: tuple):
        """Set table font configuration."""
        self.font_config = font_config
        print(f"DEBUG: TableWidget - font set to: {font_config}")
    
    def set_cell_width(self, width: int):
        """Set default cell width."""
        self.default_cellwidth = width
        print(f"DEBUG: TableWidget - cell width set to: {width}")
    
    def set_row_height(self, height: int):
        """Set default row height."""
        self.default_rowheight = height
        print(f"DEBUG: TableWidget - row height set to: {height}")
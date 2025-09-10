"""
models/data_model.py
Core data models for sheet and file data structures.
These models will replace the data structures currently in main_gui.py.
"""

from typing import Dict, List, Optional, Any
import pandas as pd
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class SheetData:
    """Model for individual sheet data and metadata."""
    name: str
    data: pd.DataFrame
    processed_data: Optional[pd.DataFrame] = None
    is_plotting_sheet: bool = False
    sheet_type: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    created_at: datetime = field(default_factory=datetime.now)
    modified_at: Optional[datetime] = None
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created SheetData for '{self.name}' with {len(self.data)} rows")
        if self.data.empty:
            print(f"WARNING: SheetData '{self.name}' has empty data")
    
    def update_data(self, new_data: pd.DataFrame):
        """Update the sheet data and mark as modified."""
        self.data = new_data
        self.modified_at = datetime.now()
        print(f"DEBUG: Updated SheetData '{self.name}' with {len(new_data)} rows at {self.modified_at}")
    
    def set_processed_data(self, processed_data: pd.DataFrame):
        """Set the processed data for this sheet."""
        self.processed_data = processed_data
        print(f"DEBUG: Set processed data for '{self.name}' with {len(processed_data)} rows")
    
    def is_empty(self) -> bool:
        """Check if the sheet data is empty."""
        return self.data.empty if self.data is not None else True
    
    def get_row_count(self) -> int:
        """Get the number of rows in the sheet."""
        return len(self.data) if self.data is not None else 0
    
    def get_column_count(self) -> int:
        """Get the number of columns in the sheet."""
        return len(self.data.columns) if self.data is not None else 0


@dataclass 
class FileData:
    """Model for file-level data and metadata."""
    filename: str
    filepath: str
    sheets: Dict[str, SheetData] = field(default_factory=dict)
    filtered_sheets: Dict[str, SheetData] = field(default_factory=dict)
    full_sample_data: Optional[pd.DataFrame] = None
    file_type: str = "excel"  # excel, vap3, etc.
    is_modified: bool = False
    created_at: datetime = field(default_factory=datetime.now)
    modified_at: Optional[datetime] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created FileData for '{self.filename}' with {len(self.sheets)} sheets")
        print(f"DEBUG: File path: {self.filepath}")
        print(f"DEBUG: File type: {self.file_type}")
    
    def add_sheet(self, sheet_name: str, sheet_data: SheetData):
        """Add a sheet to this file."""
        self.sheets[sheet_name] = sheet_data
        self.modified_at = datetime.now()
        print(f"DEBUG: Added sheet '{sheet_name}' to file '{self.filename}'")
    
    def add_filtered_sheet(self, sheet_name: str, sheet_data: SheetData):
        """Add a filtered sheet to this file."""
        self.filtered_sheets[sheet_name] = sheet_data
        self.modified_at = datetime.now()
        print(f"DEBUG: Added filtered sheet '{sheet_name}' to file '{self.filename}'")
    
    def mark_modified(self):
        """Mark the file as modified."""
        self.is_modified = True
        self.modified_at = datetime.now()
        print(f"DEBUG: Marked file '{self.filename}' as modified at {self.modified_at}")
    
    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names."""
        return list(self.sheets.keys())
    
    def get_filtered_sheet_names(self) -> List[str]:
        """Get list of filtered sheet names."""
        return list(self.filtered_sheets.keys())
    
    def has_sheet(self, sheet_name: str) -> bool:
        """Check if file has a specific sheet."""
        return sheet_name in self.sheets
    
    def has_filtered_sheet(self, sheet_name: str) -> bool:
        """Check if file has a specific filtered sheet.""" 
        return sheet_name in self.filtered_sheets
    
    def get_sheet(self, sheet_name: str) -> Optional[SheetData]:
        """Get a specific sheet by name."""
        return self.sheets.get(sheet_name)
    
    def get_filtered_sheet(self, sheet_name: str) -> Optional[SheetData]:
        """Get a specific filtered sheet by name."""
        return self.filtered_sheets.get(sheet_name)


class DataModel:
    """Main data model that manages all file and sheet data."""
    
    def __init__(self):
        """Initialize the data model."""
        self.current_file: Optional[FileData] = None
        self.all_files: List[FileData] = []
        self.selected_sheet_name: Optional[str] = None
        self.selected_plot_type: str = "TPM"
        
        print("DEBUG: DataModel initialized")
        print("DEBUG: Ready to manage file and sheet data")
    
    def add_file(self, file_data: FileData) -> None:
        """Add a file to the data model."""
        self.all_files.append(file_data)
        print(f"DEBUG: Added file '{file_data.filename}' to DataModel")
        print(f"DEBUG: Total files in model: {len(self.all_files)}")
    
    def set_current_file(self, file_data: FileData) -> None:
        """Set the current active file."""
        self.current_file = file_data
        print(f"DEBUG: Set current file to '{file_data.filename}'")
    
    def get_current_file(self) -> Optional[FileData]:
        """Get the current active file."""
        return self.current_file
    
    def get_all_files(self) -> List[FileData]:
        """Get all files in the model."""
        return self.all_files
    
    def find_file_by_name(self, filename: str) -> Optional[FileData]:
        """Find a file by its filename."""
        for file_data in self.all_files:
            if file_data.filename == filename:
                return file_data
        return None
    
    def remove_file(self, filename: str) -> bool:
        """Remove a file from the model."""
        for i, file_data in enumerate(self.all_files):
            if file_data.filename == filename:
                del self.all_files[i]
                if self.current_file == file_data:
                    self.current_file = None
                print(f"DEBUG: Removed file '{filename}' from DataModel")
                return True
        print(f"WARNING: File '{filename}' not found for removal")
        return False
    
    def clear_all_files(self) -> None:
        """Clear all files from the model."""
        file_count = len(self.all_files)
        self.all_files.clear()
        self.current_file = None
        self.selected_sheet_name = None
        print(f"DEBUG: Cleared {file_count} files from DataModel")
    
    def get_current_sheets(self) -> Dict[str, SheetData]:
        """Get sheets from the current file."""
        if self.current_file:
            return self.current_file.sheets
        return {}
    
    def get_current_filtered_sheets(self) -> Dict[str, SheetData]:
        """Get filtered sheets from the current file."""
        if self.current_file:
            return self.current_file.filtered_sheets
        return {}
    
    def set_selected_sheet(self, sheet_name: str) -> None:
        """Set the currently selected sheet."""
        self.selected_sheet_name = sheet_name
        print(f"DEBUG: Set selected sheet to '{sheet_name}'")
    
    def get_selected_sheet(self) -> Optional[SheetData]:
        """Get the currently selected sheet data."""
        if self.current_file and self.selected_sheet_name:
            return self.current_file.get_filtered_sheet(self.selected_sheet_name)
        return None
    
    def set_plot_type(self, plot_type: str) -> None:
        """Set the selected plot type."""
        self.selected_plot_type = plot_type
        print(f"DEBUG: Set plot type to '{plot_type}'")
    
    def get_plot_type(self) -> str:
        """Get the selected plot type."""
        return self.selected_plot_type
    
    def get_stats(self) -> Dict[str, Any]:
        """Get statistics about the data model."""
        stats = {
            'total_files': len(self.all_files),
            'current_file': self.current_file.filename if self.current_file else None,
            'selected_sheet': self.selected_sheet_name,
            'selected_plot_type': self.selected_plot_type
        }
        
        if self.current_file:
            stats.update({
                'current_file_sheets': len(self.current_file.sheets),
                'current_file_filtered_sheets': len(self.current_file.filtered_sheets),
                'current_file_modified': self.current_file.is_modified
            })
        
        print(f"DEBUG: DataModel stats: {stats}")
        return stats
# services/processing_service.py
"""
services/processing_service.py
Core data processing service.
This will contain the processing logic from processing.py.
"""

from typing import Optional, Dict, Any, List, Tuple, Callable
import pandas as pd
import numpy as np


class ProcessingService:
    """Service for data processing operations."""
    
    def __init__(self):
        """Initialize the processing service."""
        self.processing_functions: Dict[str, Callable] = {}
        self.valid_plot_options: Dict[str, List[str]] = {}
        
        # Initialize default processing functions
        self._initialize_default_functions()
        
        print("DEBUG: ProcessingService initialized")
        print(f"DEBUG: Registered {len(self.processing_functions)} processing functions")
    
    def _initialize_default_functions(self):
        """Initialize default processing functions."""
        # Register default processing functions
        self.register_processing_function("default", self._default_processing)
        self.register_processing_function("TPM", self._tpm_processing)
        self.register_processing_function("resistance", self._resistance_processing)
        self.register_processing_function("draw_pressure", self._draw_pressure_processing)
        
        print("DEBUG: ProcessingService initialized default functions")
    
    def register_processing_function(self, sheet_type: str, function: Callable):
        """Register a processing function for a sheet type."""
        self.processing_functions[sheet_type.lower()] = function
        print(f"DEBUG: ProcessingService registered function for '{sheet_type}'")
    
    def process_sheet_data(self, raw_data: pd.DataFrame, sheet_name: str) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process raw sheet data based on sheet type."""
        print(f"DEBUG: ProcessingService processing sheet: {sheet_name}")
        
        try:
            if raw_data.empty:
                print("WARNING: Empty data provided for processing")
                return pd.DataFrame(), {}, pd.DataFrame()
            
            # Determine processing function
            processing_func = self._get_processing_function(sheet_name)
            
            # Process the data
            processed_data, metadata, full_sample_data = processing_func(raw_data)
            
            print(f"DEBUG: ProcessingService processed {sheet_name} - "
                  f"output: {len(processed_data)} rows, {len(processed_data.columns)} columns")
            
            return processed_data, metadata, full_sample_data
            
        except Exception as e:
            error_msg = f"Failed to process sheet {sheet_name}: {e}"
            print(f"ERROR: ProcessingService - {error_msg}")
            return pd.DataFrame(), {'error': error_msg}, pd.DataFrame()
    
    def _get_processing_function(self, sheet_name: str) -> Callable:
        """Get the appropriate processing function for a sheet."""
        sheet_lower = sheet_name.lower()
        
        # Check for exact match first
        if sheet_lower in self.processing_functions:
            return self.processing_functions[sheet_lower]
        
        # Check for partial matches
        for key, func in self.processing_functions.items():
            if key in sheet_lower or any(keyword in sheet_lower for keyword in key.split()):
                return func
        
        # Default to generic processing
        return self.processing_functions.get('default', self._default_processing)
    
    def _default_processing(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Default processing for unknown sheet types."""
        print("DEBUG: ProcessingService applying default processing")
        
        try:
            # Clean column names
            processed_data = data.copy()
            processed_data.columns = [str(col).strip() for col in processed_data.columns]
            
            # Remove empty rows and columns
            processed_data = processed_data.dropna(how='all')  # Remove empty rows
            processed_data = processed_data.loc[:, processed_data.notna().any()]  # Remove empty columns
            
            metadata = {
                'processing_type': 'default',
                'original_shape': data.shape,
                'processed_shape': processed_data.shape,
                'columns': list(processed_data.columns)
            }
            
            return processed_data, metadata, processed_data
            
        except Exception as e:
            print(f"ERROR: Default processing failed: {e}")
            return data, {'error': str(e)}, data
    
    def _tpm_processing(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Processing for TPM (Total Particulate Matter) data."""
        print("DEBUG: ProcessingService applying TPM processing")
        
        try:
            processed_data = data.copy()
            
            # TPM-specific processing logic would go here
            # For example: calculate averages, remove outliers, format columns
            
            metadata = {
                'processing_type': 'TPM',
                'original_shape': data.shape,
                'processed_shape': processed_data.shape,
                'plot_options': ['TPM', 'TPM (Bar)']
            }
            
            return processed_data, metadata, processed_data
            
        except Exception as e:
            print(f"ERROR: TPM processing failed: {e}")
            return self._default_processing(data)
    
    def _resistance_processing(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Processing for resistance measurement data."""
        print("DEBUG: ProcessingService applying resistance processing")
        
        try:
            processed_data = data.copy()
            
            # Resistance-specific processing logic
            
            metadata = {
                'processing_type': 'resistance',
                'original_shape': data.shape,
                'processed_shape': processed_data.shape,
                'plot_options': ['Resistance', 'Draw Pressure']
            }
            
            return processed_data, metadata, processed_data
            
        except Exception as e:
            print(f"ERROR: Resistance processing failed: {e}")
            return self._default_processing(data)
    
    def _draw_pressure_processing(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Processing for draw pressure data."""
        print("DEBUG: ProcessingService applying draw pressure processing")
        
        try:
            processed_data = data.copy()
            
            # Draw pressure-specific processing logic
            
            metadata = {
                'processing_type': 'draw_pressure',
                'original_shape': data.shape,
                'processed_shape': processed_data.shape,
                'plot_options': ['Draw Pressure', 'Power Efficiency']
            }
            
            return processed_data, metadata, processed_data
            
        except Exception as e:
            print(f"ERROR: Draw pressure processing failed: {e}")
            return self._default_processing(data)
    
    def get_valid_plot_options(self, sheet_name: str) -> List[str]:
        """Get valid plot options for a sheet type."""
        # Process a small sample to get metadata
        try:
            sample_data = pd.DataFrame({'col1': [1], 'col2': [2]})  # Minimal sample
            _, metadata, _ = self.process_sheet_data(sample_data, sheet_name)
            return metadata.get('plot_options', ['TPM', 'Draw Pressure', 'Resistance'])
        except:
            return ['TPM', 'Draw Pressure', 'Resistance', 'Power Efficiency']
    
    def validate_data(self, data: pd.DataFrame) -> Tuple[bool, List[str]]:
        """Validate data quality and return issues."""
        issues = []
        
        if data.empty:
            issues.append("Data is empty")
            return False, issues
        
        if data.shape[0] < 2:
            issues.append("Insufficient rows (minimum 2 required)")
        
        if data.shape[1] < 2:
            issues.append("Insufficient columns (minimum 2 required)")
        
        # Check for all-null columns
        null_columns = data.columns[data.isnull().all()].tolist()
        if null_columns:
            issues.append(f"Columns with all null values: {', '.join(null_columns)}")
        
        # Check data types
        non_numeric_columns = []
        for col in data.columns:
            if not pd.api.types.is_numeric_dtype(data[col]):
                # Try to convert to numeric
                try:
                    pd.to_numeric(data[col], errors='raise')
                except:
                    non_numeric_columns.append(col)
        
        if non_numeric_columns:
            issues.append(f"Non-numeric columns: {', '.join(non_numeric_columns)}")
        
        is_valid = len(issues) == 0
        print(f"DEBUG: ProcessingService validation - Valid: {is_valid}, Issues: {len(issues)}")
        
        return is_valid, issues
# models/data_model.py
"""
models/data_model.py
Core data models and calculation services for the DataViewer application.
This consolidates data structures, TPM calculations, statistical analysis, viscosity calculations,
and data processing logic from processing.py, calculation_service.py, and viscosity_calculator.
"""

import os
import json
import re
import math
import statistics
import copy
from typing import Dict, List, Optional, Any, Tuple, Union
import pandas as pd
import numpy as np
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


def debug_print(message: str):
    """Debug print function for data model operations."""
    print(f"DEBUG: DataModel - {message}")


def round_values(value: float, decimal_places: int = 3) -> float:
    """Round values for display consistency."""
    try:
        return round(float(value), decimal_places)
    except (ValueError, TypeError):
        return 0.0


# ===================== CORE DATA STRUCTURES =====================

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


# ===================== MAIN DATA MODEL WITH CALCULATIONS =====================

class DataModel:
    """Main data model with integrated calculation services."""
    
    def __init__(self):
        """Initialize the data model with calculation capabilities."""
        debug_print("Initializing DataModel with calculation services")
        
        # Core data management
        self.current_file: Optional[FileData] = None
        self.all_files: List[FileData] = []
        self.selected_sheet_name: Optional[str] = None
        self.selected_plot_type: str = "TPM"
        
        # TPM calculation parameters
        self.default_puff_time = 3.0  # Default puff time in seconds
        self.default_puffs = 10  # Default puff count for invalid intervals
        
        # Viscosity calculation data
        self.formulation_database = {}
        self.viscosity_models = {}
        self.consolidated_models = {}
        self.base_models = {}
        
        # Statistical calculation settings
        self.precision_digits = 3
        self.statistical_threshold = 2  # Minimum data points for statistics
        
        # Data processing parameters
        self.sample_columns_per_sample = 12
        self.headers_row = 3
        self.data_start_row = 4
        
        # Load existing databases if available
        self._load_formulation_database()
        self._load_viscosity_models()
        
        debug_print("DataModel initialized successfully")
        debug_print(f"Loaded {len(self.formulation_database)} formulation entries")
        debug_print(f"Loaded {len(self.viscosity_models)} viscosity models")
    
    # ===================== CORE DATA MANAGEMENT =====================
    
    def add_file(self, file_data: FileData) -> None:
        """Add a file to the data model."""
        self.all_files.append(file_data)
        debug_print(f"Added file '{file_data.filename}' to DataModel")
        debug_print(f"Total files in model: {len(self.all_files)}")
    
    def set_current_file(self, file_data: FileData) -> None:
        """Set the current active file."""
        self.current_file = file_data
        debug_print(f"Set current file to '{file_data.filename}'")
    
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
                debug_print(f"Removed file '{filename}' from DataModel")
                return True
        debug_print(f"WARNING: File '{filename}' not found for removal")
        return False
    
    def clear_all_files(self) -> None:
        """Clear all files from the model."""
        file_count = len(self.all_files)
        self.all_files.clear()
        self.current_file = None
        self.selected_sheet_name = None
        debug_print(f"Cleared {file_count} files from DataModel")
    
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
        debug_print(f"Set selected sheet to '{sheet_name}'")
    
    def get_selected_sheet(self) -> Optional[SheetData]:
        """Get the currently selected sheet data."""
        if self.current_file and self.selected_sheet_name:
            return self.current_file.get_sheet(self.selected_sheet_name)
        return None
    
    # ===================== TPM CALCULATIONS =====================
    
    def calculate_tpm_from_weights(self, puffs: pd.Series, before_weights: pd.Series, 
                                   after_weights: pd.Series, test_type: str = "standard") -> pd.Series:
        """
        Calculate TPM from weight differences and puffing intervals.
        Consolidated from processing.py TPM calculation logic.
        """
        debug_print("Calculating TPM from weight differences")
        
        try:
            # Ensure all series are numeric and aligned
            puffs_numeric = pd.to_numeric(puffs, errors='coerce')
            before_numeric = pd.to_numeric(before_weights, errors='coerce')  
            after_numeric = pd.to_numeric(after_weights, errors='coerce')
            
            # Calculate weight differences
            weight_diff = before_numeric - after_numeric
            
            # Calculate puff intervals (difference between consecutive puff counts)
            puff_intervals = puffs_numeric.diff().fillna(puffs_numeric.iloc[0] if len(puffs_numeric) > 0 else self.default_puffs)
            
            # Handle invalid intervals
            puff_intervals = puff_intervals.replace(0, self.default_puffs)
            puff_intervals = puff_intervals.fillna(self.default_puffs)
            
            # Calculate TPM = (weight difference / puff interval) * 1000
            tpm_values = (weight_diff / puff_intervals) * 1000
            
            # Clean up invalid values
            tpm_values = tpm_values.replace([np.inf, -np.inf], np.nan)
            
            debug_print(f"Calculated TPM for {len(tpm_values)} data points")
            return tpm_values
            
        except Exception as e:
            debug_print(f"Error calculating TPM: {e}")
            return pd.Series(dtype=float)
    
    def calculate_normalized_tpm(self, tpm_data: pd.Series, puffing_regime: str = None) -> pd.Series:
        """
        Calculate normalized TPM by dividing TPM by puff time.
        Consolidated from processing.py normalization logic.
        """
        debug_print("Calculating normalized TPM")
        
        try:
            # Extract puff time from puffing regime
            puff_time = self._extract_puff_time_from_regime(puffing_regime)
            
            if puff_time and puff_time > 0:
                normalized_tpm = tpm_data / puff_time
                debug_print(f"Normalized TPM using puff time: {puff_time}s")
                return normalized_tpm
            else:
                debug_print("No valid puff time found, returning original TPM")
                return tpm_data
                
        except Exception as e:
            debug_print(f"Error calculating normalized TPM: {e}")
            return tpm_data
    
    def _extract_puff_time_from_regime(self, puffing_regime: str) -> Optional[float]:
        """Extract puff time from puffing regime string."""
        if not puffing_regime or pd.isna(puffing_regime):
            return self.default_puff_time
            
        try:
            # Pattern to match "mL/X" format where X is the puff time
            pattern = r'mL/(\d+(?:\.\d+)?)'
            match = re.search(pattern, str(puffing_regime))
            
            if match:
                puff_time = float(match.group(1))
                debug_print(f"Extracted puff time: {puff_time}s from '{puffing_regime}'")
                return puff_time
            else:
                debug_print(f"No puff time pattern found in '{puffing_regime}', using default")
                return self.default_puff_time
                
        except Exception as e:
            debug_print(f"Error extracting puff time: {e}")
            return self.default_puff_time
    
    # ===================== STATISTICAL CALCULATIONS =====================
    
    def calculate_statistics(self, data: Union[List, pd.Series, np.ndarray]) -> Dict[str, float]:
        """
        Calculate comprehensive statistics for a dataset.
        Consolidated from calculation_service.py statistical functions.
        """
        debug_print("Calculating comprehensive statistics")
        
        try:
            # Convert to pandas Series for consistent handling
            if isinstance(data, (list, np.ndarray)):
                data = pd.Series(data)
            elif not isinstance(data, pd.Series):
                data = pd.Series([data])
            
            # Remove NaN values
            clean_data = data.dropna()
            
            if len(clean_data) < self.statistical_threshold:
                debug_print(f"Insufficient data points ({len(clean_data)}) for statistics")
                return {
                    'count': len(clean_data),
                    'mean': np.nan,
                    'median': np.nan,
                    'std': np.nan,
                    'min': np.nan,
                    'max': np.nan,
                    'q1': np.nan,
                    'q3': np.nan,
                    'iqr': np.nan,
                    'cv': np.nan
                }
            
            # Calculate basic statistics
            stats = {
                'count': len(clean_data),
                'mean': round_values(clean_data.mean(), self.precision_digits),
                'median': round_values(clean_data.median(), self.precision_digits),
                'std': round_values(clean_data.std(), self.precision_digits),
                'min': round_values(clean_data.min(), self.precision_digits),
                'max': round_values(clean_data.max(), self.precision_digits)
            }
            
            # Calculate quartiles
            try:
                stats['q1'] = round_values(clean_data.quantile(0.25), self.precision_digits)
                stats['q3'] = round_values(clean_data.quantile(0.75), self.precision_digits)
                stats['iqr'] = round_values(stats['q3'] - stats['q1'], self.precision_digits)
            except Exception:
                stats['q1'] = np.nan
                stats['q3'] = np.nan
                stats['iqr'] = np.nan
            
            # Calculate coefficient of variation
            try:
                if stats['mean'] != 0:
                    stats['cv'] = round_values((stats['std'] / abs(stats['mean'])) * 100, self.precision_digits)
                else:
                    stats['cv'] = np.nan
            except Exception:
                stats['cv'] = np.nan
            
            debug_print(f"Calculated statistics for {stats['count']} data points")
            return stats
            
        except Exception as e:
            debug_print(f"Error calculating statistics: {e}")
            return {'error': str(e)}
    
    def calculate_usage_efficiency(self, sample_data: pd.DataFrame) -> Tuple[str, Optional[float]]:
        """
        Calculate usage efficiency from sample data.
        Consolidated from processing.py efficiency calculation logic.
        """
        debug_print("Calculating usage efficiency")
        
        try:
            if sample_data.shape[0] <= 3 or sample_data.shape[1] <= 8:
                debug_print("Insufficient data dimensions for efficiency calculation")
                return "", None
            
            # Get initial oil mass from H3 (column 7, row 1 with -1 indexing)
            initial_oil_mass_val = None
            if sample_data.shape[1] > 7:
                initial_oil_mass_val = sample_data.iloc[1, 7]
            
            # Get puffs values from column A (column 0) starting from row 4
            puffs_values = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce').dropna()
            
            # Get TPM values from column I (column 8) starting from row 4  
            tpm_values = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
            
            if (initial_oil_mass_val is None or pd.isna(initial_oil_mass_val) or
                puffs_values.empty or tpm_values.empty):
                debug_print("Missing required data for efficiency calculation")
                return "", None
            
            try:
                initial_oil_mass = float(initial_oil_mass_val)
                if initial_oil_mass <= 0:
                    debug_print("Invalid initial oil mass for efficiency calculation")
                    return "", None
                    
            except (ValueError, TypeError):
                debug_print("Cannot convert initial oil mass to float")
                return "", None
            
            # Calculate total mass consumed: sum of (puff_interval * tpm / 1000)
            # Get puff intervals
            puff_intervals = puffs_values.diff().fillna(puffs_values.iloc[0] if len(puffs_values) > 0 else 1)
            puff_intervals = puff_intervals.replace(0, 1)  # Avoid division by zero
            
            # Align data lengths
            min_length = min(len(puff_intervals), len(tpm_values))
            puff_intervals = puff_intervals.iloc[:min_length]
            tpm_values = tpm_values.iloc[:min_length]
            
            # Calculate mass consumed per interval
            mass_per_interval = (puff_intervals * tpm_values) / 1000.0
            total_mass_consumed = mass_per_interval.sum()
            
            if total_mass_consumed <= 0:
                debug_print("Total mass consumed is zero or negative")
                return "", None
            
            # Calculate efficiency as percentage
            efficiency_ratio = total_mass_consumed / initial_oil_mass
            efficiency_percentage = efficiency_ratio * 100.0
            
            calculated_efficiency = round_values(efficiency_percentage, self.precision_digits)
            usage_efficiency = f"{calculated_efficiency}%"
            
            debug_print(f"Calculated usage efficiency: {usage_efficiency}")
            debug_print(f"Initial mass: {initial_oil_mass}mg, Consumed: {total_mass_consumed}mg")
            
            return usage_efficiency, calculated_efficiency
            
        except Exception as e:
            debug_print(f"Error calculating usage efficiency: {e}")
            return "", None
    
    # ===================== VISCOSITY CALCULATIONS =====================
    
    def calculate_viscosity_formulation(self, media: str, terpene: str, mass: float, 
                                        target: float) -> Tuple[float, float]:
        """
        Calculate viscosity formulation based on database ratios.
        Consolidated from viscosity_calculator calculation logic.
        """
        debug_print(f"Calculating viscosity formulation: {media}_{terpene}")
        
        try:
            key = f"{media}_{terpene}"
            
            if key in self.formulation_database and 'ratio' in self.formulation_database[key]:
                ratio = self.formulation_database[key]['ratio']
                debug_print(f"Using database ratio: {ratio}")
            else:
                # Default ratios based on media type
                default_ratios = {
                    'MCT': 0.1,
                    'PG': 0.15,
                    'VG': 0.08,
                    'PEG': 0.12
                }
                ratio = default_ratios.get(media, 0.1)
                debug_print(f"Using default ratio for {media}: {ratio}")
            
            # Calculate terpene mass needed
            result_mass = mass * ratio
            result_percent = ratio * 100.0
            
            debug_print(f"Formulation result: {result_percent:.2f}% = {result_mass:.3f}mg")
            return result_percent, result_mass
            
        except Exception as e:
            debug_print(f"Error calculating viscosity formulation: {e}")
            return 0.0, 0.0
    
    def predict_viscosity_from_models(self, media: str, terpene_pct: float, 
                                      temperature: float, potency: float = None) -> Optional[float]:
        """
        Predict viscosity using trained models.
        Consolidated from viscosity_calculator model prediction logic.
        """
        debug_print(f"Predicting viscosity: {media} at {temperature}C")
        
        try:
            # Check for base model first
            model_key = f"{media}_base"
            if model_key in self.base_models:
                return self._predict_base_model_viscosity(model_key, terpene_pct, temperature, potency)
            
            # Check for consolidated model
            model_key = f"{media}_consolidated"  
            if model_key in self.consolidated_models:
                return self._predict_consolidated_model_viscosity(model_key, terpene_pct, temperature, potency)
            
            # Check for standard viscosity model
            if media in self.viscosity_models:
                return self._predict_standard_model_viscosity(media, terpene_pct, temperature)
            
            debug_print(f"No viscosity model found for {media}")
            return None
            
        except Exception as e:
            debug_print(f"Error predicting viscosity: {e}")
            return None
    
    def _predict_base_model_viscosity(self, model_key: str, terpene_pct: float, 
                                      temperature: float, potency: float) -> Optional[float]:
        """Predict viscosity using base model."""
        debug_print(f"Using base model {model_key} for prediction")
        # Placeholder for base model prediction logic
        # Would implement actual sklearn model prediction here
        return None
    
    def _predict_consolidated_model_viscosity(self, model_key: str, terpene_pct: float,
                                              temperature: float, potency: float) -> Optional[float]:
        """Predict viscosity using consolidated model.""" 
        debug_print(f"Using consolidated model {model_key} for prediction")
        # Placeholder for consolidated model prediction logic
        # Would implement actual sklearn model prediction here
        return None
    
    def _predict_standard_model_viscosity(self, media: str, terpene_pct: float, 
                                          temperature: float) -> Optional[float]:
        """Predict viscosity using standard model."""
        debug_print(f"Using standard model for {media} prediction")
        # Placeholder for standard model prediction logic
        # Would implement actual model prediction here
        return None
    
    def arrhenius_function(self, x: np.ndarray, a: float, b: float) -> np.ndarray:
        """Arrhenius function for viscosity modeling."""
        return a + b * x
    
    def save_viscosity_calculation(self, media: str, terpene: str, mass: float, 
                                   target: float, result_percent: float, result_mass: float):
        """Save viscosity calculation to database."""
        debug_print("Saving viscosity calculation to database")
        
        try:
            key = f"{media}_{terpene}"
            
            if key not in self.formulation_database:
                self.formulation_database[key] = {
                    'media': media,
                    'terpene': terpene,
                    'calculations': []
                }
            
            calculation = {
                'mass_oil': mass,
                'target_viscosity': target,
                'result_percent': result_percent,
                'result_mass': result_mass,
                'ratio': result_percent / 100.0,
                'timestamp': datetime.now().isoformat()
            }
            
            self.formulation_database[key]['calculations'].append(calculation)
            
            # Keep only last 10 calculations per formulation
            if len(self.formulation_database[key]['calculations']) > 10:
                self.formulation_database[key]['calculations'] = self.formulation_database[key]['calculations'][-10:]
            
            # Update average ratio
            ratios = [calc['ratio'] for calc in self.formulation_database[key]['calculations']]
            self.formulation_database[key]['ratio'] = sum(ratios) / len(ratios)
            
            self._save_formulation_database()
            debug_print("Viscosity calculation saved successfully")
            
        except Exception as e:
            debug_print(f"Failed to save viscosity calculation: {e}")
    
    # ===================== SAMPLE DATA PROCESSING =====================
    
    def extract_sample_data(self, raw_data: pd.DataFrame, sample_index: int) -> Dict[str, Any]:
        """
        Extract comprehensive sample data from raw Excel data.
        Consolidated from processing.py sample extraction logic.
        """
        debug_print(f"Extracting sample data for sample {sample_index + 1}")
        
        try:
            # Calculate column range for this sample
            col_start = 1 + (sample_index * self.sample_columns_per_sample)
            col_end = col_start + self.sample_columns_per_sample - 1
            
            if col_end >= len(raw_data.columns):
                debug_print(f"Sample {sample_index + 1} extends beyond available columns")
                return self._create_empty_sample_data(sample_index)
            
            # Extract sample data block
            sample_data = raw_data.iloc[:, col_start:col_end + 1].copy()
            
            # Extract sample name with enhanced logic
            sample_name = self._extract_enhanced_sample_name(sample_data, sample_index, raw_data)
            
            # Extract TPM data and statistics
            tpm_stats = self._extract_tpm_statistics(sample_data)
            
            # Extract metadata (media, viscosity, voltage, etc.)
            metadata = self._extract_sample_metadata(sample_data)
            
            # Calculate usage efficiency
            usage_efficiency, calculated_efficiency = self.calculate_usage_efficiency(sample_data)
            
            # Extract burn/clog/leak status
            burn, clog, leak = self._extract_burn_clog_leak(raw_data, sample_index)
            
            # Compile comprehensive sample data
            sample_info = {
                'Sample Name': sample_name,
                'Media': metadata.get('media', ''),
                'Viscosity': metadata.get('viscosity', ''),
                'Voltage, Resistance, Power': metadata.get('voltage_info', ''),
                'Average TPM': tpm_stats.get('avg_tpm'),
                'Standard Deviation': tpm_stats.get('std_tpm'),
                'Usage Efficiency': usage_efficiency,
                'Calculated Efficiency': calculated_efficiency,
                'Normalized TPM': self._calculate_sample_normalized_tpm(sample_data, tpm_stats.get('tpm_data')),
                'Burn?': burn,
                'Clog?': clog,
                'Leak?': leak,
                'raw_sample_data': sample_data,
                'sample_index': sample_index,
                'column_range': (col_start, col_end)
            }
            
            debug_print(f"Extracted sample: '{sample_name}' with {len(sample_info)} attributes")
            return sample_info
            
        except Exception as e:
            debug_print(f"Error extracting sample {sample_index + 1}: {e}")
            return self._create_empty_sample_data(sample_index)
    
    def _extract_enhanced_sample_name(self, sample_data: pd.DataFrame, sample_index: int, 
                                      raw_data: pd.DataFrame) -> str:
        """Enhanced sample name extraction with old/new format support."""
        sample_name = f"Sample {sample_index + 1}"  # Default fallback
        
        try:
            if sample_data.shape[1] <= 8:
                return sample_name
            
            headers = sample_data.columns.astype(str)
            project_value = None
            sample_value = None
            
            # Enhanced pattern matching
            for i, header in enumerate(headers):
                header_lower = str(header).lower().strip()
                
                # New format "Sample ID:" pattern
                if header_lower == "sample id:" and i + 1 < len(headers):
                    sample_value = str(headers[i + 1]).strip()
                    if sample_value and sample_value.lower() not in ['nan', 'none', '']:
                        return sample_value
                
                # Old format "Project:" patterns with suffix support
                elif (header_lower.startswith("project:") and 
                      (header_lower == "project:" or 
                       (header_lower.startswith("project:.") and header_lower[9:].isdigit()))):
                    if i + 1 < len(headers):
                        project_value = str(headers[i + 1]).strip()
                        # Remove pandas suffix if exists
                        if project_value.endswith(f'.{sample_index}'):
                            project_value = project_value[:-len(f'.{sample_index}')]
                
                # Old format "Sample:" patterns
                elif (header_lower == "sample:" or
                      (header_lower.startswith("sample:.") and header_lower[8:].isdigit())):
                    if i + 1 < len(headers):
                        temp_sample_value = str(headers[i + 1]).strip()
                        if not sample_value:  # Don't overwrite Sample ID
                            sample_value = temp_sample_value
            
            # Determine final sample name
            if sample_value and sample_value.lower() not in ['nan', 'none', '', f'unnamed: {5}']:
                if project_value and project_value.lower() not in ['nan', 'none', '']:
                    sample_name = f"{project_value} {sample_value}".strip()
                else:
                    sample_name = sample_value.strip()
            elif project_value and project_value.lower() not in ['nan', 'none', '']:
                sample_name = project_value.strip()
            else:
                # Try new format direct column 5 access
                if len(headers) > 5:
                    fallback_value = str(headers[5]).strip()
                    if fallback_value and fallback_value.lower() not in ['nan', 'none', '', 'unnamed: 5']:
                        sample_name = fallback_value
            
            debug_print(f"Final sample name for sample {sample_index + 1}: '{sample_name}'")
            return sample_name
            
        except Exception as e:
            debug_print(f"Error extracting sample name for sample {sample_index + 1}: {e}")
            return sample_name
    
    def _extract_tpm_statistics(self, sample_data: pd.DataFrame) -> Dict[str, Any]:
        """Extract TPM data and calculate statistics."""
        try:
            tpm_data = pd.Series(dtype=float)
            if sample_data.shape[0] > 3 and sample_data.shape[1] > 8:
                tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
            
            stats = self.calculate_statistics(tpm_data)
            
            return {
                'tpm_data': tpm_data,
                'avg_tpm': round_values(tpm_data.mean()) if not tpm_data.empty else None,
                'std_tpm': round_values(tpm_data.std()) if not tpm_data.empty else None,
                'stats': stats
            }
        except Exception as e:
            debug_print(f"Error extracting TPM statistics: {e}")
            return {'tpm_data': pd.Series(dtype=float), 'avg_tpm': None, 'std_tpm': None}
    
    def _extract_sample_metadata(self, sample_data: pd.DataFrame) -> Dict[str, str]:
        """Extract sample metadata (media, viscosity, voltage, etc.)."""
        try:
            metadata = {}
            
            # Extract media and viscosity
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 1:
                metadata['media'] = str(sample_data.iloc[0, 1]) if not pd.isna(sample_data.iloc[0, 1]) else ""
                metadata['viscosity'] = str(sample_data.iloc[1, 1]) if not pd.isna(sample_data.iloc[1, 1]) else ""
            
            # Extract voltage, resistance, power
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
                voltage = str(sample_data.iloc[1, 5]) if not pd.isna(sample_data.iloc[1, 5]) else ""
                resistance = round_values(sample_data.iloc[0, 3]) if (sample_data.shape[1] > 3 and not pd.isna(sample_data.iloc[0, 3])) else ""
                power = round_values(sample_data.iloc[0, 5]) if not pd.isna(sample_data.iloc[0, 5]) else ""
                
                metadata['voltage_info'] = f"{voltage} V, {resistance} ohm, {power} W" if voltage else ""
            
            return metadata
            
        except Exception as e:
            debug_print(f"Error extracting sample metadata: {e}")
            return {}
    
    def _calculate_sample_normalized_tpm(self, sample_data: pd.DataFrame, tpm_data: pd.Series) -> Optional[str]:
        """Calculate normalized TPM for a sample."""
        try:
            if tpm_data is None or tpm_data.empty:
                return None
            
            # Extract puffing regime from sample data
            puffing_regime = None
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
                puffing_regime_cell = sample_data.iloc[0, 7]
                if pd.notna(puffing_regime_cell):
                    puffing_regime = str(puffing_regime_cell).strip()
            
            # Calculate normalized TPM
            normalized_tpm = self.calculate_normalized_tpm(tpm_data, puffing_regime)
            
            if not normalized_tpm.empty:
                avg_normalized = round_values(normalized_tpm.mean())
                return f"{avg_normalized}"
            
            return None
            
        except Exception as e:
            debug_print(f"Error calculating sample normalized TPM: {e}")
            return None
    
    def _extract_burn_clog_leak(self, raw_data: pd.DataFrame, sample_index: int) -> Tuple[str, str, str]:
        """Extract burn/clog/leak status from raw data."""
        try:
            # Calculate column positions for burn/clog/leak
            col_start = 1 + (sample_index * self.sample_columns_per_sample)
            burn_col = col_start + 9   # Column J relative to sample start
            clog_col = col_start + 10  # Column K relative to sample start  
            leak_col = col_start + 11  # Column L relative to sample start
            
            burn = ""
            clog = ""
            leak = ""
            
            # Check if columns exist in raw data
            if burn_col < len(raw_data.columns) and len(raw_data) > 0:
                burn_val = raw_data.iloc[0, burn_col] if not pd.isna(raw_data.iloc[0, burn_col]) else ""
                burn = str(burn_val).strip() if burn_val else ""
            
            if clog_col < len(raw_data.columns) and len(raw_data) > 0:
                clog_val = raw_data.iloc[0, clog_col] if not pd.isna(raw_data.iloc[0, clog_col]) else ""
                clog = str(clog_val).strip() if clog_val else ""
                
            if leak_col < len(raw_data.columns) and len(raw_data) > 0:
                leak_val = raw_data.iloc[0, leak_col] if not pd.isna(raw_data.iloc[0, leak_col]) else ""
                leak = str(leak_val).strip() if leak_val else ""
            
            debug_print(f"Extracted burn/clog/leak for sample {sample_index + 1}: {burn}/{clog}/{leak}")
            return burn, clog, leak
            
        except Exception as e:
            debug_print(f"Error extracting burn/clog/leak for sample {sample_index + 1}: {e}")
            return "", "", ""
    
    def _create_empty_sample_data(self, sample_index: int) -> Dict[str, Any]:
        """Create empty sample data structure."""
        return {
            'Sample Name': f"Sample {sample_index + 1}",
            'Media': "",
            'Viscosity': "",
            'Voltage, Resistance, Power': "",
            'Average TPM': None,
            'Standard Deviation': None,
            'Usage Efficiency': "",
            'Calculated Efficiency': None,
            'Normalized TPM': None,
            'Burn?': "",
            'Clog?': "",
            'Leak?': "",
            'raw_sample_data': pd.DataFrame(),
            'sample_index': sample_index,
            'column_range': (0, 0)
        }
    
    # ===================== DATA PERSISTENCE =====================
    
    def _load_formulation_database(self):
        """Load formulation database from file."""
        db_file = Path("data/formulation_database.json")
        
        try:
            if db_file.exists():
                with open(db_file, 'r') as f:
                    self.formulation_database = json.load(f)
                debug_print(f"Loaded {len(self.formulation_database)} formulation entries")
            else:
                debug_print("No formulation database found")
        except Exception as e:
            debug_print(f"Error loading formulation database: {e}")
            self.formulation_database = {}
    
    def _save_formulation_database(self):
        """Save formulation database to file."""
        db_file = Path("data/formulation_database.json")
        
        try:
            db_file.parent.mkdir(parents=True, exist_ok=True)
            with open(db_file, 'w') as f:
                json.dump(self.formulation_database, f, indent=2)
            debug_print("Formulation database saved successfully")
        except Exception as e:
            debug_print(f"Error saving formulation database: {e}")
    
    def _load_viscosity_models(self):
        """Load viscosity models from files."""
        models_dir = Path("data/models")
        
        try:
            if models_dir.exists():
                for model_file in models_dir.glob("*.json"):
                    try:
                        with open(model_file, 'r') as f:
                            model_data = json.load(f)
                        
                        model_name = model_file.stem
                        if "consolidated" in model_name:
                            self.consolidated_models[model_name] = model_data
                        elif "base" in model_name:
                            self.base_models[model_name] = model_data
                        else:
                            self.viscosity_models[model_name] = model_data
                            
                    except Exception as e:
                        debug_print(f"Error loading model {model_file}: {e}")
                        
                debug_print(f"Loaded {len(self.consolidated_models)} consolidated models")
                debug_print(f"Loaded {len(self.base_models)} base models")
                debug_print(f"Loaded {len(self.viscosity_models)} other models")
            else:
                debug_print("No models directory found")
                
        except Exception as e:
            debug_print(f"Error loading viscosity models: {e}")
    
    # ===================== SERVICE STATUS AND UTILITIES =====================
    
    def get_calculation_service_status(self) -> Dict[str, Any]:
        """Get comprehensive calculation service status."""
        return {
            'initialized': True,
            'formulation_entries': len(self.formulation_database),
            'consolidated_models': len(self.consolidated_models),
            'base_models': len(self.base_models),
            'other_models': len(self.viscosity_models),
            'default_puff_time': self.default_puff_time,
            'default_puffs': self.default_puffs,
            'precision_digits': self.precision_digits,
            'statistical_threshold': self.statistical_threshold,
            'files_loaded': len(self.all_files),
            'current_file': self.current_file.filename if self.current_file else None,
            'selected_sheet': self.selected_sheet_name
        }
    
    def reset_calculation_service(self):
        """Reset the calculation service to initial state."""
        debug_print("Resetting calculation service")
        
        self.formulation_database.clear()
        self.viscosity_models.clear()
        self.consolidated_models.clear()
        self.base_models.clear()
        
        # Reload from files
        self._load_formulation_database()
        self._load_viscosity_models()
        
        debug_print("Calculation service reset completed")
    
    def process_sheet_data(self, sheet_data: SheetData, processing_function: callable = None) -> Tuple[pd.DataFrame, Dict, pd.DataFrame]:
        """
        Process sheet data using appropriate processing function.
        Consolidated processing dispatcher from processing.py.
        """
        debug_print(f"Processing sheet data for '{sheet_data.name}'")
        
        try:
            data = sheet_data.data
            
            if data.empty:
                debug_print(f"Sheet '{sheet_data.name}' is empty")
                return self._create_empty_processed_data()
            
            # Use provided processing function or get default
            if processing_function is None:
                processing_function = self._get_default_processing_function(sheet_data.name)
            
            # Apply processing function
            processed_data, sample_arrays, full_sample_data = processing_function(data)
            
            # Update sheet with processed data
            sheet_data.set_processed_data(processed_data)
            
            debug_print(f"Processed sheet '{sheet_data.name}': {len(processed_data)} rows")
            return processed_data, sample_arrays, full_sample_data
            
        except Exception as e:
            debug_print(f"Error processing sheet '{sheet_data.name}': {e}")
            return self._create_empty_processed_data()
    
    def _get_default_processing_function(self, sheet_name: str) -> callable:
        """Get default processing function for a sheet."""
        # Legacy sheet detection
        if "legacy" in sheet_name.lower():
            return self._process_legacy_sheet
        
        # Standard processing for all other sheets
        return self._process_standard_sheet
    
    def _process_standard_sheet(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict, pd.DataFrame]:
        """Standard sheet processing logic."""
        debug_print("Processing standard sheet")
        
        try:
            # Determine number of samples based on data width
            total_cols = len(data.columns)
            num_samples = max(0, (total_cols - 1) // self.sample_columns_per_sample)
            
            debug_print(f"Detected {num_samples} samples from {total_cols} columns")
            
            if num_samples == 0:
                return self._create_empty_processed_data()
            
            # Process each sample
            processed_samples = []
            sample_arrays = {}
            all_sample_data = []
            
            for sample_index in range(num_samples):
                sample_info = self.extract_sample_data(data, sample_index)
                processed_samples.append(sample_info)
                
                # Store raw sample data for plotting
                if 'raw_sample_data' in sample_info and not sample_info['raw_sample_data'].empty:
                    all_sample_data.append(sample_info['raw_sample_data'])
            
            # Create processed DataFrame for display
            display_columns = [
                'Sample Name', 'Media', 'Viscosity', 'Voltage, Resistance, Power',
                'Average TPM', 'Standard Deviation', 'Usage Efficiency', 
                'Normalized TPM', 'Burn?', 'Clog?', 'Leak?'
            ]
            
            processed_df = pd.DataFrame(processed_samples)[display_columns]
            
            # Combine all sample data for full dataset
            if all_sample_data:
                full_sample_data = pd.concat(all_sample_data, axis=1, ignore_index=True)
            else:
                full_sample_data = pd.DataFrame()
            
            debug_print(f"Standard processing complete: {len(processed_df)} samples")
            return processed_df, sample_arrays, full_sample_data
            
        except Exception as e:
            debug_print(f"Error in standard sheet processing: {e}")
            return self._create_empty_processed_data()
    
    def _process_legacy_sheet(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict, pd.DataFrame]:
        """Legacy sheet processing logic."""
        debug_print("Processing legacy sheet")
        
        # Simplified legacy processing - would expand based on legacy format requirements
        try:
            # Basic legacy processing structure
            processed_data = data.copy()
            sample_arrays = {}
            full_sample_data = data.copy()
            
            debug_print(f"Legacy processing complete: {len(processed_data)} rows")
            return processed_data, sample_arrays, full_sample_data
            
        except Exception as e:
            debug_print(f"Error in legacy sheet processing: {e}")
            return self._create_empty_processed_data()
    
    def _create_empty_processed_data(self) -> Tuple[pd.DataFrame, Dict, pd.DataFrame]:
        """Create empty processed data structure."""
        empty_df = pd.DataFrame()
        empty_dict = {}
        return empty_df, empty_dict, empty_df
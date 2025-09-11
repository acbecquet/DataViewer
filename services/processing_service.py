"""
services/processing_service.py
Consolidated data processing service for sheet processing and data analysis.
This consolidates all processing functionality from processing.py and related modules.
"""

# Standard library imports
import re
import math
import threading
from typing import Optional, Dict, Any, List, Tuple, Callable, Union
from dataclasses import dataclass

# Third party imports
import pandas as pd
import numpy as np


def debug_print(message: str):
    """Debug print function for processing operations."""
    print(f"DEBUG: ProcessingService - {message}")


def round_values(value: Union[float, int], decimal_places: int = 3) -> float:
    """Round values for display consistency."""
    try:
        return round(float(value), decimal_places)
    except (ValueError, TypeError):
        return 0.0


@dataclass
class ProcessingConfiguration:
    """Configuration for data processing operations."""
    headers_row: int = 3
    data_start_row: int = 4
    num_columns_per_sample: int = 12
    enable_validation: bool = True
    min_required_rows: int = 4
    replace_zeros_with_nan: bool = True


class ProcessingService:
    """
    Consolidated service for data processing and sheet analysis.
    Handles all sheet types, data validation, cleaning, and extraction operations.
    """
    
    def __init__(self, calculation_service=None):
        """Initialize the processing service."""
        debug_print("Initializing ProcessingService")
        
        # Service dependencies
        self.calculation_service = calculation_service
        
        # Processing configuration
        self.default_config = ProcessingConfiguration()
        
        # Processing function registry
        self.processing_functions: Dict[str, Callable] = {}
        self.plot_type_mapping: Dict[str, List[str]] = {}
        
        # Sheet type configurations
        self.sheet_configurations: Dict[str, ProcessingConfiguration] = {}
        
        # Threading locks
        self.processing_lock = threading.Lock()
        
        # Initialize all processing functions
        self._initialize_processing_functions()
        self._initialize_sheet_configurations()
        
        debug_print("ProcessingService initialized successfully")
        debug_print(f"Registered {len(self.processing_functions)} processing functions")
    
    # ===================== INITIALIZATION =====================
    
    def _initialize_processing_functions(self):
        """Initialize all sheet-specific processing functions."""
        self.processing_functions = {
            # Test planning and inspection
            "test plan": self._process_test_plan,
            "initial state inspection": self._process_initial_state_inspection,
            
            # Standard testing functions
            "quick screening test": self._process_plotting_sheet,
            "lifetime test": self._process_plotting_sheet,
            "device life test": self._process_plotting_sheet,
            "aerosol temperature": self._process_plotting_sheet,
            "user test - full cycle": self._process_plotting_sheet,
            "horizontal puffing test": self._process_plotting_sheet,
            "extended test": self._process_plotting_sheet,
            "long puff test": self._process_plotting_sheet,
            "rapid puff test": self._process_plotting_sheet,
            "intense test": self._process_plotting_sheet,
            "sheet1": self._process_plotting_sheet,
            
            # Special test types
            "user test simulation": self._process_user_test_simulation,
            "user simulation test": self._process_user_test_simulation,
            
            # Environmental tests
            "big headspace low t test": self._process_big_headspace_test,
            "big headspace high t test": self._process_big_headspace_test,
            "big headspace serial test": self._process_big_headspace_test,
            "big headspace pocket test": self._process_big_headspace_test,
            "anti-burn protection test": self._process_generic_sheet,
            "upside down test": self._process_big_headspace_test,
            "temperature cycling test": self._process_plotting_sheet,
            "high t high humidity test": self._process_plotting_sheet,
            "low temperature stability": self._process_plotting_sheet,
            "vacuum test": self._process_plotting_sheet,
            "negative pressure test": self._process_plotting_sheet,
            
            # Compatibility tests
            "viscosity compatibility": self._process_plotting_sheet,
            "various oil compatibility": self._process_plotting_sheet,
            
            # Sensory tests
            "quick sensory test": self._process_generic_sheet,
            "off-odor score": self._process_generic_sheet,
            "sensory consistency": self._process_generic_sheet,
            
            # Chemical tests
            "heavy metal leaching test": self._process_generic_sheet,
            
            # Legacy and fallback
            "legacy": self._process_legacy_test,
            "default": self._process_generic_sheet
        }
        
        debug_print(f"Initialized {len(self.processing_functions)} processing functions")
    
    def _initialize_sheet_configurations(self):
        """Initialize sheet-specific configurations."""
        # Test Plan has different row structure
        self.sheet_configurations["test plan"] = ProcessingConfiguration(
            headers_row=5, data_start_row=6, num_columns_per_sample=12
        )
        
        # User Test Simulation uses 8 columns per sample
        self.sheet_configurations["user test simulation"] = ProcessingConfiguration(
            headers_row=3, data_start_row=4, num_columns_per_sample=8
        )
        self.sheet_configurations["user simulation test"] = ProcessingConfiguration(
            headers_row=3, data_start_row=4, num_columns_per_sample=8
        )
        
        debug_print(f"Initialized {len(self.sheet_configurations)} sheet configurations")
    
    # ===================== MAIN PROCESSING INTERFACE =====================
    
    def process_sheet_data(self, raw_data: pd.DataFrame, sheet_name: str, 
                          config: Optional[ProcessingConfiguration] = None) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """
        Main entry point for processing sheet data.
        
        Args:
            raw_data: Raw DataFrame from Excel sheet
            sheet_name: Name of the sheet being processed
            config: Optional processing configuration
            
        Returns:
            Tuple of (processed_data, metadata, full_sample_data)
        """
        debug_print(f"Processing sheet: {sheet_name}, data shape: {raw_data.shape}")
        
        try:
            with self.processing_lock:
                # Get configuration
                if config is None:
                    config = self.get_sheet_configuration(sheet_name)
                
                # Validate input data
                if not self._validate_input_data(raw_data, sheet_name, config):
                    return self._create_empty_structure(raw_data, sheet_name, config)
                
                # Get processing function
                processing_func = self._get_processing_function(sheet_name)
                
                # Process the data
                processed_data, metadata, full_sample_data = processing_func(raw_data, config)
                
                # Add processing metadata
                metadata.update({
                    'sheet_name': sheet_name,
                    'processing_function': processing_func.__name__,
                    'original_shape': raw_data.shape,
                    'processed_shape': processed_data.shape,
                    'config': config
                })
                
                debug_print(f"Successfully processed {sheet_name}: {processed_data.shape}")
                return processed_data, metadata, full_sample_data
                
        except Exception as e:
            error_msg = f"Failed to process sheet {sheet_name}: {e}"
            debug_print(f"ERROR: {error_msg}")
            import traceback
            traceback.print_exc()
            
            # Return empty structure on error
            return self._create_empty_structure(raw_data, sheet_name, config or self.default_config)
    
    def get_processing_function(self, sheet_name: str) -> Callable:
        """Get the appropriate processing function for a sheet (public interface)."""
        return self._get_processing_function(sheet_name)
    
    def get_sheet_configuration(self, sheet_name: str) -> ProcessingConfiguration:
        """Get configuration for a specific sheet type."""
        sheet_key = sheet_name.lower().strip()
        return self.sheet_configurations.get(sheet_key, self.default_config)
    
    def get_valid_plot_options(self, plot_options: List[str], full_sample_data: pd.DataFrame) -> List[str]:
        """
        Check which plot options have valid, non-empty data.
        
        Args:
            plot_options: List of plot types to check
            full_sample_data: DataFrame containing sample data
            
        Returns:
            List of valid plot options
        """
        debug_print(f"Validating plot options: {plot_options}")
        
        try:
            valid_options = []
            
            for plot_type in plot_options:
                try:
                    y_data = self._get_y_data_for_plot_type(full_sample_data, plot_type)
                    
                    # Check if there are valid, non-zero, non-NaN values
                    if y_data is not None and not y_data.empty:
                        numeric_data = pd.to_numeric(y_data, errors='coerce').dropna()
                        if not numeric_data.empty and (numeric_data != 0).any():
                            valid_options.append(plot_type)
                            debug_print(f"Plot type '{plot_type}' is valid")
                        else:
                            debug_print(f"Plot type '{plot_type}' has no valid data")
                    else:
                        debug_print(f"Plot type '{plot_type}' has empty data")
                        
                except Exception as e:
                    debug_print(f"Error validating plot type '{plot_type}': {e}")
                    continue
            
            debug_print(f"Valid plot options: {valid_options}")
            return valid_options
            
        except Exception as e:
            debug_print(f"ERROR: Failed to validate plot options: {e}")
            return []
    
    # ===================== PROCESSING FUNCTION DISPATCHER =====================
    
    def _get_processing_function(self, sheet_name: str) -> Callable:
        """Get the appropriate processing function for a sheet."""
        if not sheet_name:
            return self.processing_functions["default"]
        
        sheet_key = sheet_name.lower().strip()
        
        # Check for exact match first
        if sheet_key in self.processing_functions:
            return self.processing_functions[sheet_key]
        
        # Check for partial matches
        for key, func in self.processing_functions.items():
            if key in sheet_key:
                return func
        
        # Check for legacy sheets
        if "legacy" in sheet_key:
            return self.processing_functions["legacy"]
        
        # Default to generic processing
        return self.processing_functions["default"]
    
    # ===================== SHEET-SPECIFIC PROCESSING FUNCTIONS =====================
    
    def _process_plotting_sheet(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process plotting sheets with sample data extraction."""
        debug_print(f"Processing plotting sheet with config: {config}")
        
        try:
            # Store reference to raw data
            raw_data = data.copy()
            
            # Clean data for processing
            if config.replace_zeros_with_nan:
                data = data.replace(0, np.nan)
            
            data = self._remove_empty_columns(data)
            
            # Initialize results
            samples = []
            full_sample_data = []
            sample_arrays = {}
            
            # Calculate number of samples
            num_samples = data.shape[1] // config.num_columns_per_sample
            debug_print(f"Calculated {num_samples} samples with {config.num_columns_per_sample} columns each")
            
            if num_samples == 0:
                debug_print("No samples detected, creating empty structure")
                return self._create_empty_plotting_structure(data, config)
            
            # Process each sample
            for i in range(num_samples):
                start_col = i * config.num_columns_per_sample
                end_col = start_col + config.num_columns_per_sample
                sample_data = data.iloc[:, start_col:end_col]
                
                if sample_data.empty:
                    debug_print(f"Sample {i+1} is empty, skipping")
                    continue
                
                # Extract sample information
                sample_info = self._extract_sample_data(sample_data, i, config, raw_data)
                if sample_info:
                    samples.append(sample_info)
                    
                    # Store sample arrays for plotting
                    sample_id = f"Sample_{i+1}"
                    self._extract_sample_arrays(sample_data, sample_id, sample_arrays, config)
                
                # Add to full sample data
                full_sample_data.append(sample_data)
            
            # Create processed display data
            processed_data = self._create_processed_display_data(samples)
            
            # Concatenate full sample data
            if full_sample_data:
                concatenated_data = pd.concat(full_sample_data, axis=1, ignore_index=True)
            else:
                concatenated_data = data
            
            metadata = {
                'processing_type': 'plotting_sheet',
                'num_samples': num_samples,
                'sample_arrays': sample_arrays,
                'valid_samples': len(samples)
            }
            
            debug_print(f"Processed plotting sheet: {len(samples)} valid samples")
            return processed_data, metadata, concatenated_data
            
        except Exception as e:
            debug_print(f"ERROR: Failed to process plotting sheet: {e}")
            return self._create_empty_plotting_structure(data, config)
    
    def _process_user_test_simulation(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process User Test Simulation sheets (8 columns per sample)."""
        debug_print(f"Processing User Test Simulation with data shape: {data.shape}")
        
        try:
            # User Test Simulation specific validation
            min_required_rows = max(config.headers_row + 1, config.min_required_rows)
            if data.shape[0] < min_required_rows:
                debug_print(f"Insufficient rows: {data.shape[0]} < {min_required_rows}")
                return self._create_empty_user_test_simulation_structure(data)
            
            # Clean data
            data = self._remove_empty_columns(data).replace(0, np.nan)
            
            samples = []
            full_sample_data = []
            sample_arrays = {}
            
            # Calculate potential samples (8 columns per sample)
            potential_samples = data.shape[1] // 8
            debug_print(f"User Test Simulation: {potential_samples} potential samples")
            
            # Process each sample block
            for i in range(potential_samples):
                start_col = i * 8
                end_col = start_col + 8
                sample_data = data.iloc[:, start_col:end_col]
                
                if sample_data.empty:
                    continue
                
                # Extract User Test Simulation specific data
                sample_info = self._extract_user_test_simulation_sample(sample_data, i)
                if sample_info:
                    samples.append(sample_info)
                
                # Extract arrays for plotting
                self._extract_user_test_simulation_arrays(sample_data, f"Sample_{i+1}", sample_arrays)
                full_sample_data.append(sample_data)
            
            # Create processed data
            if samples:
                processed_data = pd.DataFrame(samples)
                concatenated_data = pd.concat(full_sample_data, axis=1, ignore_index=True)
            else:
                processed_data = pd.DataFrame([{
                    "Sample Name": "Sample 1",
                    "Media": "",
                    "Viscosity": "",
                    "Voltage, Resistance, Power": "",
                    "Average TPM": "No data",
                    "Standard Deviation": "No data",
                    "Initial Oil Mass": "",
                    "Usage Efficiency": "",
                    "Test Type": "User Test Simulation"
                }])
                concatenated_data = data
            
            metadata = {
                'processing_type': 'user_test_simulation',
                'num_samples': len(samples),
                'sample_arrays': sample_arrays,
                'columns_per_sample': 8
            }
            
            debug_print(f"Processed User Test Simulation: {len(samples)} samples")
            return processed_data, metadata, concatenated_data
            
        except Exception as e:
            debug_print(f"ERROR: Failed to process User Test Simulation: {e}")
            return self._create_empty_user_test_simulation_structure(data)
    
    def _process_big_headspace_test(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process Big Headspace tests (no efficiency calculation)."""
        debug_print("Processing Big Headspace test")
        
        # Use standard plotting processing but modify the extraction function
        return self._process_plotting_sheet(data, config)
    
    def _process_generic_sheet(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process generic sheets that don't match specific patterns."""
        debug_print("Processing generic sheet")
        
        try:
            # Check if data is empty or invalid
            if data.empty or data.isna().all().all():
                return pd.DataFrame(), {}, pd.DataFrame()
            
            # Check for sufficient rows
            if data.shape[0] <= config.headers_row or data.shape[0] <= config.data_start_row:
                debug_print(f"Insufficient rows for processing")
                return pd.DataFrame(), {}, pd.DataFrame()
            
            # Extract headers and data
            column_names = data.iloc[config.headers_row].fillna("").tolist()
            table_data = data.iloc[config.data_start_row:].copy()
            table_data.columns = column_names
            
            # Convert to strings for display
            table_data = table_data.astype(str)
            
            # Create display DataFrame
            display_df = table_data.copy()
            
            # Full sample data is the same as table data for generic sheets
            full_sample_data_df = pd.concat([table_data], axis=1)
            
            metadata = {
                'processing_type': 'generic',
                'headers_extracted': len(column_names),
                'data_rows': len(table_data)
            }
            
            debug_print(f"Processed generic sheet: {display_df.shape}")
            return display_df, metadata, full_sample_data_df
            
        except Exception as e:
            debug_print(f"ERROR: Generic sheet processing failed: {e}")
            return pd.DataFrame(), {'error': str(e)}, pd.DataFrame()
    
    def _process_test_plan(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process Test Plan sheet with custom row configuration."""
        debug_print("Processing Test Plan sheet")
        
        # Test Plan uses rows 5-6 instead of 3-4
        plan_config = ProcessingConfiguration(
            headers_row=5, data_start_row=6, 
            num_columns_per_sample=config.num_columns_per_sample
        )
        
        return self._process_generic_sheet(data, plan_config)
    
    def _process_initial_state_inspection(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process Initial State Inspection sheet."""
        debug_print("Processing Initial State Inspection sheet")
        return self._process_generic_sheet(data, config)
    
    def _process_legacy_test(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process legacy test data (already converted to standard format)."""
        debug_print("Processing legacy test data")
        
        # Legacy data uses standard plotting processing
        processed_data, metadata, full_sample_data = self._process_plotting_sheet(data, config)
        
        # Add legacy-specific metadata
        metadata['processing_type'] = 'legacy'
        metadata['legacy_converted'] = True
        
        return processed_data, metadata, full_sample_data
    
    # ===================== DATA EXTRACTION FUNCTIONS =====================
    
    def _extract_sample_data(self, sample_data: pd.DataFrame, sample_index: int, 
                           config: ProcessingConfiguration, raw_data: pd.DataFrame) -> Optional[Dict[str, Any]]:
        """Extract comprehensive sample information."""
        try:
            sample_name = self._extract_sample_name(sample_data, sample_index)
            
            # Extract metadata from specific positions
            metadata = self._extract_sample_metadata(sample_data, raw_data, sample_index, config)
            
            # Calculate TPM statistics
            tpm_stats = self._calculate_tpm_statistics(sample_data, config)
            
            # Extract burn/clog/leak information
            burn_clog_leak = self._extract_burn_clog_leak(raw_data, sample_index, config)
            
            sample_info = {
                "Sample Name": sample_name,
                "Media": metadata.get('media', ''),
                "Viscosity": metadata.get('viscosity', ''),
                "Voltage, Resistance, Power": metadata.get('voltage_resistance_power', ''),
                "Average TPM": tpm_stats.get('average', 'No data'),
                "Standard Deviation": tpm_stats.get('std_dev', 'No data'),
                "Initial Oil Mass": metadata.get('initial_oil_mass', ''),
                "Usage Efficiency": tpm_stats.get('usage_efficiency', ''),
                "Burn?": burn_clog_leak.get('burn', ''),
                "Clog?": burn_clog_leak.get('clog', ''),
                "Leak?": burn_clog_leak.get('leak', '')
            }
            
            debug_print(f"Extracted sample data for {sample_name}")
            return sample_info
            
        except Exception as e:
            debug_print(f"ERROR: Failed to extract sample data: {e}")
            return None
    
    def _extract_user_test_simulation_sample(self, sample_data: pd.DataFrame, sample_index: int) -> Optional[Dict[str, Any]]:
        """Extract sample data specific to User Test Simulation format."""
        try:
            # User Test Simulation specific extraction
            sample_name = self._extract_sample_name(sample_data, sample_index)
            
            # Extract from User Test Simulation positions
            media = str(sample_data.iloc[2, 0]) if sample_data.shape[0] > 2 else ""
            viscosity = str(sample_data.iloc[2, 2]) if sample_data.shape[0] > 2 and sample_data.shape[1] > 2 else ""
            
            # Calculate TPM for User Test Simulation (puffs in column 1, weights in columns 2-3)
            tpm_stats = self._calculate_user_test_simulation_tpm(sample_data)
            
            sample_info = {
                "Sample Name": sample_name,
                "Media": media,
                "Viscosity": viscosity,
                "Voltage, Resistance, Power": "",
                "Average TPM": tpm_stats.get('average', 'No data'),
                "Standard Deviation": tpm_stats.get('std_dev', 'No data'),
                "Initial Oil Mass": "",
                "Usage Efficiency": tpm_stats.get('usage_efficiency', ''),
                "Test Type": "User Test Simulation"
            }
            
            return sample_info
            
        except Exception as e:
            debug_print(f"ERROR: Failed to extract User Test Simulation sample: {e}")
            return None
    
    def _extract_sample_arrays(self, sample_data: pd.DataFrame, sample_id: str, 
                             sample_arrays: Dict[str, Any], config: ProcessingConfiguration) -> None:
        """Extract arrays for plotting from sample data."""
        try:
            # Extract puffs (column 0)
            puffs = pd.to_numeric(sample_data.iloc[config.data_start_row:, 0], errors='coerce').dropna()
            sample_arrays[f"{sample_id}_Puffs"] = puffs.tolist()
            
            # Extract before weights (column 1)
            before_weights = pd.to_numeric(sample_data.iloc[config.data_start_row:, 1], errors='coerce').dropna()
            sample_arrays[f"{sample_id}_Before_Weight"] = before_weights.tolist()
            
            # Extract after weights (column 2)
            after_weights = pd.to_numeric(sample_data.iloc[config.data_start_row:, 2], errors='coerce').dropna()
            sample_arrays[f"{sample_id}_After_Weight"] = after_weights.tolist()
            
            # Calculate and store TPM
            if self.calculation_service:
                tpm_values = self.calculation_service.calculate_tpm_from_weights(
                    puffs, before_weights, after_weights
                )
                sample_arrays[f"{sample_id}_TPM"] = tpm_values.tolist()
            
            # Extract other measurements if available
            if sample_data.shape[1] > 3:
                draw_pressure = pd.to_numeric(sample_data.iloc[config.data_start_row:, 3], errors='coerce').dropna()
                sample_arrays[f"{sample_id}_Draw_Pressure"] = draw_pressure.tolist()
            
            if sample_data.shape[1] > 4:
                resistance = pd.to_numeric(sample_data.iloc[config.data_start_row:, 4], errors='coerce').dropna()
                sample_arrays[f"{sample_id}_Resistance"] = resistance.tolist()
            
        except Exception as e:
            debug_print(f"ERROR: Failed to extract sample arrays for {sample_id}: {e}")
    
    def _extract_user_test_simulation_arrays(self, sample_data: pd.DataFrame, sample_id: str, 
                                           sample_arrays: Dict[str, Any]) -> None:
        """Extract arrays for User Test Simulation plotting."""
        try:
            # In User Test Simulation, puffs are in column 1
            puffs = pd.to_numeric(sample_data.iloc[4:, 1], errors='coerce').dropna()
            sample_arrays[f"{sample_id}_Puffs"] = puffs.tolist()
            
            # Weights in columns 2-3
            before_weights = pd.to_numeric(sample_data.iloc[4:, 2], errors='coerce').dropna()
            after_weights = pd.to_numeric(sample_data.iloc[4:, 3], errors='coerce').dropna()
            
            sample_arrays[f"{sample_id}_Before_Weight"] = before_weights.tolist()
            sample_arrays[f"{sample_id}_After_Weight"] = after_weights.tolist()
            
            # Calculate TPM
            if self.calculation_service:
                tpm_values = self.calculation_service.calculate_tpm_from_weights(
                    puffs, before_weights, after_weights
                )
                sample_arrays[f"{sample_id}_TPM"] = tpm_values.tolist()
            
        except Exception as e:
            debug_print(f"ERROR: Failed to extract User Test Simulation arrays for {sample_id}: {e}")
    
    # ===================== DATA CALCULATION FUNCTIONS =====================
    
    def _calculate_tpm_statistics(self, sample_data: pd.DataFrame, config: ProcessingConfiguration) -> Dict[str, str]:
        """Calculate TPM statistics for a sample."""
        try:
            # Get weight data
            before_weights = pd.to_numeric(sample_data.iloc[config.data_start_row:, 1], errors='coerce').dropna()
            after_weights = pd.to_numeric(sample_data.iloc[config.data_start_row:, 2], errors='coerce').dropna()
            puffs = pd.to_numeric(sample_data.iloc[config.data_start_row:, 0], errors='coerce').dropna()
            
            if before_weights.empty or after_weights.empty or puffs.empty:
                return {'average': 'No data', 'std_dev': 'No data', 'usage_efficiency': ''}
            
            # Calculate TPM
            if self.calculation_service:
                tpm_values = self.calculation_service.calculate_tpm_from_weights(
                    puffs, before_weights, after_weights
                )
            else:
                # Simple TPM calculation
                weight_diff = before_weights - after_weights
                tpm_values = (weight_diff * 1000) / puffs  # Convert to mg and normalize
            
            tpm_numeric = pd.to_numeric(tpm_values, errors='coerce').dropna()
            
            if tpm_numeric.empty:
                return {'average': 'No data', 'std_dev': 'No data', 'usage_efficiency': ''}
            
            avg_tpm = round_values(tpm_numeric.mean())
            std_tpm = round_values(tpm_numeric.std())
            
            # Calculate usage efficiency if we have oil mass
            usage_efficiency = self._calculate_usage_efficiency(sample_data, tpm_numeric)
            
            return {
                'average': f"{avg_tpm:.3f}" if avg_tpm > 0 else "No data",
                'std_dev': f"{std_tpm:.3f}" if std_tpm > 0 else "No data",
                'usage_efficiency': usage_efficiency
            }
            
        except Exception as e:
            debug_print(f"ERROR: Failed to calculate TPM statistics: {e}")
            return {'average': 'No data', 'std_dev': 'No data', 'usage_efficiency': ''}
    
    def _calculate_user_test_simulation_tpm(self, sample_data: pd.DataFrame) -> Dict[str, str]:
        """Calculate TPM statistics for User Test Simulation format."""
        try:
            # User Test Simulation: puffs in column 1, weights in columns 2-3
            puffs = pd.to_numeric(sample_data.iloc[4:, 1], errors='coerce').dropna()
            before_weights = pd.to_numeric(sample_data.iloc[4:, 2], errors='coerce').dropna()
            after_weights = pd.to_numeric(sample_data.iloc[4:, 3], errors='coerce').dropna()
            
            if puffs.empty or before_weights.empty or after_weights.empty:
                return {'average': 'No data', 'std_dev': 'No data', 'usage_efficiency': ''}
            
            # Calculate TPM
            if self.calculation_service:
                tpm_values = self.calculation_service.calculate_tpm_from_weights(
                    puffs, before_weights, after_weights
                )
            else:
                weight_diff = before_weights - after_weights
                tpm_values = (weight_diff * 1000) / puffs
            
            tpm_numeric = pd.to_numeric(tpm_values, errors='coerce').dropna()
            
            if tpm_numeric.empty:
                return {'average': 'No data', 'std_dev': 'No data', 'usage_efficiency': ''}
            
            avg_tpm = round_values(tpm_numeric.mean())
            std_tpm = round_values(tpm_numeric.std())
            
            return {
                'average': f"{avg_tpm:.3f}" if avg_tpm > 0 else "No data",
                'std_dev': f"{std_tpm:.3f}" if std_tpm > 0 else "No data",
                'usage_efficiency': ''  # Not calculated for User Test Simulation
            }
            
        except Exception as e:
            debug_print(f"ERROR: Failed to calculate User Test Simulation TPM: {e}")
            return {'average': 'No data', 'std_dev': 'No data', 'usage_efficiency': ''}
    
    def _calculate_usage_efficiency(self, sample_data: pd.DataFrame, tpm_values: pd.Series) -> str:
        """Calculate usage efficiency from sample data."""
        try:
            # Try to get initial oil mass from header area
            if sample_data.shape[0] > 3 and sample_data.shape[1] > 7:
                initial_oil_mass_val = sample_data.iloc[1, 7]  # H3 position
                if pd.notna(initial_oil_mass_val):
                    initial_oil_mass = float(initial_oil_mass_val)
                    total_mass_vaporized = tpm_values.sum() / 1000  # Convert mg to g
                    efficiency = (total_mass_vaporized / initial_oil_mass) * 100
                    return f"{round_values(efficiency, 1)}%"
            
            return ""
            
        except Exception as e:
            debug_print(f"WARNING: Could not calculate usage efficiency: {e}")
            return ""
    
    def _get_y_data_for_plot_type(self, sample_data: pd.DataFrame, plot_type: str) -> Optional[pd.Series]:
        """Extract Y-data for plotting based on plot type."""
        try:
            if plot_type == "TPM" or plot_type == "TPM (Bar)":
                # Calculate TPM from weight differences
                if sample_data.shape[1] < 3:
                    return pd.Series()
                
                puffs = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce')
                before_weights = pd.to_numeric(sample_data.iloc[3:, 1], errors='coerce')
                after_weights = pd.to_numeric(sample_data.iloc[3:, 2], errors='coerce')
                
                if self.calculation_service:
                    return self.calculation_service.calculate_tpm_from_weights(
                        puffs, before_weights, after_weights
                    )
                else:
                    weight_diff = before_weights - after_weights
                    return (weight_diff * 1000) / puffs
                    
            elif plot_type == "Draw Pressure":
                if sample_data.shape[1] > 3:
                    return pd.to_numeric(sample_data.iloc[3:, 3], errors='coerce')
                    
            elif plot_type == "Resistance":
                if sample_data.shape[1] > 4:
                    return pd.to_numeric(sample_data.iloc[3:, 4], errors='coerce')
                    
            elif plot_type == "Power Efficiency":
                # Calculate TPM/Power ratio
                tpm_data = self._get_y_data_for_plot_type(sample_data, "TPM")
                voltage, resistance = self._extract_electrical_parameters(sample_data)
                
                if voltage and resistance and voltage > 0 and resistance > 0:
                    power = (voltage ** 2) / resistance
                    return tpm_data / power
                    
            elif plot_type == "Normalized TPM":
                tpm_data = self._get_y_data_for_plot_type(sample_data, "TPM")
                puff_time = self._extract_puff_time(sample_data)
                return tpm_data / puff_time if puff_time > 0 else tmp_data / 3.0
            
            return pd.Series()
            
        except Exception as e:
            debug_print(f"ERROR: Failed to get Y-data for {plot_type}: {e}")
            return pd.Series()
    
    # ===================== HELPER FUNCTIONS =====================
    
    def _validate_input_data(self, data: pd.DataFrame, sheet_name: str, config: ProcessingConfiguration) -> bool:
        """Validate input data for processing."""
        if data.empty:
            debug_print(f"Sheet {sheet_name} is empty")
            return False
        
        if not config.enable_validation:
            return True
        
        # Check minimum rows
        if data.shape[0] < config.min_required_rows:
            debug_print(f"Sheet {sheet_name} has insufficient rows: {data.shape[0]} < {config.min_required_rows}")
            return False
        
        # Check for data start row
        if data.shape[0] <= config.data_start_row:
            debug_print(f"Sheet {sheet_name} doesn't have data rows")
            return False
        
        return True
    
    def _remove_empty_columns(self, data: pd.DataFrame) -> pd.DataFrame:
        """Remove columns that are completely empty."""
        if data.empty:
            return data
        
        # Remove columns where all values are NaN or 0
        mask = ~((data.isna() | (data == 0)).all())
        cleaned_data = data.loc[:, mask]
        
        debug_print(f"Removed empty columns: {data.shape} -> {cleaned_data.shape}")
        return cleaned_data
    
    def _extract_sample_name(self, sample_data: pd.DataFrame, sample_index: int) -> str:
        """Extract sample name from data."""
        try:
            # Try to extract from various positions
            potential_positions = [(1, 0), (0, 0), (2, 0)]
            
            for row, col in potential_positions:
                if sample_data.shape[0] > row and sample_data.shape[1] > col:
                    value = sample_data.iloc[row, col]
                    if pd.notna(value) and str(value).strip():
                        return str(value).strip()
            
            return f"Sample {sample_index + 1}"
            
        except Exception as e:
            debug_print(f"WARNING: Could not extract sample name: {e}")
            return f"Sample {sample_index + 1}"
    
    def _extract_sample_metadata(self, sample_data: pd.DataFrame, raw_data: pd.DataFrame, 
                                sample_index: int, config: ProcessingConfiguration) -> Dict[str, str]:
        """Extract metadata from sample headers."""
        try:
            metadata = {}
            
            # Extract from standard positions
            if sample_data.shape[0] > 1:
                # Media from row 1, col 1
                if sample_data.shape[1] > 1:
                    media = sample_data.iloc[0, 1]
                    metadata['media'] = str(media) if pd.notna(media) else ""
                
                # Viscosity from row 2, col 1
                if sample_data.shape[0] > 2:
                    viscosity = sample_data.iloc[1, 1]
                    metadata['viscosity'] = str(viscosity) if pd.notna(viscosity) else ""
                
                # Voltage and resistance
                voltage, resistance = self._extract_electrical_parameters(sample_data)
                if voltage or resistance:
                    metadata['voltage_resistance_power'] = f"V:{voltage or 'N/A'}, R:{resistance or 'N/A'}"
                
                # Initial oil mass
                if sample_data.shape[1] > 7:
                    oil_mass = sample_data.iloc[1, 7]
                    metadata['initial_oil_mass'] = str(oil_mass) if pd.notna(oil_mass) else ""
            
            return metadata
            
        except Exception as e:
            debug_print(f"WARNING: Could not extract metadata: {e}")
            return {}
    
    def _extract_electrical_parameters(self, sample_data: pd.DataFrame) -> Tuple[Optional[float], Optional[float]]:
        """Extract voltage and resistance from sample data."""
        voltage = None
        resistance = None
        
        try:
            # Voltage from row 1, column 5
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
                voltage_val = sample_data.iloc[1, 5]
                if pd.notna(voltage_val):
                    voltage = float(voltage_val)
            
            # Resistance from row 0, column 3
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 3:
                resistance_val = sample_data.iloc[0, 3]
                if pd.notna(resistance_val):
                    resistance = float(resistance_val)
                    
        except (ValueError, TypeError) as e:
            debug_print(f"WARNING: Could not extract electrical parameters: {e}")
        
        return voltage, resistance
    
    def _extract_puff_time(self, sample_data: pd.DataFrame) -> float:
        """Extract puff time from puffing regime data."""
        try:
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
                puffing_regime_cell = sample_data.iloc[0, 7]
                if pd.notna(puffing_regime_cell):
                    puffing_regime = str(puffing_regime_cell).strip()
                    
                    # Extract puff time using regex
                    import re
                    pattern = r'mL/(\d+(?:\.\d+)?)'
                    match = re.search(pattern, puffing_regime)
                    if match:
                        return float(match.group(1))
            
            return 3.0  # Default puff time
            
        except Exception as e:
            debug_print(f"WARNING: Could not extract puff time: {e}")
            return 3.0
    
    def _extract_burn_clog_leak(self, raw_data: pd.DataFrame, sample_index: int, 
                              config: ProcessingConfiguration) -> Dict[str, str]:
        """Extract burn/clog/leak information from raw data."""
        try:
            # This would extract from specific positions in the raw data
            # Implementation depends on where this data is stored
            return {'burn': '', 'clog': '', 'leak': ''}
        except Exception as e:
            debug_print(f"WARNING: Could not extract burn/clog/leak data: {e}")
            return {'burn': '', 'clog': '', 'leak': ''}
    
    def _create_processed_display_data(self, samples: List[Dict[str, Any]]) -> pd.DataFrame:
        """Create processed data for display."""
        if not samples:
            return pd.DataFrame([{
                "Sample Name": "Sample 1",
                "Media": "",
                "Viscosity": "",
                "Voltage, Resistance, Power": "",
                "Average TPM": "No data",
                "Standard Deviation": "No data",
                "Initial Oil Mass": "",
                "Usage Efficiency": "",
                "Burn?": "",
                "Clog?": "",
                "Leak?": ""
            }])
        
        return pd.DataFrame(samples)
    
    # ===================== EMPTY STRUCTURE CREATION =====================
    
    def _create_empty_structure(self, data: pd.DataFrame, sheet_name: str, 
                              config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Create empty structure for data collection or when processing fails."""
        debug_print(f"Creating empty structure for {sheet_name}")
        
        if "user test simulation" in sheet_name.lower():
            return self._create_empty_user_test_simulation_structure(data)
        else:
            return self._create_empty_plotting_structure(data, config)
    
    def _create_empty_plotting_structure(self, data: pd.DataFrame, config: ProcessingConfiguration) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Create empty structure for plotting sheets."""
        processed_data = pd.DataFrame([{
            "Sample Name": "Sample 1",
            "Media": "",
            "Viscosity": "",
            "Voltage, Resistance, Power": "",
            "Average TPM": "No data",
            "Standard Deviation": "No data",
            "Initial Oil Mass": "",
            "Usage Efficiency": "",
            "Burn?": "",
            "Clog?": "",
            "Leak?": ""
        }])
        
        sample_arrays = {}
        
        if data.empty:
            full_sample_data = pd.DataFrame(
                index=range(10),
                columns=range(config.num_columns_per_sample)
            )
        else:
            full_sample_data = data
        
        metadata = {
            'processing_type': 'empty_plotting',
            'num_samples': 0,
            'sample_arrays': sample_arrays
        }
        
        return processed_data, metadata, full_sample_data
    
    def _create_empty_user_test_simulation_structure(self, data: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Create empty structure for User Test Simulation."""
        processed_data = pd.DataFrame([{
            "Sample Name": "Sample 1",
            "Media": "",
            "Viscosity": "",
            "Voltage, Resistance, Power": "",
            "Average TPM": "No data",
            "Standard Deviation": "No data",
            "Initial Oil Mass": "",
            "Usage Efficiency": "",
            "Test Type": "User Test Simulation"
        }])
        
        sample_arrays = {}
        
        if data.empty:
            full_sample_data = pd.DataFrame(
                index=range(10),
                columns=range(8)  # User Test Simulation uses 8 columns
            )
        else:
            full_sample_data = data
        
        metadata = {
            'processing_type': 'empty_user_test_simulation',
            'num_samples': 0,
            'sample_arrays': sample_arrays,
            'columns_per_sample': 8
        }
        
        return processed_data, metadata, full_sample_data
    
    # ===================== SERVICE STATUS AND UTILITIES =====================
    
    def get_service_status(self) -> Dict[str, Any]:
        """Get comprehensive service status information."""
        return {
            'registered_functions': list(self.processing_functions.keys()),
            'sheet_configurations': list(self.sheet_configurations.keys()),
            'calculation_service_available': self.calculation_service is not None,
            'default_config': {
                'headers_row': self.default_config.headers_row,
                'data_start_row': self.default_config.data_start_row,
                'num_columns_per_sample': self.default_config.num_columns_per_sample,
                'enable_validation': self.default_config.enable_validation
            }
        }
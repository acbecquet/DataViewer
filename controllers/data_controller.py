"""
controllers/data_controller.py
Data processing controller that coordinates data operations.
This handles Excel file processing, data validation, filtering, and sheet management.
Migrated from processing.py functionality.
"""

import os
import pandas as pd
import numpy as np
import copy
import math
import re
from typing import Optional, Dict, Any, List, Tuple, Union
from models.data_model import DataModel, SheetData


class DataController:
    """Controller for data processing operations."""
    
    def __init__(self, data_model: DataModel, processing_service: Any):
        """Initialize the data controller."""
        self.data_model = data_model
        self.processing_service = processing_service
        
        # Processing configuration from processing.py
        self.processing_functions = self._get_processing_functions()
        self.valid_plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        
        print("DEBUG: DataController initialized")
        print(f"DEBUG: Connected to DataModel and ProcessingService")
        print(f"DEBUG: Loaded {len(self.processing_functions)} processing functions")
    
    def _get_processing_functions(self) -> Dict[str, Any]:
        """
        Get a dictionary mapping sheet names to their processing functions.
        This replicates the processing.py get_processing_functions() dictionary.
        """
        return {
            "Test Plan": self._process_test_plan,
            "Initial State Inspection": self._process_initial_state_inspection,
            "Quick Screening Test": self._process_quick_screening_test,
            "Lifetime Test": self._process_device_life_test,
            "Device Life Test": self._process_device_life_test,
            "Aerosol Temperature": self._process_aerosol_temp_test,
            "User Test - Full Cycle": self._process_user_test,
            "User Test Simulation": self._process_user_test_simulation,
            "User Simulation Test": self._process_user_test_simulation,
            "Horizontal Puffing Test": self._process_horizontal_test,
            "Extended Test": self._process_extended_test,
            "Long Puff Test": self._process_long_puff_test,
            "Rapid Puff Test": self._process_rapid_puff_test,
            "Intense Test": self._process_intense_test,
            "Big Headspace Low T Test": self._process_big_head_low_t_test,
            "Big Headspace Serial Test": self._process_big_head_serial_test,
            "Anti-Burn Protection Test": self._process_burn_protection_test,
            "Big Headspace High T Test": self._process_big_head_high_t_test,
            "Upside Down Test": self._process_upside_down_test,
            "Big Headspace Pocket Test": self._process_pocket_test,
            "Temperature Cycling Test": self._process_temperature_cycling_test,
            "High T High Humidity Test": self._process_high_t_high_humidity_test,
            "Low Temperature Stability": self._process_cold_storage_test,
            "Vacuum Test": self._process_vacuum_test,
            "Negative Pressure Test": self._process_vacuum_test,
            "Viscosity Compatibility": self._process_viscosity_compatibility_test,
            "Various Oil Compatibility": self._process_various_oil_test,
            "Quick Sensory Test": self._process_quick_sensory_test,
            "Off-odor Score": self._process_off_odor_score,
            "Sensory Consistency": self._process_sensory_consistency,
            "Heavy Metal Leaching Test": self._process_leaching_test,
            "Sheet1": self._process_sheet1,
            "default": self._process_generic_sheet
        }
    
    def process_excel_file(self, excel_file, file_path: str) -> Dict[str, Any]:
        """Process an entire Excel file and its sheets."""
        print(f"DEBUG: DataController processing Excel file: {file_path}")
        
        try:
            result = {
                'filtered_sheets': {},
                'file_info': {
                    'path': file_path,
                    'filename': os.path.basename(file_path),
                    'sheet_count': len(excel_file.sheet_names)
                },
                'processing_stats': {
                    'processed_sheets': 0,
                    'plotting_sheets': 0,
                    'error_sheets': 0
                }
            }
            
            # Process each sheet
            for sheet_name in excel_file.sheet_names:
                print(f"DEBUG: Processing sheet: {sheet_name}")
                
                try:
                    # Read sheet data
                    raw_data = excel_file.parse(sheet_name)
                    
                    # Skip completely empty sheets
                    if raw_data.empty:
                        print(f"DEBUG: Skipping empty sheet: {sheet_name}")
                        continue
                    
                    # Process the sheet
                    sheet_result = self.process_sheet_data(sheet_name, raw_data)
                    
                    if sheet_result:
                        result['filtered_sheets'][sheet_name] = sheet_result
                        result['processing_stats']['processed_sheets'] += 1
                        
                        if sheet_result.get('is_plotting'):
                            result['processing_stats']['plotting_sheets'] += 1
                        
                        print(f"DEBUG: Successfully processed sheet {sheet_name}")
                    else:
                        print(f"DEBUG: Failed to process sheet {sheet_name}")
                        result['processing_stats']['error_sheets'] += 1
                        
                except Exception as e:
                    print(f"ERROR: Failed to process sheet {sheet_name}: {e}")
                    result['processing_stats']['error_sheets'] += 1
                    continue
            
            print(f"DEBUG: DataController processed Excel file - {result['processing_stats']['processed_sheets']} sheets successful")
            return result
            
        except Exception as e:
            print(f"ERROR: DataController failed to process Excel file: {e}")
            return {'filtered_sheets': {}, 'file_info': {}, 'processing_stats': {}}
    
    def process_sheet_data(self, sheet_name: str, raw_data: pd.DataFrame) -> Optional[Dict[str, Any]]:
        """Process raw sheet data through appropriate processing function."""
        print(f"DEBUG: DataController processing sheet data: {sheet_name}")
        
        try:
            # Determine if this is a plotting sheet
            is_plotting = self._plotting_sheet_test(sheet_name, raw_data)
            print(f"DEBUG: Sheet {sheet_name} is_plotting: {is_plotting}")
            
            # Get appropriate processing function
            process_function = self._get_processing_function(sheet_name)
            print(f"DEBUG: Using processing function: {process_function.__name__}")
            
            # Process the data
            if is_plotting:
                processed_data, metadata, full_sample_data = process_function(raw_data)
                print(f"DEBUG: Processed plotting data - shape: {processed_data.shape}")
            else:
                # Convert to string for non-plotting sheets
                string_data = raw_data.astype(str).replace([pd.NA, 'nan', 'None'], '')
                processed_data, metadata, full_sample_data = process_function(string_data)
                print(f"DEBUG: Processed non-plotting data - shape: {processed_data.shape}")
            
            # Create result dictionary
            result = {
                'data': processed_data,
                'metadata': metadata or {},
                'full_sample_data': full_sample_data,
                'is_plotting': is_plotting,
                'is_empty': self._check_if_empty(processed_data),
                'sheet_name': sheet_name,
                'processing_function': process_function.__name__
            }
            
            # Add plotting-specific information
            if is_plotting and not full_sample_data.empty:
                result['valid_plot_options'] = self.get_valid_plot_options_for_sheet(sheet_name, full_sample_data)
                result['plot_statistics'] = self._calculate_plot_statistics(full_sample_data)
            
            print(f"DEBUG: DataController successfully processed sheet {sheet_name}")
            return result
            
        except Exception as e:
            print(f"ERROR: DataController failed to process sheet {sheet_name}: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _plotting_sheet_test(self, sheet_name: str, data: pd.DataFrame) -> bool:
        """
        Determine if a sheet should have plotting capabilities.
        Replicates the plotting_sheet_test function from utils.py.
        """
        print(f"DEBUG: Checking if {sheet_name} is a plotting sheet")
        
        try:
            # Check sheet name for plotting keywords
            plotting_keywords = [
                "test", "puff", "draw", "pressure", "resistance", "power", 
                "flow", "temperature", "battery", "activation", "simulation",
                "life", "extended", "rapid", "intense", "horizontal", "vacuum",
                "viscosity", "compatibility", "upside", "headspace", "cycling"
            ]
            
            sheet_lower = sheet_name.lower()
            name_indicates_plotting = any(keyword in sheet_lower for keyword in plotting_keywords)
            
            # Check for exclusion keywords
            exclusion_keywords = [
                "plan", "inspection", "summary", "overview", "notes", 
                "instructions", "config", "setup", "metadata", "sensory",
                "leaching", "odor", "consistency"
            ]
            
            name_excludes_plotting = any(keyword in sheet_lower for keyword in exclusion_keywords)
            
            # Check data structure for plotting indicators
            has_plotting_columns = False
            if not data.empty:
                column_names = [str(col).lower() for col in data.columns]
                plotting_column_keywords = [
                    "puff", "tpm", "mg/puff", "draw", "pressure", "resistance", 
                    "power", "efficiency", "temp", "temperature", "voltage", "current"
                ]
                has_plotting_columns = any(
                    any(keyword in col for keyword in plotting_column_keywords) 
                    for col in column_names
                )
            
            # Final determination
            is_plotting = (name_indicates_plotting and not name_excludes_plotting) or has_plotting_columns
            
            print(f"DEBUG: Sheet {sheet_name} plotting determination: {is_plotting}")
            print(f"DEBUG: Name indicates: {name_indicates_plotting}, excludes: {name_excludes_plotting}, has columns: {has_plotting_columns}")
            
            return is_plotting
            
        except Exception as e:
            print(f"ERROR: Failed to determine plotting status for {sheet_name}: {e}")
            return False
    
    def _get_processing_function(self, sheet_name: str):
        """Get the appropriate processing function for a sheet based on its name."""
        if "legacy" in sheet_name.lower():
            return self._process_legacy_test
        return self.processing_functions.get(sheet_name, self.processing_functions["default"])
    
    def get_valid_plot_options_for_sheet(self, sheet_name: str, full_sample_data: pd.DataFrame) -> List[str]:
        """
        Check which plot options have valid, non-empty data and return the valid options.
        Replicates the get_valid_plot_options function from processing.py.
        """
        print(f"DEBUG: Getting valid plot options for {sheet_name}")
        
        try:
            valid_options = []
            
            for plot_type in self.valid_plot_options:
                y_data = self._get_y_data_for_plot_type(full_sample_data, plot_type)
                if not y_data.empty and y_data.dropna().astype(bool).any():
                    valid_options.append(plot_type)
            
            print(f"DEBUG: Valid plot options for {sheet_name}: {valid_options}")
            return valid_options
            
        except Exception as e:
            print(f"ERROR: Failed to get valid plot options for {sheet_name}: {e}")
            return []
    
    def _get_y_data_for_plot_type(self, sample_data: pd.DataFrame, plot_type: str) -> pd.Series:
        """
        Extract y-data for the specified plot type.
        Replicates the get_y_data_for_plot_type function from processing.py.
        """
        try:
            if sample_data.empty:
                return pd.Series()
            
            # Handle different plot types
            if plot_type == "TPM":
                if "Average TPM" in sample_data.columns:
                    return pd.to_numeric(sample_data["Average TPM"], errors='coerce')
                elif "TPM (mg/puff)" in sample_data.columns:
                    return pd.to_numeric(sample_data["TPM (mg/puff)"], errors='coerce')
                else:
                    return pd.Series()
            
            elif plot_type == "Draw Pressure":
                if "Draw Pressure" in sample_data.columns:
                    return pd.to_numeric(sample_data["Draw Pressure"], errors='coerce')
                else:
                    return pd.Series()
            
            elif plot_type == "Resistance":
                if "Resistance" in sample_data.columns:
                    return pd.to_numeric(sample_data["Resistance"], errors='coerce')
                else:
                    return pd.Series()
            
            elif plot_type == "Power Efficiency":
                # Calculate power efficiency on the fly
                if "Average TPM" in sample_data.columns and "Power" in sample_data.columns:
                    tpm = pd.to_numeric(sample_data["Average TPM"], errors='coerce')
                    power = pd.to_numeric(sample_data["Power"], errors='coerce')
                    return tpm / power.replace(0, np.nan)
                else:
                    return pd.Series()
            
            elif plot_type == "TPM (Bar)":
                # Same as TPM but for bar chart
                return self._get_y_data_for_plot_type(sample_data, "TPM")
            
            else:
                return pd.Series()
                
        except Exception as e:
            print(f"ERROR: Failed to get y-data for plot type {plot_type}: {e}")
            return pd.Series()
    
    def _check_if_empty(self, data: pd.DataFrame) -> bool:
        """Check if processed data is effectively empty."""
        try:
            if data.empty:
                return True
            
            # Check if all cells are effectively empty
            string_data = data.astype(str)
            empty_values = ['', 'nan', 'None', 'NaN', '0', '0.0']
            
            for _, row in string_data.iterrows():
                for value in row:
                    if str(value).strip() not in empty_values:
                        return False
            
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to check if data is empty: {e}")
            return True
    
    def _calculate_plot_statistics(self, full_sample_data: pd.DataFrame) -> Dict[str, Any]:
        """Calculate basic statistics for plotting data."""
        try:
            stats = {
                'sample_count': len(full_sample_data),
                'column_count': len(full_sample_data.columns),
                'has_tpm_data': 'Average TPM' in full_sample_data.columns,
                'has_pressure_data': 'Draw Pressure' in full_sample_data.columns,
                'has_resistance_data': 'Resistance' in full_sample_data.columns
            }
            
            # Calculate TPM statistics if available
            if stats['has_tpm_data']:
                tpm_data = pd.to_numeric(full_sample_data['Average TPM'], errors='coerce').dropna()
                if not tpm_data.empty:
                    stats['tpm_mean'] = float(tpm_data.mean())
                    stats['tpm_std'] = float(tpm_data.std())
                    stats['tpm_min'] = float(tpm_data.min())
                    stats['tpm_max'] = float(tpm_data.max())
            
            return stats
            
        except Exception as e:
            print(f"ERROR: Failed to calculate plot statistics: {e}")
            return {}
    
    # ==================== SHEET PROCESSING FUNCTIONS ====================
    # These functions replicate the processing functions from processing.py
    
    def _process_generic_sheet(self, data: pd.DataFrame, headers_row: int = 3, data_start_row: int = 4) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process a generic sheet with standard layout."""
        print(f"DEBUG: Processing generic sheet with headers_row={headers_row}, data_start_row={data_start_row}")
        
        try:
            if data.empty:
                return pd.DataFrame(), {}, pd.DataFrame()
            
            # Extract headers
            if headers_row < len(data):
                headers = data.iloc[headers_row].fillna('').astype(str).tolist()
            else:
                headers = [f"Column_{i}" for i in range(len(data.columns))]
            
            # Extract data
            if data_start_row < len(data):
                processed_data = data.iloc[data_start_row:].copy()
                processed_data.columns = headers[:len(processed_data.columns)]
            else:
                processed_data = pd.DataFrame(columns=headers)
            
            # Clean data
            processed_data = processed_data.dropna(how='all')
            processed_data = processed_data.loc[:, processed_data.columns != '']
            
            metadata = {
                'headers_row': headers_row,
                'data_start_row': data_start_row,
                'original_shape': data.shape,
                'processed_shape': processed_data.shape
            }
            
            return processed_data, metadata, processed_data
            
        except Exception as e:
            print(f"ERROR: Failed to process generic sheet: {e}")
            return pd.DataFrame(), {}, pd.DataFrame()
    
    def _process_plot_sheet(self, data: pd.DataFrame, headers_row: int = 3, data_start_row: int = 4, 
                           num_columns_per_sample: int = 12) -> Tuple[pd.DataFrame, Dict[str, Any], pd.DataFrame]:
        """Process a sheet with plotting data."""
        print(f"DEBUG: Processing plot sheet with {num_columns_per_sample} columns per sample")
        
        try:
            # First process as generic sheet
            processed_data, metadata, _ = self._process_generic_sheet(data, headers_row, data_start_row)
            
            if processed_data.empty:
                return pd.DataFrame(), metadata, pd.DataFrame()
            
            # Extract sample data for plotting
            full_sample_data = self._extract_sample_data(processed_data, num_columns_per_sample)
            
            # Add plotting-specific metadata
            metadata.update({
                'num_columns_per_sample': num_columns_per_sample,
                'is_plotting_sheet': True,
                'sample_count': len(full_sample_data) if not full_sample_data.empty else 0
            })
            
            return processed_data, metadata, full_sample_data
            
        except Exception as e:
            print(f"ERROR: Failed to process plot sheet: {e}")
            return pd.DataFrame(), {}, pd.DataFrame()
    
    def _extract_sample_data(self, data: pd.DataFrame, num_columns_per_sample: int) -> pd.DataFrame:
        """Extract sample data from processed sheet for plotting."""
        print(f"DEBUG: Extracting sample data with {num_columns_per_sample} columns per sample")
        
        try:
            if data.empty:
                return pd.DataFrame()
            
            # Create sample data structure
            sample_data = []
            
            # Calculate number of samples based on columns
            total_columns = len(data.columns)
            num_samples = max(1, total_columns // num_columns_per_sample)
            
            for sample_idx in range(num_samples):
                start_col = sample_idx * num_columns_per_sample
                end_col = start_col + num_columns_per_sample
                
                if start_col < total_columns:
                    sample_columns = data.columns[start_col:min(end_col, total_columns)]
                    sample_df = data[sample_columns].copy()
                    
                    # Extract key metrics for plotting
                    sample_info = self._extract_sample_metrics(sample_df, sample_idx)
                    if sample_info:
                        sample_data.append(sample_info)
            
            if sample_data:
                result_df = pd.DataFrame(sample_data)
                print(f"DEBUG: Extracted {len(result_df)} samples for plotting")
                return result_df
            else:
                return pd.DataFrame()
                
        except Exception as e:
            print(f"ERROR: Failed to extract sample data: {e}")
            return pd.DataFrame()
    
    def _extract_sample_metrics(self, sample_df: pd.DataFrame, sample_idx: int) -> Optional[Dict[str, Any]]:
        """Extract key metrics from a sample for plotting."""
        try:
            sample_info = {
                'Sample': f"Sample_{sample_idx + 1}",
                'Sample_Index': sample_idx
            }
            
            # Look for key columns
            for _, row in sample_df.iterrows():
                for col, value in row.items():
                    col_str = str(col).lower()
                    value_str = str(value).strip()
                    
                    if 'tpm' in col_str and 'mg/puff' in col_str:
                        try:
                            sample_info['Average TPM'] = float(value_str)
                        except (ValueError, TypeError):
                            pass
                    
                    elif 'draw pressure' in col_str or 'pressure' in col_str:
                        try:
                            sample_info['Draw Pressure'] = float(value_str)
                        except (ValueError, TypeError):
                            pass
                    
                    elif 'resistance' in col_str:
                        try:
                            sample_info['Resistance'] = float(value_str)
                        except (ValueError, TypeError):
                            pass
                    
                    elif 'power' in col_str and 'efficiency' not in col_str:
                        try:
                            sample_info['Power'] = float(value_str)
                        except (ValueError, TypeError):
                            pass
            
            # Only return if we have at least one metric
            metrics = ['Average TPM', 'Draw Pressure', 'Resistance', 'Power']
            if any(metric in sample_info for metric in metrics):
                return sample_info
            else:
                return None
                
        except Exception as e:
            print(f"ERROR: Failed to extract sample metrics: {e}")
            return None
    
    # Individual processing functions for each test type
    def _process_test_plan(self, data):
        """Process data for the Test Plan sheet."""
        return self._process_generic_sheet(data, headers_row=5, data_start_row=6)
    
    def _process_initial_state_inspection(self, data):
        """Process data for the Initial State Inspection sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_quick_screening_test(self, data):
        """Process data for the Quick Screening Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_device_life_test(self, data):
        """Process data for the Device Life Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_aerosol_temp_test(self, data):
        """Process data for the Aerosol Temperature sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_user_test(self, data):
        """Process data for the User Test sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_user_test_simulation(self, data):
        """Process data for the User Test Simulation sheet."""
        print("DEBUG: Processing User Test Simulation with 8-column format")
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=8)
    
    def _process_horizontal_test(self, data):
        """Process data for the Horizontal Puffing Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_extended_test(self, data):
        """Process data for the Extended Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_long_puff_test(self, data):
        """Process data for the Long Puff Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_rapid_puff_test(self, data):
        """Process data for the Rapid Puff Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_intense_test(self, data):
        """Process data for the Intense Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_big_head_low_t_test(self, data):
        """Process data for the Big Headspace Low T Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_big_head_serial_test(self, data):
        """Process data for the Big Headspace Serial Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_burn_protection_test(self, data):
        """Process data for the Anti-Burn Protection Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_big_head_high_t_test(self, data):
        """Process data for the Big Headspace High T Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_upside_down_test(self, data):
        """Process data for the Upside Down Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_pocket_test(self, data):
        """Process data for the Big Headspace Pocket Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_temperature_cycling_test(self, data):
        """Process data for the Temperature Cycling Test sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_high_t_high_humidity_test(self, data):
        """Process data for the High T High Humidity Test sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_cold_storage_test(self, data):
        """Process data for the Low Temperature Stability sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_vacuum_test(self, data):
        """Process data for the Vacuum Test sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_viscosity_compatibility_test(self, data):
        """Process data for the Viscosity Compatibility sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_various_oil_test(self, data):
        """Process data for the Various Oil Compatibility sheet."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_quick_sensory_test(self, data):
        """Process data for the Quick Sensory Test sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_off_odor_score(self, data):
        """Process data for the Off-odor Score sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_sensory_consistency(self, data):
        """Process data for the Sensory Consistency sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_leaching_test(self, data):
        """Process data for the Heavy Metal Leaching Test sheet."""
        return self._process_generic_sheet(data, headers_row=3, data_start_row=4)
    
    def _process_sheet1(self, data):
        """Process data for 'Sheet1' similarly to 'process_extended_test'."""
        return self._process_plot_sheet(data, headers_row=3, data_start_row=4, num_columns_per_sample=12)
    
    def _process_legacy_test(self, data, headers_row=3, data_start_row=4, num_columns_per_sample=12):
        """Process legacy test data."""
        print("DEBUG: Processing legacy test data")
        return self._process_plot_sheet(data, headers_row, data_start_row, num_columns_per_sample)
    
    # ==================== VALIDATION AND UTILITY METHODS ====================
    
    def validate_sheet_data(self, sheet_data: SheetData) -> List[str]:
        """Validate sheet data and return list of issues."""
        issues = []
        
        try:
            # Check basic data integrity
            if sheet_data.data.empty:
                issues.append("Sheet contains no data")
                return issues
            
            # Check for required columns if plotting sheet
            if sheet_data.is_plotting_sheet:
                required_columns = ['Average TPM', 'Draw Pressure', 'Resistance']
                missing_columns = [col for col in required_columns 
                                 if col not in sheet_data.data.columns]
                if len(missing_columns) == len(required_columns):
                    issues.append("No plotting data columns found")
            
            # Check for data quality issues
            total_cells = sheet_data.data.size
            if total_cells > 0:
                null_cells = sheet_data.data.isnull().sum().sum()
                null_percentage = (null_cells / total_cells) * 100
                if null_percentage > 80:
                    issues.append(f"High percentage of missing data: {null_percentage:.1f}%")
            
            print(f"DEBUG: DataController found {len(issues)} validation issues")
            return issues
            
        except Exception as e:
            print(f"ERROR: DataController validation failed: {e}")
            return [f"Validation error: {e}"]
    
    def get_sheet_summary(self, sheet_name: str) -> Dict[str, Any]:
        """Get summary statistics for a sheet."""
        print(f"DEBUG: DataController getting summary for sheet: {sheet_name}")
        
        try:
            sheet_data = self.data_model.get_sheet_data(sheet_name)
            if not sheet_data:
                return {}
            
            summary = {
                'sheet_name': sheet_name,
                'row_count': len(sheet_data.data),
                'column_count': len(sheet_data.data.columns),
                'is_plotting_sheet': sheet_data.is_plotting_sheet,
                'is_empty': self._check_if_empty(sheet_data.data)
            }
            
            # Add plotting-specific summary
            if sheet_data.is_plotting_sheet and hasattr(sheet_data, 'full_sample_data'):
                if not sheet_data.full_sample_data.empty:
                    summary['sample_count'] = len(sheet_data.full_sample_data)
                    summary['valid_plot_options'] = self.get_valid_plot_options_for_sheet(
                        sheet_name, sheet_data.full_sample_data)
            
            return summary
            
        except Exception as e:
            print(f"ERROR: DataController failed to get sheet summary: {e}")
            return {}
    
    def refresh_sheet_data(self, sheet_name: str) -> bool:
        """Refresh/reprocess sheet data."""
        print(f"DEBUG: DataController refreshing sheet data: {sheet_name}")
        
        try:
            # This would require re-processing from original data
            # For now, just update metadata
            current_file = self.data_model.get_current_file()
            if not current_file:
                return False
            
            sheet_data = current_file.get_sheet_data(sheet_name)
            if sheet_data:
                # Update timestamp
                if not hasattr(sheet_data, 'metadata'):
                    sheet_data.metadata = {}
                sheet_data.metadata['last_refreshed'] = self._get_current_timestamp()
                
                print(f"DEBUG: DataController refreshed {sheet_name}")
                return True
            
            return False
            
        except Exception as e:
            print(f"ERROR: DataController failed to refresh sheet data: {e}")
            return False
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp."""
        from datetime import datetime
        return datetime.now().isoformat()
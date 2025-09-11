# services/calculation_service.py
"""
services/calculation_service.py
Consolidated calculation service for mathematical operations.
This consolidates calculation logic from processing.py, viscosity_calculator, and data_collection_window.py.
"""

import os
import json
import re
import math
import statistics
from typing import Optional, Dict, Any, List, Tuple, Union, Callable
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime


def debug_print(message: str):
    """Debug print function for calculation operations."""
    print(f"DEBUG: CalculationService - {message}")


def round_values(value: float, decimal_places: int = 3) -> float:
    """Round values for display consistency."""
    try:
        return round(float(value), decimal_places)
    except (ValueError, TypeError):
        return 0.0


class CalculationService:
    """Service for all mathematical calculations and data analysis operations."""
    
    def __init__(self):
        """Initialize the calculation service."""
        debug_print("Initializing CalculationService")
        
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
        
        # Load existing databases if available
        self._load_formulation_database()
        self._load_viscosity_models()
        
        debug_print("CalculationService initialized successfully")
        debug_print(f"Loaded {len(self.formulation_database)} formulation entries")
        debug_print(f"Loaded {len(self.viscosity_models)} viscosity models")
    
    # ===================== TPM CALCULATIONS =====================
    
    def calculate_tpm_from_weights(self, puffs: pd.Series, before_weights: pd.Series, 
                                   after_weights: pd.Series, test_type: str = "standard") -> pd.Series:
        """
        Calculate TPM from weight differences and puffing intervals.
        
        Args:
            puffs: Series of puff counts
            before_weights: Series of before weights in grams
            after_weights: Series of after weights in grams
            test_type: Type of test for interval calculation logic
            
        Returns:
            Series of calculated TPM values in mg/puff
        """
        debug_print(f"Calculating TPM from weights for {test_type} test")
        debug_print(f"Processing {len(puffs)} data points")
        
        try:
            # Convert to numeric and handle NaN values
            puffs_numeric = pd.to_numeric(puffs, errors='coerce')
            before_numeric = pd.to_numeric(before_weights, errors='coerce')
            after_numeric = pd.to_numeric(after_weights, errors='coerce')
            
            # Calculate puffing intervals
            puffing_intervals = self._calculate_puffing_intervals(puffs_numeric, test_type)
            
            # Calculate weight difference in mg
            weight_diff_mg = (before_numeric - after_numeric) * 1000  # Convert g to mg
            
            # Calculate TPM: weight_diff / puffing_interval
            calculated_tpm = pd.Series(index=weight_diff_mg.index, dtype=float)
            
            for idx in weight_diff_mg.index:
                interval = puffing_intervals.loc[idx] if idx in puffing_intervals.index else np.nan
                weight_diff = weight_diff_mg.loc[idx] if idx in weight_diff_mg.index else np.nan
                
                if pd.notna(interval) and pd.notna(weight_diff) and interval > 0:
                    calculated_tpm.loc[idx] = weight_diff / interval
                    debug_print(f"Row {idx}: {weight_diff:.3f}mg / {interval} puffs = {calculated_tpm.loc[idx]:.3f}mg/puff")
                else:
                    calculated_tpm.loc[idx] = np.nan
                    if pd.notna(interval):
                        debug_print(f"Row {idx}: Skipping TPM calculation - invalid interval: {interval}")
            
            valid_count = calculated_tpm.dropna().count()
            debug_print(f"Successfully calculated TPM for {valid_count}/{len(calculated_tpm)} rows")
            
            return calculated_tpm
            
        except Exception as e:
            debug_print(f"Error calculating TPM from weights: {e}")
            return pd.Series(dtype=float)
    
    def _calculate_puffing_intervals(self, puffs: pd.Series, test_type: str = "standard") -> pd.Series:
        """Calculate puffing intervals based on cumulative puff counts."""
        debug_print(f"Calculating puffing intervals for {test_type} test")
        
        puffing_intervals = pd.Series(index=puffs.index, dtype=float)
        
        for i, idx in enumerate(puffs.index):
            if i == 0:
                # First row: use current puffs value
                current_puffs = puffs.loc[idx] if pd.notna(puffs.loc[idx]) else self.default_puffs
                puffing_intervals.loc[idx] = current_puffs
                debug_print(f"First interval: {current_puffs} puffs")
            else:
                # Subsequent rows: current_puffs - previous_puffs
                prev_idx = puffs.index[i-1]
                current_puffs = puffs.loc[idx] if pd.notna(puffs.loc[idx]) else 0
                prev_puffs = puffs.loc[prev_idx] if pd.notna(puffs.loc[prev_idx]) else 0
                puff_interval = current_puffs - prev_puffs
                
                # Handle invalid intervals
                if puff_interval <= 0 or pd.isna(current_puffs):
                    if pd.isna(current_puffs) or current_puffs == 0:
                        # Use default fallback
                        puffing_intervals.loc[idx] = self.default_puffs
                        debug_print(f"Row {i}: Using default {self.default_puffs} puffs (invalid data)")
                    else:
                        # Use current puffs as interval
                        puffing_intervals.loc[idx] = current_puffs
                        debug_print(f"Row {i}: Using current puffs {current_puffs} (invalid interval)")
                else:
                    puffing_intervals.loc[idx] = puff_interval
        
        return puffing_intervals
    
    def calculate_normalized_tpm(self, tpm_data: pd.Series, sample_data: pd.DataFrame) -> pd.Series:
        """
        Calculate normalized TPM by dividing TPM by puff time.
        
        Args:
            tpm_data: Series of TPM values
            sample_data: DataFrame containing sample metadata
            
        Returns:
            Series of normalized TPM values in mg/s
        """
        debug_print("Calculating Normalized TPM")
        
        try:
            # Convert TPM to numeric
            tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')
            valid_tpm_count = tpm_numeric.dropna().count()
            debug_print(f"Got {valid_tpm_count} valid TPM values for normalization")
            
            if valid_tpm_count == 0:
                debug_print("No valid TPM data for normalization")
                return pd.Series(dtype=float)
            
            # Extract puff time from puffing regime
            puff_time = self._extract_puff_time_from_sample(sample_data)
            
            # Apply normalization
            if puff_time is not None and puff_time > 0:
                normalized_tpm = tpm_numeric / puff_time
                debug_print(f"Successfully normalized TPM by puff time {puff_time}s")
                debug_print(f"Original TPM range: {tpm_numeric.min():.3f} - {tmp_numeric.max():.3f} mg/puff")
                debug_print(f"Normalized TPM range: {normalized_tpm.min():.3f} - {normalized_tpm.max():.3f} mg/s")
                return normalized_tpm
            else:
                debug_print(f"Using default puff time of {self.default_puff_time}s for normalization")
                normalized_tpm = tpm_numeric / self.default_puff_time
                debug_print(f"Normalized TPM with default: range {normalized_tpm.min():.3f} - {normalized_tpm.max():.3f} mg/s")
                return normalized_tpm
                
        except Exception as e:
            debug_print(f"Error calculating normalized TPM: {e}")
            return pd.Series(dtype=float)
    
    def _extract_puff_time_from_sample(self, sample_data: pd.DataFrame) -> Optional[float]:
        """Extract puff time from puffing regime cell."""
        debug_print("Extracting puff time from sample data")
        
        try:
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
                # Check row 1, column 8 (index 0, 7) for puffing regime
                puffing_regime_cell = sample_data.iloc[0, 7]
                if pd.notna(puffing_regime_cell):
                    puffing_regime = str(puffing_regime_cell).strip()
                    debug_print(f"Found puffing regime: '{puffing_regime}'")
                    
                    # Extract puff time using regex pattern
                    pattern = r'mL/(\d+(?:\.\d+)?)s/'
                    match = re.search(pattern, puffing_regime, re.IGNORECASE)
                    if match:
                        puff_time = float(match.group(1))
                        debug_print(f"Extracted puff time: {puff_time}s")
                        return puff_time
                    else:
                        debug_print(f"Could not extract puff time from pattern: '{puffing_regime}'")
                else:
                    debug_print("No puffing regime found at expected position [0,7]")
            else:
                debug_print(f"Insufficient data shape {sample_data.shape} for puff time extraction")
                
        except Exception as e:
            debug_print(f"Error extracting puff time: {e}")
        
        return None
    
    def calculate_normalized_tpm_for_sample(self, sample_data: pd.DataFrame, tpm_data: pd.Series) -> str:
        """
        Calculate normalized TPM for data extraction - returns formatted string.
        
        Args:
            sample_data: DataFrame containing sample data
            tpm_data: Series of TPM values
            
        Returns:
            Formatted normalized TPM value or empty string
        """
        debug_print("Calculating Normalized TPM for data extraction")
        
        try:
            normalized_tpm_series = self.calculate_normalized_tpm(tpm_data, sample_data)
            
            if not normalized_tpm_series.empty:
                avg_normalized_tpm = normalized_tpm_series.mean()
                result = f"{round_values(avg_normalized_tpm, 2):.2f}"
                debug_print(f"Calculated average normalized TPM: {result} mg/s")
                return result
            else:
                debug_print("No valid normalized TPM data")
                return ""
                
        except Exception as e:
            debug_print(f"Error in normalized TPM calculation for sample: {e}")
            return ""
    
    # ===================== USAGE EFFICIENCY CALCULATIONS =====================
    
    def calculate_usage_efficiency(self, sample_data: pd.DataFrame) -> str:
        """
        Calculate usage efficiency using Excel formula logic.
        Formula: ((first_tpm * first_puffs + sum(tpm * incremental_puffs)) / 1000) / initial_oil_mass * 100
        
        Args:
            sample_data: DataFrame containing sample data
            
        Returns:
            Formatted usage efficiency percentage or empty string
        """
        debug_print("Calculating usage efficiency for sample")
        
        try:
            if sample_data.shape[0] < 4 or sample_data.shape[1] < 9:
                debug_print(f"Insufficient data shape {sample_data.shape} for usage efficiency calculation")
                return ""
            
            # Get initial oil mass from H3 (column 7, row 1 with -1 indexing = row 2)
            initial_oil_mass_val = sample_data.iloc[1, 7]
            if pd.isna(initial_oil_mass_val) or initial_oil_mass_val == 0:
                debug_print(f"Invalid initial oil mass: {initial_oil_mass_val}")
                return ""
            
            # Convert to mg
            initial_oil_mass_mg = float(initial_oil_mass_val) * 1000
            debug_print(f"Initial oil mass: {initial_oil_mass_val}g ({initial_oil_mass_mg}mg)")
            
            # Get puffs and TPM values starting from row 4 (row 3 with -1 indexing)
            puffs_values = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce').dropna()
            tpm_values = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
            
            if len(puffs_values) == 0 or len(tpm_values) == 0:
                debug_print(f"No valid puffs ({len(puffs_values)}) or TPM ({len(tpm_values)}) data")
                return ""
            
            # Calculate total aerosol mass following Excel formula logic
            total_aerosol_mass_mg = 0
            
            # Align arrays to same length
            min_length = min(len(puffs_values), len(tpm_values))
            puffs_aligned = puffs_values.iloc[:min_length]
            tpm_aligned = tpm_values.iloc[:min_length]
            
            debug_print(f"Processing {min_length} data points for efficiency calculation")
            
            for i in range(min_length):
                tpm_val = tpm_aligned.iloc[i]
                puffs_val = puffs_aligned.iloc[i]
                
                if not pd.isna(tpm_val) and not pd.isna(puffs_val):
                    if i == 0:
                        # First measurement: TPM * total puffs
                        aerosol_mass = tpm_val * puffs_val
                        debug_print(f"Row {i}: First measurement - {tpm_val} * {puffs_val} = {aerosol_mass}")
                    else:
                        # Subsequent measurements: TPM * (current_puffs - previous_puffs)
                        previous_puffs = puffs_aligned.iloc[i-1]
                        if not pd.isna(previous_puffs):
                            incremental_puffs = puffs_val - previous_puffs
                            aerosol_mass = tpm_val * incremental_puffs
                            debug_print(f"Row {i}: {tpm_val} * ({puffs_val} - {previous_puffs}) = {aerosol_mass}")
                        else:
                            aerosol_mass = tpm_val * puffs_val
                            debug_print(f"Row {i}: Previous puffs NaN, using {tpm_val} * {puffs_val} = {aerosol_mass}")
                    
                    total_aerosol_mass_mg += aerosol_mass
            
            # Calculate usage efficiency: (total aerosol mass / initial oil mass) * 100
            if initial_oil_mass_mg > 0:
                calculated_efficiency = (total_aerosol_mass_mg / initial_oil_mass_mg) * 100
                usage_efficiency = f"{round_values(calculated_efficiency, 1):.1f}%"
                
                debug_print(f"Usage efficiency calculation complete:")
                debug_print(f"  - Total aerosol mass: {round_values(total_aerosol_mass_mg, 2):.2f}mg")
                debug_print(f"  - Initial oil mass: {initial_oil_mass_mg}mg")
                debug_print(f"  - Calculated usage efficiency: {usage_efficiency}")
                
                return usage_efficiency
            else:
                debug_print(f"Invalid initial oil mass for efficiency calculation: {initial_oil_mass_mg}")
                return ""
                
        except Exception as e:
            debug_print(f"Error calculating usage efficiency: {e}")
            import traceback
            traceback.print_exc()
            return ""
    
    def extract_initial_oil_mass(self, sample_data: pd.DataFrame) -> Optional[float]:
        """Extract initial oil mass from sample data."""
        debug_print("Extracting initial oil mass from sample data")
        
        try:
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 7:
                # Check H3 (column 7, row 1 with -1 indexing = row 2)
                oil_mass_val = sample_data.iloc[1, 7]
                if pd.notna(oil_mass_val):
                    oil_mass = float(oil_mass_val)
                    debug_print(f"Extracted initial oil mass: {oil_mass}g")
                    return oil_mass
                else:
                    debug_print("No oil mass found at expected position [1,7]")
            else:
                debug_print(f"Insufficient data shape {sample_data.shape} for oil mass extraction")
                
        except Exception as e:
            debug_print(f"Error extracting initial oil mass: {e}")
        
        return None
    
    # ===================== POWER EFFICIENCY CALCULATIONS =====================
    
    def calculate_power_efficiency(self, tpm_data: pd.Series, sample_data: pd.DataFrame, 
                                   test_type: str = "standard") -> pd.Series:
        """
        Calculate power efficiency from TPM/Power ratio.
        
        Args:
            tpm_data: Series of TPM values
            sample_data: DataFrame containing sample metadata
            test_type: Type of test for metadata extraction
            
        Returns:
            Series of power efficiency values
        """
        debug_print(f"Calculating Power Efficiency for {test_type} test")
        
        try:
            # Convert TPM to numeric
            tpm_numeric = pd.to_numeric(tpm_data, errors='coerce')
            
            # Extract voltage and resistance from metadata
            voltage, resistance = self._extract_electrical_parameters(sample_data, test_type)
            
            # Calculate power and power efficiency
            if voltage and resistance and voltage > 0 and resistance > 0:
                power = (voltage ** 2) / resistance
                debug_print(f"Calculated power: {power:.3f}W (V={voltage}V, R={resistance}Ω)")
                
                power_efficiency = tpm_numeric / power
                valid_values = power_efficiency.dropna()
                debug_print(f"Calculated Power Efficiency for {len(valid_values)} valid data points")
                debug_print(f"Power Efficiency range: {valid_values.min():.3f} - {valid_values.max():.3f}")
                
                return power_efficiency
            else:
                debug_print("Cannot calculate power efficiency - missing or invalid voltage/resistance")
                return pd.Series(dtype=float)
                
        except Exception as e:
            debug_print(f"Error calculating power efficiency: {e}")
            return pd.Series(dtype=float)
    
    def _extract_electrical_parameters(self, sample_data: pd.DataFrame, 
                                       test_type: str = "standard") -> Tuple[Optional[float], Optional[float]]:
        """Extract voltage and resistance from sample metadata."""
        debug_print(f"Extracting electrical parameters for {test_type} test")
        
        voltage = None
        resistance = None
        
        try:
            if test_type == "user_simulation":
                # User Test Simulation layout
                voltage_cell = sample_data.iloc[0, 5] if sample_data.shape[1] > 5 else None
                resistance_cell = sample_data.iloc[0, 3] if sample_data.shape[1] > 3 else None
            else:
                # Standard layout
                voltage_cell = sample_data.iloc[1, 5] if sample_data.shape[1] > 5 else None
                resistance_cell = sample_data.iloc[0, 3] if sample_data.shape[1] > 3 else None
            
            # Extract voltage
            if pd.notna(voltage_cell):
                voltage = float(voltage_cell)
                debug_print(f"Extracted voltage: {voltage}V")
            else:
                debug_print("Could not extract voltage")
            
            # Extract resistance
            if pd.notna(resistance_cell):
                resistance = float(resistance_cell)
                debug_print(f"Extracted resistance: {resistance}Ω")
            else:
                debug_print("Could not extract resistance")
                
        except (ValueError, IndexError, TypeError) as e:
            debug_print(f"Error extracting electrical parameters: {e}")
        
        return voltage, resistance
    
    # ===================== VISCOSITY CALCULATIONS =====================
    
    def calculate_viscosity_prediction(self, media: str, terpene_percent: float, 
                                       temperature: float, potency: float = 0.0, 
                                       terpene_name: str = "Raw") -> Optional[float]:
        """
        Calculate viscosity prediction using available models.
        
        Args:
            media: Media type
            terpene_percent: Terpene percentage (0-100 or 0-1)
            temperature: Temperature in Celsius
            potency: Potency percentage (optional)
            terpene_name: Name of terpene (optional)
            
        Returns:
            Predicted viscosity value or None
        """
        debug_print(f"Calculating viscosity prediction for {media} with {terpene_percent}% terpene at {temperature}°C")
        
        try:
            # Normalize terpene percentage to decimal if needed
            terpene_decimal = terpene_percent / 100.0 if terpene_percent > 1.0 else terpene_percent
            potency_decimal = potency / 100.0 if potency > 1.0 else potency
            
            # Check for available models
            model_key = f"{media}_consolidated"
            
            if hasattr(self, 'consolidated_models') and model_key in self.consolidated_models:
                return self._predict_with_consolidated_model(
                    model_key, terpene_decimal, temperature, potency_decimal, terpene_name
                )
            elif hasattr(self, 'base_models') and f"{media}_base" in self.base_models:
                return self._predict_with_base_model(
                    f"{media}_base", terpene_decimal, temperature, potency_decimal, terpene_name
                )
            else:
                debug_print(f"No viscosity model available for media: {media}")
                return None
                
        except Exception as e:
            debug_print(f"Error in viscosity prediction: {e}")
            return None
    
    def _predict_with_consolidated_model(self, model_key: str, terpene_decimal: float, 
                                         temperature: float, potency_decimal: float, 
                                         terpene_name: str) -> Optional[float]:
        """Predict viscosity using consolidated model."""
        # Placeholder for consolidated model prediction
        # Would implement actual model prediction logic here
        debug_print(f"Using consolidated model {model_key} for prediction")
        return None
    
    def _predict_with_base_model(self, model_key: str, terpene_decimal: float, 
                                 temperature: float, potency_decimal: float, 
                                 terpene_name: str) -> Optional[float]:
        """Predict viscosity using base model."""
        # Placeholder for base model prediction
        # Would implement actual model prediction logic here
        debug_print(f"Using base model {model_key} for prediction")
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
            debug_print("Viscosity calculation saved to database successfully")
            
        except Exception as e:
            debug_print(f"Failed to save viscosity calculation to database: {e}")
    
    # ===================== STATISTICAL CALCULATIONS =====================
    
    def calculate_statistics(self, data: Union[List, pd.Series, np.ndarray]) -> Dict[str, float]:
        """
        Calculate comprehensive statistics for a dataset.
        
        Args:
            data: Input data for statistical analysis
            
        Returns:
            Dictionary containing statistical measures
        """
        debug_print(f"Calculating statistics for {len(data)} data points")
        
        try:
            # Convert to numeric and remove NaN values
            if isinstance(data, pd.Series):
                clean_data = pd.to_numeric(data, errors='coerce').dropna().tolist()
            else:
                clean_data = [float(x) for x in data if pd.notna(x) and str(x).strip() != '']
            
            if len(clean_data) < self.statistical_threshold:
                debug_print(f"Insufficient data points ({len(clean_data)}) for statistics")
                return {}
            
            stats = {
                'count': len(clean_data),
                'mean': statistics.mean(clean_data),
                'median': statistics.median(clean_data),
                'std_dev': statistics.stdev(clean_data) if len(clean_data) > 1 else 0.0,
                'variance': statistics.variance(clean_data) if len(clean_data) > 1 else 0.0,
                'min': min(clean_data),
                'max': max(clean_data),
                'range': max(clean_data) - min(clean_data)
            }
            
            # Add percentiles
            sorted_data = sorted(clean_data)
            stats['q25'] = self._calculate_percentile(sorted_data, 25)
            stats['q75'] = self._calculate_percentile(sorted_data, 75)
            stats['iqr'] = stats['q75'] - stats['q25']
            
            # Round all values
            for key, value in stats.items():
                if isinstance(value, float):
                    stats[key] = round_values(value, self.precision_digits)
            
            debug_print(f"Statistics calculated: mean={stats['mean']}, std_dev={stats['std_dev']}")
            return stats
            
        except Exception as e:
            debug_print(f"Error calculating statistics: {e}")
            return {}
    
    def _calculate_percentile(self, sorted_data: List[float], percentile: float) -> float:
        """Calculate specific percentile from sorted data."""
        if not sorted_data:
            return 0.0
        
        index = (percentile / 100.0) * (len(sorted_data) - 1)
        if index.is_integer():
            return sorted_data[int(index)]
        else:
            lower = sorted_data[int(index)]
            upper = sorted_data[int(index) + 1]
            return lower + (upper - lower) * (index - int(index))
    
    def calculate_single_row_tpm(self, sample_data: Dict[str, List], sample_id: str, row_idx: int):
        """Calculate TPM for a single row in data collection context."""
        debug_print(f"Calculating TPM for {sample_id}, row {row_idx}")
        
        try:
            data = sample_data[sample_id]
            
            if row_idx >= len(data.get("puffs", [])) or row_idx >= len(data.get("before_weight", [])):
                debug_print(f"Row index {row_idx} out of range for {sample_id}")
                return
            
            # Get the data for this row
            current_puffs = data["puffs"][row_idx]
            before_weight = data["before_weight"][row_idx]
            after_weight = data["after_weight"][row_idx]
            
            # Validate data
            if not all([current_puffs, before_weight, after_weight]):
                debug_print(f"Missing data for TPM calculation in row {row_idx}")
                data["tpm"][row_idx] = None
                return
            
            try:
                current_puffs = float(current_puffs)
                before_weight = float(before_weight)
                after_weight = float(after_weight)
            except (ValueError, TypeError):
                debug_print(f"Invalid numeric data in row {row_idx}")
                data["tpm"][row_idx] = None
                return
            
            # Calculate puffing interval
            if row_idx == 0:
                puffing_interval = current_puffs
            else:
                prev_puffs = data["puffs"][row_idx - 1]
                if prev_puffs and str(prev_puffs).strip():
                    try:
                        prev_puffs = float(prev_puffs)
                        puffing_interval = current_puffs - prev_puffs
                        if puffing_interval <= 0:
                            puffing_interval = current_puffs
                    except (ValueError, TypeError):
                        puffing_interval = current_puffs
                else:
                    puffing_interval = current_puffs
            
            # Calculate TPM
            if puffing_interval > 0:
                weight_diff_mg = (before_weight - after_weight) * 1000  # Convert to mg
                tpm = weight_diff_mg / puffing_interval
                data["tpm"][row_idx] = round_values(tpm, 5)
                debug_print(f"TPM calculated: {tpm:.5f} mg/puff")
            else:
                debug_print(f"Invalid puffing interval: {puffing_interval}")
                data["tpm"][row_idx] = None
                
        except Exception as e:
            debug_print(f"Error calculating single row TPM: {e}")
            if sample_id in sample_data and "tpm" in sample_data[sample_id]:
                sample_data[sample_id]["tpm"][row_idx] = None
    
    # ===================== DATA CONVERSION UTILITIES =====================
    
    def convert_to_numeric(self, value: Any, default: float = 0.0) -> float:
        """Convert value to numeric with error handling."""
        try:
            if pd.isna(value) or value == '' or value is None:
                return default
            return float(str(value).replace(',', ''))
        except (ValueError, TypeError):
            debug_print(f"Could not convert '{value}' to numeric, using default {default}")
            return default
    
    def validate_percentage(self, value: float, convert_if_over_one: bool = True) -> float:
        """Validate and normalize percentage values."""
        try:
            numeric_value = float(value)
            if convert_if_over_one and numeric_value > 1.0:
                return numeric_value / 100.0
            return numeric_value
        except (ValueError, TypeError):
            debug_print(f"Invalid percentage value: {value}")
            return 0.0
    
    def parse_puffing_regime(self, regime_text: str) -> Dict[str, Optional[float]]:
        """Parse puffing regime text to extract parameters."""
        debug_print(f"Parsing puffing regime: '{regime_text}'")
        
        result = {
            'volume': None,
            'duration': None,
            'interval': None
        }
        
        try:
            if not regime_text or pd.isna(regime_text):
                return result
            
            regime_str = str(regime_text).strip()
            
            # Pattern for volume/duration/interval: e.g., "55mL/3.0s/30s"
            pattern = r'(\d+(?:\.\d+)?)mL/(\d+(?:\.\d+)?)s/(\d+(?:\.\d+)?)s'
            match = re.search(pattern, regime_str, re.IGNORECASE)
            
            if match:
                result['volume'] = float(match.group(1))
                result['duration'] = float(match.group(2))
                result['interval'] = float(match.group(3))
                debug_print(f"Parsed: {result['volume']}mL, {result['duration']}s duration, {result['interval']}s interval")
            else:
                debug_print(f"Could not parse puffing regime pattern: '{regime_str}'")
                
        except Exception as e:
            debug_print(f"Error parsing puffing regime: {e}")
        
        return result
    
    # ===================== DATABASE OPERATIONS =====================
    
    def _load_formulation_database(self):
        """Load formulation database from file."""
        db_file = Path("data/formulation_database.json")
        
        try:
            if db_file.exists():
                with open(db_file, 'r') as f:
                    self.formulation_database = json.load(f)
                debug_print(f"Loaded formulation database with {len(self.formulation_database)} entries")
            else:
                debug_print("No existing formulation database found, creating new one")
                self.formulation_database = {}
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
    
    def get_service_status(self) -> Dict[str, Any]:
        """Get comprehensive service status information."""
        return {
            'initialized': True,
            'formulation_entries': len(self.formulation_database),
            'consolidated_models': len(self.consolidated_models),
            'base_models': len(self.base_models),
            'other_models': len(self.viscosity_models),
            'default_puff_time': self.default_puff_time,
            'default_puffs': self.default_puffs,
            'precision_digits': self.precision_digits,
            'statistical_threshold': self.statistical_threshold
        }
    
    def reset_service(self):
        """Reset the calculation service to initial state."""
        debug_print("Resetting CalculationService")
        
        self.formulation_database.clear()
        self.viscosity_models.clear()
        self.consolidated_models.clear()
        self.base_models.clear()
        
        # Reload from files
        self._load_formulation_database()
        self._load_viscosity_models()
        
        debug_print("CalculationService reset completed")
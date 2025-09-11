# models/report_model.py
"""
models/report_model.py
Comprehensive report generation model with consolidated calculation services.
Consolidates functionality from report_generator.py, processing.py, viscosity_calculator/models.py,
and all calculation-related services into a unified report model.
"""

import os
import re
import json
import math
import statistics
from typing import Dict, List, Optional, Any, Union, Tuple, Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import pandas as pd
import numpy as np


def debug_print(message: str):
    """Debug print function for report operations."""
    print(f"DEBUG: ReportModel - {message}")


def round_values(value: float, decimal_places: int = 3) -> float:
    """Round values for display consistency."""
    try:
        return round(float(value), decimal_places)
    except (ValueError, TypeError):
        return 0.0


@dataclass
class ReportConfig:
    """Configuration for report generation."""
    report_type: str = "test"  # test, full, comparison
    include_plots: bool = True
    include_images: bool = True
    include_raw_data: bool = False
    output_format: str = "excel"  # excel, powerpoint, pdf
    template_path: Optional[str] = None
    output_directory: str = "reports"
    filename_prefix: str = "report"
    created_at: datetime = field(default_factory=datetime.now)
    
    # TPM calculation parameters
    default_puff_time: float = 3.0
    default_puffs: int = 10
    precision_digits: int = 3
    statistical_threshold: int = 2
    
    def __post_init__(self):
        """Post-initialization processing."""
        debug_print(f"Created ReportConfig for {self.report_type} report")
        debug_print(f"Output format: {self.output_format}, Include plots: {self.include_plots}")


@dataclass
class ReportData:
    """Container for report data and metadata."""
    title: str
    sheets_data: Dict[str, Any] = field(default_factory=dict)
    plot_data: Dict[str, Any] = field(default_factory=dict)
    image_data: Dict[str, Any] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)
    processed_data: Dict[str, pd.DataFrame] = field(default_factory=dict)
    statistical_data: Dict[str, Dict[str, float]] = field(default_factory=dict)
    generation_time: Optional[datetime] = None
    file_path: Optional[str] = None
    
    def __post_init__(self):
        """Post-initialization processing."""
        debug_print(f"Created ReportData '{self.title}' with {len(self.sheets_data)} sheets")


@dataclass
class ViscosityCalculationData:
    """Data structure for viscosity calculations."""
    media: str
    terpene: str
    mass_oil: float
    target_viscosity: float
    result_percent: float
    result_mass: float
    ratio: float
    timestamp: datetime = field(default_factory=datetime.now)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for storage."""
        return {
            'media': self.media,
            'terpene': self.terpene,
            'mass_oil': self.mass_oil,
            'target_viscosity': self.target_viscosity,
            'result_percent': self.result_percent,
            'result_mass': self.result_mass,
            'ratio': self.ratio,
            'timestamp': self.timestamp.isoformat()
        }


class ReportModel:
    """
    Comprehensive report model that consolidates all calculation services and report generation.
    This model replaces and consolidates functionality from:
    - report_generator.py
    - processing.py calculation functions
    - viscosity_calculator/models.py
    - data_collection_window.py calculations
    """
    
    def __init__(self):
        """Initialize the report model with all calculation services."""
        debug_print("Initializing comprehensive ReportModel")
        
        # Core report configuration
        self.config = ReportConfig()
        self.report_data: Optional[ReportData] = None
        self.generation_progress = 0
        self.is_generating = False
        
        # Calculation databases and models
        self.formulation_database = {}
        self.viscosity_models = {}
        self.consolidated_models = {}
        self.base_models = {}
        
        # Statistical calculation cache
        self.statistics_cache = {}
        self.tpm_calculations_cache = {}
        
        # Load existing databases
        self._load_formulation_database()
        self._load_viscosity_models()
        
        debug_print("ReportModel initialization complete")
        debug_print(f"Loaded {len(self.formulation_database)} formulation entries")
        debug_print(f"Loaded {len(self.consolidated_models)} consolidated models")
    
    # ===================== REPORT CONFIGURATION AND STATUS =====================
    
    def set_config(self, config: ReportConfig):
        """Set the report configuration."""
        self.config = config
        debug_print(f"Set report config for {config.report_type} report")
    
    def start_generation(self, report_data: ReportData):
        """Start report generation process."""
        self.report_data = report_data
        self.is_generating = True
        self.generation_progress = 0
        debug_print(f"Started report generation for '{report_data.title}'")
    
    def update_progress(self, progress: int):
        """Update generation progress."""
        self.generation_progress = progress
        debug_print(f"Report generation progress: {progress}%")
    
    def complete_generation(self, file_path: str):
        """Complete report generation."""
        if self.report_data:
            self.report_data.file_path = file_path
            self.report_data.generation_time = datetime.now()
        self.is_generating = False
        self.generation_progress = 100
        debug_print(f"Report generation completed: {file_path}")
    
    # ===================== TPM CALCULATIONS =====================
    
    def calculate_tpm_from_weights(self, puffs: pd.Series, before_weights: pd.Series, 
                                   after_weights: pd.Series, test_type: str = "standard") -> pd.Series:
        """
        Calculate TPM from weight differences and puffing intervals.
        Consolidated from processing.py and calculation_service.py.
        """
        debug_print("Calculating TPM from weight differences")
        
        try:
            # Convert to numeric and handle errors
            puffs_numeric = pd.to_numeric(puffs, errors='coerce')
            before_numeric = pd.to_numeric(before_weights, errors='coerce')
            after_numeric = pd.to_numeric(after_weights, errors='coerce')
            
            # Calculate weight differences
            weight_diff = before_numeric - after_numeric
            
            # Calculate TPM (weight difference / puffs)
            tpm_values = weight_diff / puffs_numeric
            
            # Handle special test types
            if test_type == "extended":
                # Apply extended test scaling factor
                tpm_values = tpm_values * 1.2
            elif test_type == "intense":
                # Apply intense test scaling factor
                tpm_values = tpm_values * 1.5
            
            # Cache calculation
            cache_key = f"tpm_{test_type}_{len(puffs)}"
            self.tpm_calculations_cache[cache_key] = tpm_values
            
            debug_print(f"Calculated TPM for {len(tpm_values)} data points")
            return tpm_values
            
        except Exception as e:
            debug_print(f"Error calculating TPM: {e}")
            return pd.Series(dtype=float)
    
    def calculate_normalized_tpm(self, sample_data: pd.DataFrame, tpm_data: pd.Series) -> str:
        """
        Calculate normalized TPM by dividing TPM by puff time.
        Consolidated from processing.py.
        """
        debug_print("Calculating Normalized TPM")
        
        try:
            # Convert TPM data to numeric
            tpm_numeric = pd.to_numeric(tpm_data, errors='coerce').dropna()
            if tpm_numeric.empty:
                debug_print("No valid TPM data for normalization")
                return ""
            
            # Extract puffing regime from sample data
            puff_time = None
            puffing_regime = None
            
            if sample_data.shape[0] > 0 and sample_data.shape[1] > 7:
                puffing_regime_cell = sample_data.iloc[0, 7]  # Row 1, Column H
                if pd.notna(puffing_regime_cell):
                    puffing_regime = str(puffing_regime_cell).strip()
                    debug_print(f"Found puffing regime: '{puffing_regime}'")
                    
                    # Extract puff time using regex pattern
                    pattern = r'mL/(\d+(?:\.\d+)?)'
                    match = re.search(pattern, puffing_regime)
                    if match:
                        puff_time = float(match.group(1))
                        debug_print(f"Extracted puff time: {puff_time} seconds")
            
            # Use default puff time if not found
            if puff_time is None:
                puff_time = self.config.default_puff_time
                debug_print(f"Using default puff time: {puff_time} seconds")
            
            # Calculate normalized TPM
            avg_tpm = tpm_numeric.mean()
            normalized_tpm = avg_tpm / puff_time
            
            debug_print(f"Normalized TPM: {normalized_tpm:.6f}")
            return round_values(normalized_tpm, 6)
            
        except Exception as e:
            debug_print(f"Error calculating normalized TPM: {e}")
            return ""
    
    def calculate_usage_efficiency(self, sample_data: pd.DataFrame, tpm_data: pd.Series) -> str:
        """
        Calculate usage efficiency using Excel formula logic.
        Consolidated from processing.py.
        """
        debug_print("Calculating usage efficiency")
        
        try:
            if sample_data.shape[0] <= 3 or sample_data.shape[1] <= 8:
                debug_print("Insufficient data for usage efficiency calculation")
                return ""
            
            # Get initial oil mass from H3 (column 7, row 1 with -1 indexing)
            initial_oil_mass_val = None
            if sample_data.shape[1] > 7:
                initial_oil_mass_val = sample_data.iloc[1, 7]
            
            # Get puffs values from column A starting from row 4
            puffs_values = pd.to_numeric(sample_data.iloc[3:, 0], errors='coerce').dropna()
            
            # Get TPM values from column I starting from row 4
            tpm_values = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
            
            if (initial_oil_mass_val is not None and 
                not pd.isna(initial_oil_mass_val) and 
                not puffs_values.empty and 
                not tpm_values.empty):
                
                initial_oil_mass = float(initial_oil_mass_val)
                
                # Calculate cumulative consumption
                cumulative_tpm = tpm_values.cumsum()
                
                # Calculate efficiency for each point
                efficiency_values = []
                for i, (puffs, cum_tpm) in enumerate(zip(puffs_values, cumulative_tpm)):
                    if initial_oil_mass > 0:
                        efficiency = (cum_tpm / initial_oil_mass) * 100
                        efficiency_values.append(efficiency)
                
                if efficiency_values:
                    final_efficiency = efficiency_values[-1]
                    debug_print(f"Calculated usage efficiency: {final_efficiency:.3f}%")
                    return f"{round_values(final_efficiency, 3)}%"
            
            debug_print("Could not calculate usage efficiency - missing data")
            return ""
            
        except Exception as e:
            debug_print(f"Error calculating usage efficiency: {e}")
            return ""
    
    # ===================== STATISTICAL CALCULATIONS =====================
    
    def calculate_statistics(self, data: Union[List, pd.Series, np.ndarray]) -> Dict[str, float]:
        """
        Calculate comprehensive statistics for a dataset.
        Consolidated from calculation_service.py.
        """
        debug_print("Calculating comprehensive statistics")
        
        try:
            # Convert to numpy array for calculations
            if isinstance(data, (list, pd.Series)):
                data_array = np.array(data, dtype=float)
            else:
                data_array = data.astype(float)
            
            # Remove NaN and infinite values
            clean_data = data_array[np.isfinite(data_array)]
            
            if len(clean_data) < self.config.statistical_threshold:
                debug_print(f"Insufficient data points for statistics: {len(clean_data)}")
                return {}
            
            # Calculate basic statistics
            stats = {
                'count': len(clean_data),
                'mean': float(np.mean(clean_data)),
                'median': float(np.median(clean_data)),
                'std': float(np.std(clean_data, ddof=1)) if len(clean_data) > 1 else 0.0,
                'variance': float(np.var(clean_data, ddof=1)) if len(clean_data) > 1 else 0.0,
                'min': float(np.min(clean_data)),
                'max': float(np.max(clean_data)),
                'range': float(np.max(clean_data) - np.min(clean_data)),
                'q25': float(np.percentile(clean_data, 25)),
                'q75': float(np.percentile(clean_data, 75)),
                'iqr': float(np.percentile(clean_data, 75) - np.percentile(clean_data, 25))
            }
            
            # Calculate coefficient of variation
            if stats['mean'] != 0:
                stats['cv'] = abs(stats['std'] / stats['mean'])
            else:
                stats['cv'] = 0.0
            
            # Round all values
            for key, value in stats.items():
                if isinstance(value, float):
                    stats[key] = round_values(value, self.config.precision_digits)
            
            debug_print(f"Calculated statistics for {stats['count']} data points")
            return stats
            
        except Exception as e:
            debug_print(f"Error calculating statistics: {e}")
            return {}
    
    def calculate_sample_statistics(self, sample_data: pd.DataFrame) -> Dict[str, Any]:
        """
        Calculate comprehensive sample statistics from sample data.
        Consolidated from processing.py and data_collection_window.py.
        """
        debug_print("Calculating sample statistics")
        
        try:
            # Extract TPM data
            tpm_data = pd.Series(dtype=float)
            if sample_data.shape[0] > 3 and sample_data.shape[1] > 8:
                tpm_data = pd.to_numeric(sample_data.iloc[3:, 8], errors='coerce').dropna()
            
            # Calculate basic TPM statistics
            stats = {
                'avg_tpm': round_values(tpm_data.mean()) if not tpm_data.empty else None,
                'std_tpm': round_values(tpm_data.std()) if not tpm_data.empty else None,
                'latest_tpm': round_values(tpm_data.iloc[-1]) if not tpm_data.empty else None,
                'tpm_count': len(tpm_data)
            }
            
            # Extract metadata with bounds checking
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 1:
                stats['media'] = str(sample_data.iloc[0, 1]) if not pd.isna(sample_data.iloc[0, 1]) else ""
                stats['viscosity'] = str(sample_data.iloc[1, 1]) if not pd.isna(sample_data.iloc[1, 1]) else ""
            else:
                stats['media'] = ""
                stats['viscosity'] = ""
            
            # Extract voltage, resistance, power
            if sample_data.shape[0] > 1 and sample_data.shape[1] > 5:
                stats['voltage'] = str(sample_data.iloc[1, 5]) if not pd.isna(sample_data.iloc[1, 5]) else ""
                if sample_data.shape[1] > 3:
                    stats['resistance'] = round_values(sample_data.iloc[0, 3]) if not pd.isna(sample_data.iloc[0, 3]) else ""
                stats['power'] = round_values(sample_data.iloc[0, 5]) if not pd.isna(sample_data.iloc[0, 5]) else ""
            else:
                stats['voltage'] = ""
                stats['resistance'] = ""
                stats['power'] = ""
            
            # Calculate advanced metrics
            stats['normalized_tpm'] = self.calculate_normalized_tpm(sample_data, tpm_data)
            stats['usage_efficiency'] = self.calculate_usage_efficiency(sample_data, tpm_data)
            
            # Calculate comprehensive statistics if TPM data exists
            if not tpm_data.empty:
                comprehensive_stats = self.calculate_statistics(tpm_data)
                stats.update(comprehensive_stats)
            
            debug_print(f"Calculated sample statistics: TPM count={stats['tpm_count']}")
            return stats
            
        except Exception as e:
            debug_print(f"Error calculating sample statistics: {e}")
            return {}
    
    # ===================== VISCOSITY CALCULATIONS =====================
    
    def calculate_viscosity_prediction(self, media: str, terpene: str, mass_oil: float, 
                                       target_viscosity: float) -> Dict[str, float]:
        """
        Calculate viscosity formulation prediction.
        Consolidated from viscosity_calculator/models.py.
        """
        debug_print(f"Calculating viscosity prediction for {media} with {terpene}")
        
        try:
            # Check if we have a model for this media
            model_key = f"{media}_{terpene}"
            
            # Use formulation database if available
            if model_key in self.formulation_database:
                formulation_data = self.formulation_database[model_key]
                avg_ratio = formulation_data.get('ratio', 0.15)  # Default 15% ratio
                
                result_mass = mass_oil * avg_ratio
                result_percent = avg_ratio * 100
                
                debug_print(f"Used database ratio {avg_ratio:.3f} for prediction")
            else:
                # Use default calculation if no historical data
                # Basic viscosity mixing calculation
                base_ratio = 0.15  # Default 15%
                
                # Adjust based on target viscosity (simplified model)
                if target_viscosity > 100:
                    base_ratio *= 1.2  # Higher viscosity needs more terpene
                elif target_viscosity < 50:
                    base_ratio *= 0.8  # Lower viscosity needs less terpene
                
                result_mass = mass_oil * base_ratio
                result_percent = base_ratio * 100
                
                debug_print(f"Used default ratio {base_ratio:.3f} for prediction")
            
            prediction = {
                'result_mass': round_values(result_mass, 3),
                'result_percent': round_values(result_percent, 2),
                'confidence': 0.8 if model_key in self.formulation_database else 0.5
            }
            
            debug_print(f"Viscosity prediction: {prediction['result_mass']}g ({prediction['result_percent']}%)")
            return prediction
            
        except Exception as e:
            debug_print(f"Error calculating viscosity prediction: {e}")
            return {'result_mass': 0.0, 'result_percent': 0.0, 'confidence': 0.0}
    
    def save_viscosity_calculation(self, media: str, terpene: str, mass_oil: float, 
                                   target_viscosity: float, result_percent: float, result_mass: float):
        """
        Save viscosity calculation to database.
        Consolidated from calculation_service.py.
        """
        debug_print("Saving viscosity calculation to database")
        
        try:
            key = f"{media}_{terpene}"
            
            if key not in self.formulation_database:
                self.formulation_database[key] = {
                    'media': media,
                    'terpene': terpene,
                    'calculations': []
                }
            
            calculation = ViscosityCalculationData(
                media=media,
                terpene=terpene,
                mass_oil=mass_oil,
                target_viscosity=target_viscosity,
                result_percent=result_percent,
                result_mass=result_mass,
                ratio=result_percent / 100.0
            )
            
            self.formulation_database[key]['calculations'].append(calculation.to_dict())
            
            # Keep only last 10 calculations per formulation
            if len(self.formulation_database[key]['calculations']) > 10:
                self.formulation_database[key]['calculations'] = self.formulation_database[key]['calculations'][-10:]
            
            # Update average ratio
            calculations = self.formulation_database[key]['calculations']
            ratios = [calc['ratio'] for calc in calculations]
            self.formulation_database[key]['ratio'] = sum(ratios) / len(ratios)
            
            self._save_formulation_database()
            debug_print("Viscosity calculation saved to database successfully")
            
        except Exception as e:
            debug_print(f"Failed to save viscosity calculation to database: {e}")
    
    def analyze_viscosity_models(self) -> List[str]:
        """
        Analyze consolidated viscosity models and generate report.
        Consolidated from viscosity_calculator/models.py.
        """
        debug_print("Analyzing consolidated viscosity models")
        
        report = []
        report.append("Consolidated Viscosity Model Analysis Report")
        report.append("==========================================")
        
        if not self.consolidated_models:
            report.append("No consolidated models available for analysis.")
            debug_print("No consolidated models found for analysis")
            return report
        
        report.append(f"Total Consolidated Models: {len(self.consolidated_models)}")
        report.append("-" * 50)
        
        # Analyze each consolidated model
        for model_key, model in self.consolidated_models.items():
            report.append(f"\nModel: {model_key}")
            
            try:
                # Extract media type from model key
                media = model_key.split('_')[0] if '_' in model_key else model_key
                
                # Get model components
                baseline_model = model.get('baseline_model')
                residual_model = model.get('residual_model')
                baseline_features = model.get('baseline_features', [])
                residual_features = model.get('residual_features', [])
                metadata = model.get('metadata', {})
                
                # Analyze baseline model (Arrhenius temperature relationship)
                report.append("1. Temperature baseline model (Arrhenius)")
                
                if baseline_model and hasattr(baseline_model, 'coef_'):
                    coef = baseline_model.coef_[0] if len(baseline_model.coef_) > 0 else 0
                    intercept = baseline_model.intercept_
                    
                    # Calculate activation energy
                    R = 8.314  # Gas constant (J/mol·K)
                    Ea = coef * R
                    Ea_kJ = Ea / 1000  # Convert to kJ/mol
                    
                    report.append(f"  - Equation: log(viscosity) = {intercept:.4f} + {coef:.4f} * (1/T)")
                    report.append(f"  - Activation energy: {Ea_kJ:.2f} kJ/mol")
                    
                    # Categorize temperature sensitivity
                    if Ea_kJ < 20:
                        report.append("  - Low temperature sensitivity")
                    elif Ea_kJ < 40:
                        report.append("  - Medium temperature sensitivity")
                    else:
                        report.append("  - High temperature sensitivity")
                
                # Analyze residual model
                report.append("\n2. Residual model (terpene and potency effects)")
                if residual_features:
                    report.append(f"  - Features: {', '.join(residual_features[:5])}")
                    if len(residual_features) > 5:
                        report.append(f"  - ... and {len(residual_features) - 5} more features")
                
                # Add performance metrics if available
                if 'r2_score' in metadata:
                    report.append(f"  - R² Score: {metadata['r2_score']:.4f}")
                if 'mse' in metadata:
                    report.append(f"  - MSE: {metadata['mse']:.4f}")
                
            except Exception as e:
                report.append(f"  - Error analyzing model: {str(e)}")
                debug_print(f"Error analyzing model {model_key}: {e}")
        
        # Add analysis summary
        report.append("\n" + "=" * 50)
        report.append("ANALYSIS SUMMARY:")
        report.append("1. Focus on high R² and low MSE values to identify quality models")
        report.append("2. Compare activation energies across media types for temperature sensitivity")
        report.append("3. Review feature importance to understand what drives viscosity")
        report.append("4. For media with poor models, consider collecting more diverse data")
        report.append("5. Physical constraint features should have significant importance")
        
        debug_print("Viscosity model analysis completed")
        return report
    
    # ===================== DATA PROCESSING FOR REPORTS =====================
    
    def process_sheet_for_report(self, sheet_data: pd.DataFrame, sheet_name: str, 
                                 headers_row: int = 3, data_start_row: int = 4) -> Dict[str, Any]:
        """
        Process sheet data for report generation.
        Consolidated from processing.py process_sheet_data and process_generic_sheet.
        """
        debug_print(f"Processing sheet '{sheet_name}' for report")
        
        try:
            # Determine if this is a standard sheet or generic
            is_standard = self._is_standard_sheet(sheet_name)
            
            if is_standard:
                return self._process_standard_sheet(sheet_data, sheet_name)
            else:
                return self._process_generic_sheet(sheet_data, headers_row, data_start_row)
                
        except Exception as e:
            debug_print(f"Error processing sheet {sheet_name}: {e}")
            return {
                'display_df': pd.DataFrame(),
                'additional_output': {},
                'full_sample_data': pd.DataFrame(),
                'statistics': {}
            }
    
    def _is_standard_sheet(self, sheet_name: str) -> bool:
        """Check if sheet name matches standard sheet patterns."""
        standard_sheets = [
            "Quick Screening Test", "Lifetime Test", "Device Life Test", 
            "Horizontal Puffing Test", "Extended Test", "Long Puff Test",
            "Rapid Puff Test", "Intense Test", "Big Headspace Low T Test", 
            "Big Headspace High T Test", "Big Headspace Serial Test",
            "Viscosity Compatibility", "Upside Down Test", "Big Headspace Pocket Test",
            "Low Temperature Stability", "Vacuum Test", "Negative Pressure Test", 
            "User Test Simulation", "User Simulation Test", "Various Oil Compatibility"
        ]
        return sheet_name in standard_sheets or sheet_name == "Sheet1"
    
    def _process_standard_sheet(self, data: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """Process standard format sheet."""
        debug_print(f"Processing standard sheet: {sheet_name}")
        
        try:
            # Extract sample data chunks (every 11 columns)
            samples_data = []
            additional_output = {}
            
            num_samples = data.shape[1] // 11
            debug_print(f"Found {num_samples} samples in sheet")
            
            for sample_index in range(num_samples):
                start_col = sample_index * 11
                end_col = start_col + 11
                
                if end_col <= data.shape[1]:
                    sample_data = data.iloc[:, start_col:end_col].copy()
                    
                    # Calculate sample statistics
                    sample_stats = self.calculate_sample_statistics(sample_data)
                    
                    # Create sample summary
                    sample_summary = self._create_sample_summary(sample_data, sample_stats, sample_index)
                    samples_data.append(sample_summary)
            
            # Create display DataFrame
            if samples_data:
                display_df = pd.DataFrame(samples_data)
            else:
                display_df = pd.DataFrame()
            
            # Combine all sample data for full dataset
            full_sample_data = data.copy()
            
            # Calculate overall statistics
            overall_stats = self._calculate_overall_statistics(samples_data)
            
            return {
                'display_df': display_df,
                'additional_output': additional_output,
                'full_sample_data': full_sample_data,
                'statistics': overall_stats,
                'sample_count': num_samples
            }
            
        except Exception as e:
            debug_print(f"Error processing standard sheet: {e}")
            return {
                'display_df': pd.DataFrame(),
                'additional_output': {},
                'full_sample_data': pd.DataFrame(),
                'statistics': {}
            }
    
    def _process_generic_sheet(self, data: pd.DataFrame, headers_row: int = 3, 
                               data_start_row: int = 4) -> Dict[str, Any]:
        """Process generic format sheet."""
        debug_print("Processing generic sheet")
        
        try:
            # Simple processing for generic sheets
            display_df = data.copy()
            
            # Calculate basic statistics for numeric columns
            numeric_columns = display_df.select_dtypes(include=[np.number]).columns
            statistics = {}
            
            for col in numeric_columns:
                col_stats = self.calculate_statistics(display_df[col].dropna())
                if col_stats:
                    statistics[col] = col_stats
            
            return {
                'display_df': display_df,
                'additional_output': {'type': 'generic'},
                'full_sample_data': display_df,
                'statistics': statistics
            }
            
        except Exception as e:
            debug_print(f"Error processing generic sheet: {e}")
            return {
                'display_df': pd.DataFrame(),
                'additional_output': {},
                'full_sample_data': pd.DataFrame(),
                'statistics': {}
            }
    
    def _create_sample_summary(self, sample_data: pd.DataFrame, sample_stats: Dict[str, Any], 
                               sample_index: int) -> Dict[str, Any]:
        """Create summary dictionary for a single sample."""
        debug_print(f"Creating summary for sample {sample_index + 1}")
        
        # Extract sample name with enhanced logic
        sample_name = self._extract_sample_name(sample_data, sample_index)
        
        # Extract burn/clog/leak status
        burn, clog, leak = self._extract_burn_clog_leak(sample_data, sample_index)
        
        summary = {
            "Sample Name": sample_name,
            "Media": sample_stats.get('media', ''),
            "Viscosity": sample_stats.get('viscosity', ''),
            "Voltage, Resistance, Power": f"{sample_stats.get('voltage', '')} V, "
                                         f"{sample_stats.get('resistance', '')} ohm, "
                                         f"{sample_stats.get('power', '')} W",
            "Average TPM": sample_stats.get('avg_tpm'),
            "Standard Deviation": sample_stats.get('std_tpm'),
            "Usage Efficiency": sample_stats.get('usage_efficiency', ''),
            "Normalized TPM": sample_stats.get('normalized_tpm', ''),
            "Burn?": burn,
            "Clog?": clog,
            "Leak?": leak
        }
        
        debug_print(f"Created summary for sample: {sample_name}")
        return summary
    
    def _extract_sample_name(self, sample_data: pd.DataFrame, sample_index: int) -> str:
        """
        Extract sample name using enhanced logic.
        Consolidated from processing.py extract_sample_name.
        """
        try:
            if sample_data.shape[0] < 2:
                return f"Sample {sample_index + 1}"
            
            # Get headers from first two rows
            headers = []
            for row_idx in range(min(2, sample_data.shape[0])):
                for col_idx in range(sample_data.shape[1]):
                    cell_value = sample_data.iloc[row_idx, col_idx]
                    if pd.notna(cell_value):
                        headers.append(str(cell_value).strip())
            
            sample_name = f"Sample {sample_index + 1}"  # Default
            sample_value = ""
            project_value = ""
            
            # Search for sample name patterns
            for i, header in enumerate(headers):
                header_lower = header.lower()
                
                # New format: "Sample ID:" or "Sample ID.1" patterns
                if (header_lower == "sample id:" or 
                    (header_lower.startswith("sample id.") and header_lower[10:].isdigit())):
                    if i + 1 < len(headers):
                        project_value = str(headers[i + 1]).strip()
                        # Remove pandas suffix if present
                        if '.' in project_value and project_value.split('.')[-1].isdigit():
                            project_value = '.'.join(project_value.split('.')[:-1])
                
                # Old format: "Sample:" patterns
                elif (header_lower == "sample:" or
                      (header_lower.startswith("sample:.") and header_lower[8:].isdigit())):
                    if i + 1 < len(headers):
                        temp_sample_value = str(headers[i + 1]).strip()
                        if not sample_value:
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
                # Fallback: try direct column 5 access
                if sample_data.shape[1] > 5 and sample_data.shape[0] > 0:
                    fallback_value = str(sample_data.iloc[0, 5]).strip()
                    if fallback_value and fallback_value.lower() not in ['nan', 'none', '', 'unnamed: 5']:
                        sample_name = fallback_value
            
            debug_print(f"Extracted sample name: '{sample_name}'")
            return sample_name
            
        except Exception as e:
            debug_print(f"Error extracting sample name: {e}")
            return f"Sample {sample_index + 1}"
    
    def _extract_burn_clog_leak(self, sample_data: pd.DataFrame, sample_index: int) -> Tuple[str, str, str]:
        """Extract burn/clog/leak status from sample data."""
        try:
            # Default values
            burn = "No"
            clog = "No"
            leak = "No"
            
            # Look for status indicators in the data
            # This is a simplified implementation - adjust based on actual data format
            if sample_data.shape[0] > 10:  # Ensure we have enough rows
                # Check for burn indicators
                if any("burn" in str(cell).lower() for cell in sample_data.iloc[-3:, :].values.flatten() if pd.notna(cell)):
                    burn = "Yes"
                
                # Check for clog indicators
                if any("clog" in str(cell).lower() for cell in sample_data.iloc[-3:, :].values.flatten() if pd.notna(cell)):
                    clog = "Yes"
                
                # Check for leak indicators
                if any("leak" in str(cell).lower() for cell in sample_data.iloc[-3:, :].values.flatten() if pd.notna(cell)):
                    leak = "Yes"
            
            return burn, clog, leak
            
        except Exception as e:
            debug_print(f"Error extracting burn/clog/leak status: {e}")
            return "No", "No", "No"
    
    def _calculate_overall_statistics(self, samples_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Calculate overall statistics across all samples."""
        debug_print("Calculating overall statistics")
        
        try:
            if not samples_data:
                return {}
            
            # Extract TPM values
            tpm_values = []
            for sample in samples_data:
                avg_tpm = sample.get("Average TPM")
                if avg_tpm is not None and not pd.isna(avg_tpm):
                    tpm_values.append(float(avg_tpm))
            
            if not tpm_values:
                return {}
            
            # Calculate overall statistics
            overall_stats = self.calculate_statistics(tpm_values)
            overall_stats['sample_count'] = len(samples_data)
            overall_stats['valid_tpm_count'] = len(tpm_values)
            
            debug_print(f"Calculated overall statistics for {len(tpm_values)} valid TPM values")
            return overall_stats
            
        except Exception as e:
            debug_print(f"Error calculating overall statistics: {e}")
            return {}
    
    # ===================== DATA PERSISTENCE =====================
    
    def _load_formulation_database(self):
        """Load formulation database from file."""
        db_file = Path("data/formulation_database.json")
        
        try:
            if db_file.exists():
                with open(db_file, 'r') as f:
                    self.formulation_database = json.load(f)
                debug_print(f"Loaded formulation database with {len(self.formulation_database)} entries")
            else:
                debug_print("No formulation database file found, starting with empty database")
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
            'is_generating': self.is_generating,
            'generation_progress': self.generation_progress,
            'formulation_entries': len(self.formulation_database),
            'consolidated_models': len(self.consolidated_models),
            'base_models': len(self.base_models),
            'other_models': len(self.viscosity_models),
            'statistics_cache_size': len(self.statistics_cache),
            'tpm_cache_size': len(self.tpm_calculations_cache),
            'config': {
                'report_type': self.config.report_type,
                'output_format': self.config.output_format,
                'include_plots': self.config.include_plots,
                'default_puff_time': self.config.default_puff_time,
                'precision_digits': self.config.precision_digits
            }
        }
    
    def clear_cache(self):
        """Clear all calculation caches."""
        self.statistics_cache.clear()
        self.tpm_calculations_cache.clear()
        debug_print("Cleared all calculation caches")
    
    def reset_service(self):
        """Reset the report model to initial state."""
        debug_print("Resetting ReportModel service")
        
        self.report_data = None
        self.generation_progress = 0
        self.is_generating = False
        self.clear_cache()
        
        # Reset to default config
        self.config = ReportConfig()
        
        debug_print("ReportModel service reset complete")
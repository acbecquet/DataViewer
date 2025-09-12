"""
Processing package for the DataViewer Application.
Handles all data processing, sheet-specific processing, and legacy file conversion.
Developed by Charlie Becquet.
"""

# Import core processing functions - these are the main public interface
from .core_processing import (
    get_processing_function,
    get_processing_functions,
    get_valid_plot_options,
    get_y_data_for_plot_type,
    process_generic_sheet,
    process_plot_sheet,
    create_empty_plot_structure,
    create_empty_user_test_simulation_structure
)

# Import all sheet processors - these are used by the dispatcher
from .sheet_processors import (
    process_test_plan,
    process_initial_state_inspection,
    process_quick_screening_test,
    process_device_life_test,
    process_aerosol_temp_test,
    process_user_test,
    process_horizontal_test,
    process_extended_test,
    process_long_puff_test,
    process_rapid_puff_test,
    process_intense_test,
    process_legacy_test,
    process_big_head_low_t_test,
    process_burn_protection_test,
    process_big_head_high_t_test,
    process_big_head_serial_test,
    process_upside_down_test,
    process_pocket_test,
    process_temperature_cycling_test,
    process_high_t_high_humidity_test,
    process_cold_storage_test,
    process_vacuum_test,
    process_viscosity_compatibility_test,
    process_various_oil_test,
    process_quick_sensory_test,
    process_leaching_test,
    process_sheet1,
    process_user_test_simulation,
    process_off_odor_score,
    process_sensory_consistency
)

# Import legacy processing functions
from .legacy_processing import (
    extract_samples_from_old_file,
    extract_samples_from_cart_format,
    process_legacy_file_auto_detect,
    detect_template_format,
    convert_legacy_file_using_template,
    convert_legacy_file_using_template_v2,
    convert_cart_format_to_template
)

# Import data extraction utilities
from .data_extraction import (
    aggregate_sheet_metrics,
    extract_burn_clog_leak_from_raw_data,
    no_efficiency_extracted_data,
    updated_extracted_data_function_with_raw_data,
    updated_extracted_data_function_with_raw_data_old,
    calculate_normalized_tpm_for_sample,
    calculate_usage_efficiency_for_sample,
    extract_initial_oil_mass,
    get_y_data_for_user_test_simulation_plot_type,
    fix_x_axis_sequence
)

# Define what gets imported when someone does "from processing import *"
__all__ = [
    # Core processing
    'get_processing_function',
    'get_processing_functions', 
    'get_valid_plot_options',
    'get_y_data_for_plot_type',
    'process_generic_sheet',
    'process_plot_sheet',
    
    # Sheet processors (most commonly used)
    'process_test_plan',
    'process_initial_state_inspection',
    'process_quick_screening_test',
    'process_device_life_test',
    'process_user_test_simulation',
    
    # Legacy processing
    'extract_samples_from_old_file',
    'process_legacy_file_auto_detect',
    
    # Data extraction
    'aggregate_sheet_metrics',
    'extract_burn_clog_leak_from_raw_data',
    'no_efficiency_extracted_data'
]

# Package metadata
__version__ = "1.0.0"
__author__ = "Charlie Becquet"
__description__ = "Processing package for DataViewer Application"
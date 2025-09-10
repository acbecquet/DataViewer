# utils/debug.py
"""
utils/debug.py
Enhanced debugging utilities for the DataViewer application.
Provides consistent logging and debug output across all modules.
"""

import sys
import traceback
from datetime import datetime
from typing import Any, Optional

# Import the debug flag from helpers
try:
    from .helpers import DEBUG_ENABLED, debug_print as base_debug_print
except ImportError:
    DEBUG_ENABLED = True
    def base_debug_print(*args, **kwargs):
        if DEBUG_ENABLED:
            print(*args, **kwargs)

def debug_print(message: str, *args, **kwargs):
    """Enhanced debug print with timestamp and caller info."""
    if DEBUG_ENABLED:
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        caller_frame = sys._getframe(1)
        caller_name = caller_frame.f_code.co_name
        caller_file = caller_frame.f_code.co_filename.split('/')[-1]
        
        formatted_message = f"[{timestamp}] {caller_file}:{caller_name}() - {message}"
        if args:
            formatted_message += f" {' '.join(map(str, args))}"
        
        print(formatted_message, **kwargs)

def error_print(message: str, exception: Optional[Exception] = None):
    """Print error messages with traceback information."""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    caller_frame = sys._getframe(1)
    caller_name = caller_frame.f_code.co_name
    caller_file = caller_frame.f_code.co_filename.split('/')[-1]
    
    error_msg = f"[{timestamp}] ERROR in {caller_file}:{caller_name}() - {message}"
    
    if exception:
        error_msg += f"\nException: {type(exception).__name__}: {str(exception)}"
        if DEBUG_ENABLED:
            error_msg += f"\nTraceback:\n{traceback.format_exc()}"
    
    print(error_msg, file=sys.stderr)

def success_print(message: str):
    """Print success messages with consistent formatting."""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    caller_frame = sys._getframe(1)
    caller_name = caller_frame.f_code.co_name
    caller_file = caller_frame.f_code.co_filename.split('/')[-1]
    
    success_msg = f"[{timestamp}] SUCCESS in {caller_file}:{caller_name}() - {message}"
    print(success_msg)

def log_function_entry(func_name: str, *args, **kwargs):
    """Log function entry with parameters."""
    if DEBUG_ENABLED:
        args_str = ', '.join(map(str, args)) if args else ''
        kwargs_str = ', '.join(f"{k}={v}" for k, v in kwargs.items()) if kwargs else ''
        params = ', '.join(filter(None, [args_str, kwargs_str]))
        debug_print(f"ENTER {func_name}({params})")

def log_function_exit(func_name: str, result: Any = None):
    """Log function exit with return value."""
    if DEBUG_ENABLED:
        result_str = f" -> {result}" if result is not None else ""
        debug_print(f"EXIT {func_name}{result_str}")

def log_timing_checkpoint(checkpoint_name: str, start_time: float):
    """Log timing checkpoint with elapsed time."""
    import time
    elapsed = time.time() - start_time
    debug_print(f"TIMING: {checkpoint_name}: {elapsed:.3f}s")

def print_dataframe_info(df, name: str = "DataFrame"):
    """Print detailed information about a DataFrame for debugging."""
    if not DEBUG_ENABLED:
        return
    
    debug_print(f"{name} Info:")
    debug_print(f"  Shape: {df.shape}")
    debug_print(f"  Columns: {list(df.columns)}")
    debug_print(f"  Data types: {df.dtypes.to_dict()}")
    debug_print(f"  Memory usage: {df.memory_usage(deep=True).sum()} bytes")
    debug_print(f"  Has NaN: {df.isnull().any().any()}")
    
    if len(df) > 0:
        debug_print(f"  First row: {df.iloc[0].to_dict()}")

def print_exception_details(exception: Exception, context: str = ""):
    """Print detailed exception information for debugging."""
    error_print(f"Exception in {context}", exception)
    
    if DEBUG_ENABLED:
        debug_print("Exception details:")
        debug_print(f"  Type: {type(exception).__name__}")
        debug_print(f"  Message: {str(exception)}")
        debug_print(f"  Args: {exception.args}")
        
        # Print the full traceback
        debug_print("Full traceback:")
        traceback.print_exc()

print("DEBUG: debug.py - Enhanced debugging utilities loaded")
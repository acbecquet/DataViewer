"""
import_manager.py
Lazy loading utilities for optional dependencies.
Handles deferred imports to improve startup time.
"""

from utils import debug_print

# Lazy loading helper functions
def lazy_import_pandas():
    """Lazy import pandas when needed."""
    try:
        import pandas as pd
        return pd
    except ImportError as e:
        debug_print(f"Error importing pandas: {e}")
        return None

def lazy_import_numpy():
    """Lazy import numpy when needed."""
    try:
        import numpy as np
        return np
    except ImportError as e:
        debug_print(f"Error importing numpy: {e}")
        return None

def lazy_import_matplotlib():
    """Lazy import matplotlib when needed."""
    try:
        import matplotlib
        matplotlib.use('TkAgg')
        return matplotlib
    except ImportError as e:
        debug_print(f"Error importing matplotlib: {e}")
        return None

def lazy_import_tkintertable():
    """Lazy import tkintertable when needed."""
    try:
        from tkintertable import TableCanvas, TableModel
        return TableCanvas, TableModel
    except ImportError as e:
        debug_print(f"Error importing tkintertable: {e}")
        return None, None

def lazy_import_tksheet():
    """Lazy import tksheet when needed."""
    try:
        from tksheet import Sheet
        return Sheet
    except ImportError as e:
        debug_print(f"Error importing tksheet: {e}")
        return None

def lazy_import_requests():
    """Lazy import requests when needed."""
    try:
        import requests
        return requests
    except ImportError as e:
        debug_print(f"Error importing requests: {e}")
        return None

def lazy_import_packaging():
    """Lazy import packaging when needed."""
    try:
        from packaging import version
        return version
    except ImportError as e:
        debug_print(f"Error importing packaging: {e}")
        return None

def _lazy_import_processing():
    """Lazy import processing module."""
    try:
        import processing
        from processing import get_valid_plot_options
        return processing, get_valid_plot_options
    except ImportError as e:
        debug_print(f"Error importing processing: {e}")
        return None, None

def lazy_import_viscosity_gui():
    """Lazy import viscosity GUI when needed."""
    try:
        from viscosity_gui import ViscosityGUI
        return ViscosityGUI
    except ImportError as e:
        debug_print(f"Error importing viscosity GUI: {e}")
        return None
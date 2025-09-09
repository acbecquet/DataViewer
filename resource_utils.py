# resource_utils.py - Modern resource utilities (pkg_resources-free)
import os
import sys
from pathlib import Path

def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for development and PyInstaller."""
    
    # Check if running as PyInstaller executable
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
        print(f"DEBUG: Using PyInstaller path: {base_path}")
    else:
        # Development mode - use current directory
        base_path = os.path.abspath(".")
        print(f"DEBUG: Using development path: {base_path}")

    full_path = os.path.join(base_path, relative_path)
    # Normalize path separators for cross-platform compatibility
    full_path = os.path.normpath(full_path)
    
    print(f"DEBUG: Resource path for {relative_path}: {full_path}")
    print(f"DEBUG: Resource exists: {os.path.exists(full_path)}")
    
    return full_path

def get_resource_dir() -> str:
    """Get the base resource directory."""
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    else:
        return os.path.abspath(".")

def resource_exists(relative_path: str) -> bool:
    """Check if a resource exists."""
    return os.path.exists(get_resource_path(relative_path))
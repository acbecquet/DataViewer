# resource_utils.py - Lightweight resource utilities
import os
import sys

def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for development and installed package."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        print(f"DEBUG: Using PyInstaller path: {base_path}")
    except AttributeError:
        # Check if running as installed package
        try:
            import pkg_resources
            base_path = pkg_resources.resource_filename(__name__, '.')
            print(f"DEBUG: Using package resource path: {base_path}")
        except ImportError:
            # Development mode
            base_path = os.path.abspath(".")
            print(f"DEBUG: Using development path: {base_path}")

    full_path = os.path.join(base_path, relative_path)
    # Normalize path separators for cross-platform compatibility
    full_path = os.path.normpath(full_path)
    
    print(f"DEBUG: Resource path for {relative_path}: {full_path}")
    print(f"DEBUG: Resource exists: {os.path.exists(full_path)}")
    
    return full_path
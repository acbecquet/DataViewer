"""
Sensory Utility Functions
Contains utility functions for lazy importing and helper operations
"""

import tkinter as tk
from tkinter import messagebox
from utils import debug_print


def _lazy_import_cv2():
    """Lazy import opencv."""
    try:
        import cv2
        debug_print("TIMING: Lazy loaded cv2 for image processing")
        return cv2
    except ImportError as e:
        debug_print(f"Error importing cv2: {e}")
        messagebox.showerror("Missing Dependency",
                            "OpenCV is required for image processing.\nPlease install: pip install opencv-python")
        return None

def _lazy_import_pil():
    """Lazy import PIL."""
    try:
        from PIL import Image, ImageTk
        debug_print("TIMING: Lazy loaded PIL for image processing")
        return Image, ImageTk
    except ImportError as e:
        debug_print(f"Error importing PIL: {e}")
        messagebox.showerror("Missing Dependency",
                            "PIL is required for image processing.\nPlease install: pip install Pillow")
        return None, None

def _lazy_import_sklearn():
    """Lazy import scikit-learn."""
    try:
        from sklearn.cluster import DBSCAN
        debug_print("TIMING: Lazy loaded sklearn for image processing")
        return DBSCAN
    except ImportError as e:
        debug_print(f"Error importing sklearn: {e}")
        messagebox.showerror("Missing Dependency",
                            "Scikit-learn is required for advanced image processing.\nPlease install: pip install scikit-learn")
        return None

def _lazy_import_pytesseract():
    """Lazy import pytesseract."""
    try:
        import pytesseract
        debug_print("TIMING: Lazy loaded pytesseract for OCR")
        return pytesseract
    except ImportError as e:
        debug_print(f"Error importing pytesseract: {e}")
        messagebox.showerror("Missing Dependency",
                            "Pytesseract is required for text recognition.\nPlease install: pip install pytesseract")
        return None
"""
controllers/__init__.py
Controllers package for the DataViewer application.
Controllers coordinate between views and services, handling business logic flow.
"""

# Import all controller classes for easy access
from .main_controller import MainController
from .file_controller import FileController
from .plot_controller import PlotController
from .report_controller import ReportController
from .data_controller import DataController
from .image_controller import ImageController
from .calculation_controller import CalculationController

__all__ = [
    'MainController',
    'FileController', 
    'PlotController',
    'ReportController',
    'DataController',
    'ImageController',
    'CalculationController'
]

# Debug output for architecture tracking
print("DEBUG: Controllers package initialized successfully")
print(f"DEBUG: Available controllers: {', '.join(__all__)}")
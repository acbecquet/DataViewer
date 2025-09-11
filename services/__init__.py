"""
services/__init__.py
Services package for the DataViewer application.
Services contain the core business logic and external operations.
"""

# Import all service classes for easy access
from .file_service import FileService
from .plot_service import PlotService
from .report_service import ReportService
from .database_service import DatabaseService
from .image_service import ImageService
from .update_service import UpdateService
from .calculation_service import CalculationService

# Note: ProcessingService functionality has been consolidated into FileService
# for better organization and to eliminate redundancy.

__all__ = [
    'FileService',
    'PlotService', 
    'ReportService',
    'DatabaseService',
    'ImageService',
    'UpdateService',
    'CalculationService'
]

# Debug output for architecture tracking
print("DEBUG: Services package initialized successfully")
print(f"DEBUG: Available services: {', '.join(__all__)}")
print("DEBUG: FileService now includes consolidated file operations, VAP3 handling, and calculations")
print("DEBUG: ProcessingService functionality consolidated into FileService for better organization")
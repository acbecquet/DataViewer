"""
models/__init__.py
Data models and structures for the DataViewer application.
This package contains all data models, state management, and data structures.
"""

# Import all model classes for easy access
from .data_model import DataModel, SheetData, FileData
from .file_model import FileModel, FileState, FileCache
from .plot_model import PlotModel, PlotConfig, PlotSettings
from .report_model import ReportModel, ReportConfig
from .viscosity_model import ViscosityModel, ViscosityData
from .image_model import ImageModel, ImageMetadata, CropState
from .database_model import DatabaseModel, DatabaseConfig

__all__ = [
    'DataModel', 'SheetData', 'FileData',
    'FileModel', 'FileState', 'FileCache', 
    'PlotModel', 'PlotConfig', 'PlotSettings',
    'ReportModel', 'ReportConfig',
    'ViscosityModel', 'ViscosityData',
    'ImageModel', 'ImageMetadata', 'CropState',
    'DatabaseModel', 'DatabaseConfig'
]

# Debug output for architecture tracking
print("DEBUG: Models package initialized successfully")
print(f"DEBUG: Available models: {', '.join(__all__)}")
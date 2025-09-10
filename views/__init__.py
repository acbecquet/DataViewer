"""
views/__init__.py
Views package for the DataViewer application.
Views handle UI components and presentation logic only.
"""

# Import main view components
from .main_window import MainWindow

# Import dialog components
from .dialogs.progress_dialog import ProgressDialog
from .dialogs.file_dialogs import FileDialogs
from .dialogs.trend_dialog import TrendDialog
from .dialogs.startup_dialog import StartupDialog

# Import widget components  
from .widgets.plot_widget import PlotWidget
from .widgets.table_widget import TableWidget
from .widgets.image_widget import ImageWidget
from .widgets.menu_widget import MenuWidget

# Import calculator views
from .calculators.viscosity_view import ViscosityView

__all__ = [
    'MainWindow',
    'ProgressDialog',
    'FileDialogs', 
    'TrendDialog',
    'StartupDialog',
    'PlotWidget',
    'TableWidget',
    'ImageWidget',
    'MenuWidget',
    'ViscosityView'
]

# Debug output for architecture tracking
print("DEBUG: Views package initialized successfully")
print(f"DEBUG: Available views: {', '.join(__all__)}")
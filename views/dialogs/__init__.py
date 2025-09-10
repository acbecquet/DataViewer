# views/dialogs/__init__.py
"""
views/dialogs/__init__.py
Dialog views package initialization.
"""

from .progress_dialog import ProgressDialog
from .file_dialogs import FileDialogs
from .trend_dialog import TrendDialog
from .startup_dialog import StartupDialog

__all__ = ['ProgressDialog', 'FileDialogs', 'TrendDialog', 'StartupDialog']

print("DEBUG: Views dialogs package initialized")
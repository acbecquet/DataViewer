# views/dialogs/__init__.py
"""
views/dialogs/__init__.py
Dialog views package initialization.
"""

from .progress_dialog import ProgressDialog
from .file_dialogs import FileDialogs, FileSelectionDialog
from .trend_dialog import TrendDialog
from .startup_dialog import StartupDialog, TestSelectionDialog

__all__ = [
    'ProgressDialog',
    'FileDialogs', 
    'FileSelectionDialog',
    'TrendDialog',
    'StartupDialog',
    'TestSelectionDialog'
]

print("DEBUG: Views dialogs package initialized")
print(f"DEBUG: Available dialogs: {', '.join(__all__)}")
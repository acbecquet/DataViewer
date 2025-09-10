# views/widgets/__init__.py
"""
views/widgets/__init__.py
Widget views package initialization.
"""

from .plot_widget import PlotWidget
from .table_widget import TableWidget
from .image_widget import ImageWidget
from .menu_widget import MenuWidget

__all__ = ['PlotWidget', 'TableWidget', 'ImageWidget', 'MenuWidget']

print("DEBUG: Views widgets package initialized")
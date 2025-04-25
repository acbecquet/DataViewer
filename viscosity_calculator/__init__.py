# viscosity_calculator/__init__.py
from .core import ViscosityCalculator

# Import methods from other modules
from .ui import UI_Methods
from .calculations import Calculation_Methods
from .models import Model_Methods
from .data_processing import DataProcessing_Methods
from .terpene_profiles import TerpeneProfile_Methods
from .temperature_blocks import TemperatureBlock_Methods
from .data_management import DataManagement_Methods

# Add methods from each module to the ViscosityCalculator class
for module in [UI_Methods, Calculation_Methods, Model_Methods, 
               DataProcessing_Methods, TerpeneProfile_Methods, 
               TemperatureBlock_Methods, DataManagement_Methods]:
    for name, method in module.__dict__.items():
        if not name.startswith('__') and callable(method):
            setattr(ViscosityCalculator, name, method)
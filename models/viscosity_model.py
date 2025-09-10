# models/viscosity_model.py
"""
models/viscosity_model.py
Viscosity calculation models and data structures.
These models will replace data structures from viscosity_calculator.py.
"""

from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field
from datetime import datetime
import numpy as np


@dataclass
class ViscosityData:
    """Container for viscosity measurement data."""
    temperature: float
    viscosity: float
    sample_id: str
    measurement_time: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created ViscosityData for sample '{self.sample_id}' at {self.temperature}°C")


@dataclass
class ViscosityModel:
    """Model for viscosity calculations and predictions."""
    model_name: str
    model_type: str = "polynomial"  # polynomial, arrhenius, custom
    coefficients: List[float] = field(default_factory=list)
    training_data: List[ViscosityData] = field(default_factory=list)
    accuracy_score: Optional[float] = None
    trained_at: Optional[datetime] = None
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created ViscosityModel '{self.model_name}' of type {self.model_type}")
    
    def add_training_data(self, data: ViscosityData):
        """Add training data to the model."""
        self.training_data.append(data)
        print(f"DEBUG: Added training data to {self.model_name} - total: {len(self.training_data)} points")
    
    def set_trained(self, coefficients: List[float], accuracy: float):
        """Mark model as trained with results."""
        self.coefficients = coefficients
        self.accuracy_score = accuracy
        self.trained_at = datetime.now()
        print(f"DEBUG: Model {self.model_name} trained with accuracy {accuracy:.3f}")


class ViscosityCalculatorModel:
    """Main model for viscosity calculations and model management."""
    
    def __init__(self):
        """Initialize the viscosity calculator model."""
        self.models: Dict[str, ViscosityModel] = {}
        self.raw_data: List[ViscosityData] = []
        self.current_model: Optional[str] = None
        
        print("DEBUG: ViscosityCalculatorModel initialized")
    
    def add_model(self, model: ViscosityModel):
        """Add a viscosity model."""
        self.models[model.model_name] = model
        print(f"DEBUG: Added viscosity model '{model.model_name}'")
    
    def set_current_model(self, model_name: str):
        """Set the currently active model."""
        if model_name in self.models:
            self.current_model = model_name
            print(f"DEBUG: Set current viscosity model to '{model_name}'")
        else:
            print(f"WARNING: Model '{model_name}' not found")
    
    def add_raw_data(self, data: ViscosityData):
        """Add raw viscosity data."""
        self.raw_data.append(data)
        print(f"DEBUG: Added raw viscosity data - total: {len(self.raw_data)} points")
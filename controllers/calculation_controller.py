# controllers/calculation_controller.py
"""
controllers/calculation_controller.py
Specialized calculations controller (viscosity, etc.).
This replaces the coordination logic currently in viscosity_calculator.py.
"""

from typing import Optional, Dict, Any, List
from models.data_model import DataModel
from models.viscosity_model import ViscosityCalculatorModel, ViscosityData, ViscosityModel


class CalculationController:
    """Controller for specialized calculations like viscosity."""
    
    def __init__(self, data_model: DataModel):
        """Initialize the calculation controller."""
        self.data_model = data_model
        self.viscosity_model = ViscosityCalculatorModel()
        
        print("DEBUG: CalculationController initialized")
        print(f"DEBUG: Connected to DataModel and ViscosityCalculatorModel")
    
    def train_viscosity_model(self, model_name: str, training_data: List[ViscosityData]) -> bool:
        """Train a viscosity model with provided data."""
        print(f"DEBUG: CalculationController training viscosity model: {model_name}")
        
        try:
            # Create viscosity model
            model = ViscosityModel(
                model_name=model_name,
                model_type="polynomial"
            )
            
            # Add training data
            for data_point in training_data:
                model.add_training_data(data_point)
            
            # Train model (placeholder)
            # coefficients, accuracy = self.calculation_service.train_model(training_data)
            coefficients = [1.0, 0.5, 0.1]  # Placeholder
            accuracy = 0.95  # Placeholder
            
            model.set_trained(coefficients, accuracy)
            
            # Add to viscosity model
            self.viscosity_model.add_model(model)
            self.viscosity_model.set_current_model(model_name)
            
            print(f"DEBUG: CalculationController trained model {model_name} with accuracy {accuracy}")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to train model: {e}")
            return False
    
    def predict_viscosity(self, temperature: float) -> Optional[float]:
        """Predict viscosity at a given temperature."""
        print(f"DEBUG: CalculationController predicting viscosity at {temperature}°C")
        
        try:
            current_model_name = self.viscosity_model.current_model
            if not current_model_name:
                print("WARNING: No viscosity model selected")
                return None
            
            model = self.viscosity_model.models.get(current_model_name)
            if not model or not model.coefficients:
                print("WARNING: Model not trained")
                return None
            
            # Predict using model coefficients (placeholder calculation)
            # In real implementation, would use proper mathematical model
            prediction = sum(coef * (temperature ** i) for i, coef in enumerate(model.coefficients))
            
            print(f"DEBUG: CalculationController predicted viscosity: {prediction}")
            return prediction
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to predict viscosity: {e}")
            return None
    
    def add_viscosity_data(self, temperature: float, viscosity: float, sample_id: str) -> bool:
        """Add viscosity measurement data."""
        print(f"DEBUG: CalculationController adding viscosity data for {sample_id}")
        
        try:
            data = ViscosityData(
                temperature=temperature,
                viscosity=viscosity,
                sample_id=sample_id
            )
            
            self.viscosity_model.add_raw_data(data)
            
            print(f"DEBUG: CalculationController added viscosity data for {sample_id}")
            return True
            
        except Exception as e:
            print(f"ERROR: CalculationController failed to add viscosity data: {e}")
            return False
    
    def get_available_models(self) -> List[str]:
        """Get list of available viscosity models."""
        return list(self.viscosity_model.models.keys())
    
    def set_active_model(self, model_name: str) -> bool:
        """Set the active viscosity model."""
        if model_name in self.viscosity_model.models:
            self.viscosity_model.set_current_model(model_name)
            return True
        return False
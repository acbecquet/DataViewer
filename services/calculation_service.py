# services/calculation_service.py
"""
services/calculation_service.py
Specialized calculation engines service.
This will contain the calculation logic from viscosity_calculator.py.
"""

from typing import Optional, Dict, Any, List, Tuple
import numpy as np
from scipy import optimize
from sklearn.linear_model import LinearRegression, PolynomialFeatures
from sklearn.metrics import r2_score
import pandas as pd


class CalculationService:
    """Service for specialized calculations like viscosity modeling."""
    
    def __init__(self):
        """Initialize the calculation service."""
        self.trained_models: Dict[str, Dict[str, Any]] = {}
        self.model_types = ['polynomial', 'arrhenius', 'linear', 'exponential']
        
        print("DEBUG: CalculationService initialized")
        print(f"DEBUG: Available model types: {', '.join(self.model_types)}")
    
    def train_polynomial_model(self, temperature_data: List[float], 
                             viscosity_data: List[float], degree: int = 2) -> Tuple[bool, Dict[str, Any], str]:
        """Train a polynomial viscosity model."""
        print(f"DEBUG: CalculationService training polynomial model (degree {degree})")
        
        try:
            if len(temperature_data) != len(viscosity_data):
                return False, {}, "Mismatched data lengths"
            
            if len(temperature_data) < degree + 1:
                return False, {}, f"Insufficient data points for degree {degree} polynomial"
            
            # Convert to numpy arrays
            X = np.array(temperature_data).reshape(-1, 1)
            y = np.array(viscosity_data)
            
            # Create polynomial features
            poly_features = PolynomialFeatures(degree=degree)
            X_poly = poly_features.fit_transform(X)
            
            # Train linear regression on polynomial features
            model = LinearRegression()
            model.fit(X_poly, y)
            
            # Calculate accuracy
            y_pred = model.predict(X_poly)
            r2 = r2_score(y, y_pred)
            
            # Store model information
            model_info = {
                'type': 'polynomial',
                'degree': degree,
                'coefficients': model.coef_.tolist(),
                'intercept': model.intercept_,
                'r2_score': r2,
                'poly_features': poly_features,
                'model': model,
                'training_size': len(temperature_data)
            }
            
            print(f"DEBUG: CalculationService trained polynomial model - R� = {r2:.4f}")
            return True, model_info, "Model trained successfully"
            
        except Exception as e:
            error_msg = f"Failed to train polynomial model: {e}"
            print(f"ERROR: CalculationService - {error_msg}")
            return False, {}, error_msg
    
    def train_arrhenius_model(self, temperature_data: List[float], 
                            viscosity_data: List[float]) -> Tuple[bool, Dict[str, Any], str]:
        """Train an Arrhenius viscosity model."""
        print("DEBUG: CalculationService training Arrhenius model")
        
        try:
            if len(temperature_data) != len(viscosity_data):
                return False, {}, "Mismatched data lengths"
            
            if len(temperature_data) < 3:
                return False, {}, "Insufficient data points for Arrhenius model"
            
            # Convert to numpy arrays
            T = np.array(temperature_data) + 273.15  # Convert to Kelvin
            eta = np.array(viscosity_data)
            
            # Arrhenius equation: eta = A * exp(E/(R*T))
            # Linearized: ln(eta) = ln(A) + E/(R*T)
            # Let x = 1/T, y = ln(eta), then y = ln(A) + (E/R)*x
            
            x = 1.0 / T
            y = np.log(eta)
            
            # Linear regression on linearized form
            model = LinearRegression()
            model.fit(x.reshape(-1, 1), y)
            
            # Extract Arrhenius parameters
            E_over_R = model.coef_[0]  # Activation energy / gas constant
            ln_A = model.intercept_    # Natural log of pre-exponential factor
            A = np.exp(ln_A)          # Pre-exponential factor
            
            # Calculate accuracy
            y_pred = model.predict(x.reshape(-1, 1))
            r2 = r2_score(y, y_pred)
            
            model_info = {
                'type': 'arrhenius',
                'A': A,
                'E_over_R': E_over_R,
                'ln_A': ln_A,
                'r2_score': r2,
                'model': model,
                'training_size': len(temperature_data)
            }
            
            print(f"DEBUG: CalculationService trained Arrhenius model - R� = {r2:.4f}")
            print(f"DEBUG: A = {A:.3e}, E/R = {E_over_R:.1f} K")
            return True, model_info, "Arrhenius model trained successfully"
            
        except Exception as e:
            error_msg = f"Failed to train Arrhenius model: {e}"
            print(f"ERROR: CalculationService - {error_msg}")
            return False, {}, error_msg
    
    def predict_viscosity(self, model_info: Dict[str, Any], temperature: float) -> Tuple[bool, Optional[float], str]:
        """Predict viscosity using a trained model."""
        print(f"DEBUG: CalculationService predicting viscosity at {temperature}캜")
        
        try:
            model_type = model_info.get('type')
            
            if model_type == 'polynomial':
                return self._predict_polynomial(model_info, temperature)
            elif model_type == 'arrhenius':
                return self._predict_arrhenius(model_info, temperature)
            else:
                return False, None, f"Unsupported model type: {model_type}"
            
        except Exception as e:
            error_msg = f"Failed to predict viscosity: {e}"
            print(f"ERROR: CalculationService - {error_msg}")
            return False, None, error_msg
    
    def _predict_polynomial(self, model_info: Dict[str, Any], temperature: float) -> Tuple[bool, float, str]:
        """Predict using polynomial model."""
        try:
            X = np.array([[temperature]])
            poly_features = model_info['poly_features']
            model = model_info['model']
            
            X_poly = poly_features.transform(X)
            prediction = model.predict(X_poly)[0]
            
            print(f"DEBUG: Polynomial prediction: {prediction:.6f}")
            return True, prediction, "Success"
            
        except Exception as e:
            return False, 0.0, f"Polynomial prediction failed: {e}"
    
    def _predict_arrhenius(self, model_info: Dict[str, Any], temperature: float) -> Tuple[bool, float, str]:
        """Predict using Arrhenius model."""
        try:
            A = model_info['A']
            E_over_R = model_info['E_over_R']
            
            T_kelvin = temperature + 273.15
            prediction = A * np.exp(E_over_R / T_kelvin)
            
            print(f"DEBUG: Arrhenius prediction: {prediction:.6f}")
            return True, prediction, "Success"
            
        except Exception as e:
            return False, 0.0, f"Arrhenius prediction failed: {e}"
    
    def analyze_model_performance(self, model_info: Dict[str, Any], 
                                temperature_data: List[float], 
                                viscosity_data: List[float]) -> Tuple[bool, Dict[str, Any], str]:
        """Analyze model performance on test data."""
        print("DEBUG: CalculationService analyzing model performance")
        
        try:
            predictions = []
            for temp in temperature_data:
                success, pred, _ = self.predict_viscosity(model_info, temp)
                if success:
                    predictions.append(pred)
                else:
                    return False, {}, "Failed to generate predictions"
            
            # Calculate performance metrics
            y_true = np.array(viscosity_data)
            y_pred = np.array(predictions)
            
            r2 = r2_score(y_true, y_pred)
            mae = np.mean(np.abs(y_true - y_pred))
            mse = np.mean((y_true - y_pred) ** 2)
            rmse = np.sqrt(mse)
            
            # Calculate relative error
            relative_errors = np.abs((y_true - y_pred) / y_true) * 100
            mean_relative_error = np.mean(relative_errors)
            max_relative_error = np.max(relative_errors)
            
            performance = {
                'r2_score': r2,
                'mean_absolute_error': mae,
                'mean_squared_error': mse,
                'root_mean_squared_error': rmse,
                'mean_relative_error_percent': mean_relative_error,
                'max_relative_error_percent': max_relative_error,
                'num_predictions': len(predictions)
            }
            
            print(f"DEBUG: Model performance - R� = {r2:.4f}, RMSE = {rmse:.6f}")
            return True, performance, "Analysis completed"
            
        except Exception as e:
            error_msg = f"Failed to analyze model performance: {e}"
            print(f"ERROR: CalculationService - {error_msg}")
            return False, {}, error_msg
    
    def compare_models(self, models: Dict[str, Dict[str, Any]], 
                      test_temperature: List[float], 
                      test_viscosity: List[float]) -> Tuple[bool, Dict[str, Dict[str, Any]], str]:
        """Compare multiple models on test data."""
        print(f"DEBUG: CalculationService comparing {len(models)} models")
        
        try:
            comparison_results = {}
            
            for model_name, model_info in models.items():
                success, performance, msg = self.analyze_model_performance(
                    model_info, test_temperature, test_viscosity
                )
                
                if success:
                    comparison_results[model_name] = performance
                    comparison_results[model_name]['model_type'] = model_info.get('type', 'unknown')
                else:
                    print(f"WARNING: Failed to analyze model {model_name}: {msg}")
            
            # Rank models by R� score
            if comparison_results:
                best_model = max(comparison_results.keys(), 
                               key=lambda x: comparison_results[x]['r2_score'])
                
                for model_name in comparison_results:
                    comparison_results[model_name]['is_best'] = (model_name == best_model)
                
                print(f"DEBUG: Best model: {best_model} (R� = {comparison_results[best_model]['r2_score']:.4f})")
            
            return True, comparison_results, "Model comparison completed"
            
        except Exception as e:
            error_msg = f"Failed to compare models: {e}"
            print(f"ERROR: CalculationService - {error_msg}")
            return False, {}, error_msg
    
    def export_model(self, model_info: Dict[str, Any], file_path: str) -> Tuple[bool, str]:
        """Export a trained model to file."""
        print(f"DEBUG: CalculationService exporting model to: {file_path}")
        
        try:
            # Create exportable model data (without sklearn objects)
            exportable_data = {
                'type': model_info.get('type'),
                'r2_score': model_info.get('r2_score'),
                'training_size': model_info.get('training_size'),
                'created_at': str(pd.Timestamp.now())
            }
            
            if model_info.get('type') == 'polynomial':
                exportable_data.update({
                    'degree': model_info.get('degree'),
                    'coefficients': model_info.get('coefficients'),
                    'intercept': model_info.get('intercept')
                })
            elif model_info.get('type') == 'arrhenius':
                exportable_data.update({
                    'A': model_info.get('A'),
                    'E_over_R': model_info.get('E_over_R'),
                    'ln_A': model_info.get('ln_A')
                })
            
            # Save to JSON file
            import json
            with open(file_path, 'w') as f:
                json.dump(exportable_data, f, indent=2)
            
            print(f"DEBUG: CalculationService exported model to: {file_path}")
            return True, "Model exported successfully"
            
        except Exception as e:
            error_msg = f"Failed to export model: {e}"
            print(f"ERROR: CalculationService - {error_msg}")
            return False, error_msg
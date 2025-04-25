import numpy as np
import pandas as pd
import math
import traceback
from scipy.optimize import minimize_scalar
from sklearn.impute import SimpleImputer
import tkinter as tk
from tkinter import messagebox

class Calculation_Methods:
    def calculate_terpene_percentage(self):
        """
        Calculate the terpene percentage needed to achieve target viscosity
        using the two-level model system.
        """
        try:
            # Load models if not already loaded
            if not hasattr(self, 'base_models') or not self.base_models:
                self.load_models()
        
            # Extract input values
            media = self.media_var.get()
            terpene = self.terpene_var.get() or "Raw"
            terpene_brand = self.terpene_brand_var.get()
        
            # Combine terpene name with brand if provided
            if terpene_brand:
                terpene = f"{terpene}_{terpene_brand}"
            
            target_viscosity = float(self.target_viscosity_var.get())
            mass_of_oil = float(self.mass_of_oil_var.get())
        
            # Get potency values (if available)
            potency = self._total_potency_var.get()
            d9_thc = self._d9_thc_var.get()
            d8_thc = self._d8_thc_var.get()
        
            # Calculate total potency if not provided directly
            if potency == 0 and (d9_thc > 0 or d8_thc > 0):
                potency = d9_thc + d8_thc
        
            # Check if we have a model for this media type
            base_model_key = f"{media}_base"
            if base_model_key not in self.base_models:
                raise ValueError(f"No model found for {media}. Please train models first.")
        
            # Use optimization to find optimal terpene percentage
            from scipy.optimize import minimize_scalar
        
            def objective(terpene_pct):
                return abs(self.predict_model_viscosity(media, terpene_pct, 25.0, potency, terpene) - target_viscosity)
        
            # Find optimal terpene percentage (bounded between 0.1% and 15%)
            result = minimize_scalar(objective, bounds=(0.1, 15.0), method='bounded')
            exact_value = result.x
        
            # Calculate mass values
            exact_mass = mass_of_oil * (exact_value / 100)
            start_percent = min(exact_value * 1.1, 15.0)  # Cap at 15%
            start_mass = mass_of_oil * (start_percent / 100)
        
            # Update result variables
            self.exact_percent_var.set(f"{exact_value:.1f}%")
            self.exact_mass_var.set(f"{exact_mass:.2f}g")
            self.start_percent_var.set(f"{start_percent:.1f}%")
            self.start_mass_var.set(f"{start_mass:.2f}g")
        
            # Add constraint check and warning
            if potency > 0:
                theoretical_max_terpene = 100 * (1 - potency/100)
                # Warn if optimization finds a value beyond physical possibility
                if exact_value > theoretical_max_terpene:
                    messagebox.showinfo(
                        "Physical Constraint Notice",
                        f"The calculated terpene percentage ({exact_value:.1f}%) exceeds the "
                        f"theoretical maximum ({theoretical_max_terpene:.1f}%) for a formulation "
                        f"with {potency:.1f}% potency.\n\n"
                        f"Consider either reducing potency or accepting a higher viscosity."
                    )
        
            # Show completion message
            messagebox.showinfo(
                "Calculation Complete", 
                f"Calculation performed using two-level model for {media}\n\n"
                f"For {exact_value:.1f}% terpenes, estimated viscosity: {target_viscosity:.1f}"
            )

        except Exception as e:
            import traceback
            traceback_str = traceback.format_exc()
            print(f"Error during calculation: {e}\n{traceback_str}")
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")

    def predict_model_viscosity(self, media_type_or_model, terpene_pct, temp, potency=None, terpene_name=None):
        """
        Enhanced viscosity prediction using either two-level or consolidated model approach.
    
        Args:
            media_type_or_model: Either a media type string (for two-level models) or model dict (for consolidated)
            terpene_pct: Terpene percentage (0-100 or 0-1)
            temp: Temperature in Celsius
            potency: Potency percentage (0-100 or 0-1), optional
            terpene_name: Name of terpene being used, optional
    
        Returns:
            float: Predicted viscosity in centipoise
        """
        import numpy as np
        import pandas as pd
        from sklearn.impute import SimpleImputer
    
        # Check if we're using a consolidated model (dict) or media type string
        if isinstance(media_type_or_model, dict):
            # Using consolidated model
            model = media_type_or_model
        
            # Extract models and features
            baseline_model = model['baseline_model']
            residual_model = model['residual_model']
            baseline_features = model.get('baseline_features', ['inverse_temp'])
            residual_features = model.get('residual_features', [])
        
            # Get metadata about transformations
            metadata = model.get('metadata', {})
            use_arrhenius = metadata.get('use_arrhenius', True)
        
            # Ensure terpene_pct is in the right range (0-1 for calculations)
            terpene_decimal = terpene_pct / 100.0 if terpene_pct > 1.0 else terpene_pct
        
            # Ensure potency is in the right range (0-1 for calculations)
            if potency is None:
                # Use inverse relationship: potency + terpene = 100%
                potency = 100.0 - (terpene_pct if terpene_pct > 1.0 else terpene_pct * 100.0)
        
            potency_decimal = potency / 100.0 if potency > 1.0 else potency
        
            # Calculate temperature features
            temp_kelvin = temp + 273.15
            inverse_temp = 1.0 / temp_kelvin
        
            # Create baseline feature vector
            baseline_vector = []
        
            for feature in baseline_features:
                if feature == 'inverse_temp':
                    baseline_vector.append(inverse_temp)
                elif feature in ['total_potency']:
                    baseline_vector.append(potency_decimal)
                elif feature in ['terpene_pct']:
                    baseline_vector.append(terpene_decimal)
                else:
                    baseline_vector.append(0)  # Default value
        
            # Get baseline prediction
            baseline_prediction = baseline_model.predict([baseline_vector])[0]
        
            # Create residual feature vector
            residual_vector = []
        
            for feature in residual_features:
                if feature == 'total_potency':
                    residual_vector.append(potency_decimal)
                elif feature == 'terpene_pct':
                    residual_vector.append(terpene_decimal)
                elif feature == 'is_raw':
                    residual_vector.append(1 if terpene_name == 'Raw' else 0)
                elif feature.startswith('terpene_') and feature != 'terpene_pct':
                    # Handle terpene one-hot encoding
                    feature_name = feature[8:]  # Remove 'terpene_' prefix
                    residual_vector.append(1 if feature_name == terpene_name else 0)
                elif feature == 'potency_terpene_ratio':
                    # Calculate ratio with safety
                    residual_vector.append(potency_decimal / max(0.01, terpene_decimal))
                elif feature == 'terpene_headroom':
                    # Calculate headroom (theoretical_max - current)
                    theoretical_max = 1.0 - potency_decimal
                    residual_vector.append(theoretical_max - terpene_decimal)
                elif feature == 'theoretical_max_terpene':
                    # Calculate theoretical max
                    residual_vector.append(1.0 - potency_decimal)
                elif feature == 'terpene_max_ratio':
                    # Calculate as proportion of theoretical maximum
                    theoretical_max = 1.0 - potency_decimal
                    residual_vector.append(terpene_decimal / max(0.01, theoretical_max))
                else:
                    residual_vector.append(0)  # Default value
        
            # Convert to numpy array for prediction
            residual_vector = np.array(residual_vector).reshape(1, -1)
        
            # Check for size mismatch
            if hasattr(residual_model, 'n_features_in_') and residual_model.n_features_in_ != residual_vector.shape[1]:
                # Handle mismatch by padding with zeros
                if residual_vector.shape[1] < residual_model.n_features_in_:
                    padding = np.zeros((1, residual_model.n_features_in_ - residual_vector.shape[1]))
                    residual_vector = np.hstack((residual_vector, padding))
        
            # Check for NaN values
            if np.isnan(residual_vector).any():
                # Use SimpleImputer to replace NaN values
                imputer = SimpleImputer(strategy='constant', fill_value=0)
                residual_vector = imputer.fit_transform(residual_vector)
        
            # Get residual prediction
            residual_prediction = residual_model.predict(residual_vector)[0]
        
            # Combine predictions
            combined_prediction = baseline_prediction + residual_prediction
        
            # Transform back if using Arrhenius
            if use_arrhenius:
                return np.exp(combined_prediction)
            else:
                return combined_prediction
    
        else:
            # Using two-level model system (media_type_or_model is a media type string)
            media = media_type_or_model
        
            # Ensure we have the base models loaded
            if not hasattr(self, 'base_models') or not self.base_models:
                self.load_models()
        
            # Normalize terpene percentage to decimal (0-1)
            terpene_decimal = terpene_pct / 100.0 if terpene_pct > 1.0 else terpene_pct
        
            # Calculate potency if not provided
            if potency is None:
                # Use inverse relationship: potency + terpene = 100%
                potency = 100.0 - (terpene_pct if terpene_pct > 1.0 else terpene_pct * 100.0)
        
            # Normalize potency to decimal (0-1)
            potency_decimal = potency / 100.0 if potency > 1.0 else potency
        
            # Create the model key
            base_model_key = f"{media}_base"
            composition_model_key = f"{media}_composition"
        
            # Check if we have a base model for this media type
            if base_model_key not in self.base_models:
                raise ValueError(f"No base model found for media type: {media}")
        
            # Get the base model
            base_model = self.base_models[base_model_key]
        
            # Calculate temperature features
            temp_kelvin = temp + 273.15
            inverse_temp = 1.0 / temp_kelvin
        
            # Get baseline prediction (temperature effect)
            baseline_model = base_model['baseline_model']
            baseline_pred = baseline_model.predict([[inverse_temp]])[0]
        
            # Get residual prediction (concentration effect)
            residual_model = base_model['residual_model']
            residual_inputs = pd.DataFrame({
                'total_potency': [potency_decimal],
                'terpene_pct': [terpene_decimal]
            })
        
            # Predict using residual model
            base_residual_pred = residual_model.predict(residual_inputs)[0]
        
            # Combine for level 1 prediction
            level1_prediction = baseline_pred + base_residual_pred
        
            # LEVEL 2: COMPOSITION ENHANCEMENT (if available)
            composition_adjustment = 0.0
        
            if terpene_name:
                # First check if we have actual measured profile data
                if hasattr(self, 'composition_models') and hasattr(self, 'terpene_profiles'):
                    # Check if we have a composition model for this media type
                    if composition_model_key in self.composition_models:
                        composition_model = self.composition_models[composition_model_key]
            
                        # Check if we have a profile for this terpene name
                        if media in self.terpene_profiles and terpene_name in self.terpene_profiles[media]:
                            # Get the composition profile
                            profile = self.terpene_profiles[media][terpene_name]
                            has_measured_profile = True
                        else:
                            has_measured_profile = False
                
                        # If no measured profile exists, try using default profiles
                        if not has_measured_profile and hasattr(self, 'default_terpene_profiles'):
                            # Try exact match first
                            if terpene_name in self.default_terpene_profiles:
                                # Found a default profile for this specific terpene
                                default_profile = self.default_terpene_profiles[terpene_name]
                    
                                # Create a profile with values scaled by terpene percentage
                                profile = {}
                                for terpene_compound, percent in default_profile.items():
                                    # Convert the percentage (0-100) to decimal (0-1) and multiply by terpene_pct
                                    scaled_value = (percent / 100.0) * terpene_decimal
                                    profile[terpene_compound] = scaled_value
                    
                                print(f"Using default profile for {terpene_name}, scaled to {terpene_pct}% total terpenes")
                                has_measured_profile = True
                            else:
                                # Try to find a suitable default profile based on substring matching
                                matched_profile = None
                                for profile_name, profile_data in self.default_terpene_profiles.items():
                                    # Skip the 'Generic' profile initially to prefer more specific matches
                                    if profile_name == 'Generic':
                                        continue
                            
                                    # Check if the profile name appears in the terpene name or vice versa
                                    if profile_name.lower() in terpene_name.lower() or terpene_name.lower() in profile_name.lower():
                                        matched_profile = profile_name
                                        break
                    
                                # If no match, use 'Indica', 'Sativa', or 'Generic' based on name
                                if not matched_profile:
                                    if 'indica' in terpene_name.lower():
                                        matched_profile = 'Indica'
                                    elif 'sativa' in terpene_name.lower():
                                        matched_profile = 'Sativa'
                                    else:
                                        matched_profile = 'Generic'
                    
                                # Use the matched profile
                                if matched_profile:
                                    default_profile = self.default_terpene_profiles[matched_profile]
                        
                                    # Create a profile with values scaled by terpene percentage
                                    profile = {}
                                    for terpene_compound, percent in default_profile.items():
                                        # Convert the percentage (0-100) to decimal (0-1) and multiply by terpene_pct
                                        scaled_value = (percent / 100.0) * terpene_decimal
                                        profile[terpene_compound] = scaled_value
                        
                                    print(f"Using default '{matched_profile}' profile for {terpene_name}, scaled to {terpene_pct}% total terpenes")
                                    has_measured_profile = True
            
                        # If we have a valid profile (measured or default), use it
                        if has_measured_profile:
                            # Create feature vector for the composition model
                            comp_features = composition_model['features']
                            comp_inputs = pd.DataFrame(columns=comp_features)
                
                            # Fill in values from the profile
                            for feature in comp_features:
                                if feature in profile:
                                    comp_inputs.loc[0, feature] = profile[feature]
                                else:
                                    comp_inputs.loc[0, feature] = 0.0
                
                            # Get composition adjustment
                            comp_model = composition_model['model']
                            composition_adjustment = comp_model.predict(comp_inputs)[0]
        
            # Combine all components
            final_log_prediction = level1_prediction + composition_adjustment
        
            # Return viscosity by converting from log scale
            return np.exp(final_log_prediction)

    def calculate_step1(self):
        """Calculate the first step amount of terpenes to add."""
        try:
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
            
            # Get estimated raw oil viscosity based on media type
            raw_oil_viscosity = self.get_raw_oil_viscosity(self.media_var.get())
            
            # Calculate amount to add for step 1
            if 2 * target_viscosity < raw_oil_viscosity:
                # Add 1% terpene if target is much lower than raw oil
                percent_to_add = 1.0
            else:
                # Add 0.1% terpene if target is closer to raw oil
                percent_to_add = 0.1
            
            step1_amount = (percent_to_add / 100) * mass_of_oil
            self.step1_amount_var.set(f"{step1_amount:.2f}g")
            
            messagebox.showinfo("Step 1", 
                               f"Add {step1_amount:.2f}g of {self.terpene_var.get()} terpenes.\n"
                               f"Mix thoroughly and then measure the viscosity at 25C.\n"
                               f"Enter the measured viscosity in the field.")
            
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")
    
    def calculate_step2(self):
        """
        Calculate the second step amount of terpenes to add using the consolidated model.
        This method is for the iterative approach when no direct formula is available.
        """
        try:
            # Get the inputs
            mass_of_oil = float(self.mass_of_oil_var.get())
            target_viscosity = float(self.target_viscosity_var.get())
            step1_viscosity_text = self.step1_viscosity_var.get()
    
            if not step1_viscosity_text:
                messagebox.showerror("Input Error", 
                                    "Please enter the measured viscosity from Step 1.")
                return
    
            step1_viscosity = float(step1_viscosity_text)
            step1_amount = float(self.step1_amount_var.get().replace('g', ''))
    
            # Get raw oil viscosity based on media type
            raw_oil_viscosity = self.get_raw_oil_viscosity(self.media_var.get())
    
            # Convert to percentages for calculations
            step1_percent = (step1_amount / mass_of_oil) * 100
    
            # Get media and terpene values
            media = self.media_var.get()
            terpene = self.terpene_var.get() or "Raw"
        
            # Get potency values (if available)
            potency = self._total_potency_var.get()
            # Calculate total potency if not provided directly
            if potency == 0:
                d9_thc = self._d9_thc_var.get()
                d8_thc = self._d8_thc_var.get()
                if d9_thc > 0 or d8_thc > 0:
                    potency = d9_thc + d8_thc
                else:
                    # If no potency is explicitly provided, estimate it from terpene
                    potency = 100 - step1_percent
        
            # First check if we have consolidated models
            have_model = False
            if hasattr(self, 'consolidated_models') and self.consolidated_models:
                have_model = True
        
            # Try to use model-based prediction if available
            if have_model:
                try:
                    # Look for consolidated model for this media
                    model_key = f"{media}_consolidated"
                
                    if model_key not in self.consolidated_models:
                        raise ValueError(f"No consolidated model found for {media}")
                
                    model_info = self.consolidated_models[model_key]
                
                    # Debug the model inputs
                    print(f"Using consolidated model for {media}")
                    print(f"Step 1: {step1_percent}%, viscosity = {step1_viscosity}")
                
                    # Use optimization to find optimal terpene percentage
                    from scipy.optimize import minimize_scalar
                
                    def objective(terpene_pct):
                        # Calculate predicted viscosity at this terpene percentage
                        predicted_visc = self.predict_model_viscosity(model_info, terpene_pct, 25.0, potency, terpene)
                        # Return the absolute error relative to target
                        return abs(predicted_visc - target_viscosity)
                
                    # Start optimization from the current terpene percentage
                    result = minimize_scalar(objective, 
                                            bounds=(step1_percent, 15.0), 
                                            method='bounded')
                
                    # Get the optimal terpene percentage
                    total_percent_needed = result.x
                
                    # Calculate additional percentage needed
                    percent_needed = max(0, total_percent_needed - step1_percent)
                
                    # Calculate amount for step 2
                    step2_amount = (percent_needed / 100) * mass_of_oil
                
                    # Update the UI
                    self.step2_amount_var.set(f"{step2_amount:.2f}g")
                
                    # Predict final viscosity
                    expected_viscosity = self.predict_model_viscosity(model_info, total_percent_needed, 25.0, potency, terpene)
                    self.expected_viscosity_var.set(f"{expected_viscosity:.2f}")
                
                    # Show information about the model used
                    messagebox.showinfo("Step 2 (Model-based)", 
                                       f"Add an additional {step2_amount:.2f}g of {terpene} terpenes.\n"
                                       f"Mix thoroughly and then measure the final viscosity at 25C.\n"
                                       f"The expected final viscosity is {expected_viscosity:.2f}.\n\n"
                                       f"(Using consolidated model for {media})")
                
                    return
                
                except Exception as e:
                    import traceback
                    print(f"Error using model for prediction: {str(e)}")
                    print(traceback.format_exc())
                    # Fall back to exponential calculation below
        
            # Fallback: Use exponential model based on the two measurements
            import math
        
            # Solve for the decay constant in the exponential model
            # viscosity = raw_oil_viscosity * exp(-k * terpene_percentage)
            if step1_viscosity <= 0 or raw_oil_viscosity <= 0:
                raise ValueError("Viscosity values must be positive")
        
            # Calculate decay constant k
            k = -math.log(step1_viscosity / raw_oil_viscosity) / step1_percent
        
            # Solve for the total percentage needed to reach target viscosity
            if target_viscosity <= 0:
                raise ValueError("Target viscosity must be positive")
        
            total_percent_needed = -math.log(target_viscosity / raw_oil_viscosity) / k
        
            # Calculate additional percentage needed
            percent_needed = max(0, total_percent_needed - step1_percent)
        
            # Calculate amount for step 2
            step2_amount = (percent_needed / 100) * mass_of_oil
        
            # Update the UI
            self.step2_amount_var.set(f"{step2_amount:.2f}g")
        
            # Calculate expected final viscosity using the model
            expected_viscosity = raw_oil_viscosity * math.exp(-k * (step1_percent + percent_needed))
            self.expected_viscosity_var.set(f"{expected_viscosity:.2f}")
        
            messagebox.showinfo("Step 2 (Exponential Model)", 
                               f"Add an additional {step2_amount:.2f}g of {terpene} terpenes.\n"
                               f"Mix thoroughly and then measure the final viscosity at 25C.\n"
                               f"The expected final viscosity is {expected_viscosity:.2f}.\n\n"
                               f"(Using exponential model based on your measurements)")
    
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            messagebox.showerror("Calculation Error", f"An error occurred: {str(e)}")

    def predict_with_confidence(self, model, terpene_pct, temp, potency, terpene_profile):
        """
        Make a viscosity prediction with confidence adjustment based on terpene profile.
    
        Args:
            model: The viscosity model
            terpene_pct: Overall terpene percentage
            temp: Temperature in Celsius
            potency: Total potency
            terpene_profile: Terpene profile dict or strain name
        
        Returns:
            tuple: (predicted_viscosity, confidence_factor)
        """
        # If terpene_profile is a string (strain name), get the profile
        if isinstance(terpene_profile, str):
            profile = self.get_terpene_profile(terpene_profile)
        else:
            profile = self.handle_other_terpenes(terpene_profile)
    
        # Extract confidence factor
        confidence = profile.pop('_confidence', 1.0)
    
        # Make prediction (using your existing method)
        predicted_viscosity = self.predict_model_viscosity(model, terpene_pct, temp, potency, profile)
    
        # Return both the prediction and the confidence
        return predicted_viscosity, confidence
import os
import pickle
import threading
import traceback
import math
# Keep only lightweight imports at module level
import tkinter as tk
from tkinter import ttk, Toplevel, Frame, Label, Text, Scrollbar, StringVar, DoubleVar, IntVar, BooleanVar, Scale, Button, HORIZONTAL
from tkinter import messagebox

# Import constants from core module - these should be lightweight
from .core import APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT

try:
    from utils import debug_print
except ImportError:
    def debug_print(msg):
        print(f"DEBUG: {msg}")

class Model_Methods:
    def train_unified_models(self, data=None, alpha=1.0):
        """
        Trains a two-level viscosity prediction system with L2 regularization:
        1. Base model using temperature, potency, and total terpene percentage
        2. Composition enhancement model using detailed terpene profiles
    
        Args:
            data: Optional dataframe to use for training
            alpha: Regularization strength (default=1.0)
        """
        import pandas as pd
        import numpy as np
        from sklearn.linear_model import Ridge
        from sklearn.ensemble import RandomForestRegressor
        from sklearn.impute import SimpleImputer
        from sklearn.model_selection import KFold, cross_val_score
        
        debug_print("Loaded heavy modules for train_unified_models")
        debug_print(f"numpy version: {np.__version__}")
        debug_print(f"pandas version: {pd.__version__}")
        debug_print(f"Starting model training with alpha={alpha}")

        # Create configuration window
        config_window = Toplevel(self.root)
        config_window.title("Train Regularized Two-Level Models")
        config_window.geometry("600x400")
        config_window.transient(self.root)
        config_window.grab_set()
        config_window.configure(bg=APP_BACKGROUND_COLOR)

        # Center the window
        self.gui.center_window(config_window)

        # Configuration variables
        model_type_var = StringVar(value="Ridge")
        alpha_var = tk.DoubleVar(value=1.0)  # Default regularization strength
        features_var = StringVar(value="both")
        cv_folds_var = tk.IntVar(value=5)  # Cross-validation folds
    
        # Create a frame for options
        options_frame = Frame(config_window, bg=APP_BACKGROUND_COLOR, padx=20, pady=20)
        options_frame.pack(fill="both", expand=True)

        # Model type selection
        Label(options_frame, text="Residual Model Type:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=0, column=0, sticky="w", pady=10)

        model_types = ["Ridge", "RandomForest"]
        model_dropdown = ttk.Combobox(
            options_frame, 
            textvariable=model_type_var,
            values=model_types,
            state="readonly",
            width=12
        )
        model_dropdown.grid(row=0, column=1, sticky="w", pady=10)
    
        # L2 Regularization strength (alpha)
        Label(options_frame, text="Regularization Strength (α):", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=1, column=0, sticky="w", pady=10)
    
        alpha_slider = Scale(options_frame, variable=alpha_var, from_=0.001, to=10.0, 
                            resolution=0.001, orient=HORIZONTAL, length=200,
                            bg=APP_BACKGROUND_COLOR, fg="white", troughcolor=BUTTON_COLOR)
        alpha_slider.grid(row=1, column=1, sticky="w", pady=10)
    
        # Option to use both potency and terpene percentage or just one
        Label(options_frame, text="Feature Selection:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=2, column=0, sticky="w", pady=10)
    
        features_frame = Frame(options_frame, bg=APP_BACKGROUND_COLOR)
        features_frame.grid(row=2, column=1, sticky="w", pady=10)
    
        
        tk.Radiobutton(features_frame, text="Use both features", variable=features_var, 
                      value="both", bg=APP_BACKGROUND_COLOR, fg="white",
                      selectcolor=APP_BACKGROUND_COLOR).pack(anchor="w")
    
        tk.Radiobutton(features_frame, text="Use only potency", variable=features_var, 
                      value="potency", bg=APP_BACKGROUND_COLOR, fg="white",
                      selectcolor=APP_BACKGROUND_COLOR).pack(anchor="w")
    
        tk.Radiobutton(features_frame, text="Use only terpene %", variable=features_var, 
                      value="terpene", bg=APP_BACKGROUND_COLOR, fg="white",
                      selectcolor=APP_BACKGROUND_COLOR).pack(anchor="w")
    
        # Cross-validation folds
        Label(options_frame, text="Cross-validation folds:", bg=APP_BACKGROUND_COLOR, fg="white", 
              font=FONT, anchor="w").grid(row=3, column=0, sticky="w", pady=10)
    
        cv_slider = Scale(options_frame, variable=cv_folds_var, from_=3, to=10, 
                         resolution=1, orient=HORIZONTAL, length=200,
                         bg=APP_BACKGROUND_COLOR, fg="white", troughcolor=BUTTON_COLOR)
        cv_slider.grid(row=3, column=1, sticky="w", pady=10)

        # Status label
        status_label = Label(options_frame, text="Click Train Models to begin training regularized models",
                         bg=APP_BACKGROUND_COLOR, fg="white", font=FONT)
        status_label.grid(row=4, column=0, columnspan=2, pady=5)

        # Training thread function
        def train_models_thread(config, status_label, window):
            try:
                # Load data if not provided
                nonlocal data
                if data is None:
                    specific_file = './data/Master_Viscosity_Data_processed.csv'

                    try:
                        if os.path.exists(specific_file):
                            data = pd.read_csv(specific_file)
                            self.root.after(0, lambda: status_label.config(
                                text=f"Loaded {len(data)} data points..."
                            ))
                        else:
                            self.root.after(0, lambda: messagebox.showerror(
                                "Error", f"Master data file not found: {specific_file}"
                            ))
                            window.after(0, window.destroy)
                            return
                    except Exception as e:
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error", f"Failed to load data: {str(e)}"
                        ))
                        window.after(0, window.destroy)
                        return

                # Clean and prepare data
                self.root.after(0, lambda: status_label.config(text="Performing initial data cleaning..."))
                data_cleaned = data.copy()

                # Convert columns to numeric and handle errors
                num_columns = ['terpene_pct', 'temperature', 'viscosity', 'd9_thc', 'd8_thc', 'thca', 'total_potency']
                for col in num_columns:
                    if col in data_cleaned.columns:
                        data_cleaned[col] = pd.to_numeric(data_cleaned[col], errors='coerce')

                # Clean terpene values
                data_cleaned.loc[data_cleaned['terpene'].isna(), 'terpene'] = 'Raw'
                data_cleaned.loc[data_cleaned['terpene'] == '', 'terpene'] = 'Raw'

                # Improved raw media detection
                data_cleaned['is_raw'] = False  # Default to False

                # Check for exact matches between media and terpene (raw case)
                exact_match_mask = data_cleaned['terpene'].str.lower() == data_cleaned['media'].str.lower()
                data_cleaned.loc[exact_match_mask, 'is_raw'] = True

                # Special case for D9/D9 Distillate and D8/D8 Distillate
                d9_mask = (data_cleaned['media'].str.lower() == 'd9') & (data_cleaned['terpene'].str.lower().str.contains('d9.*distillate'))
                d8_mask = (data_cleaned['media'].str.lower() == 'd8') & (data_cleaned['terpene'].str.lower().str.contains('d8.*distillate'))
                data_cleaned.loc[d9_mask | d8_mask, 'is_raw'] = True

                # Check for cases where terpene contains the media name with no other terms
                media_in_terpene_mask = data_cleaned.apply(
                    lambda row: row['media'].lower() in row['terpene'].lower() and 
                                len(row['terpene'].split('/')) == 1,  # No slash separator
                    axis=1
                )
                data_cleaned.loc[media_in_terpene_mask, 'is_raw'] = True

                # Also mark explicit "Raw" values as raw
                raw_mask = data_cleaned['terpene'] == 'Raw'
                data_cleaned.loc[raw_mask, 'is_raw'] = True

                # Log detection results
                raw_count = data_cleaned['is_raw'].sum()
                total_count = len(data_cleaned)
                self.root.after(0, lambda: status_label.config(
                    text=f"Detected {raw_count} raw samples out of {total_count} total samples ({raw_count/total_count*100:.1f}%)"
                ))

                # Fill missing terpene_pct with 0 for raw data
                data_cleaned.loc[data_cleaned['is_raw'] & data_cleaned['terpene_pct'].isna(), 'terpene_pct'] = 0.0
        
                # Handle missing values in potency or terpene percentage
                if 'total_potency' in data_cleaned.columns and 'terpene_pct' in data_cleaned.columns:
                    # Fill missing potency using terpene percentage
                    potency_missing = data_cleaned['total_potency'].isna()
                    terpene_available = ~data_cleaned['terpene_pct'].isna()
                    mask = potency_missing & terpene_available
                    if mask.any():
                        # Ensure terpene_pct is in percentage (0-100)
                        terpene_values = data_cleaned.loc[mask, 'terpene_pct']
                        if terpene_values.max() <= 1.0:  # If decimal format
                            data_cleaned.loc[mask, 'total_potency'] = 1.0 - terpene_values
                        else:  # If percentage format
                            data_cleaned.loc[mask, 'total_potency'] = 100.0 - terpene_values
                    
                        self.root.after(0, lambda n=mask.sum(): status_label.config(
                            text=f"Filled {n} missing potency values using terpene percentage"
                        ))
                
                    # Fill missing terpene percentage using potency
                    terpene_missing = data_cleaned['terpene_pct'].isna()
                    potency_available = ~data_cleaned['total_potency'].isna()
                    mask = terpene_missing & potency_available
                    if mask.any():
                        potency_values = data_cleaned.loc[mask, 'total_potency']
                        if potency_values.max() <= 1.0:  # If decimal format
                            data_cleaned.loc[mask, 'terpene_pct'] = 1.0 - potency_values
                        else:  # If percentage format
                            data_cleaned.loc[mask, 'terpene_pct'] = 100.0 - potency_values
                    
                        self.root.after(0, lambda n=mask.sum(): status_label.config(
                            text=f"Filled {n} missing terpene values using potency"
                        ))
                
                    # Apply feature selection based on user choice
                    if config["features"] == "potency":
                        # Use only potency as a feature
                        data_cleaned['concentration'] = data_cleaned['total_potency']
                        self.root.after(0, lambda: status_label.config(
                            text="Using only potency as composition feature"
                        ))
                    elif config["features"] == "terpene":
                        # Use only terpene percentage as a feature
                        data_cleaned['concentration'] = data_cleaned['terpene_pct']
                        self.root.after(0, lambda: status_label.config(
                            text="Using only terpene percentage as composition feature"
                        ))
                    else:
                        # Use both features (default)
                        self.root.after(0, lambda: status_label.config(
                            text="Using both potency and terpene percentage as features"
                        ))

                    # Add physical constraint features
                    # Calculate theoretical maximum terpene percentage
                    data_cleaned['theoretical_max_terpene'] = 1.0 - data_cleaned['total_potency']

                    # Calculate how close the formulation is to theoretical maximum
                    data_cleaned['terpene_headroom'] = data_cleaned['theoretical_max_terpene'] - data_cleaned['terpene_pct']

                    # Flag physically impossible formulations (allowing for small measurement error)
                    data_cleaned['physically_valid'] = data_cleaned['terpene_pct'] <= (1.05 * data_cleaned['theoretical_max_terpene'])

                    # Calculate ratio as a proportion of theoretical maximum
                    data_cleaned['terpene_max_ratio'] = data_cleaned['terpene_pct'] / data_cleaned['theoretical_max_terpene'].clip(lower=0.01)

                    # Calculate potency to terpene ratio
                    data_cleaned['potency_terpene_ratio'] = data_cleaned['total_potency'] / data_cleaned['terpene_pct'].clip(lower=0.01)

                    # Log information about the constraints
                    valid_pct = 100 * data_cleaned['physically_valid'].mean()
                    self.root.after(0, lambda p=valid_pct: status_label.config(
                        text=f"Added physical constraints. {p:.1f}% of formulations are physically valid."
                    ))

                # Apply Arrhenius transformation
                data_cleaned['temperature_kelvin'] = data_cleaned['temperature'] + 273.15
                data_cleaned['inverse_temp'] = 1 / data_cleaned['temperature_kelvin']
                data_cleaned['log_viscosity'] = np.log(data_cleaned['viscosity'])

                # Create one model per media type
                self.root.after(0, lambda: status_label.config(text="Creating base models by media type..."))

                # Get unique media types
                media_types = data_cleaned['media'].unique()

                # Model creation helper function
                def build_residual_model(config):
                    if config["model_type"] == "Ridge":
                        return Ridge(alpha=config["alpha"])
                    else:
                        return RandomForestRegressor(
                            n_estimators=100,
                            max_depth=4,  # Slightly deeper for better complexity modeling
                            min_samples_leaf=5,
                            random_state=42
                        )

                # Initialize model dictionaries
                base_models = {}
                composition_models = {}
                terpene_profiles = {}
                cv_scores = {}  # To store cross-validation scores
        
                # Identify terpene composition columns if they exist in the data
                possible_terpene_columns = [
                    'alpha-Pinene', 'Camphene', 'beta-Pinene', 'beta-Myrcene', '3-Carene', 
                    'alpha-Terpinene', 'p-Cymene', 'D-Limonene', 'Ocimene 1', 'Ocimene 2',
                    'gamma-Terpinene', 'Terpinolene', 'Linalool', 'Isopulegol', 'Geraniol',
                    'Caryophyllene', 'alpha-Humulene', 'Nerolidol 1', 'Nerolidol 2', 
                    'Guaiol', 'alpha-Bisabolol'
                ]
        
                # Filter to columns that actually exist in the data
                terpene_composition_columns = [col for col in possible_terpene_columns if col in data_cleaned.columns]
                has_composition_data = len(terpene_composition_columns) > 0
        
                # Flag rows that have composition data
                if has_composition_data:
                    data_cleaned['has_composition'] = data_cleaned[terpene_composition_columns].notna().any(axis=1)
                    composition_count = data_cleaned['has_composition'].sum()
                    self.root.after(0, lambda: status_label.config(
                        text=f"Found {composition_count} samples with detailed terpene composition data"
                    ))
                else:
                    data_cleaned['has_composition'] = False
                    self.root.after(0, lambda: status_label.config(
                        text="No detailed terpene composition data found. Using basic model only."
                    ))

                # Set up cross-validation
                cv = KFold(n_splits=config["cv_folds"], shuffle=True, random_state=42)

                # Process each media type
                for media_idx, media in enumerate(media_types):
                    # Update progress
                    progress = f"Training model {media_idx+1}/{len(media_types)}: {media}"
                    self.root.after(0, lambda p=progress: status_label.config(text=p))
        
                    # Filter data for this media type
                    media_data = data_cleaned[data_cleaned['media'] == media].copy()
        
                    # Skip if not enough data
                    if len(media_data) < 10:
                        self.root.after(0, lambda m=media: status_label.config(
                            text=f"Skipping {m} - insufficient data ({len(media_data)} points)"
                        ))
                        continue
        
                    # Drop rows with NaN in critical columns
                    critical_cols = ['temperature', 'viscosity', 'inverse_temp', 'log_viscosity']
                    media_data = media_data.dropna(subset=critical_cols)
        
                    if len(media_data) < 10:
                        continue
        
                    # --------- LEVEL 1: BASE MODEL ---------
            
                    # Features for temperature baseline
                    base_X_temp = media_data[['inverse_temp']]
                    base_y_temp = media_data['log_viscosity']
            
                    # Create temperature baseline model
                    temp_model = Ridge(alpha=0.1)  # Light regularization for temperature model
                    temp_model.fit(base_X_temp, base_y_temp)
            
                    # Calculate baseline predictions and residuals
                    baseline_preds = temp_model.predict(base_X_temp)
                    media_data['baseline_prediction'] = baseline_preds
                    media_data['residual'] = base_y_temp - baseline_preds
            
                    # Determine residual model features based on user selection
                    if config["features"] == "potency":
                        residual_features = ['total_potency']
                    elif config["features"] == "terpene":
                        residual_features = ['terpene_pct']
                    else:
                        # Use both with physical constraints
                        residual_features = ['total_potency', 'terpene_pct', 
                                           'terpene_headroom', 'potency_terpene_ratio']
                
                    # Add is_raw flag to all models
                    residual_features.append('is_raw')
            
                    # Train residual model for base level
                    base_X_residual = media_data[residual_features].copy()
            
                    # Handle NaN values
                    base_X_residual = self.check_features_for_nan(base_X_residual)
                    base_y_residual = media_data['residual']
            
                    # Train residual model with cross-validation
                    residual_model = build_residual_model(config)
                
                    # Perform cross-validation
                    try:
                        cv_results = cross_val_score(
                            residual_model, 
                            base_X_residual, 
                            base_y_residual,
                            cv=cv,
                            scoring='r2'
                        )
                    
                        # Store CV scores
                        cv_scores[media] = {
                            'mean_r2': cv_results.mean(),
                            'std_r2': cv_results.std(),
                            'all_scores': cv_results.tolist()
                        }
                    
                        self.root.after(0, lambda m=media, r=cv_results.mean(): status_label.config(
                            text=f"{m}: Cross-validation R² = {r:.3f}"
                        ))
                    
                    except Exception as e:
                        self.root.after(0, lambda m=media, err=str(e): status_label.config(
                            text=f"Warning: CV failed for {m} - {err}"
                        ))
                
                    # Fit final model on all data
                    residual_model.fit(base_X_residual, base_y_residual)
            
                    # Store base model
                    base_models[f"{media}_base"] = {
                        'baseline_model': temp_model,
                        'residual_model': residual_model,
                        'baseline_features': ['inverse_temp'],
                        'residual_features': residual_features,
                        'metadata': {
                            'use_arrhenius': True,
                            'temperature_feature': 'inverse_temp',
                            'target_feature': 'log_viscosity',
                            'use_two_step': True,
                            'model_type': 'base',
                            'regularization': {
                                'type': config["model_type"],
                                'alpha': config["alpha"] if config["model_type"] == "Ridge" else None
                            },
                            'cv_results': cv_scores.get(media, None)
                        }
                    }
            
                    # --------- LEVEL 2: COMPOSITION MODEL ---------
            
                    # Check if we have composition data for this media type
                    if has_composition_data:
                        # Filter to samples with composition data
                        comp_data = media_data[media_data['has_composition']].copy()
                
                        if len(comp_data) >= 5:  # Need some minimum number of samples
                            self.root.after(0, lambda m=media, n=len(comp_data): status_label.config(
                                text=f"Training composition model for {m} using {n} detailed profiles"
                            ))
                    
                            # Get base model predictions for these samples
                            comp_X_base = comp_data[residual_features]
                            comp_base_residual_preds = residual_model.predict(comp_X_base)
                    
                            # Calculate new residuals after base model
                            comp_data['base_residual_prediction'] = comp_base_residual_preds
                            comp_data['composition_residual'] = comp_data['residual'] - comp_base_residual_preds
                    
                            # Train composition model on the new residuals
                            comp_X = comp_data[terpene_composition_columns].copy()
                            comp_X = self.check_features_for_nan(comp_X)  # Handle NaN values
                            comp_y = comp_data['composition_residual']
                    
                            # Use a simpler model for composition effects
                            if config["model_type"] == "Ridge":
                                comp_model = Ridge(alpha=config["alpha"])
                            else:
                                comp_model = RandomForestRegressor(
                                    n_estimators=50,
                                    max_depth=2,
                                    min_samples_leaf=2,
                                    random_state=42
                                )
                        
                            # Try cross-validation for composition model too
                            try:
                                if len(comp_data) >= 10:  # Only if enough data
                                    comp_cv = KFold(n_splits=min(config["cv_folds"], len(comp_data) // 2), 
                                                  shuffle=True, random_state=42)
                                
                                    comp_cv_results = cross_val_score(
                                        comp_model, 
                                        comp_X, 
                                        comp_y,
                                        cv=comp_cv,
                                        scoring='r2'
                                    )
                                
                                    comp_r2 = comp_cv_results.mean()
                                    self.root.after(0, lambda m=media, r=comp_r2: status_label.config(
                                        text=f"{m}: Composition model R² = {r:.3f}"
                                    ))
                                else:
                                    comp_cv_results = None
                            except Exception:
                                comp_cv_results = None
                        
                            # Fit the model
                            comp_model.fit(comp_X, comp_y)
                    
                            # Store composition model
                            composition_models[f"{media}_composition"] = {
                                'model': comp_model,
                                'features': terpene_composition_columns,
                                'metadata': {
                                    'model_type': 'composition',
                                    'sample_count': len(comp_data),
                                    'regularization': {
                                        'type': config["model_type"],
                                        'alpha': config["alpha"] if config["model_type"] == "Ridge" else None
                                    },
                                    'cv_results': comp_cv_results.tolist() if comp_cv_results is not None else None
                                }
                            }
                    
                    # --------- CREATE TERPENE PROFILE DATABASE ---------
            
                    if has_composition_data:
                        # Build profile database for this media type
                        media_profiles = {}
                
                        # Group by terpene name
                        for terpene in media_data['terpene'].unique():
                            terpene_data = media_data[(media_data['terpene'] == terpene) & 
                                                     media_data['has_composition']].copy()
                    
                            if len(terpene_data) > 0:
                                # Calculate average profile
                                profile = terpene_data[terpene_composition_columns].mean().to_dict()
                                media_profiles[terpene] = profile
                        
                        # Store in global database
                        terpene_profiles[media] = media_profiles
                
                        profile_count = len(media_profiles)
                        self.root.after(0, lambda m=media, n=profile_count: status_label.config(
                            text=f"Created {n} terpene profiles for {m}"
                        ))

                # Save all models
                os.makedirs('models', exist_ok=True)
        
                if base_models:
                    self.root.after(0, lambda: status_label.config(
                        text=f"Saving {len(base_models)} base models..."
                    ))
            
                    # Save base models
                    with open('models/viscosity_base_models.pkl', 'wb') as f:
                        pickle.dump(base_models, f)
                
                    # Store in class attribute for immediate use
                    self.base_models = base_models
            
                    # Save composition models if we have any
                    if composition_models:
                        with open('models/viscosity_composition_models.pkl', 'wb') as f:
                            pickle.dump(composition_models, f)
                
                        self.composition_models = composition_models
            
                    # Save terpene profiles
                    if terpene_profiles:
                        with open('models/terpene_profiles.pkl', 'wb') as f:
                            pickle.dump(terpene_profiles, f)
                
                        self.terpene_profiles = terpene_profiles
                
                    # Save CV scores separately for easier analysis
                    with open('models/viscosity_cv_scores.pkl', 'wb') as f:
                        pickle.dump(cv_scores, f)
            
                    # Show success message with cross-validation results
                    message = f"Training complete with L2 regularization (α = {config['alpha']})!\n\n"
                    message += f"Created {len(base_models)} base models\n"
            
                    if composition_models:
                        message += f"Created {len(composition_models)} composition enhancement models\n"
            
                    if terpene_profiles:
                        profile_count = sum(len(profiles) for profiles in terpene_profiles.values())
                        message += f"Created {profile_count} terpene composition profiles\n"
            
                    message += "\nCross-validation results (R²):\n"
                    for media, scores in cv_scores.items():
                        message += f"- {media}: {scores['mean_r2']:.3f} ± {scores['std_r2']:.3f}\n"
                
                    # Analyze feature importance for Ridge models
                    if config["model_type"] == "Ridge":
                        message += "\nFeature importance analysis:\n"
                        feature_importances = {}
                    
                        for model_key, model_data in base_models.items():
                            media = model_key.split('_')[0]
                            residual_model = model_data['residual_model']
                            residual_features = model_data['residual_features']
                        
                            if hasattr(residual_model, 'coef_'):
                                # Get absolute coefficient values for importance
                                coeffs = np.abs(residual_model.coef_)
                                # Normalize to sum to 100%
                                if coeffs.sum() > 0:
                                    importances = 100 * coeffs / coeffs.sum()
                                
                                    feature_importances[media] = {
                                        feature: importance for feature, importance in 
                                        zip(residual_features, importances)
                                    }
                    
                        # Report average importance across models for key features
                        if feature_importances:
                            key_features = ['total_potency', 'terpene_pct', 'is_raw', 
                                          'terpene_headroom', 'potency_terpene_ratio']
                        
                            avg_importance = {feature: [] for feature in key_features}
                        
                            for media, importances in feature_importances.items():
                                for feature in key_features:
                                    if feature in importances:
                                        avg_importance[feature].append(importances[feature])
                        
                            for feature, values in avg_importance.items():
                                if values:
                                    message += f"- {feature}: {np.mean(values):.1f}%\n"
            
                    self.root.after(0, lambda: messagebox.showinfo("Success", message))
                else:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Warning", "No models were created. Check data quality and availability."
                    ))
    
                # Close window
                window.after(0, window.destroy)
    
            except Exception as e:
                import traceback
                error_msg = f"Training error: {str(e)}\n\n{traceback.format_exc()}"
                print(error_msg)
                self.root.after(0, lambda: messagebox.showerror("Error", f"Training failed: {str(e)}"))
                window.after(0, window.destroy)

        # Function to start the training thread
        def start_training():
            # Disable the run button while training is running
            import threading
            train_button.config(state="disabled")
            status_label.config(text="Starting model training with L2 regularization...")
    
            # Collect configuration
            training_config = {
                "model_type": model_type_var.get(),
                "alpha": alpha_var.get(),
                "features": features_var.get(),
                "cv_folds": cv_folds_var.get()
            }
    
            # Start training in a background thread
            training_thread = threading.Thread(
                target=lambda: train_models_thread(
                    training_config, 
                    status_label,
                    config_window
                )
            )
            training_thread.daemon = True
            training_thread.start()

        # Create button frame
        button_frame = Frame(config_window, bg=APP_BACKGROUND_COLOR)
        button_frame.pack(pady=10)

        # Add buttons
        train_button = ttk.Button(button_frame, text="Train Models", command=start_training)
        train_button.pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=config_window.destroy).pack(side="left", padx=10)

    def analyze_models(self, show_dialog=True):
        """
        Analyze consolidated viscosity models with a focus on residual performance.
        Provides insights into temperature sensitivity, feature importance, and predictive accuracy.
        """
        import pandas as pd
        import numpy as np
        import matplotlib.pyplot as plt
        from sklearn.metrics import mean_squared_error, r2_score
        from sklearn.impute import SimpleImputer
        
        debug_print("Loaded heavy modules for analyze_models")
        debug_print(f"matplotlib backend: {plt.get_backend()}")
        debug_print("Starting model analysis...")

        # Check if consolidated models exist
        consolidated_models_exist = hasattr(self, 'consolidated_models') and self.consolidated_models

        if not consolidated_models_exist:
            messagebox.showinfo("No Models", "No consolidated models found. Please train models first.")
            return

        # Create analysis report
        report = []
        report.append("Consolidated Model Analysis Report")
        report.append("===================================")

        # Load validation data
        validation_data = None
        try:
            data_dir = 'data'
            if os.path.exists(data_dir):
                validation_files = [f for f in os.listdir(data_dir) 
                                  if (f.startswith('viscosity_data_') or f.startswith('Master_Viscosity_Data_')) 
                                  and f.endswith('.csv')]

                if validation_files:
                    # Use most recent file
                    validation_files.sort(reverse=True)
                    latest_file = os.path.join('data', validation_files[0])
                    validation_data = pd.read_csv(latest_file)
                    report.append(f"Using {os.path.basename(latest_file)} for validation")

                    # Prepare data for Arrhenius models
                    validation_data['temperature_kelvin'] = validation_data['temperature'] + 273.15
                    validation_data['inverse_temp'] = 1 / validation_data['temperature_kelvin']
                    validation_data['log_viscosity'] = np.log(validation_data['viscosity'])

                    # Add is_raw flag if missing
                    if 'is_raw' not in validation_data.columns:
                        validation_data['is_raw'] = validation_data['terpene'].isna() | (validation_data['terpene'] == '') | (validation_data['terpene'] == 'Raw')
    
                    # Clean up terpene values
                    validation_data.loc[validation_data['terpene'].isna(), 'terpene'] = 'Raw'
                    validation_data.loc[validation_data['terpene'] == '', 'terpene'] = 'Raw'
        
                    # Convert terpene_pct to decimal if over 1
                    if 'terpene_pct' in validation_data.columns and (validation_data['terpene_pct'] > 1).any():
                        validation_data.loc[validation_data['terpene_pct'] > 1, 'terpene_pct'] = validation_data.loc[validation_data['terpene_pct'] > 1, 'terpene_pct'] / 100
                
                    # Convert total_potency to decimal if over 1
                    if 'total_potency' in validation_data.columns and (validation_data['total_potency'] > 1).any():
                        validation_data.loc[validation_data['total_potency'] > 1, 'total_potency'] = validation_data.loc[validation_data['total_potency'] > 1, 'total_potency'] / 100
                
                    # Add physical constraint features
                    if 'total_potency' in validation_data.columns and 'terpene_pct' in validation_data.columns:
                        validation_data['theoretical_max_terpene'] = 1.0 - validation_data['total_potency']
                        validation_data['terpene_headroom'] = validation_data['theoretical_max_terpene'] - validation_data['terpene_pct']
                        validation_data['terpene_max_ratio'] = validation_data['terpene_pct'] / validation_data['theoretical_max_terpene'].clip(lower=0.01)
                        validation_data['potency_terpene_ratio'] = validation_data['total_potency'] / validation_data['terpene_pct'].clip(lower=0.01)
        except Exception as e:
            report.append(f"Error loading validation data: {str(e)}")

        # Analyze consolidated models
        report.append(f"\nConsolidated Media Models: {len(self.consolidated_models)}")
        report.append("-" * 50)
    
        # Process each consolidated model
        for model_key, model in self.consolidated_models.items():
            report.append(f"\nModel: {model_key}")
        
            try:
                # Extract media type from model key
                media = model_key.split('_')[0]
            
                # Extract model components
                baseline_model = model['baseline_model']
                residual_model = model['residual_model']
                baseline_features = model['baseline_features']
                residual_features = model['residual_features']
                residual_features = [f for f in residual_features if f != 'terpene_brand']
            
                # Get metadata
                metadata = model.get('metadata', {})
            
                # Analyze baseline model (Arrhenius temperature relationship)
                report.append("1. Temperature baseline model (Arrhenius)")
            
                if hasattr(baseline_model, 'coef_'):
                    # Extract Arrhenius parameters
                    coef = baseline_model.coef_[0]
                    intercept = baseline_model.intercept_
                
                    # Calculate activation energy
                    R = 8.314  # Gas constant (J/mol·K)
                    Ea = coef * R
                    Ea_kJ = Ea / 1000  # Convert to kJ/mol
                
                    report.append(f"  - Equation: log(viscosity) = {intercept:.4f} + {coef:.4f} * (1/T)")
                    report.append(f"  - Activation energy: {Ea_kJ:.2f} kJ/mol")
                
                    # Categorize temperature sensitivity
                    if Ea_kJ < 20:
                        report.append("  - Low temperature sensitivity")
                    elif Ea_kJ < 40:
                        report.append("  - Medium temperature sensitivity")
                    else:
                        report.append("  - High temperature sensitivity")
            
                # Analyze residual model (terpene and potency effects)
                report.append("\n2. Residual model analysis")
                report.append(f"  - Model type: {type(residual_model).__name__}")
            
                # Extract one-hot encoded terpene features
                terpene_features = [f for f in residual_features if f.startswith('terpene_')]
                report.append(f"  - Model handles {len(terpene_features)} distinct terpenes")
            
                # Extract other feature categories
                physical_features = [f for f in residual_features if f in ['theoretical_max_terpene', 'terpene_headroom', 'terpene_max_ratio']]
                primary_features = ['terpene_pct', 'total_potency', 'is_raw']
                interaction_features = [f for f in residual_features if f == 'potency_terpene_ratio']
            
                # Analyze feature importance if available
                if hasattr(residual_model, 'feature_importances_'):
                    importances = residual_model.feature_importances_
                    features_with_importance = list(zip(residual_features, importances))
                    sorted_features = sorted(features_with_importance, key=lambda x: x[1], reverse=True)
                
                    # Report top features
                    report.append("  - Top 5 most important features:")
                    for feature, importance in sorted_features[:5]:
                        report.append(f"    * {feature}: {importance:.4f}")
                
                    # Calculate importance by feature type
                    total_importance = sum(importances)
                    terpene_total_importance = sum(importance for feature, importance in features_with_importance 
                                              if feature in terpene_features)
                    physical_total_importance = sum(importance for feature, importance in features_with_importance 
                                               if feature in physical_features)
                    primary_total_importance = sum(importance for feature, importance in features_with_importance 
                                              if feature in primary_features)
                
                    report.append("\n  - Feature importance by category:")
                    report.append(f"    * Primary features: {primary_total_importance:.4f} ({primary_total_importance/total_importance*100:.1f}%)")
                    report.append(f"    * Terpene-specific features: {terpene_total_importance:.4f} ({terpene_total_importance/total_importance*100:.1f}%)")
                    report.append(f"    * Physical constraint features: {physical_total_importance:.4f} ({physical_total_importance/total_importance*100:.1f}%)")
                
                    # Analyze key feature importance
                    if 'terpene_pct' in residual_features and 'total_potency' in residual_features:
                        terpene_idx = residual_features.index('terpene_pct')
                        potency_idx = residual_features.index('total_potency')
                    
                        terpene_importance = importances[terpene_idx]
                        potency_importance = importances[potency_idx]
                    
                        report.append("\n  - Key feature comparison:")
                        report.append(f"    * Terpene %: {terpene_importance:.4f}")
                        report.append(f"    * Potency: {potency_importance:.4f}")
                    
                        if potency_importance > 1.5 * terpene_importance:
                            report.append("    * Potency has significantly more impact than terpene %")
                        elif terpene_importance > 1.5 * potency_importance:
                            report.append("    * Terpene % has significantly more impact than potency")
                        else:
                            report.append("    * Terpene % and potency have similar importance")
            
                # Validation with available data
                if validation_data is not None:
                    try:
                        # Filter validation data for this media type
                        media_val_data = validation_data[validation_data['media'] == media].copy()
                    
                        if len(media_val_data) >= 10:
                            report.append(f"\nValidation on {len(media_val_data)} samples for {media}:")
                        
                            # Drop NaN values in key features
                            media_val_data = media_val_data.dropna(subset=['inverse_temp', 'log_viscosity'])
                        
                            # Ensure primary features exist
                            required_features = ['terpene_pct', 'total_potency']
                            missing_features = [f for f in required_features if f not in media_val_data.columns]
                        
                            if missing_features:
                                report.append(f"  - Missing required features: {', '.join(missing_features)}")
                                report.append("  - Skipping validation due to missing features")
                                continue
                        
                            # Create combined terpene field before one-hot encoding
                            if 'terpene' in media_val_data.columns and 'terpene_brand' in media_val_data.columns:
                                media_val_data['combined_terpene'] = media_val_data.apply(
                                    lambda row: f"{row['terpene']}_{row['terpene_brand']}" if pd.notna(row['terpene_brand']) and row['terpene_brand'] != '' 
                                    else row['terpene'], 
                                    axis=1
                                )
                                # Use combined_terpene for one-hot encoding
                                encoded_val_data = pd.get_dummies(
                                    media_val_data,
                                    columns=['combined_terpene'],
                                    prefix=['terpene']
                                )
                            else:
                                # Fall back to just terpene if terpene_brand isn't available
                                encoded_val_data = pd.get_dummies(
                                    media_val_data,
                                    columns=['terpene'],
                                    prefix=['terpene']
                                )
                            
                            # Step 1: Evaluate baseline model
                            X_baseline = encoded_val_data[baseline_features]
                            y_true = encoded_val_data['log_viscosity']
                        
                            # Get baseline predictions
                            baseline_preds = baseline_model.predict(X_baseline)
                        
                            # Calculate residuals
                            encoded_val_data['baseline_prediction'] = baseline_preds
                            encoded_val_data['residual'] = y_true - baseline_preds
                        
                            # Step 2: Evaluate residual model
                        
                            # First, print more diagnostic information
                            debug_print(f"DEBUG - Validation DataFrame columns: {encoded_val_data.columns.tolist()}")
                            debug_print(f"DEBUG - Required features: {residual_features}")
                            debug_print(f"DEBUG - Missing features: {[f for f in residual_features if f not in encoded_val_data.columns]}")
    
                            # Check if 'terpene_brand' is in the required features and remove it
                            clean_residual_features = [f for f in residual_features if f != 'terpene_brand']
    
                            debug_print(f"DEBUG - Removed 'terpene_brand' from features. Original count: {len(residual_features)}, New count: {len(clean_residual_features)}")
    
                            # Create a properly aligned DataFrame with correct features
                            aligned_data = pd.DataFrame(0, index=encoded_val_data.index, columns=clean_residual_features)
    
                            # Fill in values from encoded_val_data where available
                            for col in clean_residual_features:
                                if col in encoded_val_data.columns:
                                    aligned_data[col] = encoded_val_data[col]
                                elif col == 'potency_terpene_ratio' and 'total_potency' in encoded_val_data.columns and 'terpene_pct' in encoded_val_data.columns:
                                    aligned_data[col] = encoded_val_data['total_potency'] / encoded_val_data['terpene_pct'].clip(lower=0.01)
    
                            # Use the aligned data for residual model validation
                            X_residual = aligned_data
    
                            # Debug output
                            debug_print(f"DEBUG - X_residual final shape: {X_residual.shape}, expected shape: ({len(encoded_val_data)}, {len(clean_residual_features)})")
   

                            # Debug feature alignment and duplicates
                            debug_print(f"\nDEBUG - Model: {media} validation")
                            debug_print(f"residual_features length: {len(residual_features)}")
                            debug_print(f"residual_features unique length: {len(set(residual_features))}")
                            debug_print(f"X_residual columns length: {len(X_residual.columns)}")

                            # Check for duplicates in residual_features and handle them
                            duplicates = {}
                            seen = set()
                            for feature in residual_features:
                                if feature in seen:
                                    duplicates[feature] = duplicates.get(feature, 1) + 1
                                else:
                                    seen.add(feature)

                            duplicate_features = list(duplicates.keys())
                            if duplicate_features:
                                debug_print(f"FOUND DUPLICATE FEATURES: {duplicate_features}")
    
                                # Create DataFrame with correct shape (matching residual_features exactly)
                                X_aligned = pd.DataFrame(index=X_residual.index, columns=range(len(residual_features)))
    
                                # Copy values from X_residual, repeating values for duplicated features
                                for i, feature in enumerate(residual_features):
                                    if feature in X_residual.columns:
                                        X_aligned.iloc[:, i] = X_residual[feature].values
    
                                # Now we have the exact shape required by the model, but with numeric column names
                                # Rename columns to match expected feature names for better debugging
                                X_aligned.columns = [f"{f}_{i}" if f in duplicate_features and i > 0 
                                                    else f for i, f in enumerate([
                                                        next(feat for feat in residual_features if feat == f 
                                                            or feat not in residual_features[:i]) 
                                                        for f in residual_features
                                                    ])]
    
                                report.append(f"  - Handled duplicate features: {duplicate_features}")
                                X_residual = X_aligned
                            else:
                                # Original alignment code if no duplicates
                                missing_cols = set(residual_features) - set(X_residual.columns)
                                extra_cols = set(X_residual.columns) - set(residual_features)
    
                                if missing_cols or extra_cols:
                                    aligned_data = pd.DataFrame(0, index=X_residual.index, columns=residual_features)
                                    for col in X_residual.columns:
                                        if col in residual_features:
                                            aligned_data[col] = X_residual[col]
                                    X_residual = aligned_data
                            # Final verification - ensure shape matches exactly what the model expects
                            if hasattr(residual_model, 'n_features_in_'):
                                assert X_residual.shape[1] == residual_model.n_features_in_, \
                                    f"Shape mismatch: {X_residual.shape[1]} vs expected {residual_model.n_features_in_}"

                            # Handle NaN values
                            if X_residual.isna().any().any():
                                imputer = SimpleImputer(strategy='mean')
                                X_residual_values = imputer.fit_transform(X_residual)
                                X_residual = pd.DataFrame(X_residual_values, 
                                                      index=X_residual.index, 
                                                      columns=X_residual.columns)
                        
                            # Get residual predictions
                            y_residual = encoded_val_data['residual']
                            residual_preds = residual_model.predict(X_residual)
                        
                            # Calculate metrics for residual model
                            r2_residual = r2_score(y_residual, residual_preds)
                            mse_residual = mean_squared_error(y_residual, residual_preds)
                        
                            report.append(f"  - Residual model - MSE: {mse_residual:.2f}, R²: {r2_residual:.4f}")
                        
                            # Combined prediction metrics
                            combined_preds = baseline_preds + residual_preds
                        
                            # Log scale metrics
                            r2_log = r2_score(y_true, combined_preds)
                            mse_log = mean_squared_error(y_true, combined_preds)
                        
                            # Original scale metrics
                            y_orig = np.exp(y_true)
                            preds_orig = np.exp(combined_preds)
                        
                            r2_orig = r2_score(y_orig, preds_orig)
                            mse_orig = mean_squared_error(y_orig, preds_orig)
                        
                            report.append(f"  - Log scale - MSE: {mse_log:.2f}, R²: {r2_log:.4f}")
                            report.append(f"  - Original scale - MSE: {mse_orig:.2f}, R²: {r2_orig:.4f}")
                        
                            # Quality assessment
                            if r2_orig >= 0.9:
                                report.append("  - EXCELLENT model performance (R² ≥ 0.9)")
                            elif r2_orig >= 0.8:
                                report.append("  - GOOD model performance (R² ≥ 0.8)")
                            elif r2_orig >= 0.7:
                                report.append("  - ACCEPTABLE model performance (R² ≥ 0.7)")
                            elif r2_orig >= 0.5:
                                report.append("  - FAIR model performance (R² ≥ 0.5)")
                            else:
                                report.append("  - POOR model performance (R² < 0.5)")
                        else:
                            report.append(f"  - Insufficient validation data: only {len(media_val_data)} samples for {media}")
                    except Exception as e:
                        report.append(f"  - Error during validation: {str(e)}")
            except Exception as e:
                report.append(f"Error analyzing model: {str(e)}")
                report.append(traceback.format_exc())

        # Add recommendations
        report.append("\nRecommendations:")
        report.append("---------------")
        report.append("1. Focus on high R² and low MSE values to identify quality models")
        report.append("2. Compare activation energies across media types for temperature sensitivity")
        report.append("3. Review feature importance to understand what drives viscosity")
        report.append("4. For media with poor models, consider collecting more diverse data")
        report.append("5. Physical constraint features should have significant importance")

        # Print report to console
        debug_print("\n".join(report))

        # Show dialog if requested
        if show_dialog:
            report_window = Toplevel(self.root)
            report_window.title("Consolidated Model Analysis")
            report_window.geometry("800x600")

            Label(report_window, text="Consolidated Model Analysis", 
                  font=("Arial", 14, "bold")).pack(pady=10)

            text_frame = Frame(report_window)
            text_frame.pack(fill="both", expand=True, padx=10, pady=10)

            scrollbar = Scrollbar(text_frame)
            scrollbar.pack(side="right", fill="y")

            text_widget = Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
            text_widget.pack(side="left", fill="both", expand=True)

            scrollbar.config(command=text_widget.yview)

            text_widget.insert("1.0", "\n".join(report))
            text_widget.config(state="disabled")

            Button(report_window, text="Close", 
                   command=report_window.destroy).pack(pady=10)

        return report

    def analyze_chemical_importance(self):
        """
        Analyze and visualize the importance of chemical properties in consolidated viscosity models.
        Creates bar charts and heatmaps showing the relative importance of features across media types.
        """
        # Import required libraries
        import matplotlib.pyplot as plt
        import numpy as np
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.figure import Figure
        
        debug_print("Loaded heavy modules for analyze_chemical_importance")
        debug_print(f"matplotlib Figure class: {Figure}")
        debug_print("Starting chemical importance analysis...")
    
        # Check if consolidated models exist
        if not hasattr(self, 'consolidated_models') or not self.consolidated_models:
            messagebox.showinfo(
                "No Consolidated Models",
                "No consolidated models found.\n\n"
                "Please train models first."
            )
            return
    
        # Create window for the analysis
        analysis_window = Toplevel(self.root)
        analysis_window.title("Chemical Properties Importance Analysis")
        analysis_window.geometry("800x800")  # Increased height
        analysis_window.transient(self.root)
    
        # Add title
        Label(
            analysis_window, 
            text="Impact of Chemical Properties on Viscosity",
            font=("Arial", 16, "bold")
        ).pack(pady=10)
    
        # Create a frame for the plots
        frame = Frame(analysis_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
    
        # Create figure with subplots
        fig = Figure(figsize=(10, 12))  # Taller figure
    
        # Get all unique features and media types
        all_features = set()
        media_types = set()
        terpene_features = set()
    
        # Track all feature types
        primary_features = ['terpene_pct', 'total_potency', 'is_raw']
        physical_features = ['theoretical_max_terpene', 'terpene_headroom', 'terpene_max_ratio', 'potency_terpene_ratio']
    
        for model_key, model_data in self.consolidated_models.items():
            # Extract media type from the model key
            media = model_key.split('_')[0]  
            media_types.add(media)
        
            # Add features
            if isinstance(model_data, dict) and 'residual_features' in model_data:
                residual_features = model_data['residual_features']
                all_features.update(residual_features)
            
                # Identify terpene one-hot features
                for feature in residual_features:
                    if feature.startswith('terpene_') and feature != 'terpene_pct':
                        terpene_features.add(feature)
    
        # Remove specific features from general analysis as they're handled separately
        base_features = all_features - terpene_features
    
        # Create bar plot for feature importance
        ax1 = fig.add_subplot(211)
    
        # Calculate average importance for each feature by media type
        media_list = sorted(list(media_types))
        feature_list = sorted(list(base_features))
    
        # Create arrays to store importance values
        importance_data = {media: {feature: [] for feature in feature_list} for media in media_list}
    
        # Collect importance values
        for model_key, model_data in self.consolidated_models.items():
            # Extract media type
            media = model_key.split('_')[0]
        
            # Get the model and its features with proper type checking
            if isinstance(model_data, dict) and 'residual_model' in model_data:
                residual_model = model_data['residual_model']
                residual_features = model_data['residual_features']
            else:
                continue
        
            # Skip if model doesn't have feature_importances_
            if not hasattr(residual_model, 'feature_importances_'):
                continue
        
            importances = residual_model.feature_importances_
        
            # Map importances to features
            for i, feature in enumerate(residual_features):
                if feature in feature_list:  # Only include base features
                    importance_data[media][feature].append(importances[i])
    
        # Calculate averages
        avg_importances = {media: {feature: np.mean(values) if values else 0 
                                  for feature, values in media_data.items()}
                          for media, media_data in importance_data.items()}
    
        # Plot bar chart
        bar_width = 0.8 / len(media_list)
        x = np.arange(len(feature_list))
    
        for i, media in enumerate(media_list):
            values = [avg_importances[media][feature] for feature in feature_list]
            ax1.bar(x + i * bar_width, values, bar_width, label=media)
    
        ax1.set_xlabel('Chemical Property')
        ax1.set_ylabel('Average Importance')
        ax1.set_title('Importance of Chemical Properties by Media Type')
        ax1.set_xticks(x + bar_width * (len(media_list) - 1) / 2)
        ax1.set_xticklabels(feature_list, rotation=45, ha='right')
    
        ax1.legend()
    
        # Create heatmap showing feature importance across models
        ax2 = fig.add_subplot(212)
    
        # Special handling for terpene features - combine them into a single "terpene type" feature
        # Get all model keys and organize by media type
        model_keys_by_media = {media: [] for media in media_list}
        for model_key in self.consolidated_models.keys():
            media = model_key.split('_')[0]
            model_keys_by_media[media].append(model_key)
    
        # Prepare data for heatmap - group features by category
        feature_categories = {
            'Primary': primary_features,
            'Physical': physical_features,
            'Terpene Type': list(terpene_features)
        }
    
        # Create a simpler matrix for the heatmap - media types vs feature categories
        category_importances = {media: {category: 0.0 for category in feature_categories} 
                               for media in media_list}
    
        # Combine terpene feature importances for each media
        for media in media_list:
            # For each media, calculate the average importance of each feature category
            for model_key in model_keys_by_media[media]:
                model_data = self.consolidated_models[model_key]
            
                # Skip if not a proper model or no feature importances
                if not isinstance(model_data, dict) or 'residual_model' not in model_data:
                    continue
                
                residual_model = model_data['residual_model']
                if not hasattr(residual_model, 'feature_importances_'):
                    continue
                
                residual_features = model_data['residual_features']
                importances = residual_model.feature_importances_
            
                # Calculate importance for each category
                for category, cat_features in feature_categories.items():
                    # Sum importance of features in this category
                    category_importance = 0.0
                    count = 0
                
                    for feature in cat_features:
                        if feature in residual_features:
                            idx = residual_features.index(feature)
                            if idx < len(importances):
                                category_importance += importances[idx]
                                count += 1
                
                    # Average the importance if we found any features
                    if count > 0:
                        category_importances[media][category] += category_importance / count
        
            # Average across all models for this media
            model_count = len(model_keys_by_media[media])
            if model_count > 0:
                for category in feature_categories:
                    category_importances[media][category] /= model_count
    
        # Create heatmap data matrix
        heatmap_data = []
        for media in media_list:
            row = [category_importances[media][category] for category in feature_categories]
            heatmap_data.append(row)
    
        # Create heatmap
        im = ax2.imshow(heatmap_data, cmap='viridis', aspect='auto')
    
        # Add colorbar
        fig.colorbar(im, ax=ax2)
    
        # Set labels
        ax2.set_xticks(np.arange(len(feature_categories)))
        ax2.set_yticks(np.arange(len(media_list)))
        ax2.set_xticklabels(feature_categories.keys())
        ax2.set_yticklabels(media_list)
    
        # Rotate x labels for better readability
        plt.setp(ax2.get_xticklabels(), rotation=45, ha="right", rotation_mode="anchor")
    
        ax2.set_title("Feature Category Importance by Media Type")
    
        # Adjust layout
        fig.tight_layout()
    
        # Create a canvas to display the figure
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
        # Add a button to close the window
        ttk.Button(
            analysis_window,
            text="Close",
            command=analysis_window.destroy
        ).pack(pady=10)

    def filter_and_analyze_specific_combinations(self):
        """
        Analyze and visualize Arrhenius relationships for two-level model system.
        Shows temperature sensitivity and potency effects on viscosity.
        """
        import pandas as pd
        import numpy as np
        import matplotlib
        import matplotlib.pyplot as plt
        from scipy.optimize import curve_fit
        from sklearn.metrics import r2_score
    
        # Create the main window
        progress_window = Toplevel(self.root)
        progress_window.title("Potency-Temperature Analysis")
        progress_window.geometry("800x600")
        progress_window.transient(self.root)
        progress_window.grab_set()
    
        # Main layout frames
        top_frame = Frame(progress_window, bg=APP_BACKGROUND_COLOR)
        top_frame.pack(fill="x", padx=10, pady=5)
    
        Label(top_frame, text="Potency Effects on Viscosity and Temperature Sensitivity", 
              font=("Arial", 14, "bold"), fg="white", bg=APP_BACKGROUND_COLOR).pack(pady=10)
    
        # Configuration frame
        config_frame = Frame(top_frame, bg=APP_BACKGROUND_COLOR)
        config_frame.pack(fill="x", padx=10, pady=5)
    
        # Potency configuration
        Label(config_frame, text="Analysis potency range:", 
              fg="white", bg=APP_BACKGROUND_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    
        # Radio buttons for potency analysis type
        potency_mode_var = StringVar(value="variable")
    
        
        tk.Radiobutton(config_frame, text="Fixed potency", variable=potency_mode_var, 
                      value="fixed", bg=APP_BACKGROUND_COLOR, fg="white", 
                      selectcolor=APP_BACKGROUND_COLOR).grid(row=0, column=1, padx=5, pady=5, sticky="w")
    
        tk.Radiobutton(config_frame, text="Variable potency", variable=potency_mode_var, 
                      value="variable", bg=APP_BACKGROUND_COLOR, fg="white",
                      selectcolor=APP_BACKGROUND_COLOR).grid(row=0, column=2, padx=5, pady=5, sticky="w")
    
        # Potency value slider
        Label(config_frame, text="Center potency value (%):", 
              fg="white", bg=APP_BACKGROUND_COLOR).grid(row=1, column=0, padx=5, pady=5, sticky="w")
    
        potency_var = DoubleVar(value=80.0)  # Default value of 80%
        potency_slider = Scale(config_frame, variable=potency_var, from_=60.0, to=95.0, 
                              orient=HORIZONTAL, length=200, resolution=0.5,
                              bg=APP_BACKGROUND_COLOR, fg="white", troughcolor=BUTTON_COLOR)
        potency_slider.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
    
        # Note about terpene percentage
        terpene_note = Label(config_frame, 
                            text="Note: Terpene percentage will be calculated as (100% - potency)",
                            fg="yellow", bg=APP_BACKGROUND_COLOR, font=("Arial", 9, "italic"))
        terpene_note.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky="w")
    
        # Advanced options toggle
        use_advanced_var = tk.BooleanVar(value=False)
        advanced_check = tk.Checkbutton(config_frame, text="Advanced Options", 
                                       variable=use_advanced_var, 
                                       command=lambda: toggle_advanced_options(),
                                       bg=APP_BACKGROUND_COLOR, fg="white",
                                       selectcolor=APP_BACKGROUND_COLOR)
        advanced_check.grid(row=3, column=0, padx=5, pady=5, sticky="w")
    
        # Advanced options frame (hidden by default)
        advanced_frame = Frame(config_frame, bg=APP_BACKGROUND_COLOR)
    
        # Model type selection
        Label(advanced_frame, text="Model type:", 
              fg="white", bg=APP_BACKGROUND_COLOR).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    
        model_type_var = StringVar(value="base")
        model_types = ["base", "both", "consolidated"]
    
        model_dropdown = ttk.Combobox(
            advanced_frame, 
            textvariable=model_type_var,
            values=model_types,
            state="readonly",
            width=12
        )
        model_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    
        # Specific terpene selection (for advanced usage)
        Label(advanced_frame, text="Specific terpene (optional):", 
              fg="white", bg=APP_BACKGROUND_COLOR).grid(row=1, column=0, padx=5, pady=5, sticky="w")
    
        # Get available terpenes from different sources
        available_terpenes = ["Raw"]  # Default option
    
        # Try to gather terpene names from profiles or models
        if hasattr(self, 'terpene_profiles') and self.terpene_profiles:
            for media_profiles in self.terpene_profiles.values():
                available_terpenes.extend(media_profiles.keys())
        elif hasattr(self, 'consolidated_models') and self.consolidated_models:
            for model_key, model in self.consolidated_models.items():
                if isinstance(model, dict) and 'residual_features' in model:
                    for feature in model['residual_features']:
                        if feature.startswith('terpene_') and feature != 'terpene_pct':
                            terpene_name = feature[8:]  # Remove 'terpene_' prefix
                            available_terpenes.append(terpene_name)
    
        # Remove duplicates and sort
        available_terpenes = sorted(list(set(available_terpenes)))
    
        terpene_var = StringVar(value="Raw")
        terpene_dropdown = ttk.Combobox(
            advanced_frame, 
            textvariable=terpene_var,
            values=available_terpenes,
            state="readonly",
            width=15
        )
        terpene_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    
        # Function to toggle advanced options visibility
        def toggle_advanced_options():
            if use_advanced_var.get():
                advanced_frame.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="w")
            else:
                advanced_frame.grid_forget()
    
        # Text area for showing progress
        text_frame = Frame(progress_window)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
        scrollbar = Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
    
        text_widget = Text(text_frame, wrap="word", yscrollcommand=scrollbar.set, 
                          bg="white", fg="black")
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_widget.yview)
    
        # Function to add text to the widget
        def add_text(message):
            text_widget.insert("end", message + "\n")
            text_widget.see("end")
            progress_window.update_idletasks()  # Force UI update
    
        # Add initial message
        add_text("This analysis examines how potency affects viscosity and temperature sensitivity.")
        add_text("For each media type, it will generate:")
        add_text("1. Viscosity vs temperature curves at different potency levels")
        add_text("2. Arrhenius plots showing temperature sensitivity")
        add_text("3. A comparison of activation energies across media types")
        add_text("\nConfigure the settings above and click 'Run Analysis' to start.")
    
        button_frame = Frame(progress_window)
        button_frame.pack(pady=10)

        # Define function for the analysis
        def run_analysis_thread():
            try:
                # Import required modules
                import glob
                import threading
                import os
                import pickle
                import math
                from scipy import stats
            
                # Set matplotlib to non-interactive backend
                matplotlib.use('Agg')
            
                # Helper function to sanitize filenames
                def sanitize_filename(name):
                    """Replace invalid filename characters with underscores."""
                    invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ']
                    sanitized = name
                    for char in invalid_chars:
                        sanitized = sanitized.replace(char, '_')
                    return sanitized
            
                # Get configuration values
                potency_mode = potency_mode_var.get()
                center_potency = potency_var.get()
                model_type = model_type_var.get() if use_advanced_var.get() else "base"
                terpene_name = terpene_var.get() if use_advanced_var.get() else "Raw"
            
                # Choose which model system to use
                have_base_models = hasattr(self, 'base_models') and self.base_models
                have_consolidated = hasattr(self, 'consolidated_models') and self.consolidated_models
            
                if model_type == "base" and not have_base_models:
                    add_text("No base models found. Falling back to consolidated models.")
                    model_type = "consolidated"
            
                if model_type == "both" and not have_base_models:
                    add_text("No base models found for two-level analysis. Falling back to consolidated models.")
                    model_type = "consolidated"
            
                if model_type == "consolidated" and not have_consolidated:
                    add_text("No consolidated models found. Falling back to base models.")
                    model_type = "base"
            
                # If no models are available at all
                if not have_base_models and not have_consolidated:
                    add_text("No models found. Please train models first.")
                    self.root.after(0, lambda: run_button.config(state="normal"))
                    return
            
                # Create plots directory if needed
                os.makedirs('plots', exist_ok=True)
            
                # Define potency range based on mode
                if potency_mode == "variable":
                    # Calculate a range around the selected potency value
                    offsets = [-10, -5, 0, 5, 10]
                    potency_range = [max(50, center_potency + offset) for offset in offsets]
                    # Ensure no values exceed 95%
                    potency_range = [min(95, p) for p in potency_range]
                    # Remove duplicates and sort
                    potency_range = sorted(list(set(potency_range)))
                    add_text(f"Using variable potency analysis with levels: {', '.join([f'{p:.1f}%' for p in potency_range])}")
                else:
                    # Fixed potency mode
                    potency_range = [center_potency]
                    add_text(f"Using fixed potency value of {center_potency:.1f}%")
            
                # Calculate corresponding terpene percentages for each potency
                terpene_percentages = [100.0 - p for p in potency_range]
                add_text(f"Corresponding terpene percentages: {', '.join([f'{t:.1f}%' for t in terpene_percentages])}")
            
                if terpene_name != "Raw":
                    add_text(f"Using specific terpene profile: {terpene_name}")
            
                # Temperature range for analysis
                temperature_range = np.linspace(20, 70, 11)  # 20C to 70C
            
                # Arrhenius function: ln(viscosity) = ln(A) + (Ea/R)*(1/T)
                def arrhenius_function(x, a, b):
                    """x is 1/T (inverse temperature in Kelvin)"""
                    return a + b * x
            
                # Store results for comparing activation energies
                activation_energies = []
            
                # Process each media type
                if model_type in ["base", "both"]:
                    # Using the two-level model system
                    media_types = set(key.split('_')[0] for key in self.base_models.keys())
                    add_text(f"Analyzing {len(media_types)} media types using the two-level model system.")
                else:
                    # Using consolidated models
                    media_types = set(key.split('_')[0] for key in self.consolidated_models.keys())
                    add_text(f"Analyzing {len(media_types)} media types using consolidated models.")
            
                for media in sorted(media_types):
                    try:
                        add_text(f"\nAnalyzing {media}...")
                    
                        # Create figure
                        plt.figure(figsize=(12, 10))
                    
                        # Create subplots
                        ax1 = plt.subplot(211)  # Viscosity vs Temperature
                        ax2 = plt.subplot(212)  # Arrhenius plot
                    
                        # For storing activation energies by potency
                        media_activation_energies = []
                    
                        # Process each potency level and corresponding terpene percentage
                        for potency, terpene_pct in zip(potency_range, terpene_percentages):
                            # Calculate viscosity at each temperature
                            temperatures_kelvin = temperature_range + 273.15
                            inverse_temp = 1 / temperatures_kelvin
                        
                            # Convert to decimal if needed
                            decimal_potency = potency / 100.0
                            decimal_terpene = terpene_pct / 100.0
                        
                            # Get predictions for each temperature
                            add_text(f"  - Calculating viscosity at {potency:.1f}% potency ({terpene_pct:.1f}% terpenes)...")
                            predicted_visc = []
                        
                            for temp in temperature_range:
                                try:
                                    if model_type in ["base", "both"]:
                                        # Using two-level model system
                                        visc = self.predict_model_viscosity(
                                            media, decimal_terpene, temp, decimal_potency, terpene_name
                                        )
                                    else:
                                        # Using consolidated model
                                        model_key = f"{media}_consolidated"
                                        if model_key in self.consolidated_models:
                                            model = self.consolidated_models[model_key]
                                            visc = self.predict_model_viscosity(
                                                model, terpene_pct, temp, potency, terpene_name
                                            )
                                        else:
                                            raise ValueError(f"No model found for {media}")
                                
                                    predicted_visc.append(visc)
                                except Exception as e:
                                    add_text(f"    Error at {temp}C: {str(e)}")
                                    predicted_visc.append(float('nan'))
                        
                            predicted_visc = np.array(predicted_visc)
                        
                            # Filter invalid values
                            valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                            if not any(valid_indices):
                                add_text(f"    No valid predictions for {potency:.1f}% potency. Skipping.")
                                continue
                        
                            # Get valid data for Arrhenius analysis
                            inv_temp_valid = inverse_temp[valid_indices]
                            predicted_visc_valid = predicted_visc[valid_indices]
                            temperatures_valid = temperature_range[valid_indices]
                        
                            # Calculate ln(viscosity)
                            ln_visc = np.log(predicted_visc_valid)
                        
                            # Fit Arrhenius equation
                            params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                            a, b = params
                        
                            # Calculate activation energy
                            R = 8.314  # Gas constant
                            Ea = b * R
                            Ea_kJ = Ea / 1000  # Convert to kJ/mol
                        
                            # Store for comparison
                            media_activation_energies.append((potency, Ea_kJ))
                        
                            # Calculate predicted values
                            ln_visc_pred = arrhenius_function(inv_temp_valid, a, b)
                        
                            # Calculate R-squared
                            r2 = r2_score(ln_visc, ln_visc_pred)
                        
                            # Plot viscosity vs temperature
                            ax1.semilogy(temperatures_valid, predicted_visc_valid,
                                      'o-', label=f'Potency {potency:.1f}% / Terps {terpene_pct:.1f}%')
                        
                            # Plot Arrhenius relationship
                            ax2.scatter(inv_temp_valid, ln_visc,
                                     marker='o', label=f'Potency {potency:.1f}%')
                            ax2.plot(inv_temp_valid, ln_visc_pred, '--',
                                   label=f'Fit {potency:.1f}% (Ea={Ea_kJ:.1f} kJ/mol)')
                        
                            # Report results
                            viscosity_25C = np.interp(25, temperatures_valid, predicted_visc_valid)
                            add_text(f"    Potency {potency:.1f}% / Terpenes {terpene_pct:.1f}%: Viscosity @ 25C = {viscosity_25C:.0f} cP")
                            add_text(f"    Activation Energy: {Ea_kJ:.1f} kJ/mol (R² = {r2:.4f})")
                    
                        # Configure plots
                        ax1.set_xlabel('Temperature (C)')
                        ax1.set_ylabel('Viscosity (cP)')
                        ax1.set_title(f'Viscosity vs Temperature for {media}\nModel: {model_type}')
                        ax1.grid(True)
                        ax1.legend()
                    
                        ax2.set_xlabel('1/T (K⁻¹)')
                        ax2.set_ylabel('ln(Viscosity)')
                        ax2.set_title('Arrhenius Plots at Different Potency Levels')
                        ax2.grid(True)
                        ax2.legend()
                    
                        plt.tight_layout()
                    
                        # Save plot
                        plot_path = f'plots/Potency_Analysis_{sanitize_filename(media)}_{model_type}.png'
                        plt.savefig(plot_path)
                        plt.close()
                    
                        # Create potency vs activation energy plot
                        if len(media_activation_energies) > 1:
                            plt.figure(figsize=(8, 6))
                            potency_values, ea_values = zip(*media_activation_energies)
                        
                            plt.plot(potency_values, ea_values, 'o-', linewidth=2)
                            plt.xlabel('Total Potency (%)')
                            plt.ylabel('Activation Energy (kJ/mol)')
                            plt.title(f'Effect of Potency on Activation Energy\n{media}, Model: {model_type}')
                            plt.grid(True)
                        
                            # Add trendline
                            if len(potency_values) > 2:
                                z = np.polyfit(potency_values, ea_values, 1)
                                p = np.poly1d(z)
                                plt.plot(potency_values, p(potency_values), "r--",
                                       label=f"Trend: {z[0]:.2f}x + {z[1]:.2f}")
                                plt.legend()
                        
                            # Save plot
                            trend_path = f'plots/Potency_Trend_{sanitize_filename(media)}_{model_type}.png'
                            plt.savefig(trend_path)
                            plt.close()
                        
                            # Store in global list for comparison
                            activation_energies.append({
                                'media': media,
                                'slope': z[0] if len(potency_values) > 2 else 0,
                                'intercept': z[1] if len(potency_values) > 2 else 0,
                                'potency_values': potency_values,
                                'ea_values': ea_values,
                                'avg_ea': sum(ea_values) / len(ea_values)
                            })
                        
                            add_text(f"  Plots saved to:")
                            add_text(f"  - {plot_path}")
                            add_text(f"  - {trend_path}")
                        
                            # Report trend
                            if len(potency_values) > 2:
                                if z[0] > 0:
                                    add_text(f"  Trend: Activation energy increases by {z[0]:.2f} kJ/mol per 1% increase in potency")
                                else:
                                    add_text(f"  Trend: Activation energy decreases by {abs(z[0]):.2f} kJ/mol per 1% increase in potency")
                            
                                # Interpret significance
                                slope_range = abs(z[0]) * (max(potency_values) - min(potency_values))
                                avg_ea = sum(ea_values) / len(ea_values)
                                significance = (slope_range / avg_ea) * 100
                            
                                if significance < 5:
                                    add_text("  - This represents a minimal effect on temperature sensitivity")
                                elif significance < 15:
                                    add_text("  - This represents a moderate effect on temperature sensitivity")
                                else:
                                    add_text("  - This represents a significant effect on temperature sensitivity")
                
                    except Exception as e:
                        add_text(f"Error analyzing {media}: {str(e)}")
                        import traceback
                        traceback_str = traceback.format_exc()
                        debug_print(f"Detailed error: {traceback_str}")
            
                # Create comparison plot across media types
                if len(activation_energies) > 1:
                    try:
                        # Sort by slope magnitude
                        activation_energies.sort(key=lambda x: abs(x['slope']), reverse=True)
                    
                        plt.figure(figsize=(10, 6))
                    
                        # Plot trends for each media
                        for i, result in enumerate(activation_energies):
                            media = result['media']
                            potency_values = result['potency_values']
                            ea_values = result['ea_values']
                        
                            # Use different colors
                            color = plt.cm.tab10(i % 10)
                        
                            # Plot actual values
                            plt.plot(potency_values, ea_values, 'o-', color=color,
                                   label=f"{media}", linewidth=2)
                        
                            # Add trendline if available
                            if len(potency_values) > 2:
                                # Calculate trend over consistent range for visualization
                                p_range = np.linspace(min(potency_values), max(potency_values), 10)
                                trend = result['slope'] * p_range + result['intercept']
                                plt.plot(p_range, trend, '--', color=color, alpha=0.7)
                    
                        plt.xlabel('Potency (%)')
                        plt.ylabel('Activation Energy (kJ/mol)')
                        plt.title('Potency Effect on Activation Energy Across Media Types')
                        plt.grid(True)
                        plt.legend(loc='best')
                    
                        # Save comparison plot
                        comparison_path = f'plots/Potency_Effect_Comparison_{model_type}.png'
                        plt.savefig(comparison_path)
                        plt.close()
                    
                        add_text(f"\nComparison plot saved to: {comparison_path}")
                    
                        # Table of slopes
                        add_text("\nSummary of potency effects on activation energy:")
                        add_text("Media Type | Effect Direction | Magnitude (kJ/mol per 1% potency)")
                        add_text("-" * 60)
                    
                        for result in activation_energies:
                            direction = "Increases" if result['slope'] > 0 else "Decreases"
                            add_text(f"{result['media']:<10} | {direction:<16} | {abs(result['slope']):.3f}")
                    
                        # Determine overall trend
                        avg_slope = sum(result['slope'] for result in activation_energies) / len(activation_energies)
                        if abs(avg_slope) < 0.05:
                            add_text("\nOverall: Potency has minimal effect on temperature sensitivity across media types")
                        elif avg_slope > 0:
                            add_text(f"\nOverall: Higher potency tends to increase temperature sensitivity (avg: {avg_slope:.3f})")
                        else:
                            add_text(f"\nOverall: Higher potency tends to decrease temperature sensitivity (avg: {avg_slope:.3f})")
                
                    except Exception as e:
                        add_text(f"Error creating comparison plot: {str(e)}")
            
                # Re-enable button
                self.root.after(0, lambda: run_button.config(state="normal"))
        
            except Exception as e:
                add_text(f"Error in analysis: {str(e)}")
                import traceback
                traceback_str = traceback.format_exc()
                print(f"Thread error: {traceback_str}")
                # Re-enable button
                self.root.after(0, lambda: run_button.config(state="normal"))
    
        # Add the Run Analysis button
        def start_analysis():
            # Disable the run button while analysis is running
            import threading
            run_button.config(state="disabled")
            add_text("\nStarting analysis...")
    
            # Start the analysis in a background thread
            analysis_thread = threading.Thread(target=run_analysis_thread)
            analysis_thread.daemon = True
            analysis_thread.start()

        run_button = ttk.Button(
            button_frame, 
            text="Run Analysis",
            command=start_analysis
        )
        run_button.pack(padx=10)

    def generate_activation_energy_comparison_twolevel(self, potency_value, terpene_name, model_level, log_func=None):
        """
        Generate a comparison plot of activation energies across different media types,
        compatible with the two-level model system.
    
        Args:
            potency_value: Potency value to use for predictions (as percentage)
            terpene_name: Name of terpene to use for analysis
            model_level: Level of model to use (base, composition, both)
            log_func: Optional function to log messages
        """
        import numpy as np
        import matplotlib
        import matplotlib.pyplot as plt
        from scipy.optimize import curve_fit
        from sklearn.metrics import r2_score
        import pandas as pd
        
        debug_print("Loaded heavy modules for generate_activation_energy_comparison_twolevel")
        debug_print(f"potency_value: {potency_value}, terpene_name: {terpene_name}")
        debug_print(f"model_level: {model_level}")
        
        matplotlib.use('Agg')  # Use non-interactive backend

        def sanitize_filename(name):
            """Replace invalid filename characters with underscores."""
            invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ']
            sanitized = name
            for char in invalid_chars:
                sanitized = sanitized.replace(char, '_')
            return sanitized

        # Store results for comparison
        results = []

        # Arrhenius function
        def arrhenius_function(x, a, b):
            return a + b * x

        # Temperature range
        temperature_range = np.linspace(20, 70, 11)
        temperatures_kelvin = temperature_range + 273.15
        inverse_temp = 1 / temperatures_kelvin

        # R constant
        R = 8.314  # Gas constant

        # Using base_models for the two-level system
        models_dict = None
    
        if hasattr(self, 'base_models') and self.base_models and model_level in ["base", "both"]:
            # Use base models
            models_dict = {k.split('_')[0]: k for k in self.base_models.keys()}
            if log_func:
                log_func(f"Using {len(models_dict)} media types from base models")
        elif hasattr(self, 'consolidated_models') and self.consolidated_models:
            # Fallback to consolidated models
            models_dict = {k.split('_')[0]: k for k in self.consolidated_models.keys()}
            if log_func:
                log_func(f"Using {len(models_dict)} media types from consolidated models")
        else:
            if log_func:
                log_func("No models available for comparison")
            return

        # Calculate terpene percentage based on potency (physical constraint)
        terpene_pct = 100.0 - potency_value
    
        # Convert potency and terpene to decimal if using 2-level models
        decimal_potency = potency_value / 100.0 if model_level in ["base", "both"] else potency_value
        decimal_terpene = terpene_pct / 100.0 if model_level in ["base", "both"] else terpene_pct

        # Process each media type
        for media, model_key in models_dict.items():
            try:
                # Generate predictions
                predicted_visc = []
                for temp in temperature_range:
                    try:
                        if model_level in ["base", "both"] and hasattr(self, 'base_models'):
                            # Use two-level prediction system
                            visc = self.predict_model_viscosity(
                                media, decimal_terpene, temp, decimal_potency, terpene_name
                            )
                        else:
                            # Use consolidated model prediction
                            model = self.consolidated_models[model_key]
                            visc = self.predict_model_viscosity(
                                model, terpene_pct, temp, potency_value, terpene_name
                            )
                        predicted_visc.append(visc)
                    except Exception as e:
                        if log_func:
                            log_func(f"Error predicting {media} at {temp}C: {e}")
                        predicted_visc.append(float('nan'))
        
                predicted_visc = np.array(predicted_visc)
        
                # Filter invalid values
                valid_indices = ~np.isnan(predicted_visc) & (predicted_visc > 0)
                if not any(valid_indices):
                    if log_func:
                        log_func(f"Warning: No valid predictions for {media}")
                    continue
        
                inv_temp_valid = inverse_temp[valid_indices]
                predicted_visc_valid = predicted_visc[valid_indices]
        
                # Calculate ln(viscosity)
                ln_visc = np.log(predicted_visc_valid)
        
                # Fit Arrhenius equation
                params, covariance = curve_fit(arrhenius_function, inv_temp_valid, ln_visc)
                a, b = params
        
                # Calculate activation energy
                Ea = b * R
                Ea_kJ = Ea / 1000
        
                # Calculate R-squared
                ln_visc_pred = arrhenius_function(inv_temp_valid, a, b)
                r2 = r2_score(ln_visc, ln_visc_pred)
        
                # Get viscosity at 25C
                visc_25C = None
                if 25 in temperature_range:
                    idx = list(temperature_range).index(25)
                    if idx < len(predicted_visc) and not np.isnan(predicted_visc[idx]):
                        visc_25C = predicted_visc[idx]
        
                # Store result
                results.append({
                    'media': media,
                    'Ea_kJ': Ea_kJ,
                    'ln_A': a,
                    'r2': r2,
                    'visc_25C': visc_25C,
                    'potency': potency_value,
                    'terpene_pct': terpene_pct,
                    'model_level': model_level
                })
    
            except Exception as e:
                if log_func:
                    log_func(f"Error processing {media}: {e}")

        if not results:
            if log_func:
                log_func("No valid results for comparison plot")
            return

        # Convert to DataFrame
        import pandas as pd
        results_df = pd.DataFrame(results)

        # Sort by activation energy
        results_df = results_df.sort_values('Ea_kJ', ascending=False)

        # Create figure
        plt.figure(figsize=(15, max(8, len(results_df) * 0.25)))

        # Create positions for bars
        positions = np.arange(len(results_df))
        bar_height = 0.6

        # Create colormap based on viscosity at 25C
        if 'visc_25C' in results_df.columns and not results_df['visc_25C'].isna().all():
            visc_values = results_df['visc_25C'].fillna(0)
            visc_min = visc_values.min()
            visc_max = visc_values.max()
        
            # Check if there's enough variation in viscosity to create a meaningful colormap
            if visc_max > visc_min and visc_max > 0:
                normalized_visc = (visc_values - visc_min) / (visc_max - visc_min)
                colors = plt.cm.viridis(normalized_visc)
            else:
                # Fallback to default colors if not enough variation
                colors = plt.cm.tab10(np.linspace(0, 1, len(results_df)))
        else:
            # Fallback to default colors if viscosity not available
            colors = plt.cm.tab10(np.linspace(0, 1, len(results_df)))

        # Create horizontal bars
        bars = plt.barh(
            positions, 
            results_df['Ea_kJ'], 
            height=bar_height,
            color=colors
        )

        # Add labels
        plt.yticks(positions, results_df['media'], fontsize=8)

        plt.xlabel('Activation Energy (kJ/mol)', fontsize=12)
        plt.title(f'Activation Energy Comparison by Media Type\nPotency: {potency_value:.1f}%, Terpenes: {terpene_pct:.1f}%\nModel Level: {model_level}', fontsize=14)

        # Add value labels
        for i, bar in enumerate(bars):
            plt.text(
                bar.get_width() + 0.5, 
                bar.get_y() + bar.get_height()/2, 
                f"{results_df['Ea_kJ'].iloc[i]:.1f}", 
                va='center',
                fontsize=8
            )

        # Add a color bar for viscosity with proper error handling
        try:
            if 'visc_25C' in results_df.columns and not results_df['visc_25C'].isna().all():
                visc_values = results_df['visc_25C'].fillna(0)
                # Only add colorbar if there's sufficient variation
                if visc_values.max() > visc_values.min() and visc_values.max() > 0:
                    # Create a ScalarMappable with the appropriate colormap and normalization
                    sm = plt.cm.ScalarMappable(
                        cmap=plt.cm.viridis,
                        norm=plt.Normalize(vmin=visc_values.min(), vmax=visc_values.max())
                    )
                    # This is important - matplotlib needs this for the colorbar to work
                    sm._A = []  # This line fixes the common colorbar error
                
                    # Add the colorbar
                    cbar = plt.colorbar(sm)
                    cbar.set_label('Viscosity at 25C (cP)')
        except Exception as e:
            # Log error but continue without colorbar
            print(f"Colorbar error: {str(e)}")
            if log_func:
                log_func(f"Warning: Could not create colorbar: {str(e)}")

        plt.tight_layout()

        # Save plot
        model_suffix = f"_{model_level}" if model_level != "both" else ""
        plot_path = f'plots/Activation_Energy_Comparison_{sanitize_filename(terpene_name)}_Potency{int(potency_value)}{model_suffix}.png'
        plt.savefig(plot_path, dpi=300)
        if log_func:
            log_func(f"Comparison plot saved to: {plot_path}")

        plt.close()

    def diagnose_models(self):
        """Diagnose issues with feature importance in models"""
        print("\nModel Feature Importance Analysis")
        print("=================================")
    
        if hasattr(self, 'consolidated_models') and self.consolidated_models:
            models = self.consolidated_models
            print(f"Analyzing {len(models)} consolidated models:")
        
            for model_key, model in models.items():
                print(f"\nModel: {model_key}")
            
                # Extract residual model
                residual_model = model['residual_model']
                residual_features = model['residual_features']
            
                # Check feature importance
                if hasattr(residual_model, 'feature_importances_'):
                    importances = residual_model.feature_importances_
                    for i, feature in enumerate(residual_features):
                        if i < len(importances):  # Ensure index is valid
                            importance = importances[i]
                            print(f"  * {feature}: {importance:.6f}")
                        
                            # Flag problems
                            if feature == 'total_potency' and importance < 0.01:
                                print("    WARNING: Potency has extremely low importance")
                            elif feature == 'total_potency' and importance < 0.1:
                                print("    WARNING: Potency has low importance")
                        
                elif hasattr(residual_model, 'coef_'):
                    # For linear models
                    coefs = residual_model.coef_
                    for i, feature in enumerate(residual_features):
                        if i < len(coefs):  # Ensure index is valid
                            coef = coefs[i] if len(coefs.shape) == 1 else coefs[0, i]
                            print(f"  * {feature} coefficient: {coef:.6f}")
                        
                            # Flag problems
                            if feature == 'total_potency' and abs(coef) < 0.01:
                                print("    WARNING: Potency has extremely low coefficient")
                            elif feature == 'total_potency' and abs(coef) < 0.1:
                                print("    WARNING: Potency has low coefficient")
                else:
                    print("  WARNING: No feature importance information available")
                
                # Test potency variation
                potencies = [70, 75, 80, 85, 90]
                viscosities = []
            
                for pot in potencies:
                    visc = self.predict_model_viscosity(model, 5.0, 25.0, pot)
                    viscosities.append(visc)
                
                # Check for variation
                if len(set([round(v, 2) for v in viscosities])) == 1:
                    print("  CRITICAL ERROR: Model shows no response to potency variation")
                else:
                    min_visc = min(viscosities)
                    max_visc = max(viscosities)
                    variation = (max_visc - min_visc) / min_visc * 100
                    print(f"  * Potency variation effect: {variation:.2f}% change in viscosity")
                
                    # Print values
                    for i, pot in enumerate(potencies):
                        print(f"    - {pot}%: {viscosities[i]:.0f} cP")

    def create_potency_demo_model(self):
        """Create a demonstration model with very clear potency effects"""
        import numpy as np
        import pandas as pd
        from sklearn.linear_model import Ridge
        
        debug_print("Loaded heavy modules for create_potency_demo_model")
        debug_print(f"pandas DataFrame class: {pd.DataFrame}")
        debug_print(f"Ridge model class: {Ridge}")
        debug_print("Creating enhanced potency demo model with strong effects...")
    
        # Define baseline model for temperature effects (Arrhenius)
        baseline_model = Ridge(alpha=1.0)
        baseline_log_visc = np.log(10000)  # Reference viscosity at 25C
        activation_energy = 12000  # Controls temperature sensitivity
        gas_constant = 8.314  # Physical constant
    
        # Train baseline model on temperature relationship
        temps = np.array([20, 25, 30, 40, 50, 60])
        temps_kelvin = temps + 273.15
        inverse_temps = 1 / temps_kelvin
        baseline_log_viscs = baseline_log_visc + (activation_energy / gas_constant) * (1/298.15 - inverse_temps)
        X_baseline = pd.DataFrame({'inverse_temp': inverse_temps})
        y_baseline = baseline_log_viscs
        baseline_model.fit(X_baseline, y_baseline)
    
        # Create residual model with very strong potency effect
        residual_model = Ridge(alpha=0.01)  # Lower alpha for stronger fitting
    
        # Generate synthetic training data with extreme potency effect
        # Create more potency values for smoother curve
        potencies = np.linspace(0.7, 0.95, 10)  # More granular potency range
        terpene_pcts = np.linspace(0.03, 0.20, 7)  # Wider terpene range
    
        potency_vals = []
        terpene_vals = []
        pt_ratio_vals = []  # New feature: potency/terpene ratio
        residual_vals = []
    
        for pot in potencies:
            for terp in terpene_pcts:
                # STRONG potency effect - exponential relationship
                # Higher potency = exponentially higher viscosity
                potency_effect = 3.0 * np.exp(pot * 2.0) - 15.0
            
                # Standard terpene effect - inverse relationship
                # Higher terpene = lower viscosity
                terpene_effect = -5.0 * terp
            
                # Interaction effect (potency matters more at lower terpene %)
                interaction = -2.0 * pot * terp
            
                # Combined effect
                combined_effect = potency_effect + terpene_effect + interaction
            
                # Store values
                potency_vals.append(pot)
                terpene_vals.append(terp)
                pt_ratio_vals.append(pot / max(0.01, terp))  # Avoid division by zero
                residual_vals.append(combined_effect)
    
        # Create training dataframe
        X_residual = pd.DataFrame({
            'total_potency': potency_vals,
            'terpene_pct': terpene_vals,
            'potency_terpene_ratio': pt_ratio_vals,
            'is_raw': [0] * len(potency_vals)
        })
        y_residual = np.array(residual_vals)
    
        # Train residual model
        residual_model.fit(X_residual, y_residual)
    
        # Create model dictionary
        demo_model = {
            'baseline_model': baseline_model,
            'residual_model': residual_model,
            'baseline_features': ['inverse_temp'],
            'residual_features': ['total_potency', 'terpene_pct', 'potency_terpene_ratio', 'is_raw'],
            'metadata': {
                'use_arrhenius': True,
                'temperature_feature': 'inverse_temp',
                'target_feature': 'log_viscosity',
                'use_two_step': True,
                'feature_type': 'combined',
                'primary_features': ['total_potency', 'terpene_pct']
            }
        }
    
        # Test the model with different potency values
        print("Testing enhanced demo model with varying potency:")
        for pot in [70, 75, 80, 85, 90]:
            visc = self.predict_model_viscosity(demo_model, 5.0, 25.0, pot)
            print(f"  • Potency {pot}%: viscosity = {visc:.0f} cP")
    
        print("\nTesting with varying terpene percentages (at 80% potency):")
        for terp in [3, 5, 7, 10, 15]:
            visc = self.predict_model_viscosity(demo_model, terp, 25.0, 80)
            print(f"  • Terpene {terp}%: viscosity = {visc:.0f} cP")
    
        # Store in consolidated models dictionary
        if hasattr(self, 'consolidated_models'):
            self.consolidated_models[f'Enhanced_consolidated'] = demo_model
            print("Enhanced demo model added to consolidated models")
        else:
            self.consolidated_models = {f'Enhanced_consolidated': demo_model}
    
        return demo_model

    def analyze_model_feature_response(self, model_key=None, model=None):
        """
        Analyze how a model responds to changes in feature values
        """
        import pandas as pd
        import numpy as np
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        debug_print("Loaded heavy modules for analyze_model_feature_response")
        debug_print(f"model_key: {model_key}")
        debug_print(f"model provided: {model is not None}")
        debug_print("Starting feature response analysis...")
    
        # If no specific model provided, use a model from combined models
        if model is None:
            if model_key is None:
                # Use first model in consolidated models
                if hasattr(self, 'consolidated_models') and self.consolidated_models:
                    model_key = next(iter(self.consolidated_models))
                    model = self.consolidated_models[model_key]
                else:
                    print("No consolidated models available for analysis")
                    return
            else:
                # Try to find the specified model
                if hasattr(self, 'consolidated_models') and model_key in self.consolidated_models:
                    model = self.consolidated_models[model_key]
                else:
                    print(f"Model '{model_key}' not found in consolidated models")
                    return
    
        # Extract model components
        if not isinstance(model, dict) or 'residual_model' not in model:
            print("Invalid model structure for analysis")
            return
    
        # Create analysis window
        analysis_window = Toplevel(self.root)
        analysis_window.title(f"Model Feature Response Analysis: {model_key}")
        analysis_window.geometry("800x600")
    
        # Create figure for plots
        fig = plt.Figure(figsize=(10, 8), tight_layout=True)
    
        # 1. Potency Response Plot
        ax1 = fig.add_subplot(221)
        potencies = np.linspace(0.7, 1.0, 7)  # Range as decimal
        predictions = []
    
        # Fixed values for other parameters
        temperature = 25
        terpene_pct = 5.0
    
        for pot in potencies:
            # Convert potency to percentage for prediction function
            pot_pct = pot * 100
            pred = self.predict_model_viscosity(model, terpene_pct, temperature, pot_pct)
            predictions.append(pred)
    
        ax1.plot(potencies * 100, predictions, 'o-', linewidth=2)
        ax1.set_xlabel('Potency (%)')
        ax1.set_ylabel('Viscosity (cP)')
        ax1.set_title('Viscosity vs Potency')
        ax1.grid(True)
    
        # Calculate responsiveness
        if max(predictions) > min(predictions):
            potency_response = (max(predictions) - min(predictions)) / min(predictions) * 100
            ax1.annotate(f"Δ: {potency_response:.1f}%", 
                         xy=(0.05, 0.95), xycoords='axes fraction',
                         bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))
    
        # 2. Terpene Response Plot
        ax2 = fig.add_subplot(222)
        terpenes = np.linspace(1.0, 20.0, 10)  # Range in percentage
        predictions = []
    
        # Fixed potency (80%)
        potency = 80.0
    
        for terp in terpenes:
            pred = self.predict_model_viscosity(model, terp, temperature, potency)
            predictions.append(pred)
    
        ax2.plot(terpenes, predictions, 'o-', linewidth=2, color='green')
        ax2.set_xlabel('Terpene (%)')
        ax2.set_ylabel('Viscosity (cP)')
        ax2.set_title('Viscosity vs Terpene %')
        ax2.grid(True)
    
        # Calculate responsiveness
        if max(predictions) > min(predictions):
            terpene_response = (max(predictions) - min(predictions)) / min(predictions) * 100
            ax2.annotate(f"Δ: {terpene_response:.1f}%", 
                         xy=(0.05, 0.95), xycoords='axes fraction',
                         bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))
    
        # 3. Temperature Response Plot
        ax3 = fig.add_subplot(223)
        temperatures = np.linspace(20, 70, 11)  # Range in Celsius
        predictions = []
    
        # Fixed values
        terpene_pct = 5.0
        potency = 80.0
    
        for temp in temperatures:
            pred = self.predict_model_viscosity(model, terpene_pct, temp, potency)
            predictions.append(pred)
    
        ax3.semilogy(temperatures, predictions, 'o-', linewidth=2, color='red')
        ax3.set_xlabel('Temperature (°C)')
        ax3.set_ylabel('Viscosity (cP)')
        ax3.set_title('Viscosity vs Temperature')
        ax3.grid(True)
    
        # Calculate responsiveness
        if max(predictions) > min(predictions):
            temp_response = (max(predictions) - min(predictions)) / min(predictions) * 100
            ax3.annotate(f"Δ: {temp_response:.1f}%", 
                         xy=(0.05, 0.95), xycoords='axes fraction',
                         bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))
    
        # 4. Feature Importance
        ax4 = fig.add_subplot(224)
    
        residual_model = model['residual_model']
        residual_features = model.get('residual_features', [])
    
        if hasattr(residual_model, 'feature_importances_'):
            importances = residual_model.feature_importances_
            if len(importances) == len(residual_features):
                # Create bar chart of feature importances
                y_pos = range(len(residual_features))
                ax4.barh(y_pos, importances, align='center')
                ax4.set_yticks(y_pos)
                ax4.set_yticklabels(residual_features)
                ax4.set_xlabel('Importance')
                ax4.set_title('Feature Importance')
        elif hasattr(residual_model, 'coef_'):
            coefs = residual_model.coef_
            if len(coefs) == len(residual_features):
                # Create bar chart of coefficients
                y_pos = range(len(residual_features))
                ax4.barh(y_pos, np.abs(coefs), align='center')
                ax4.set_yticks(y_pos)
                ax4.set_yticklabels(residual_features)
                ax4.set_xlabel('|Coefficient|')
                ax4.set_title('Feature Coefficients')
        else:
            ax4.text(0.5, 0.5, "No feature importance data available", 
                     ha='center', va='center', transform=ax4.transAxes)
    
        # Create canvas to display figure
        canvas = FigureCanvasTkAgg(fig, master=analysis_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
        # Create button to close window
        from tkinter import Button
        Button(analysis_window, text="Close", command=analysis_window.destroy).pack(pady=10)
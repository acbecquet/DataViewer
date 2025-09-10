import os
import pickle
import pandas as pd
import numpy as np
from tkinter import filedialog, messagebox
from sklearn.impute import SimpleImputer
import shutil
from utils import debug_print
class DataProcessing_Methods:
    def load_models(self):
        """
        Load all models needed for the two-level prediction system.
        """
        import pickle
        import os
    
        # Initialize model containers
        self.base_models = {}
        self.composition_models = {}
        self.terpene_profiles = {}
    
        # Load base models
        base_model_path = 'models/viscosity_base_models.pkl'
        if os.path.exists(base_model_path):
            try:
                with open(base_model_path, 'rb') as f:
                    self.base_models = pickle.load(f)
                debug_print(f"Loaded {len(self.base_models)} base models")
            except Exception as e:
                print(f"Error loading base models: {e}")
        else:
            debug_print("No base models found. Please train models first.")
    
        # Load composition models if available
        comp_model_path = 'models/viscosity_composition_models.pkl'
        if os.path.exists(comp_model_path):
            try:
                with open(comp_model_path, 'rb') as f:
                    self.composition_models = pickle.load(f)
                debug_print(f"Loaded {len(self.composition_models)} composition models")
            except Exception as e:
                print(f"Error loading composition models: {e}")
    
        # Load terpene profiles if available
        profile_path = 'models/terpene_profiles.pkl'
        if os.path.exists(profile_path):
            try:
                with open(profile_path, 'rb') as f:
                    self.terpene_profiles = pickle.load(f)
                profile_count = sum(len(profiles) for profiles in self.terpene_profiles.values())
                debug_print(f"Loaded {profile_count} terpene profiles")
            except Exception as e:
                print(f"Error loading terpene profiles: {e}")
    
        # Load default terpene profiles
        self.load_default_terpene_profiles()

    def load_consolidated_models(self):
        """
        Load consolidated viscosity models from disk and clean feature lists.
        """
        import pickle
        import os
    
        model_path = 'models/viscosity_models_consolidated.pkl'
    
        if os.path.exists(model_path):
            try:
                with open(model_path, 'rb') as f:
                    models = pickle.load(f)
                
                    # Clean up models by removing terpene_brand from feature lists
                    cleaned_models = {}
                    for key, model in models.items():
                        cleaned_models[key] = self.remove_terpene_brand_from_features(model)
                
                    self.consolidated_models = cleaned_models
                    debug_print(f"Loaded and cleaned {len(models)} consolidated models from {model_path}")
            except Exception as e:
                print(f"Error loading consolidated models: {e}")
                self.consolidated_models = {}
        else:
            debug_print(f"No consolidated model file found at {model_path}")
            self.consolidated_models = {}

    def upload_training_data(self):
        """
        Allow the user to upload the master CSV data file for training viscosity models.
    
        Only accepts the Master_Viscosity_Data_processed.csv file
        """
        from tkinter import filedialog
        import pandas as pd
        import os
        import shutil

        # Prompt user to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select Master Viscosity Data CSV",
            filetypes=[("CSV files", "*.csv")]
        )

        if not file_path:
            return None  # User canceled

        # Check if this is the correct file
        file_name = os.path.basename(file_path)
        if not file_name.startswith("Master_Viscosity_Data"):
            messagebox.showerror("Invalid File", 
                               "Only Master_Viscosity_Data_processed.csv is accepted.\n"
                               "Please select the correct file.")
            return None

        try:
            # Load the data
            data = pd.read_csv(file_path)
    
            # Validate the data has the required columns
            required_cols = ['media', 'terpene', 'terpene_pct', 'temperature', 'viscosity']
            missing_cols = [col for col in required_cols if col not in data.columns]
    
            if missing_cols:
                messagebox.showerror("Error", 
                                   f"CSV missing required columns: {', '.join(missing_cols)}")
                return None
    
            # Copy the file to the data directory
            os.makedirs('./data', exist_ok=True)
            dest_path = './data/Master_Viscosity_Data_processed.csv'
            
            data.to_csv(dest_path, index = False)
    
            messagebox.showinfo("Success", 
                              f"Loaded {len(data)} data points from {os.path.basename(file_path)}.\n"
                              f"File copied to {dest_path}")
    
            return data
    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
            return None

    def check_features_for_nan(self, feature_df):
        """
        Check a feature DataFrame for NaN values and handle them appropriately:
        - Replace completely empty columns with zeros
        - Impute partially empty columns with mean values
    
        Args:
            feature_df: DataFrame of features to check and clean
    
        Returns:
            DataFrame: The cleaned feature DataFrame with NaN values handled
        """
        # Make a copy to avoid modifying the original unexpectedly
        df = feature_df.copy()
    
        # Check if there are any NaN values at all
        if df.isna().any().any():
            # Find columns that are completely NaN (all values are NaN)
            all_nan_mask = df.isna().all()
            all_nan_columns = df.columns[all_nan_mask].tolist()
        
            # Fill completely empty columns with zeros
            for col in all_nan_columns:
                df.loc[:, col] = 0
        
            # Find columns with at least one non-NaN value but some NaN values
            partial_nan_mask = df.isna().any() & ~all_nan_mask
            partial_nan_columns = df.columns[partial_nan_mask].tolist()
        
            if partial_nan_columns:
                # Check for duplicate column names - this is what's causing the error
                if len(partial_nan_columns) != len(set(partial_nan_columns)):
                    # There are duplicates - we need to handle each column individually
                    for col in set(partial_nan_columns):
                        # Get all positions of this column name (might be duplicated)
                        col_indices = [i for i, c in enumerate(df.columns) if c == col]
                    
                        for idx in col_indices:
                            # Process each duplicate column separately
                            col_series = df.iloc[:, idx]
                            if col_series.isna().any() and not col_series.isna().all():
                                # Replace NaN values with the mean of non-NaN values
                                mean_value = col_series.mean()
                                # Use iloc to avoid the duplicate column issue
                                df.iloc[:, idx] = col_series.fillna(mean_value)
                else:
                    # No duplicates - original method can work
                    # Select only columns with some NaN values (but not all)
                    columns_to_impute = df.loc[:, partial_nan_columns]
                
                    # Create an imputer
                    from sklearn.impute import SimpleImputer
                    imputer = SimpleImputer(strategy='mean')
                    imputed_values = imputer.fit_transform(columns_to_impute)
                
                    # Put the imputed values back
                    df.loc[:, partial_nan_columns] = imputed_values
    
        return df

    def remove_terpene_brand_from_features(self, model_info):
        """
        Creates a copy of the model info with terpene_brand removed from features.
    
        Args:
            model_info: Original model info dictionary
        
        Returns:
            dict: Updated model info with terpene_brand removed from features
        """
        # Create a shallow copy of the model
        updated_model = model_info.copy()
    
        # Update residual_features if present
        if 'residual_features' in updated_model:
            updated_model['residual_features'] = [
                f for f in updated_model['residual_features'] 
                if f != 'terpene_brand'
            ]
    
        # If there's metadata, update feature lists there too
        if 'metadata' in updated_model and isinstance(updated_model['metadata'], dict):
            for key in updated_model['metadata']:
                if isinstance(updated_model['metadata'][key], list):
                    updated_model['metadata'][key] = [
                        f for f in updated_model['metadata'][key] 
                        if f != 'terpene_brand'
                    ]
    
        return updated_model    

    def preprocess_training_data(self, data):
        """Fill missing potency or terpene values using inverse relationship"""
        processed_data = data.copy()
    
        # For rows with missing potency but available terpene percentage
        potency_missing = processed_data['total_potency'].isna()
        terpene_available = ~processed_data['terpene_pct'].isna()
    
        # Apply the inverse relationship: potency = 1 - terpene_pct (as decimal)
        mask = potency_missing & terpene_available
        if mask.any():
            processed_data.loc[mask, 'total_potency'] = 1.0 - (processed_data.loc[mask, 'terpene_pct'] / 100.0)
            debug_print(f"Filled {mask.sum()} missing potency values using inverse relationship")
    
        # For rows with missing terpene but available potency
        terpene_missing = processed_data['terpene_pct'].isna()
        potency_available = ~processed_data['total_potency'].isna()
    
        # Apply the inverse relationship: terpene_pct = (1 - potency) * 100
        mask = terpene_missing & potency_available
        if mask.any():
            processed_data.loc[mask, 'terpene_pct'] = (1.0 - processed_data.loc[mask, 'total_potency']) * 100.0
            debug_print(f"Filled {mask.sum()} missing terpene values using inverse relationship")
    
        return processed_data


import os
import pickle
import pandas as pd
from tkinter import messagebox

class TerpeneProfile_Methods:

    def save_default_terpene_profiles(self):
        """Save default terpene profiles to a file for future use"""
        import pickle
        import os
    
        if hasattr(self, 'default_terpene_profiles'):
            try:
                os.makedirs('models', exist_ok=True)
                with open('models/default_terpene_profiles.pkl', 'wb') as f:
                    pickle.dump(self.default_terpene_profiles, f)
                print(f"Saved {len(self.default_terpene_profiles)} default terpene profiles")
                return True
            except Exception as e:
                print(f"Error saving default terpene profiles: {e}")
                return False
        return False

    def load_default_terpene_profiles(self):
        """Load default terpene profiles from file if available"""
        import pickle
        import os
    
        profile_path = 'models/default_terpene_profiles.pkl'
        if os.path.exists(profile_path):
            try:
                with open(profile_path, 'rb') as f:
                    self.default_terpene_profiles = pickle.load(f)
                print(f"Loaded {len(self.default_terpene_profiles)} default terpene profiles")
                return True
            except Exception as e:
                print(f"Error loading default terpene profiles: {e}")
                # Initialize with built-in defaults on error
                self.initialize_default_terpene_profiles()
                return False
        else:
            # If no file exists, initialize with built-in defaults
            self.initialize_default_terpene_profiles()
            return False

    def initialize_default_terpene_profiles(self):
        """
        Initialize default terpene profiles for common strains when specific breakdowns aren't available.
        Each profile contains approximate percentages of individual terpenes that make up 100% of the terpene content.
        When used, these percentages will be multiplied by the overall terpene percentage.
        """
        # Create default profiles if not already loaded
        if not hasattr(self, 'default_terpene_profiles'):
            self.default_terpene_profiles = {}
        
       
        # Format: Each profile is a dictionary where keys are terpene compounds and values are percentages (should sum to ~100%)
    
        self.default_terpene_profiles["Tiger's Blood"] = {
            'apha-Pinene': 4.2,
            'Camphene': 0,
            'beta-Pinene': 5.6,
            'beta-Myrcene': 0,
            '3-Carene': 0,
            'alpha-Terpinene': 0,
            'p-Cymene': 0,
            'D-Limonene': 0,
            'Ocimene 1': 0,
            'gamma-Terpinene': 0,
            'Terpinolene': 0,
            'Linalool': 0,
            'Isopulegol': 0,
            'Geraniol': 0,
            'Caryophyllene': 1.4,
            'alpha-Humulene': 11.8,
            'Nerolidol 1': 0,
            'Nerolidol 2': 0,
            'Guaiol': 0,
            'alpha-Bisabolol': 0,
            'other':77 
        }

        self.default_terpene_profiles["Guava Gelato"] = {
            'apha-Pinene': 0,
            'Camphene': 0,
            'beta-Pinene': 0,
            'beta-Myrcene': 18.6,
            '3-Carene': 0,
            'alpha-Terpinene': 0,
            'p-Cymene': 0,
            'D-Limonene': 12.3,
            'Ocimene 1': 13.1,
            'gamma-Terpinene': 0,
            'Terpinolene': 0,
            'Linalool': 4.1,
            'Isopulegol': 0,
            'Geraniol': 0,
            'Caryophyllene': 0,
            'alpha-Humulene': 0,
            'Nerolidol 1': 2.8,
            'Nerolidol 2': 0,
            'Guaiol': 0,
            'alpha-Bisabolol': 8.2,
            'other': 40.9
        }

        self.default_terpene_profiles["Grape Ape"] = {
            'apha-Pinene': 17.2,
            'Camphene': 0,
            'beta-Pinene': 7.5,
            'beta-Myrcene': 31.1,
            '3-Carene': 0,
            'alpha-Terpinene': 0,
            'p-Cymene': 0,
            'D-Limonene': 6.1,
            'Ocimene 1': 2.2,
            'gamma-Terpinene': 0,
            'Terpinolene': 0,
            'Linalool': 4.2,
            'Isopulegol': 0,
            'Geraniol': 0,
            'Caryophyllene': 0,
            'alpha-Humulene': 0,
            'Nerolidol 1': 0,
            'Nerolidol 2': 0,
            'Guaiol': 0,
            'alpha-Bisabolol': 0,
            'other': 31.7
        }
            
    
        print(f"Initialized {len(self.default_terpene_profiles)} default terpene profiles")

    def handle_other_terpenes(self, terpene_profile):
        """
        Process terpene profiles that contain an 'Other' category.
    
        Args:
            terpene_profile: Dictionary with terpene names and percentages
        
        Returns:
            Dictionary with adjusted terpene percentages
        """
        # Create a copy to avoid modifying the original
        processed_profile = terpene_profile.copy()
    
        # Check if "Other" exists in the profile
        if 'Other' in processed_profile:
            other_percentage = processed_profile.pop('Other')
        
            # If "Other" is the only component or dominates the profile (>80%), use generic profile
            if len(processed_profile) == 0 or other_percentage > 80:
                # Use the Generic profile but scaled to match the total percentage
                total_percentage = other_percentage + sum(processed_profile.values())
                for terpene, pct in self.default_terpene_profiles['Generic'].items():
                    # Scale generic percentages to match the total percentage
                    scaled_pct = pct * total_percentage / 100.0
                    if terpene in processed_profile:
                        processed_profile[terpene] += scaled_pct
                    else:
                        processed_profile[terpene] = scaled_pct
                return processed_profile
            
            # For profiles with moderate "Other" component (25-50%), use a hybrid approach
            total_known = sum(processed_profile.values())
        
            # 1. Redistribute 70% of "Other" proportionally among known terpenes
            redistribute_amount = other_percentage * 0.7
            for terpene in processed_profile:
                # Calculate proportion of this terpene relative to known terpenes
                if total_known > 0:
                    proportion = processed_profile[terpene] / total_known
                    # Add proportional share of the redistributed amount
                    processed_profile[terpene] += redistribute_amount * proportion
        
            # 2. Allocate 30% of "Other" to common minor terpenes not already in the profile
            minor_allocation = other_percentage * 0.3
            minor_terpenes = {
                'beta-Pinene': 30,
                'alpha-Humulene': 25,
                'Terpinolene': 15,
                'Ocimene': 15,
                'Nerolidol': 15
            }
        
            # Normalize minor terpene percentages to sum to 100%
            minor_total = sum(minor_terpenes.values())
            for terpene, weight in minor_terpenes.items():
                if terpene not in processed_profile:
                    processed_profile[terpene] = minor_allocation * (weight / minor_total)
    
        # Calculate confidence factor (what percentage of terpenes are known)
        if 'Other' in terpene_profile:
            confidence = (100 - terpene_profile['Other']) / 100
            processed_profile['_confidence'] = confidence
        else:
            processed_profile['_confidence'] = 1.0
        
        return processed_profile

    def get_terpene_profile(self, strain_name):
        """
        Get the terpene profile for a given strain, handling cases with "Other" terpenes.
    
        Args:
            strain_name: Name of the cannabis strain
        
        Returns:
            Dictionary with terpene percentages
        """
        # First try to find a specific profile for this strain
        if hasattr(self, 'terpene_profiles') and strain_name in self.terpene_profiles:
            # Found a specific profile from database
            profile = self.terpene_profiles[strain_name].copy()
        
        # Then try to use default profiles
        elif hasattr(self, 'default_terpene_profiles') and strain_name in self.default_terpene_profiles:
            # Found a default profile
            profile = self.default_terpene_profiles[strain_name].copy()
        
        # Fall back to generic profile
        else:
            # Use generic profile if no match found
            profile = self.default_terpene_profiles['Generic'].copy()
    
        # Process the profile to handle any "Other" terpenes
        processed_profile = self.handle_other_terpenes(profile)
    
        return processed_profile

    def export_terpene_profiles(self):
        """
        Export terpene composition profiles to a CSV file.
        """
        import pandas as pd
        import os
    
        if not hasattr(self, 'terpene_profiles') or not self.terpene_profiles:
            messagebox.showinfo("No Profiles", "No terpene profiles are currently loaded.")
            return
    
        # Create directory if needed
        os.makedirs('data', exist_ok=True)
    
        # Collect all profiles into a DataFrame
        rows = []
        for media_type, profiles in self.terpene_profiles.items():
            for terpene_name, profile in profiles.items():
                # Create a row with media type and terpene name
                row = {'media': media_type, 'terpene': terpene_name}
                # Add all composition values
                row.update(profile)
                rows.append(row)
    
        # Create DataFrame
        if rows:
            df = pd.DataFrame(rows)
        
            # Save to CSV
            output_path = 'data/terpene_profiles_export.csv'
            df.to_csv(output_path, index=False)
        
            messagebox.showinfo(
                "Export Complete", 
                f"Exported {len(rows)} terpene profiles to {output_path}"
            )
        else:
            messagebox.showinfo("No Data", "No valid terpene profiles found to export.")
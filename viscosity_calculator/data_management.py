import os
import pandas as pd
import numpy as np
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, StringVar
import json

# These constants should be imported from your core module if they're defined there
from .core import APP_BACKGROUND_COLOR, FONT

class DataManagement_Methods:
    def save_as_csv(self, measurements):
        """
        Save the measurements to the master CSV file.

        Args:
            measurements (dict): The measurements data structure
        """
        # Lazy import pandas only when needed
        import pandas as pd
        import os
        import datetime

        try:
            # Load existing master file if it exists
            master_file = './data/Master_Viscosity_Data_processed.csv'
            if os.path.exists(master_file):
                master_df = pd.read_csv(master_file)
            else:
                # Create directory if it doesn't exist
                os.makedirs('./data', exist_ok=True)
                master_df = pd.DataFrame()
    
            # Create rows for the CSV
            rows = []
    
            media = measurements['media']
            terpene = measurements['terpene']
            #terpene_brand = measurements.get('terpene_brand', '')
            terpene_pct = measurements['terpene_pct']
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
            # Create the combined terpene field
            combined_terpene = terpene
    
            # Process each temperature block
            for temp_block in measurements['temperature_data']:
                temperature = temp_block['temperature']
                speed = temp_block.get('speed', '')
        
                # Process each individual run
                for run in temp_block['runs']:
                    torque = run.get('torque', '')
                    viscosity = run.get('viscosity', '')
            
                    # Only add rows with valid viscosity values
                    if viscosity and viscosity.strip():
                        try:
                            viscosity_float = float(viscosity.replace(',', ''))
                            row = {
                                'media': media,
                                'media_brand': measurements.get('media_brand', ''),
                                'terpene': terpene,
                                #'terpene_brand': terpene_brand,
                                #'combined_terpene': combined_terpene,
                                'terpene_pct': terpene_pct,
                                'temperature': temperature,
                                'speed': speed,
                                'torque': torque,
                                'viscosity': viscosity_float,
                                'timestamp': timestamp
                            }
                            rows.append(row)
                        except ValueError as e:
                            print(f"Warning: Could not convert viscosity value '{viscosity}' to float: {e}")
        
                # Add the average if available
                avg_viscosity = temp_block.get('average_viscosity', '')
                if avg_viscosity and avg_viscosity.strip():
                    try:
                        avg_viscosity_float = float(avg_viscosity.replace(',', ''))
                        row = {
                            'media': media,
                            'media_brand': measurements.get('media_brand', ''),
                            'terpene': terpene,
                            #'terpene_brand': terpene_brand,
                            #'combined_terpene': combined_terpene,
                            'terpene_pct': terpene_pct,
                            'temperature': temperature,
                            'speed': speed,
                            'torque': temp_block.get('average_torque', ''),
                            'viscosity': avg_viscosity_float,
                            'is_average': True,
                            'timestamp': timestamp
                        }
                        rows.append(row)
                    except ValueError as e:
                        print(f"Warning: Could not convert average viscosity value '{avg_viscosity}' to float: {e}")
    
            # Create a DataFrame and append to master file
            if rows:
                new_df = pd.DataFrame(rows)
                if master_df.empty:
                    master_df = new_df
                else:
                    # Ensure all columns exist in both dataframes
                    for col in new_df.columns:
                        if col not in master_df.columns:
                            master_df[col] = None
                    for col in master_df.columns:
                        if col not in new_df.columns:
                            new_df[col] = None
            
                    master_df = pd.concat([master_df, new_df], ignore_index=True)
        
                # Save back to master file
                master_df.to_csv(master_file, index=False)
        
                messagebox.showinfo("Success", 
                                   f"Added {len(rows)} new measurements to master data file.")
                return master_file
            else:
                print("No valid data rows to save")
                return None
        
        except Exception as e:
            import traceback
            print(f"Error saving to master file: {e}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Failed to save measurements: {str(e)}")
            return None

    def save_formulation(self):
        """
        Save the terpene formulation to the database file and append to the master CSV.
        Creates two separate rows: one for step 1 and one for step 2, each with their specific terpene percentages.
        """
        try:
            # Verify we have all required data
            step1_viscosity_text = self.step1_viscosity_var.get()
            step2_viscosity_text = self.step2_viscosity_var.get()
        
            if not step1_viscosity_text:
                messagebox.showinfo("Input Needed", 
                                  "Please enter the measured viscosity from Step 1 before saving.")
                return
            
            if not step2_viscosity_text:
                messagebox.showinfo("Input Needed", 
                                  "Please enter the final measured viscosity from Step 2 before saving.")
                return
        
            # Get all formulation data
            media = self.media_var.get()
            media_brand = self.media_brand_var.get()
            terpene = self.terpene_var.get()
            terpene_brand = self.terpene_brand_var.get()
            target_viscosity = float(self.target_viscosity_var.get())
        
            step1_amount = float(self.step1_amount_var.get().replace('g', ''))
            step1_viscosity = float(step1_viscosity_text)
            step1_terpene_pct = (step1_amount / float(self.mass_of_oil_var.get())) * 100
        
            # For step 2, we need the total terpene amount and percentage
            step2_amount = float(self.step2_amount_var.get().replace('g', ''))
            step2_viscosity = float(step2_viscosity_text)
            total_oil_mass = float(self.mass_of_oil_var.get())
            total_terpene_mass = step1_amount + step2_amount
            total_terpene_pct = (total_terpene_mass / total_oil_mass) * 100
        
            # Calculate potency as 1 - terpene percent (in decimal)
            step1_potency = 1.0 - (step1_terpene_pct / 100.0)
            final_potency = 1.0 - (total_terpene_pct / 100.0)
        
            # Get cannabinoid content if available
            d9_thc = getattr(self, '_d9_thc_var', None)
            d9_thc_value = d9_thc.get() / 100.0 if d9_thc else None
        
            d8_thc = getattr(self, '_d8_thc_var', None)
            d8_thc_value = d8_thc.get() / 100.0 if d8_thc else None
        
            # Current timestamp
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
            # Create the database entry (for JSON formulation database)
            formulation = {
                "media": media,
                "media_brand": media_brand,
                "terpene": terpene,
                #"terpene_brand": terpene_brand,
                "target_viscosity": target_viscosity,
                "step1_amount": step1_amount,
                "step1_viscosity": step1_viscosity,
                "step1_terpene_pct": step1_terpene_pct,
                "step2_amount": step2_amount,
                "step2_viscosity": step2_viscosity,
                "expected_viscosity": float(self.expected_viscosity_var.get()),
                "total_oil_mass": total_oil_mass,
                "total_terpene_mass": total_terpene_mass,
                "total_terpene_pct": total_terpene_pct,
                "step1_potency": step1_potency * 100,  # Store as percentage
                "final_potency": final_potency * 100,  # Store as percentage
                "timestamp": timestamp
            }
        
            # Add to database
            key = f"{formulation['media']}_{formulation['media_brand']}_{formulation['terpene']}"
        
            if key not in self.formulation_db:
                self.formulation_db[key] = []
        
            self.formulation_db[key].append(formulation)
        
            # Save database to file
            self.save_formulation_database()
        
            # Now create entries for the master CSV file
            import pandas as pd
            import os
        
            # Create two separate rows - one for step 1, one for step 2 (final)
            csv_rows = []
        
             # Common fields for both rows
            common_fields = {
                'media': media,
                'media_brand': media_brand,
                'terpene': terpene,
                #'terpene_brand': terpene_brand,
                #'combined_terpene': f"{terpene}_{terpene_brand}" if terpene_brand else terpene,
                'temperature': 25.0,  # Standard measurement temperature
                'timestamp': timestamp
            }
        
            # Add cannabinoid fields if available
            if d9_thc_value is not None:
                common_fields['d9_thc'] = d9_thc_value
        
            if d8_thc_value is not None:
                common_fields['d8_thc'] = d8_thc_value
        
            # Step 1 row
            step1_row = common_fields.copy()
            step1_row.update({
                'terpene_pct': step1_terpene_pct / 100.0,  # Store as decimal in CSV
                'total_potency': step1_potency,  # Store as decimal
                'viscosity': step1_viscosity,
                'measurement_stage': 'step1'
            })
            csv_rows.append(step1_row)
        
            # Step 2 row (final formulation)
            step2_row = common_fields.copy()
            step2_row.update({
                'terpene_pct': total_terpene_pct / 100.0,  # Store as decimal in CSV
                'total_potency': final_potency,  # Store as decimal
                'viscosity': step2_viscosity,
                'measurement_stage': 'step2'
            })
            csv_rows.append(step2_row)
        
            # Append to master CSV file
            master_file = './data/Master_Viscosity_Data_processed.csv'
        
            try:
                # Create DataFrame from new rows
                new_rows_df = pd.DataFrame(csv_rows)
            
                # Check if file exists
                if os.path.exists(master_file):
                    # Load existing data
                    master_df = pd.read_csv(master_file)
                
                    # Ensure all columns are present in both dataframes
                    for col in new_rows_df.columns:
                        if col not in master_df.columns:
                            master_df[col] = None
                
                    for col in master_df.columns:
                        if col not in new_rows_df.columns:
                            new_rows_df[col] = None
                
                    # Append new rows
                    master_df = pd.concat([master_df, new_rows_df], ignore_index=True)
                else:
                    # Create directory if needed
                    os.makedirs(os.path.dirname(master_file), exist_ok=True)
                    # First time creating the file
                    master_df = new_rows_df
            
                # Save to CSV
                master_df.to_csv(master_file, index=False)
                print(f"Added {len(csv_rows)} new measurements to master CSV file.")
            except Exception as e:
                import traceback
                print(f"Error updating master CSV: {str(e)}")
                print(traceback.format_exc())
                # Don't stop execution - we've already saved to the database
        
            # Show success message
            messagebox.showinfo("Success", 
                              f"Formulation saved successfully!\n\n"
                              f"Total terpene percentage: {total_terpene_pct:.2f}%\n"
                              f"Total terpene mass: {total_terpene_mass:.2f}g\n"
                              f"Added step 1 ({step1_terpene_pct:.2f}%) and step 2 ({total_terpene_pct:.2f}%) data points.")
        
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("Input Error", 
                                f"Please ensure all numeric fields contain valid numbers: {str(e)}")
        except Exception as e:
            import traceback
            traceback_str = traceback.format_exc()
            print(f"Error in save_formulation: {str(e)}")
            print(traceback_str)
            messagebox.showerror("Save Error", f"An error occurred: {str(e)}")

    def view_formulation_data(self):
        """
        Display a window for viewing and managing data from the Master_Viscosity_Data_processed.csv file.
        Allows viewing, filtering, and deleting entries.
        """
        import tkinter as tk
        from tkinter import ttk, messagebox
        import pandas as pd
        import os
        import datetime
    
        # CSV file path
        csv_file = './data/Master_Viscosity_Data_processed.csv'
    
        # Check if the file exists
        if not os.path.exists(csv_file):
            messagebox.showerror("File Not Found", 
                               f"The master CSV file was not found at:\n{csv_file}")
            return
    
        # Create the window
        data_window = tk.Toplevel(self.root)
        data_window.title("Formulation Data Manager")
        data_window.geometry("1000x600")
        data_window.minsize(900, 500)
        data_window.configure(bg=APP_BACKGROUND_COLOR)
    
        # Center the window
        self.gui.center_window(data_window)
    
        # Load the CSV data
        try:
            df = pd.read_csv(csv_file)
            original_df = df.copy()  # Keep a copy for comparison when saving changes
        
            # Convert datetime columns if present
            if 'timestamp' in df.columns:
                df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
        
            # Ensure terpene_pct is in decimal format
            if 'terpene_pct' in df.columns:
                # If any values are >1, assume they're in percentage and convert to decimal
                if (df['terpene_pct'] > 1).any():
                    df.loc[df['terpene_pct'] > 1, 'terpene_pct'] = df.loc[df['terpene_pct'] > 1, 'terpene_pct'] / 100
        
            # Same for total_potency
            if 'total_potency' in df.columns:
                if (df['total_potency'] > 1).any():
                    df.loc[df['total_potency'] > 1, 'total_potency'] = df.loc[df['total_potency'] > 1, 'total_potency'] / 100
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
            return
    
        # Create main frames
        top_frame = tk.Frame(data_window, bg=APP_BACKGROUND_COLOR, padx=10, pady=10)
        top_frame.pack(fill="x")
    
        # Add title
        title_label = tk.Label(top_frame, text="Master Viscosity Data Manager", 
                             font=("Arial", 14, "bold"), bg=APP_BACKGROUND_COLOR, fg="white")
        title_label.pack(pady=5)
    
        # Filter frame
        filter_frame = tk.Frame(top_frame, bg=APP_BACKGROUND_COLOR, padx=10, pady=5)
        filter_frame.pack(fill="x")
    
        # Get unique values for dropdowns
        media_options = ['All'] + sorted([str(x) for x in df['media'].unique()]) if 'media' in df.columns else ['All']
        terpene_options = ['All'] + sorted([str(x) for x in df['terpene'].unique()]) if 'terpene' in df.columns else ['All']
    
        # Add step filter options if available
        step_options = ['All']
        if 'measurement_stage' in df.columns:
            step_options += sorted([str(x) for x in df['measurement_stage'].unique()])
    
        # Filter controls
        tk.Label(filter_frame, text="Filter by:", bg=APP_BACKGROUND_COLOR, fg="white", 
               font=FONT).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    
        # Media filter
        tk.Label(filter_frame, text="Media:", bg=APP_BACKGROUND_COLOR, fg="white").grid(
            row=0, column=1, sticky="w", padx=5, pady=5)
    
        media_var = tk.StringVar(value="All")
        media_combo = ttk.Combobox(filter_frame, textvariable=media_var, values=media_options, width=10)
        media_combo.grid(row=0, column=2, sticky="w", padx=5, pady=5)
    
        # Terpene filter
        tk.Label(filter_frame, text="Terpene:", bg=APP_BACKGROUND_COLOR, fg="white").grid(
            row=0, column=3, sticky="w", padx=5, pady=5)
    
        terpene_var = tk.StringVar(value="All")
        terpene_combo = ttk.Combobox(filter_frame, textvariable=terpene_var, values=terpene_options, width=15)
        terpene_combo.grid(row=0, column=4, sticky="w", padx=5, pady=5)
    
        # Step filter
        tk.Label(filter_frame, text="Stage:", bg=APP_BACKGROUND_COLOR, fg="white").grid(
            row=0, column=5, sticky="w", padx=5, pady=5)
    
        step_var = tk.StringVar(value="All")
        step_combo = ttk.Combobox(filter_frame, textvariable=step_var, values=step_options, width=10)
        step_combo.grid(row=0, column=6, sticky="w", padx=5, pady=5)
    
        # Date range filters
        date_frame = tk.Frame(filter_frame, bg=APP_BACKGROUND_COLOR)
        date_frame.grid(row=1, column=0, columnspan=7, sticky="w", pady=5)
    
        tk.Label(date_frame, text="Date Range:", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
    
        today = datetime.datetime.now()
        # Default start date is 30 days ago
        start_date = today - datetime.timedelta(days=10000)
    
        # Format dates for display
        start_date_str = start_date.strftime("%Y-%m-%d")
        end_date_str = today.strftime("%Y-%m-%d")
    
        start_date_var = tk.StringVar(value=start_date_str)
        end_date_var = tk.StringVar(value=end_date_str)
    
        tk.Label(date_frame, text="From:", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
        start_date_entry = ttk.Entry(date_frame, textvariable=start_date_var, width=12)
        start_date_entry.pack(side="left", padx=2)
    
        tk.Label(date_frame, text="To:", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
        end_date_entry = ttk.Entry(date_frame, textvariable=end_date_var, width=12)
        end_date_entry.pack(side="left", padx=2)
    
        # Info about date format
        tk.Label(date_frame, text="(YYYY-MM-DD)", bg=APP_BACKGROUND_COLOR, fg="white").pack(side="left", padx=5)
    
        # Create main treeview for data display
        tree_frame = tk.Frame(data_window)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
        # Add scrollbars
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        y_scroll.pack(side="right", fill="y")
    
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        x_scroll.pack(side="bottom", fill="x")
    
        # Define columns based on CSV headers
        columns = df.columns.tolist()
    
        # Create treeview
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                          yscrollcommand=y_scroll.set,
                          xscrollcommand=x_scroll.set)
    
        # Configure columns
        for col in columns:
            # Format column headings nicely
            display_name = col.replace('_', ' ').title()
            tree.heading(col, text=display_name)
        
            # Set column widths based on data type
            if col == 'timestamp':
                tree.column(col, width=150)
            elif 'media' in col or 'terpene' in col:
                tree.column(col, width=100)
            elif 'pct' in col or 'potency' in col:
                tree.column(col, width=80)
            else:
                tree.column(col, width=120)
    
        # Link scrollbars
        y_scroll.config(command=tree.yview)
        x_scroll.config(command=tree.xview)
    
        # Pack the treeview
        tree.pack(fill="both", expand=True)
    
        # Button frame
        button_frame = tk.Frame(data_window, bg=APP_BACKGROUND_COLOR, padx=10, pady=10)
        button_frame.pack(fill="x")
    
        # Status variable
        status_var = tk.StringVar(value=f"Loaded {len(df)} records")
        status_label = tk.Label(button_frame, textvariable=status_var, 
                              bg=APP_BACKGROUND_COLOR, fg="white")
        status_label.pack(side="left", padx=5)
    
        # Apply filter function
        def apply_filter():
            # Start with the full dataframe
            filtered_df = df.copy()
            print(f"Starting with {len(filtered_df)} records")  # Debug print
        
            # Apply media filter
            if media_var.get() != "All" and 'media' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['media'] == media_var.get()]
        
            # Apply terpene filter
            if terpene_var.get() != "All" and 'terpene' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['terpene'] == terpene_var.get()]
        
            # Apply step filter
            if step_var.get() != "All" and 'measurement_stage' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['measurement_stage'] == step_var.get()]
        
            # Apply date filter if timestamp column exists
            if 'timestamp' in filtered_df.columns:
                try:
                    start_date = pd.to_datetime(start_date_var.get())
                    end_date = pd.to_datetime(end_date_var.get())
        
                    # Add one day to end_date to include the entire end day
                    end_date = end_date + pd.Timedelta(days=1)
        
                    # Check how many records have valid timestamps versus NaN
                    valid_timestamps = filtered_df['timestamp'].notna().sum()
                    print(f"Records with valid timestamps: {valid_timestamps} out of {len(filtered_df)}")
        
                    # Only apply date filter to records with timestamps
                    # For records without timestamps, keep them regardless of date range
                    mask = (
                        ((filtered_df['timestamp'] >= start_date) & (filtered_df['timestamp'] < end_date)) | 
                        filtered_df['timestamp'].isna()
                    )
                    filtered_df = filtered_df[mask]
                    print(f"After date filter: {len(filtered_df)} records")
        
                except Exception as e:
                    print(f"Date parsing error: {e}")
                    # If date parsing fails, don't apply date filter
        
            # Clear existing data
            for i in tree.get_children():
                tree.delete(i)

            print(f"Final filtered records: {len(filtered_df)}")
        
            # Populate treeview with filtered data
            for idx, row in filtered_df.iterrows():
                values = []
            
                for col in columns:
                    # Format based on column type
                    if col == 'timestamp' and pd.notna(row[col]):
                        if isinstance(row[col], pd.Timestamp):
                            values.append(row[col].strftime('%Y-%m-%d %H:%M'))
                        else:
                            values.append(str(row[col]))
                    elif col in ['terpene_pct', 'total_potency']:
                        # Display as percentage
                        if pd.notna(row[col]):
                            values.append(f"{row[col]*100:.2f}%")
                        else:
                            values.append("")
                    else:
                        values.append(str(row[col]) if pd.notna(row[col]) else "")
            
                tree.insert('', 'end', values=values, tags=('row',))
        
            # Update status
            status_var.set(f"Displaying {len(filtered_df)} of {len(df)} records")
    
        # Function to delete selected records
        def delete_selected():
            nonlocal df
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("No Selection", "Please select records to delete.")
                return

            # Confirm deletion
            if not messagebox.askyesno("Confirm Delete", 
                                        f"Are you sure you want to delete {len(selected)} selected records?"):
                return

            # Create a list to store indices to delete
            to_delete = []
    
            # For debugging
            print(f"Attempting to delete {len(selected)} records")
    
            # Save original dataframe for potential rollback
            original_df_copy = df.copy()
    
            # Process each selected item
            for item_id in selected:
                # Get values from the treeview
                item_values = tree.item(item_id)['values']
        
                # Find matching rows using more flexible matching criteria
                if len(item_values) >= 3:  # Ensure we have enough columns to match
                    # Use key identifiable columns for matching
                    media_value = str(item_values[columns.index('media')] if 'media' in columns else '')
                    terpene_value = str(item_values[columns.index('terpene')] if 'terpene' in columns else '')
                    temp_value = item_values[columns.index('temperature')] if 'temperature' in columns else None
                    visc_value = item_values[columns.index('viscosity')] if 'viscosity' in columns else None
            
                    # Print debug info
                    print(f"Looking for: media={media_value}, terpene={terpene_value}, temp={temp_value}, visc={visc_value}")
            
                    # Create a flexible matching criteria
                    for idx, row in df.iterrows():
                        # Check if critical columns match
                        match = True
                
                        if media_value and 'media' in row and str(row['media']) != media_value:
                            match = False
                    
                        if terpene_value and 'terpene' in row and str(row['terpene']) != terpene_value:
                            match = False
                
                        # For numeric values, use approximate matching
                        if temp_value is not None and 'temperature' in row:
                            try:
                                if abs(float(row['temperature']) - float(temp_value)) > 0.01:
                                    match = False
                            except (ValueError, TypeError):
                                match = False
                
                        if visc_value is not None and 'viscosity' in row:
                            try:
                                # Handle percentage string or numeric value
                                tree_visc = visc_value
                                if isinstance(visc_value, str) and '%' in visc_value:
                                    tree_visc = float(visc_value.replace('%', ''))
                        
                                # Allow for some margin of error in floating point comparisons
                                if abs(float(row['viscosity']) - float(tree_visc)) > 0.1:
                                    match = False
                            except (ValueError, TypeError):
                                match = False
                
                        if match:
                            to_delete.append(idx)
                            print(f"Found match at index {idx}")
                            break

            # Remove duplicates
            to_delete = list(set(to_delete))
    
            if to_delete:
                print(f"Will delete {len(to_delete)} records")
        
                try:
                    # Step 1: Delete from dataframe without leaving gaps
                    # Create a copy of the dataframe with the filtered rows
                    df_updated = df.drop(to_delete).reset_index(drop=True)
            
                    # Step 2: Try to save to the master file
                    master_file = csv_file  # Use the same file path loaded earlier
            
                    # Create backup before modifying anything
                    backup_file = f"{master_file}.bak.{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    original_df.to_csv(backup_file, index=False)
                    print(f"Backup saved to {backup_file}")
            
                    # Try to save the updated dataframe
                    df_updated.to_csv(master_file, index=False)
                    print(f"Updated dataset saved to {master_file}")
            
                    # If we got here, the file save was successful
                    # Update the global dataframe reference
                    
                    df = df_updated
            
                    # Refresh the view
                    apply_filter()
            
                    # Update status
                    status_var.set(f"Deleted {len(to_delete)} records. {len(df)} records remaining. File updated.")
            
                except Exception as e:
                    # If any error occurs during the save, roll back to the original dataframe
                    error_msg = f"Error updating file: {str(e)}\nChanges have been reverted to maintain consistency."
                    print(error_msg)
            
                    # Restore the original dataframe
                    
                    df = original_df_copy
            
                    # Show error message
                    messagebox.showerror("Save Error", error_msg)
            
                    # Refresh the view with original data
                    apply_filter()
            
            else:
                print("No matching records found for deletion")
                messagebox.showwarning("No Matches", "Could not identify the selected records in the dataset.")
    
        # Function to save changes to the CSV file
        def save_changes():
            nonlocal original_df
            if len(df) == len(original_df) and all(df.eq(original_df).all()):
                messagebox.showinfo("No Changes", "No changes have been made to the data.")
                return
        
            # Confirm save
            if not messagebox.askyesno("Confirm Save", 
                                        "Save changes to the master CSV file?\nThis will overwrite the existing file."):
                return
        
            try:
                # Create backup
                backup_path = f"{csv_file}.bak"
                original_df.to_csv(backup_path, index=False)
            
                # Save changes
                df.to_csv(csv_file, index=False)
            
                messagebox.showinfo("Save Complete", 
                                    f"Changes saved successfully to {csv_file}\nBackup created at {backup_path}")
            
                # Update original_df reference
                
                original_df = df.copy()
            except Exception as e:
                messagebox.showerror("Save Error", f"Error saving changes: {str(e)}")
    
        # Function to export selected records
        def export_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("No Selection", "Please select records to export.")
                return
        
            # Ask for file location
            from tkinter import filedialog
            export_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Selected Records"
            )
        
            if not export_path:
                return  # User cancelled
        
            try:
                # Get selected indices
                selected_indices = []
                for item_id in selected:
                    item_values = tree.item(item_id)['values']
                
                    # Match with dataframe rows (same logic as delete function)
                    typed_values = []
                    for i, col in enumerate(columns):
                        val = item_values[i]
                    
                        if col in ['terpene_pct', 'total_potency'] and isinstance(val, str) and '%' in val:
                            val = float(val.replace('%', '')) / 100
                    
                        if col == 'timestamp' and isinstance(val, str):
                            try:
                                val = pd.to_datetime(val)
                            except:
                                pass
                    
                        typed_values.append(val)
                
                    for idx, row in df.iterrows():
                        match = True
                        for i, col in enumerate(columns):
                            row_val = row[col]
                            tree_val = typed_values[i]
                        
                            if col == 'timestamp':
                                if isinstance(row_val, pd.Timestamp) and isinstance(tree_val, str):
                                    if row_val.strftime('%Y-%m-%d %H:%M') != tree_val:
                                        match = False
                                        break
                                elif row_val != tree_val:
                                    match = False
                                    break
                            elif col in ['terpene_pct', 'total_potency']:
                                row_float = float(row_val) if pd.notna(row_val) else 0
                                tree_float = float(tree_val.replace('%', '')) / 100 if isinstance(tree_val, str) and '%' in tree_val else float(tree_val) if pd.notna(tree_val) else 0
                                if abs(row_float - tree_float) > 0.0001:
                                    match = False
                                    break
                            elif str(row_val) != str(tree_val):
                                match = False
                                break
                    
                        if match:
                            selected_indices.append(idx)
                            break
            
                # Remove duplicates
                selected_indices = list(set(selected_indices))
            
                if selected_indices:
                    # Create export dataframe
                    export_df = df.loc[selected_indices].copy()
                
                    # Save to CSV
                    export_df.to_csv(export_path, index=False)
                
                    messagebox.showinfo("Export Complete", 
                                        f"Successfully exported {len(export_df)} records to {export_path}")
                else:
                    messagebox.showwarning("No Matches", "Could not identify the selected records for export.")
            except Exception as e:
                messagebox.showerror("Export Error", f"Error exporting data: {str(e)}")
    
        # Add buttons for filter apply and refresh
        filter_button = ttk.Button(filter_frame, text="Apply Filter", command=apply_filter)
        filter_button.grid(row=0, column=7, padx=5, pady=5)
    
        # Add action buttons
        delete_button = ttk.Button(button_frame, text="Delete Selected", command=delete_selected)
        delete_button.pack(side="right", padx=5)
    
        save_button = ttk.Button(button_frame, text="Save Changes", command=save_changes)
        save_button.pack(side="right", padx=5)
    
        export_button = ttk.Button(button_frame, text="Export Selected", command=export_selected)
        export_button.pack(side="right", padx=5)
    
        # Apply initial filter
        apply_filter()
    
        # Add alternating row colors
        tree.tag_configure('row', background='#f0f0ff')
    
        # Set multiple selection mode
        tree.config(selectmode='extended')
    
        # Make the window modal
        data_window.transient(self.root)
        data_window.grab_set()
        data_window.focus_set()   

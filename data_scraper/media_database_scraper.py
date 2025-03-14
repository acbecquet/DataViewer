import pandas as pd
import numpy as np
import os
from datetime import datetime

def extract_media_database(excel_path):
    """
    Extract data from the complex multi-block Excel structure and format
    it for viscosity analysis with Total Potency included.
    
    Args:
        excel_path: Path to the Excel file
        
    Returns:
        DataFrame with combined data in the format ready for viscosity analysis
    """
    print(f"Loading Excel file from: {excel_path}")
    
    try:
        # Load the excel file without specifying headers
        xl = pd.ExcelFile(excel_path)
        sheet_names = xl.sheet_names
        
        if len(sheet_names) < 3:
            print(f"Warning: Expected 3 sheets, but found {len(sheet_names)}")
            print(f"Available sheets: {sheet_names}")
        
        # Read each sheet without headers
        potency_df = pd.read_excel(xl, sheet_name=sheet_names[0], header=None)
        terpene_df = pd.read_excel(xl, sheet_name=sheet_names[1], header=None)
        viscosity_df = pd.read_excel(xl, sheet_name=sheet_names[2], header=None)
        
        print(f"Processing sheets: {sheet_names}")
        print(f"Sheet dimensions - Potency: {potency_df.shape}, Terpenes: {terpene_df.shape}, Viscosity: {viscosity_df.shape}")
        
        # Extract potency data
        potency_data = extract_potency_blocks(potency_df)
        
        # Extract terpene data
        terpene_data = extract_terpene_blocks(terpene_df)
        
        # Extract viscosity data - use our new specialized function
        viscosity_data = extract_viscosity_blocks(viscosity_df)
        
        # Merge the data
        merged_data = merge_data(potency_data, terpene_data, viscosity_data)
        
        # Format for training
        formatted_data = format_for_training(merged_data)
        
        return formatted_data
        
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def extract_potency_blocks(df):
    """
    Extract potency data from the potency sheet, with special focus on Total Potency
    and specific cannabinoid columns like d9-THC.
    """
    all_potency_data = []
    
    # Find media type header rows
    media_type_rows = []
    for i in range(len(df)):
        row_str = " ".join([str(val) for val in df.iloc[i].values if str(val) != 'nan'])
        row_str = row_str.lower()
        
        for media in ["live resin", "live rosin", "d8", "d9", "distillate", "liquid diamonds"]:
            if media in row_str and len(row_str.strip()) < 30:
                media_type_rows.append((i, media))
                break
    
    print(f"Found {len(media_type_rows)} media type headers in potency sheet")
    
    # Process each block
    for block_idx, (media_row, media_type) in enumerate(media_type_rows):
        # Determine end of block
        next_block = len(df)
        if block_idx < len(media_type_rows) - 1:
            next_block = media_type_rows[block_idx + 1][0]
        
        # Find the brand/strain header row
        header_row = None
        for i in range(media_row + 1, min(media_row + 5, next_block)):
            row_str = " ".join([str(val) for val in df.iloc[i].values if str(val) != 'nan']).lower()
            if "brand" in row_str and "strain" in row_str:
                header_row = i
                break
        
        if header_row is None:
            print(f"Warning: Could not find header row for block {block_idx+1} ({media_type})")
            continue
        
        # Get column names
        headers = []
        for col in range(df.shape[1]):
            val = df.iloc[header_row, col]
            if pd.notna(val) and str(val).strip():
                headers.append(str(val).strip())
            else:
                headers.append(f"Column_{col}")
        
        # Debug: Print identified headers for this block
        print(f"Block {block_idx+1} ({media_type}) headers: {headers}")
        
        # Process data rows
        for row_idx in range(header_row + 1, next_block):
            row = df.iloc[row_idx]
            
            # Skip rows with missing brand or strain
            if pd.isna(row[0]) or pd.isna(row[1]):
                continue
            
            # Get brand and strain
            brand = str(row[0]).strip()
            strain = str(row[1]).strip()
            
            # Skip empty rows
            if not brand or not strain:
                continue
            
            # Create data dictionary
            potency_data = {
                'media_brand': brand,
                'strain': strain,
                'media_type': media_type
            }
            
            # Add data for all columns, but especially check for Total Potency and d9-THC
            total_potency_found = False
            d9_thc_found = False
            
            for col in range(2, min(len(headers), len(row))):
                if pd.notna(row[col]):
                    col_name = headers[col]
                    
                    try:
                        # Convert to float, removing any % signs
                        value = float(str(row[col]).replace('%', '').strip())
                        potency_data[col_name] = value
                        
                        # Track if we found these key columns
                        if 'total potency' in col_name.lower():
                            total_potency_found = True
                        if 'd9-thc' in col_name.lower() or 'd9 thc' in col_name.lower():
                            d9_thc_found = True
                            
                    except (ValueError, TypeError):
                        # Skip values that can't be converted to float
                        pass
            
            # Check if we found the critical columns
            if not total_potency_found:
                print(f"  Warning: No 'Total Potency' found for {brand} - {strain}")
            if not d9_thc_found:
                print(f"  Warning: No 'd9-THC' found for {brand} - {strain}")
                
            all_potency_data.append(potency_data)
    
    # Convert to DataFrame
    if all_potency_data:
        result_df = pd.DataFrame(all_potency_data)
        print(f"Extracted {len(result_df)} potency data points")
        
        # Print out column names to verify Total Potency was captured
        print(f"Columns in potency data: {result_df.columns.tolist()}")
        
        return result_df
    else:
        print("No potency data extracted")
        return pd.DataFrame()

def extract_terpene_blocks(df):
    """
    Extract terpene data from the terpene sheet, with special focus on Total Terpene column.
    """
    all_terpene_data = []
    
    # Find media type header rows
    media_type_rows = []
    for i in range(len(df)):
        row_str = " ".join([str(val) for val in df.iloc[i].values if str(val) != 'nan'])
        row_str = row_str.lower()
        
        for media in ["live resin", "live rosin", "d8", "d9", "distillate", "liquid diamonds"]:
            if media in row_str and len(row_str.strip()) < 30:
                media_type_rows.append((i, media))
                break
    
    print(f"Found {len(media_type_rows)} media type headers in terpene sheet")
    
    # Process each block
    for block_idx, (media_row, media_type) in enumerate(media_type_rows):
        # Determine end of block
        next_block = len(df)
        if block_idx < len(media_type_rows) - 1:
            next_block = media_type_rows[block_idx + 1][0]
        
        # Find header with "Name of Oil" and "Strain"
        header_row = None
        for i in range(media_row + 1, min(media_row + 5, next_block)):
            row_values = [str(val).lower() for val in df.iloc[i] if pd.notna(val)]
            row_str = " ".join(row_values)
            
            # Look for either 'name of oil' and 'strain' OR 'weight %' header
            if ('name of oil' in row_str and 'strain' in row_str) or 'weight %' in row_str:
                # Look ahead to see if the next row has actual column headers
                if i+1 < next_block:
                    next_row_values = [str(val).lower() for val in df.iloc[i+1] if pd.notna(val)]
                    if 'alpha' in " ".join(next_row_values) or 'total terpene' in " ".join(next_row_values):
                        header_row = i+1
                    else:
                        header_row = i
                else:
                    header_row = i
                break
        
        if header_row is None:
            print(f"Warning: Could not find header row for block {block_idx+1} ({media_type})")
            continue
        
        # Get column names, using the previous few rows to find "Name of Oil" and "Strain" if needed
        name_col = -1
        strain_col = -1
        
        # First look in the rows just before the header row to find the name/strain columns
        for i in range(max(0, header_row-2), header_row+1):
            row_values = [str(val).lower() for val in df.iloc[i] if pd.notna(val)]
            for col, val in enumerate(df.iloc[i]):
                if pd.notna(val):
                    val_str = str(val).lower()
                    if 'name of oil' in val_str:
                        name_col = col
                    elif 'strain' in val_str:
                        strain_col = col
        
        if name_col == -1 or strain_col == -1:
            print(f"  Warning: Could not find Name of Oil and Strain columns in block {block_idx+1}")
            # Default to first two columns
            name_col = 0
            strain_col = 1
        
        # Get all column headers
        headers = []
        for col in range(df.shape[1]):
            val = df.iloc[header_row, col]
            if pd.notna(val) and str(val).strip():
                headers.append(str(val).strip())
            else:
                headers.append(f"Column_{col}")
        
        # Debug
        print(f"Block {block_idx+1} ({media_type}) headers: {headers}")
        
        # Determine row with data start
        data_start_row = header_row + 1
        
        # Find the 'Total Terpene' column
        total_terpene_col = -1
        for col, header in enumerate(headers):
            if 'total terpene' in header.lower():
                total_terpene_col = col
                print(f"  Found 'Total Terpene' column at index {col}")
                break
        
        if total_terpene_col == -1:
            print(f"  WARNING: No 'Total Terpene' column found in block {block_idx+1}")
        
        # Process data rows
        for row_idx in range(data_start_row, next_block):
            row = df.iloc[row_idx]
            
            # Skip rows with missing brand or strain
            if pd.isna(row[name_col]) or pd.isna(row[strain_col]):
                continue
            
            # Get brand and strain
            brand = str(row[name_col]).strip()
            strain = str(row[strain_col]).strip()
            
            # Skip empty rows
            if not brand or not strain:
                continue
            
            # Create data dictionary
            terpene_data = {
                'media_brand': brand,
                'strain': strain,
                'media_type': media_type
            }
            
            # Process all columns
            for col in range(0, min(len(headers), len(row))):
                if col == name_col or col == strain_col:
                    continue  # Skip name and strain columns
                
                if pd.notna(row[col]):
                    col_name = headers[col]
                    try:
                        # Convert to float, removing any % signs
                        value = float(str(row[col]).replace('%', '').strip())
                        terpene_data[col_name] = value
                    except (ValueError, TypeError):
                        # Skip values that can't be converted to float
                        pass
            
            # Check if we have the total terpene value
            if total_terpene_col >= 0 and pd.notna(row[total_terpene_col]):
                try:
                    total_terpene_value = float(str(row[total_terpene_col]).replace('%', '').strip())
                    terpene_data['Total Terpene'] = total_terpene_value
                except (ValueError, TypeError):
                    print(f"  Warning: Could not convert Total Terpene value for {brand} - {strain}")
            
            all_terpene_data.append(terpene_data)
    
    # Convert to DataFrame
    if all_terpene_data:
        result_df = pd.DataFrame(all_terpene_data)
        print(f"Extracted {len(result_df)} terpene data points")
        
        # Print out column names to verify Total Terpene was captured
        print(f"Columns in terpene data: {result_df.columns.tolist()}")
        
        return result_df
    else:
        print("No terpene data extracted")
        return pd.DataFrame()

def merge_data(potency_df, terpene_df, viscosity_df):
    """
    Merge the potency, terpene, and viscosity data.
    """
    # Handle empty DataFrames
    if potency_df.empty and terpene_df.empty and viscosity_df.empty:
        print("All data frames are empty, nothing to merge")
        return pd.DataFrame()
    
    # Merge terpene data into potency data first (if both exist)
    if not potency_df.empty and not terpene_df.empty:
        merged = pd.merge(
            potency_df, 
            terpene_df,
            on=['media_brand', 'strain', 'media_type'],
            how='outer',
            suffixes=('_potency', '_terpene')
        )
    elif not potency_df.empty:
        merged = potency_df
    elif not terpene_df.empty:
        merged = terpene_df
    else:
        merged = pd.DataFrame()
    
    # Then merge with viscosity data
    if not merged.empty and not viscosity_df.empty:
        final_merged = pd.merge(
            merged,
            viscosity_df,
            on=['media_brand', 'strain', 'media_type'],
            how='outer'
        )
    elif not viscosity_df.empty:
        final_merged = viscosity_df
    else:
        final_merged = merged
    
    print(f"Merged data shape: {final_merged.shape}")
    return final_merged

def format_for_training(df):
    """
    Format the merged data for viscosity training, ensuring Total Potency and
    Total Terpene values make it to the final output.
    """
    if df.empty:
        return pd.DataFrame()
    
    # Debug: Print all available columns to help identify what's available
    print("\nAvailable columns in merged data:")
    for col in sorted(df.columns):
        print(f"  - {col}")
    
    # Create a new DataFrame with columns matching the expected training format
    training_df = pd.DataFrame()
    
    # Map source columns to training format columns
    if 'media_type' in df.columns:
        training_df['media'] = df['media_type']
    else:
        training_df['media'] = 'Unknown'
        
    if 'media_brand' in df.columns:
        training_df['media_brand'] = df['media_brand']
    
    if 'strain' in df.columns:
        training_df['terpene'] = df['strain']
    
    # Add total potency with flexible column name matching
    potency_candidates = ['Total Potency', 'TOTAL POTENCY', 'Total potency']
    total_potency_column = None
    
    # Find the actual column name for Total Potency
    for candidate in potency_candidates:
        if candidate in df.columns:
            total_potency_column = candidate
            break
    
    # Also check with case-insensitive search
    if total_potency_column is None:
        for col in df.columns:
            if 'total potency' in col.lower():
                total_potency_column = col
                break
    
    if total_potency_column:
        print(f"Using column '{total_potency_column}' for total_potency")
        training_df['total_potency'] = pd.to_numeric(df[total_potency_column], errors='coerce')
    else:
        print("WARNING: Could not find Total Potency column")
    
    # Add total terpenes with flexible column name matching
    terpene_candidates = ['Total Terpene', 'TOTAL TERPENE', 'Total terpene']
    total_terpene_column = None
    
    # Find the actual column name for Total Terpene
    for candidate in terpene_candidates:
        if candidate in df.columns:
            total_terpene_column = candidate
            break
    
    # Also check with case-insensitive search
    if total_terpene_column is None:
        for col in df.columns:
            if 'total terpene' in col.lower():
                total_terpene_column = col
                break
    
    if total_terpene_column:
        print(f"Using column '{total_terpene_column}' for terpene_pct")
        training_df['terpene_pct'] = pd.to_numeric(df[total_terpene_column], errors='coerce')
    else:
        print("WARNING: Could not find Total Terpene column")
    
    # Add temperature and viscosity if available
    if 'temperature' in df.columns:
        training_df['temperature'] = pd.to_numeric(df['temperature'], errors='coerce')
    
    if 'viscosity' in df.columns:
        training_df['viscosity'] = pd.to_numeric(df['viscosity'], errors='coerce')
    
    # Add key cannabinoids if available
    cannabinoid_cols = {
        'd9-THC': 'd9_thc',
        'THCA': 'thca',
        'CBD': 'cbd',
        'CBDA': 'cbda',
        'd8-THC': 'd8_thc',
        'CBG': 'cbg'
    }
    
    for source_col, target_col in cannabinoid_cols.items():
        if source_col in df.columns:
            print(f"Including cannabinoid column {source_col}")
            training_df[target_col] = pd.to_numeric(df[source_col], errors='coerce')
    
    # Add timestamp and source
    from datetime import datetime
    training_df['timestamp'] = datetime.now().isoformat()
    training_df['original_filename'] = 'Media Database V3 2023'
    
    # Remove rows without viscosity data (if focusing on training data)
    if 'viscosity' in training_df.columns:
        training_df = training_df.dropna(subset=['viscosity'])
    
    print(f"\nFinal training data shape: {training_df.shape}")
    print(f"Final training data columns: {training_df.columns.tolist()}")
    
    # Print a sample of potency and terpene values to verify
    if len(training_df) > 0:
        print("\nSample of extracted values:")
        if 'total_potency' in training_df.columns:
            print(f"total_potency sample: {training_df['total_potency'].head(3).tolist()}")
        if 'terpene_pct' in training_df.columns:
            print(f"terpene_pct sample: {training_df['terpene_pct'].head(3).tolist()}")
    
    return training_df

def extract_viscosity_blocks(df):
    """
    Extract viscosity data from the viscosity sheet, handling the complex structure
    with multiple blocks and different temperature column layouts.
    """
    all_viscosity_data = []
    
    # Debugging
    print(f"Original viscosity dataframe shape: {df.shape}")
    
    # Find all occurrences of "Name of Oil" which indicate block headers
    name_rows = []
    for i in range(len(df)):
        row_values = [str(val).strip() for val in df.iloc[i] if pd.notna(val)]
        row_str = " ".join(row_values)
        if "Name of Oil" in row_str:
            name_rows.append(i)
            print(f"Found 'Name of Oil' at row {i}: {row_str}")
    
    if not name_rows:
        print("Could not find any 'Name of Oil' headers")
        return pd.DataFrame()
    
    print(f"Found {len(name_rows)} block headers at rows: {name_rows}")
    
    # Process each block
    for block_idx, name_row in enumerate(name_rows):
        # Determine end of block (next name row or end of sheet)
        next_block = name_rows[block_idx + 1] if block_idx < len(name_rows) - 1 else len(df)
        print(f"\nProcessing block {block_idx+1} (rows {name_row}-{next_block-1})")
        
        # Find the temperature values - they could be in the next row or in column headers
        is_horizontal_temps = False
        temp_row = name_row + 1
        temp_values = []
        
        # Check if "Temperature" appears in the row after the header
        temp_row_str = " ".join([str(val) for val in df.iloc[temp_row] if pd.notna(val)])
        if "Temperature" in temp_row_str:
            print(f"Found temperature row at {temp_row}: {temp_row_str}")
            # Look for a single temperature in this row (Block 1 style)
            for col in range(2, df.shape[1]):
                cell_value = df.iloc[temp_row, col]
                if pd.notna(cell_value) and "Temperature" not in str(cell_value):
                    try:
                        # Extract numeric value, handling different formats
                        temp_str = str(cell_value).replace('C', '').replace('oC', '').strip()
                        temp_val = float(temp_str)
                        print(f"  Found temperature {temp_val} at column {col}")
                        temp_values.append((col, temp_val))
                    except (ValueError, TypeError) as e:
                        print(f"  Could not convert value to temperature: {cell_value}, error: {e}")
        else:
            # Look for temperature values in the cells of the name_row (Block 2/3 style)
            for col in range(2, df.shape[1]):
                cell_value = df.iloc[name_row, col]
                if pd.notna(cell_value) and "Temperature" in str(cell_value):
                    is_horizontal_temps = True
                    # Now check the next row for the actual temperature values
                    temp_headers_row = name_row + 1
                    for tcol in range(2, df.shape[1]):
                        temp_cell = df.iloc[temp_headers_row, tcol]
                        if pd.notna(temp_cell):
                            try:
                                temp_str = str(temp_cell).replace('C', '').replace('oC', '').strip()
                                temp_val = float(temp_str)
                                print(f"  Found horizontal temperature {temp_val} at column {tcol}")
                                temp_values.append((tcol, temp_val))
                            except (ValueError, TypeError) as e:
                                print(f"  Could not convert horizontal temp: {temp_cell}, error: {e}")
                    break
        
        if not temp_values:
            print(f"  Warning: No temperature values found for block {block_idx+1}")
            continue
        
        print(f"  Found {len(temp_values)} temperature values")
        
        # Set the start row for data (depends on whether temps are horizontal)
        data_start_row = temp_row + 1 if not is_horizontal_temps else name_row + 2
        
        # Process data rows
        for row_idx in range(data_start_row, next_block):
            row = df.iloc[row_idx]
            
            # Skip rows with missing brand or strain
            if pd.isna(row[0]) or pd.isna(row[1]):
                continue
            
            # Get brand and strain
            brand = str(row[0]).strip()
            strain = str(row[1]).strip()
            
            # Skip empty rows
            if not brand or not strain or "Name of Oil" in brand:
                continue
            
            # Detect media type from strain field
            media_type = "Unknown"
            strain_lower = strain.lower()
            for media in ["D8", "D9", "Live Resin", "Live Rosin", "Liquid Diamonds", "Distillate"]:
                if media.lower() in strain_lower:
                    media_type = media
                    break
            
            print(f"  Processing row {row_idx}: {brand} - {strain} ({media_type})")
            
            # Process each temperature column
            for col_idx, temp_val in temp_values:
                if col_idx >= len(row):
                    continue
                    
                visc_value = row[col_idx]
                
                # Skip empty viscosity values
                if pd.isna(visc_value) or str(visc_value).strip() == '':
                    continue
                
                try:
                    # Convert to numeric, handling different formats
                    if isinstance(visc_value, str):
                        visc_value = float(visc_value.replace(',', ''))
                    else:
                        visc_value = float(visc_value)
                        
                    print(f"    Found viscosity {visc_value} at temp {temp_val}")
                    
                    # Add to results
                    all_viscosity_data.append({
                        'media_brand': brand,
                        'strain': strain,
                        'media_type': media_type,
                        'temperature': temp_val,
                        'viscosity': visc_value
                    })
                except (ValueError, TypeError) as e:
                    print(f"    Could not convert viscosity: {visc_value}, error: {e}")
    
    # Convert to DataFrame
    if all_viscosity_data:
        result_df = pd.DataFrame(all_viscosity_data)
        print(f"Extracted {len(result_df)} viscosity data points")
        return result_df
    else:
        print("No viscosity data extracted")
        return pd.DataFrame()

def save_training_data(df, output_path=None):
    """Save the formatted data to a CSV file."""
    if df.empty:
        print("No data to save.")
        return None
    
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"Master_Viscosity_Data_{timestamp}.csv"
    
    df.to_csv(output_path, index=False)
    print(f"Data saved to {output_path}")
    return output_path

# Main execution
if __name__ == "__main__":
    # Replace with the actual path to your Excel file
    excel_path = "Media Database V3 2023 copy.xlsx"
    
    if not os.path.exists(excel_path):
        print(f"Error: File not found at {excel_path}")
    else:
        # Extract and process the data
        formatted_data = extract_media_database(excel_path)
        
        # Save the formatted data
        if not formatted_data.empty:
            save_path = save_training_data(formatted_data)
            
            # Print a sample of the data
            print("\nSample of formatted data:")
            print(formatted_data.head())
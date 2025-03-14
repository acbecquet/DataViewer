import pandas as pd
import numpy as np
import os
from datetime import datetime

def extract_media_database(excel_path):
    """
    Extract data from the complex multi-block Excel structure and format
    it for viscosity analysis with Total Potency included.
    
    This improved version ensures that each viscosity measurement at different
    temperatures has the associated terpene and potency values.
    
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
        
        # Extract data from each sheet
        potency_data = extract_potency_blocks(potency_df)
        terpene_data = extract_terpene_blocks(terpene_df)
        viscosity_data = extract_viscosity_blocks(viscosity_df)
        
        # Print summary of extracted data
        print("\nExtraction Summary:")
        print(f"Potency data: {len(potency_data)} rows, {len(potency_data.columns)} columns")
        print(f"Terpene data: {len(terpene_data)} rows, {len(terpene_data.columns)} columns")
        print(f"Viscosity data: {len(viscosity_data)} rows, {len(viscosity_data.columns)} columns")
        
        # Verify we have the key columns in each dataset
        if not potency_data.empty:
            print("\nPotency columns:")
            for col in sorted(potency_data.columns):
                if 'potency' in col.lower() or 'thc' in col.lower():
                    count = potency_data[col].count()
                    print(f"  - {col}: {count} non-null values")
        
        if not terpene_data.empty:
            print("\nTerpene columns:")
            for col in sorted(terpene_data.columns):
                if 'terpene' in col.lower():
                    count = terpene_data[col].count()
                    print(f"  - {col}: {count} non-null values")
        
        if not viscosity_data.empty:
            print("\nViscosity data summary:")
            count = viscosity_data['viscosity'].count()
            print(f"  - Viscosity: {count} non-null values")
            if 'temperature' in viscosity_data.columns:
                temp_count = viscosity_data['temperature'].count()
                print(f"  - Temperature: {temp_count} non-null values")
                print(f"  - Temperature range: {viscosity_data['temperature'].min()} to {viscosity_data['temperature'].max()}C")
        
        # Merge the data
        merged_data = merge_data(potency_data, terpene_data, viscosity_data)
        
        # Format for training - this now properly associates properties with each viscosity measurement
        formatted_data = format_for_training(merged_data)
        
        # Final verification - make sure each row has a complete set of data
        if not formatted_data.empty:
            print("\nFinal data verification:")
            print(f"Total rows: {len(formatted_data)}")
            print("Complete rows (with temp, viscosity, potency, and terpene):")
            complete_rows = formatted_data.dropna(subset=['temperature', 'viscosity', 'total_potency', 'terpene_pct'])
            print(f"  - {len(complete_rows)} rows ({len(complete_rows)/len(formatted_data)*100:.1f}%)")
            
            # Group by media/terpene to verify structure
            groups = formatted_data.groupby(['media', 'media_brand', 'terpene']).size()
            print(f"Number of unique media/brand/terpene combinations: {len(groups)}")
            
            # Check for repeating property values within temperature groups
            print("Verifying property value consistency within temperature groups...")
            for name, group in formatted_data.groupby(['media', 'media_brand', 'terpene']):
                if 'total_potency' in group.columns and 'terpene_pct' in group.columns:
                    potency_std = group['total_potency'].std()
                    terpene_std = group['terpene_pct'].std()
                    if pd.notna(potency_std) and pd.notna(terpene_std):
                        if potency_std > 0.001 or terpene_std > 0.001:
                            print(f"  - WARNING: Inconsistent property values for {name}")
        
        return formatted_data
        
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def extract_potency_blocks(df):
    """
    Extract potency data from the potency sheet with special pattern matching for the
    structure of the Excel file. Designed to find Total Potency and d9-THC columns.
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
    
    # Scan the entire dataframe for Total Potency and d9-THC column headers
    total_potency_locations = []
    d9_thc_locations = []
    
    # First, try to find these column headers throughout the dataframe
    for i in range(len(df)):
        for j in range(df.shape[1]):
            cell_value = df.iloc[i, j]
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip().lower()
                if 'total potency' in cell_str:
                    total_potency_locations.append((i, j))
                    print(f"DEBUG: Found 'Total Potency' at row {i}, col {j}: '{cell_value}'")
                if 'd9-thc' in cell_str or ('d9' in cell_str and 'thc' in cell_str):
                    d9_thc_locations.append((i, j))
                    print(f"DEBUG: Found 'd9-THC' at row {i}, col {j}: '{cell_value}'")
    
    # Process each block
    for block_idx, (media_row, media_type) in enumerate(media_type_rows):
        # Determine end of block
        next_block = len(df)
        if block_idx < len(media_type_rows) - 1:
            next_block = media_type_rows[block_idx + 1][0]
        
        # Find the row with Brand/Strain header
        header_row = None
        for i in range(media_row, min(media_row + 7, next_block)):
            row_str = " ".join([str(val) for val in df.iloc[i].values if str(val) != 'nan']).lower()
            if "brand" in row_str and "strain" in row_str:
                header_row = i
                break
        
        if header_row is None:
            print(f"Warning: Could not find Brand/Strain header row for block {block_idx+1} ({media_type})")
            continue
            
        # Get the main headers for this block
        headers = []
        for col in range(df.shape[1]):
            if col < len(df.iloc[header_row]):
                val = df.iloc[header_row, col]
                if pd.notna(val) and str(val).strip():
                    headers.append(str(val).strip())
                else:
                    headers.append(f"Column_{col}")
            else:
                headers.append(f"Column_{col}")
                
        # Build a map of cannabinoid column positions
        # Look for cannabinoid headers in the rows just above the data rows
        cannabinoid_columns = {}
        
        # Check the rows around the header row for the complete column structure
        # Also check a few rows above the header row for column headers
        for search_row in range(max(0, header_row - 5), header_row + 2):
            if search_row >= len(df):
                continue
                
            for col in range(df.shape[1]):
                if col >= len(df.iloc[search_row]):
                    continue
                    
                cell_value = df.iloc[search_row, col]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip().lower()
                    
                    # Look for key cannabinoid indicators
                    if any(marker in cell_str for marker in ['thc', 'cbd', 'cbg', 'cbn', 'total potency']):
                        cannabinoid_columns[col] = str(cell_value).strip()
                        print(f"DEBUG: Found cannabinoid column at row {search_row}, col {col}: '{cell_value}'")
        
        print(f"Block {block_idx+1} ({media_type}) headers: {headers}")
        print(f"  Identified cannabinoid columns: {cannabinoid_columns}")
        
        # Find the Total Potency and d9-THC column positions that are valid for this block
        # We'll identify column positions by matching the detected cannabinoid columns with
        # the positions we found during the full dataframe scan
        total_potency_col = None
        d9_thc_col = None
        
        # Find the data section for this block
        data_start_row = header_row + 1
        
        # Use our pre-identified locations to help find the right columns for this block
        for tp_row, tp_col in total_potency_locations:
            # Check if this location is relevant for the current block
            if tp_row >= media_row and tp_row < next_block:
                total_potency_col = tp_col
                print(f"DEBUG: Using 'Total Potency' column at col {tp_col} for block {block_idx+1}")
                break
        
        for d9_row, d9_col in d9_thc_locations:
            # Check if this location is relevant for the current block
            if d9_row >= media_row and d9_row < next_block:
                d9_thc_col = d9_col
                print(f"DEBUG: Using 'd9-THC' column at col {d9_col} for block {block_idx+1}")
                break
        
        # If we didn't find them from the scan, check our cannabinoid_columns
        if total_potency_col is None:
            for col, name in cannabinoid_columns.items():
                if 'total potency' in name.lower():
                    total_potency_col = col
                    print(f"DEBUG: Using 'Total Potency' column at col {col} from cannabinoid columns")
                    break
                    
        if d9_thc_col is None:
            for col, name in cannabinoid_columns.items():
                if 'd9-thc' in name.lower() or ('d9' in name.lower() and 'thc' in name.lower()):
                    d9_thc_col = col
                    print(f"DEBUG: Using 'd9-THC' column at col {col} from cannabinoid columns")
                    break
        
        # Process the data rows
        for row_idx in range(data_start_row, next_block):
            row = df.iloc[row_idx]
            
            # Skip rows with missing brand or strain
            if pd.isna(row[0]) or pd.isna(row[1]):
                continue
            
            # Get brand and strain
            brand = str(row[0]).strip()
            strain = str(row[1]).strip()
            
            # Skip empty rows or header rows
            if not brand or not strain or "Brand" in brand:
                continue
            
            # Create data dictionary
            potency_data = {
                'media_brand': brand,
                'strain': strain,
                'media_type': media_type
            }
            
            # Process the main data columns
            for col in range(2, min(len(headers), len(row))):
                if pd.notna(row[col]):
                    col_name = headers[col]
                    try:
                        # Convert to float, removing any % signs
                        value = float(str(row[col]).replace('%', '').strip())
                        potency_data[col_name] = value
                    except (ValueError, TypeError):
                        # Skip values that can't be converted to float
                        pass
            
            # Add any identified cannabinoids
            for col, col_name in cannabinoid_columns.items():
                if col < len(row) and pd.notna(row[col]):
                    try:
                        value = float(str(row[col]).replace('%', '').strip())
                        potency_data[col_name] = value
                    except (ValueError, TypeError):
                        pass
            
            # Direct check for Total Potency and d9-THC using identified columns
            if total_potency_col is not None and total_potency_col < len(row) and pd.notna(row[total_potency_col]):
                try:
                    value = float(str(row[total_potency_col]).replace('%', '').strip())
                    potency_data['Total Potency'] = value
                    print(f"DEBUG: Extracted Total Potency value: {value} for {brand} - {strain}")
                except (ValueError, TypeError):
                    pass
                    
            if d9_thc_col is not None and d9_thc_col < len(row) and pd.notna(row[d9_thc_col]):
                try:
                    value = float(str(row[d9_thc_col]).replace('%', '').strip())
                    potency_data['d9-THC'] = value
                    print(f"DEBUG: Extracted d9-THC value: {value} for {brand} - {strain}")
                except (ValueError, TypeError):
                    pass
            
            # Check if we found the critical columns
            found_total_potency = 'Total Potency' in potency_data
            found_d9_thc = 'd9-THC' in potency_data
            
            # Try another approach - scan the entire row for values
            if not found_total_potency or not found_d9_thc:
                # Manual scan of the full row for relevant values
                for col in range(len(row)):
                    if col >= len(df.columns):
                        continue
                        
                    cell_val = row[col]
                    if pd.notna(cell_val):
                        try:
                            float_val = float(str(cell_val).replace('%', '').strip())
                            
                            # Check if we're in a column that matches a cannabinoid
                            # by comparing with all rows above up to the header
                            for check_row in range(header_row, -1, -1):
                                if check_row < 0 or col >= len(df.iloc[check_row]):
                                    continue
                                    
                                check_cell = df.iloc[check_row, col]
                                if pd.notna(check_cell):
                                    check_str = str(check_cell).strip().lower()
                                    
                                    if not found_total_potency and 'total potency' in check_str:
                                        potency_data['Total Potency'] = float_val
                                        found_total_potency = True
                                        print(f"DEBUG: Found Total Potency value {float_val} at column {col} via cell scan")
                                        
                                    if not found_d9_thc and ('d9-thc' in check_str or ('d9' in check_str and 'thc' in check_str)):
                                        potency_data['d9-THC'] = float_val
                                        found_d9_thc = True
                                        print(f"DEBUG: Found d9-THC value {float_val} at column {col} via cell scan")
                                        
                                    # Exit early if we found both
                                    if found_total_potency and found_d9_thc:
                                        break
                        except (ValueError, TypeError):
                            pass
                            
                    # Exit early if we found both
                    if found_total_potency and found_d9_thc:
                        break
            
            # Final check if we found the critical columns
            if not found_total_potency:
                print(f"  Warning: No 'Total Potency' found for {brand} - {strain}")
            if not found_d9_thc:
                print(f"  Warning: No 'd9-THC' found for {brand} - {strain}")
                
            all_potency_data.append(potency_data)
    
    # Convert to DataFrame
    if all_potency_data:
        result_df = pd.DataFrame(all_potency_data)
        print(f"Extracted {len(result_df)} potency data points")
        print(f"Columns in potency data: {result_df.columns.tolist()}")
        
        # Check if we successfully extracted any Total Potency or d9-THC data
        if 'Total Potency' in result_df.columns:
            non_null_count = result_df['Total Potency'].count()
            print(f"Extracted {non_null_count} Total Potency values")
            
        if 'd9-THC' in result_df.columns:
            non_null_count = result_df['d9-THC'].count()
            print(f"Extracted {non_null_count} d9-THC values")
            
        return result_df
    else:
        print("No potency data extracted")
        return pd.DataFrame()

def extract_terpene_blocks(df):
    """
    Extract terpene data from the terpene sheet, with special focus on Total Terpene column.
    
    Improved to better identify and extract the Total Terpene values.
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
        
        # Find header with "Name of Oil" or "Brand" and "Strain"
        header_row = None
        
        # First check for 'Name of Oil' and 'Strain' headers
        for i in range(media_row + 1, min(media_row + 7, next_block)):
            row_values = [str(val).lower() for val in df.iloc[i] if pd.notna(val)]
            row_str = " ".join(row_values)
            
            # Look for various header patterns
            if ('name of oil' in row_str and 'strain' in row_str) or 'weight %' in row_str:
                # Look ahead to see if the next row has actual column headers
                if i+1 < next_block:
                    next_row_values = [str(val).lower() for val in df.iloc[i+1] if pd.notna(val)]
                    if any(terpene in " ".join(next_row_values) for terpene in 
                          ['alpha', 'pinene', 'terpene', 'limonene']):
                        header_row = i+1
                    else:
                        header_row = i
                else:
                    header_row = i
                break
            
            # Also check for 'Brand' and 'Strain' headers
            if 'brand' in row_str and 'strain' in row_str:
                if i+1 < next_block:
                    next_row_values = [str(val).lower() for val in df.iloc[i+1] if pd.notna(val)]
                    if any(terpene in " ".join(next_row_values) for terpene in 
                          ['alpha', 'pinene', 'terpene', 'limonene']):
                        header_row = i+1
                    else:
                        header_row = i
                else:
                    header_row = i
                break
        
        if header_row is None:
            print(f"Warning: Could not find header row for block {block_idx+1} ({media_type})")
            continue
        
        # Get column names, using the previous few rows to find "Name of Oil"/"Brand" and "Strain" if needed
        name_col = -1
        strain_col = -1
        
        # First look in the rows just before the header row to find the name/strain columns
        for i in range(max(0, header_row-3), header_row+1):
            for col, val in enumerate(df.iloc[i]):
                if pd.notna(val):
                    val_str = str(val).lower()
                    if 'name of oil' in val_str or 'brand' in val_str:
                        name_col = col
                    elif 'strain' in val_str:
                        strain_col = col
        
        if name_col == -1 or strain_col == -1:
            print(f"  Warning: Could not find Name of Oil/Brand and Strain columns in block {block_idx+1}")
            # Default to first two columns
            name_col = 0
            strain_col = 1
        
        # Get all column headers from the header row and surrounding rows
        headers = []
        for col in range(df.shape[1]):
            val = df.iloc[header_row, col]
            if pd.notna(val) and str(val).strip():
                headers.append(str(val).strip())
            else:
                # Check a few rows above for column headers
                for i in range(max(0, header_row-3), header_row):
                    if i < len(df):
                        above_val = df.iloc[i, col] if col < len(df.iloc[i]) else None
                        if pd.notna(above_val) and str(above_val).strip():
                            headers.append(str(above_val).strip())
                            break
                else:
                    headers.append(f"Column_{col}")
        
        # Debug
        print(f"Block {block_idx+1} ({media_type}) headers: {headers}")
        
        # Find the 'Total Terpene' column
        total_terpene_col = -1
        for col, header in enumerate(headers):
            if isinstance(header, str) and 'total terpene' in header.lower():
                total_terpene_col = col
                print(f"  Found 'Total Terpene' column at index {col}")
                break
        
        # If not found in headers, look for it in surrounding rows
        if total_terpene_col == -1:
            for col in range(df.shape[1]):
                # Check a few rows above and the header row itself
                for i in range(max(0, header_row-3), header_row+1):
                    if i < len(df):
                        val = df.iloc[i, col] if col < len(df.iloc[i]) else None
                        if pd.notna(val) and 'total terpene' in str(val).lower():
                            total_terpene_col = col
                            print(f"  Found 'Total Terpene' column at index {col} in row {i}")
                            break
                if total_terpene_col != -1:
                    break
        
        if total_terpene_col == -1:
            print(f"  WARNING: No 'Total Terpene' column found in block {block_idx+1}")
        
        # Determine row with data start
        data_start_row = header_row + 1
        
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
    Merge the potency, terpene, and viscosity data to ensure that terpene and potency
    data is properly associated with each viscosity measurement.
    
    This modified version ensures that rows with viscosity measurements include
    the corresponding terpene and potency values from their media/strain combination.
    
    Args:
        potency_df: DataFrame with potency data
        terpene_df: DataFrame with terpene data
        viscosity_df: DataFrame with viscosity data
        
    Returns:
        DataFrame with properly merged data where each viscosity measurement
        has associated terpene and potency values
    """
    print("Starting enhanced merge process...")
    
    # Handle empty DataFrames
    if potency_df.empty and terpene_df.empty and viscosity_df.empty:
        print("All data frames are empty, nothing to merge")
        return pd.DataFrame()
    
    # Make copies of the dataframes to avoid modifying the originals
    potency_copy = potency_df.copy() if not potency_df.empty else None
    terpene_copy = terpene_df.copy() if not terpene_df.empty else None
    viscosity_copy = viscosity_df.copy() if not viscosity_df.empty else None
    
    # NORMALIZE THE KEY COLUMNS IN EACH DATAFRAME
    for df in [potency_copy, terpene_copy, viscosity_copy]:
        if df is not None and not df.empty:
            # Clean up media_brand
            if 'media_brand' in df.columns:
                df['media_brand'] = df['media_brand'].str.strip().str.lower()
            
            # Clean up strain - extract the main part before any '/' character
            if 'strain' in df.columns:
                # Option 1: Keep only the part before the first '/'
                df['strain_key'] = df['strain'].str.split('/').str[0].str.strip().str.lower()
                
                # Option 2: Remove the part after and including "/"
                # df['strain_key'] = df['strain'].str.replace(r'/.*$', '', regex=True).str.strip().str.lower()
            
            # Clean up media_type
            if 'media_type' in df.columns:
                df['media_type'] = df['media_type'].str.strip().str.lower()

    # Check if we have viscosity data - this is key for our merge strategy
    if viscosity_copy is None or viscosity_copy.empty:
        print("No viscosity data available for merging")
        
        # If no viscosity data, just merge potency and terpene data as before
        if potency_copy is not None and terpene_copy is not None:
            merged = pd.merge(
                potency_copy, 
                terpene_copy,
                on=['media_brand', 'strain', 'media_type'],
                how='outer',
                suffixes=('_potency', '_terpene')
            )
        elif potency_copy is not None:
            merged = potency_copy
        elif terpene_copy is not None:
            merged = terpene_copy
        else:
            merged = pd.DataFrame()
            
        return merged
    
    # If we have viscosity data, use a different approach focusing on preserving viscosity values
    print(f"Viscosity data shape: {viscosity_copy.shape}")
    
    # First, merge potency and terpene data to create a "properties" dataframe
    if potency_copy is not None and terpene_copy is not None:
        properties_df = pd.merge(
            potency_copy, 
            terpene_copy,
            on=['media_brand', 'strain', 'media_type'],
            how='outer',
            suffixes=('_potency', '_terpene')
        )
    elif potency_copy is not None:
        properties_df = potency_copy
    elif terpene_copy is not None:
        properties_df = terpene_copy
    else:
        properties_df = pd.DataFrame()
    
    if properties_df.empty:
        print("No potency or terpene data available")
        return viscosity_copy
    
    print(f"Combined properties data shape: {properties_df.shape}")
    
    # Now, for each viscosity measurement, associate the corresponding properties
    # using media_brand, strain, and media_type as keys
    merged_data = pd.merge(
        viscosity_copy,
        properties_df,
        on=['media_brand', 'strain', 'media_type'],
        how='left'
    )
    
    # Check merge results
    print(f"Final merged data shape: {merged_data.shape}")
    print(f"Number of rows with viscosity values: {merged_data['viscosity'].count()}")
    
    # Check if key property columns made it through
    for column in ['Total Potency', 'total_potency_value', 'Total Terpene', 'terpene_pct_value']:
        if column in merged_data.columns:
            non_null_count = merged_data[column].count()
            print(f"Column '{column}' has {non_null_count} non-null values")
    
    return merged_data

def format_for_training(df):
    """
    Format the merged data for viscosity training, ensuring Total Potency and
    Total Terpene values are properly associated with each viscosity measurement.
    
    This modified version focuses on creating a clean dataset where each row
    represents a viscosity measurement with its associated properties.
    
    Args:
        df: The merged DataFrame from merge_data()
        
    Returns:
        DataFrame properly formatted for training models
    """
    if df.empty:
        return pd.DataFrame()
    
    # Start with a copy of the input DataFrame to avoid modifying the original
    working_df = df.copy()
    
    # Print all available columns to help identify what's available
    print("\nAvailable columns in merged data:")
    for col in sorted(working_df.columns):
        print(f"  - {col}")
    
    # Create a new DataFrame with columns matching the expected training format
    training_df = pd.DataFrame()
    
    # Map source columns to training format columns
    if 'media_type' in working_df.columns:
        training_df['media'] = working_df['media_type']
    else:
        training_df['media'] = 'Unknown'
        
    if 'media_brand' in working_df.columns:
        training_df['media_brand'] = working_df['media_brand']
    
    if 'strain' in working_df.columns:
        training_df['terpene'] = working_df['strain']
    
    # Add temperature and viscosity if available - these are the key measurements
    if 'temperature' in working_df.columns:
        training_df['temperature'] = pd.to_numeric(working_df['temperature'], errors='coerce')
    
    if 'viscosity' in working_df.columns:
        training_df['viscosity'] = pd.to_numeric(working_df['viscosity'], errors='coerce')
    
    # Process total potency - try multiple possible column names
    potency_found = False
    
    # First try our pre-processed column
    if 'total_potency_value' in working_df.columns:
        print(f"Using pre-processed 'total_potency_value' column for total_potency")
        training_df['total_potency'] = working_df['total_potency_value']
        potency_found = True
    elif 'Total Potency' in working_df.columns:
        print(f"Using column 'Total Potency' for total_potency")
        training_df['total_potency'] = pd.to_numeric(working_df['Total Potency'], errors='coerce')
        potency_found = True
    else:
        # Try various case/name variations
        potency_variations = [
            'TOTAL POTENCY', 'Total potency', 'total potency', 'TotalPotency', 
            'Potency, Total', 'Total_Potency', 'Potency (Total)', 'total_potency'
        ]
        
        for var in potency_variations:
            if var in working_df.columns:
                print(f"Using column '{var}' for total_potency")
                training_df['total_potency'] = pd.to_numeric(
                    working_df[var].astype(str).str.replace('%', '').str.strip(), 
                    errors='coerce'
                )
                potency_found = True
                break
                
        # If not found with exact names, try case-insensitive search
        if not potency_found:
            for col in working_df.columns:
                if isinstance(col, str) and 'total' in col.lower() and 'potency' in col.lower():
                    print(f"Using column '{col}' for total_potency using partial match")
                    training_df['total_potency'] = pd.to_numeric(
                        working_df[col].astype(str).str.replace('%', '').str.strip(), 
                        errors='coerce'
                    )
                    potency_found = True
                    break
    
    if not potency_found:
        print("WARNING: Could not find Total Potency column")
    
    # Process total terpene - similar approach as total potency
    terpene_found = False
    
    # First try our pre-processed column
    if 'terpene_pct_value' in working_df.columns:
        print(f"Using pre-processed 'terpene_pct_value' column for terpene_pct")
        training_df['terpene_pct'] = working_df['terpene_pct_value']
        terpene_found = True
    elif 'Total Terpene' in working_df.columns:
        print(f"Using column 'Total Terpene' for terpene_pct")
        training_df['terpene_pct'] = pd.to_numeric(working_df['Total Terpene'], errors='coerce')
        terpene_found = True
    else:
        # Try various case/name variations
        terpene_variations = [
            'TOTAL TERPENE', 'Total terpene', 'total terpene', 'TotalTerpene', 
            'Terpene, Total', 'Total_Terpene', 'Terpene (Total)', 'total_terpene'
        ]
        
        for var in terpene_variations:
            if var in working_df.columns:
                print(f"Using column '{var}' for terpene_pct")
                training_df['terpene_pct'] = pd.to_numeric(
                    working_df[var].astype(str).str.replace('%', '').str.strip(), 
                    errors='coerce'
                )
                terpene_found = True
                break
                
        # If not found with exact names, try case-insensitive search
        if not terpene_found:
            for col in working_df.columns:
                if isinstance(col, str) and 'total' in col.lower() and 'terpene' in col.lower():
                    print(f"Using column '{col}' for terpene_pct using partial match")
                    training_df['terpene_pct'] = pd.to_numeric(
                        working_df[col].astype(str).str.replace('%', '').str.strip(), 
                        errors='coerce'
                    )
                    terpene_found = True
                    break
    
    if not terpene_found:
        print("WARNING: Could not find Total Terpene column")
    
    # Add key cannabinoids if available (use flexible matching)
    cannabinoid_mappings = {
        'd9-THC': ['d9-THC', 'D9-THC', 'd9 THC', 'D9 THC', 'd9THC', 'D9THC', 'delta9-THC'],
        'THCA': ['THCA', 'THCa', 'thca', 'THC-A', 'THC A'],
        'CBD': ['CBD', 'cbd'],
        'CBDA': ['CBDA', 'CBDa', 'cbda', 'CBD-A', 'CBD A'],
        'd8-THC': ['d8-THC', 'D8-THC', 'd8 THC', 'D8 THC', 'd8THC', 'D8THC', 'delta8-THC'],
        'CBG': ['CBG', 'cbg']
    }
    
    target_mappings = {
        'd9-THC': 'd9_thc',
        'THCA': 'thca',
        'CBD': 'cbd',
        'CBDA': 'cbda',
        'd8-THC': 'd8_thc',
        'CBG': 'cbg'
    }
    
    # Special case for d9-THC with pre-processed column
    if 'd9_thc_value' in working_df.columns:
        print(f"Using pre-processed 'd9_thc_value' column for d9_thc")
        training_df['d9_thc'] = working_df['d9_thc_value']
        
        # Remove d9-THC from the mappings since we've handled it
        if 'd9-THC' in cannabinoid_mappings:
            del cannabinoid_mappings['d9-THC']
            del target_mappings['d9-THC']
    
    # Process remaining cannabinoids
    for cannabinoid, variations in cannabinoid_mappings.items():
        target_col = target_mappings[cannabinoid]
        
        # Try exact matches first
        found_col = None
        for var in variations:
            if var in working_df.columns:
                found_col = var
                break
                
        # If not found, try case-insensitive matching
        if found_col is None:
            for col in working_df.columns:
                if isinstance(col, str) and col.lower() in [v.lower() for v in variations]:
                    found_col = col
                    break
                    
        # If still not found, try partial matching
        if found_col is None:
            for col in working_df.columns:
                if isinstance(col, str):
                    col_lower = col.lower()
                    # For d8-THC, look for both 'd8' and 'thc' in the column name
                    if cannabinoid == 'd8-THC' and 'd8' in col_lower and 'thc' in col_lower:
                        found_col = col
                        break
                    # For other cannabinoids, just check if the lowercase name is contained
                    elif cannabinoid.lower() in col_lower:
                        found_col = col
                        break
        
        # Use the column if found
        if found_col:
            print(f"Using column '{found_col}' for {target_col}")
            training_df[target_col] = pd.to_numeric(working_df[found_col], errors='coerce')
    
    # Add timestamp and source
    from datetime import datetime
    training_df['timestamp'] = datetime.now().isoformat()
    training_df['original_filename'] = 'Media Database V3 2023'
    
    # CRITICAL FIX: Filter out rows without a temperature or viscosity value
    # These rows are not useful for training viscosity models
    final_data = training_df.dropna(subset=['temperature', 'viscosity'])
    
    print(f"\nFinal training data shape: {final_data.shape}")
    print(f"Final training data columns: {final_data.columns.tolist()}")
    
    # Print a sample of the data to verify the fix worked
    if len(final_data) > 0:
        print("\nSample of final data (first 3 rows):")
        print(final_data[['media', 'media_brand', 'terpene', 'temperature', 'viscosity', 'total_potency', 'terpene_pct']].head(3))
        
        # Add detailed statistics 
        print("\nStatistics for key columns:")
        for col in ['temperature', 'viscosity', 'total_potency', 'terpene_pct']:
            if col in final_data.columns:
                non_null = final_data[col].count()
                total = len(final_data)
                print(f"{col}: {non_null}/{total} non-null values ({non_null/total*100:.1f}%)")
                if non_null > 0:
                    print(f"  - Min: {final_data[col].min()}, Max: {final_data[col].max()}, Mean: {final_data[col].mean():.2f}")
    
    return final_data

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
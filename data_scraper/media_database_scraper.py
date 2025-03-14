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
    Merge the potency, terpene, and viscosity data.
    Enhanced to better preserve percentage values during the merge operation.
    """
    # Handle empty DataFrames
    if potency_df.empty and terpene_df.empty and viscosity_df.empty:
        print("All data frames are empty, nothing to merge")
        return pd.DataFrame()
    
    # Make copies of the dataframes to avoid modifying the originals
    potency_copy = potency_df.copy() if not potency_df.empty else None
    terpene_copy = terpene_df.copy() if not terpene_df.empty else None
    
    # Pre-merge processing: Make sure key columns have the right data types
    # And create explicit copies of the total potency and total terpene columns
    if potency_copy is not None:
        if 'Total Potency' in potency_copy.columns:
            # Create a specific column for total potency that won't be lost in the merge
            potency_copy['total_potency_value'] = pd.to_numeric(potency_copy['Total Potency'], errors='coerce')
            print(f"Created 'total_potency_value' column with {potency_copy['total_potency_value'].count()} non-null values")
            print(f"Sample potency values: {potency_copy['total_potency_value'].head(3).tolist()}")
        
        if 'd9-THC' in potency_copy.columns:
            # Create a specific column for d9-THC that won't be lost in the merge
            potency_copy['d9_thc_value'] = pd.to_numeric(potency_copy['d9-THC'], errors='coerce')
            print(f"Created 'd9_thc_value' column with {potency_copy['d9_thc_value'].count()} non-null values")
    
    if terpene_copy is not None:
        if 'Total Terpene' in terpene_copy.columns:
            # Create a specific column for total terpene that won't be lost in the merge
            terpene_copy['terpene_pct_value'] = pd.to_numeric(terpene_copy['Total Terpene'], errors='coerce')
            print(f"Created 'terpene_pct_value' column with {terpene_copy['terpene_pct_value'].count()} non-null values")
            print(f"Sample terpene values: {terpene_copy['terpene_pct_value'].head(3).tolist()}")
    
    # Merge terpene data into potency data first (if both exist)
    if potency_copy is not None and terpene_copy is not None:
        merged = pd.merge(
            potency_copy, 
            terpene_copy,
            on=['media_brand', 'strain', 'media_type'],
            how='outer',
            suffixes=('_potency', '_terpene')
        )
        
        # Check if our special columns made it through the merge
        for col in ['total_potency_value', 'd9_thc_value', 'terpene_pct_value']:
            if col in merged.columns:
                print(f"After merge 1: '{col}' has {merged[col].count()} non-null values")
            else:
                print(f"WARNING: '{col}' did not survive the merge!")
                
    elif potency_copy is not None:
        merged = potency_copy
    elif terpene_copy is not None:
        merged = terpene_copy
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
        
        # Check if our special columns made it through the final merge
        for col in ['total_potency_value', 'd9_thc_value', 'terpene_pct_value']:
            if col in final_merged.columns:
                print(f"After merge 2: '{col}' has {final_merged[col].count()} non-null values")
            else:
                print(f"WARNING: '{col}' did not survive the final merge!")
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
    
    Enhanced to handle complex Excel data structures and different column name variations.
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
    
    # First check for our special pre-processed columns created in merge_data
    if 'total_potency_value' in df.columns:
        print(f"Using pre-processed 'total_potency_value' column for total_potency")
        training_df['total_potency'] = df['total_potency_value']
        print(f"Sample from total_potency_value: {df['total_potency_value'].head(3).tolist()}")
        
        # Add debugging info
        non_null_count = df['total_potency_value'].count()
        total_count = len(df)
        print(f"total_potency_value has {non_null_count}/{total_count} non-null values")
    elif 'Total Potency' in df.columns:
        print(f"Using column 'Total Potency' for total_potency")
        training_df['total_potency'] = pd.to_numeric(df['Total Potency'], errors='coerce')
    else:
        # Try various case variations and naming patterns
        potency_variations = [
            'TOTAL POTENCY', 'Total potency', 'total potency', 'TotalPotency', 
            'Potency, Total', 'Total_Potency', 'Potency (Total)', 'total_potency'
        ]
        
        # Try exact matches first
        total_potency_col = None
        for var in potency_variations:
            if var in df.columns:
                total_potency_col = var
                break
                
        # If not found, try case-insensitive matching
        if total_potency_col is None:
            for col in df.columns:
                if isinstance(col, str) and col.lower() in [v.lower() for v in potency_variations]:
                    total_potency_col = col
                    break
                    
        # If still not found, try partial matching
        if total_potency_col is None:
            for col in df.columns:
                if isinstance(col, str) and 'total' in col.lower() and 'potency' in col.lower():
                    total_potency_col = col
                    break
        
        # Use the column if found
        if total_potency_col:
            print(f"Using column '{total_potency_col}' for total_potency")
            # First convert to string, then remove % symbols, then convert to numeric
            training_df['total_potency'] = pd.to_numeric(
                df[total_potency_col].astype(str).str.replace('%', '').str.strip(), 
                errors='coerce'
            )
        else:
            print("WARNING: Could not find Total Potency column")
            # Suggest potential columns that might be related
            potency_candidates = []
            for col in df.columns:
                if isinstance(col, str) and ('potency' in col.lower() or 'thc' in col.lower()):
                    potency_candidates.append(col)
            
            if potency_candidates:
                print(f"Potential potency columns: {potency_candidates}")
    
    # Check for special pre-processed terpene column first
    if 'terpene_pct_value' in df.columns:
        print(f"Using pre-processed 'terpene_pct_value' column for terpene_pct")
        training_df['terpene_pct'] = df['terpene_pct_value']
        print(f"Sample from terpene_pct_value: {df['terpene_pct_value'].head(3).tolist()}")
        
        # Add debugging info
        non_null_count = df['terpene_pct_value'].count()
        total_count = len(df)
        print(f"terpene_pct_value has {non_null_count}/{total_count} non-null values")
    elif 'Total Terpene' in df.columns:
        print(f"Using column 'Total Terpene' for terpene_pct")
        training_df['terpene_pct'] = pd.to_numeric(df['Total Terpene'], errors='coerce')
    else:
        # Try various case variations and naming patterns
        terpene_variations = [
            'TOTAL TERPENE', 'Total terpene', 'total terpene', 'TotalTerpene', 
            'Terpene, Total', 'Total_Terpene', 'Terpene (Total)', 'total_terpene'
        ]
        
        # Try exact matches first
        total_terpene_col = None
        for var in terpene_variations:
            if var in df.columns:
                total_terpene_col = var
                break
                
        # If not found, try case-insensitive matching
        if total_terpene_col is None:
            for col in df.columns:
                if isinstance(col, str) and col.lower() in [v.lower() for v in terpene_variations]:
                    total_terpene_col = col
                    break
                    
        # If still not found, try partial matching
        if total_terpene_col is None:
            for col in df.columns:
                if isinstance(col, str) and 'total' in col.lower() and 'terpene' in col.lower():
                    total_terpene_col = col
                    break
        
        # Use the column if found
        if total_terpene_col:
            print(f"Using column '{total_terpene_col}' for terpene_pct")
            training_df['terpene_pct'] = pd.to_numeric(
                df[total_terpene_col].astype(str).str.replace('%', '').str.strip(), 
                errors='coerce'
            )
        else:
            print("WARNING: Could not find Total Terpene column")
            # Suggest potential columns that might be related
            terpene_candidates = []
            for col in df.columns:
                if isinstance(col, str) and 'terpene' in col.lower():
                    terpene_candidates.append(col)
            
            if terpene_candidates:
                print(f"Potential terpene columns: {terpene_candidates}")
    
    # Add temperature and viscosity if available
    if 'temperature' in df.columns:
        training_df['temperature'] = pd.to_numeric(df['temperature'], errors='coerce')
    
    if 'viscosity' in df.columns:
        training_df['viscosity'] = pd.to_numeric(df['viscosity'], errors='coerce')
    
    # Add key cannabinoids with flexible matching
    cannabinoid_mappings = {
        'd9-THC': ['d9-THC', 'D9-THC', 'd9 THC', 'D9 THC', 'd9THC', 'D9THC', 'delta9-THC', 'Delta9-THC', 'delta 9 THC'],
        'THCA': ['THCA', 'THCa', 'thca', 'THC-A', 'THC A'],
        'CBD': ['CBD', 'cbd'],
        'CBDA': ['CBDA', 'CBDa', 'cbda', 'CBD-A', 'CBD A'],
        'd8-THC': ['d8-THC', 'D8-THC', 'd8 THC', 'D8 THC', 'd8THC', 'D8THC', 'delta8-THC', 'Delta8-THC', 'delta 8 THC'],
        'CBG': ['CBG', 'cbg']
    }
    
    # Target column names in our output DataFrame
    target_mappings = {
        'd9-THC': 'd9_thc',
        'THCA': 'thca',
        'CBD': 'cbd',
        'CBDA': 'cbda',
        'd8-THC': 'd8_thc',
        'CBG': 'cbg'
    }
    
    # Special case for d9-THC which has our pre-processed column
    if 'd9_thc_value' in df.columns:
        print(f"Using pre-processed 'd9_thc_value' column for d9_thc")
        training_df['d9_thc'] = df['d9_thc_value']
        print(f"Sample from d9_thc_value: {df['d9_thc_value'].head(3).tolist()}")
        
        # Remove d9-THC from the cannabinoid mapping since we've already handled it
        if 'd9-THC' in cannabinoid_mappings:
            del cannabinoid_mappings['d9-THC']
            del target_mappings['d9-THC']
    
    # For each remaining cannabinoid, try all possible variations of the column name
    for cannabinoid, variations in cannabinoid_mappings.items():
        target_col = target_mappings[cannabinoid]
        
        # Try exact matches first
        found_col = None
        for var in variations:
            if var in df.columns:
                found_col = var
                break
                
        # If not found, try case-insensitive matching
        if found_col is None:
            for col in df.columns:
                if isinstance(col, str) and col.lower() in [v.lower() for v in variations]:
                    found_col = col
                    break
                    
        # If still not found, try partial matching
        if found_col is None:
            for col in df.columns:
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
            training_df[target_col] = pd.to_numeric(df[found_col], errors='coerce')
    
    # Add timestamp and source
    from datetime import datetime
    training_df['timestamp'] = datetime.now().isoformat()
    training_df['original_filename'] = 'Media Database V3 2023'
    
    # Create a copy of all data before filtering
    all_data = training_df.copy()
    
    # Create a filtered dataset for viscosity analysis
    viscosity_data = None
    if 'viscosity' in training_df.columns:
        viscosity_data = training_df.dropna(subset=['viscosity'])
        print(f"Created filtered dataset for viscosity analysis: {len(viscosity_data)} rows")
    
    # IMPORTANT: For this fix, we're returning ALL data, not just rows with viscosity values
    # This ensures we don't lose potency and terpene values
    final_data = all_data
    
    print(f"\nFinal training data shape: {final_data.shape}")
    print(f"Final training data columns: {final_data.columns.tolist()}")
    
    # Print a sample of potency and terpene values to verify
    if len(final_data) > 0:
        print("\nSample of extracted values:")
        if 'total_potency' in final_data.columns:
            print(f"total_potency sample: {final_data['total_potency'].head(3).tolist()}")
        if 'terpene_pct' in final_data.columns:
            print(f"terpene_pct sample: {final_data['terpene_pct'].head(3).tolist()}")
            
        # Add more detailed statistics to help troubleshoot
        if 'total_potency' in final_data.columns:
            non_null = final_data['total_potency'].count()
            total = len(final_data)
            print(f"Total potency non-null values: {non_null}/{total} ({non_null/total*100:.1f}%)")
            
        if 'terpene_pct' in final_data.columns:
            non_null = final_data['terpene_pct'].count()
            total = len(final_data)
            print(f"Terpene percent non-null values: {non_null}/{total} ({non_null/total*100:.1f}%)")
    
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
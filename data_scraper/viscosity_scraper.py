import os
import pandas as pd
import glob
import re
from datetime import datetime

def extract_viscosity_data(file_path, media="Unknown", terpene="Unknown", terpene_pct=0.0, original_filename=None):
    """
    Extract viscosity data from an Excel file, focusing on temperature (column L) and 
    average viscosity (column N) for each spindle.
    Only processes the first 15 rows to avoid duplicate headers further down.
    
    Args:
        file_path: Path to the Excel file
        media: Media type extracted from filename
        terpene: Terpene type extracted from filename
        terpene_pct: Terpene percentage
        original_filename: Original filename for reference
        
    Returns:
        DataFrame with the extracted data in the required format
    """
    try:
        # Read only the first 15 rows of the Excel file to find header
        df_headers = pd.read_excel(file_path, nrows=15, header=None)
        
        # Look for the header row with "Spindle"
        header_row_idx = None
        for i in range(len(df_headers)):
            row = df_headers.iloc[i].astype(str)
            if row.str.contains('Spindle').any():
                header_row_idx = i
                break
        
        if header_row_idx is None:
            print(f"Warning: Could not find header row in {file_path}")
            return pd.DataFrame()
        
        # Read the data with the identified header
        df = pd.read_excel(file_path, header=header_row_idx, nrows=15)
        
        # Drop rows where all values are NaN
        df = df.dropna(how='all')
        
        # Get column mapping - try to identify key columns
        columns = df.columns.tolist()
        
        # Find columns by name first
        spindle_col = columns[0]  # Usually first column
        speed_col = next((col for col in columns if 'Speed' in str(col)), None)
        
        # Try column L for temperature (0-indexed position 11)
        temp_col = next((col for col in columns if 'Temperature' in str(col) and 'Average' in str(col)), None)
        if temp_col is None and len(columns) > 11:
            temp_col = columns[11]  # Column L (0-indexed = 11)
        
        # Try column N for viscosity (0-indexed position 13)
        visc_col = next((col for col in columns if 'Viscosity' in str(col) and 'Average' in str(col)), None)
        if visc_col is None and len(columns) > 13:
            visc_col = columns[13]  # Column N (0-indexed = 13)
        
        # Verify we have all the columns we need
        if spindle_col is None or speed_col is None or temp_col is None or visc_col is None:
            print(f"Warning: Could not identify all required columns in {file_path}")
            print(f"Found columns: spindle={spindle_col}, speed={speed_col}, temp={temp_col}, visc={visc_col}")
            return pd.DataFrame()
        
        # Find rows with spindle data
        spindle_rows = df[df[spindle_col].astype(str).str.contains('CPA-')]
        
        if spindle_rows.empty:
            print(f"Warning: No spindle data found in {file_path}")
            return pd.DataFrame()
        
        # Process each spindle row to extract data
        spindle_data = []
        for _, row in spindle_rows.iterrows():
            spindle = row[spindle_col]
            speed = row[speed_col] if pd.notna(row[speed_col]) else None
            temp = row[temp_col] if pd.notna(row[temp_col]) else None
            visc = row[visc_col] if pd.notna(row[visc_col]) else None
            
            # Handle scientific notation in viscosity
            if visc is not None:
                try:
                    visc = float(visc)
                except (ValueError, TypeError):
                    if isinstance(visc, str):
                        # Remove commas and try again
                        visc = visc.replace(',', '')
                        try:
                            visc = float(visc)
                        except:
                            # Try parsing scientific notation
                            try:
                                import re
                                scientific_match = re.match(r'(\d+\.\d+)[eE]([+-]?\d+)', visc)
                                if scientific_match:
                                    visc = float(visc)
                                else:
                                    visc = None
                            except:
                                visc = None
            
            # Create a data row
            data_row = {
                'media': media,
                'terpene': terpene,
                'terpene_pct': terpene_pct,
                'spindle': spindle,
                'speed': speed,
                'temperature': temp,
                'viscosity': visc,
                'original_filename': original_filename
            }
            
            spindle_data.append(data_row)
        
        return pd.DataFrame(spindle_data)
        
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()
def extract_percentage(filename):
    """
    Extract percentage values from filename - looking for digits (including decimals)
    right before a % sign.
    """
    # Look for digits with optional decimal part followed by % sign
    pattern = r'(\d{1,2}(?:\.\d+)?)%'
    match = re.search(pattern, filename)
    
    if match:
        try:
            percent_value = float(match.group(1))
            # If value < 50, it's terpene %; if > 50, it's THC % (convert to terpene %)
            if percent_value < 50:
                return percent_value
            else:
                return 100 - percent_value
        except ValueError:
            pass
    
    return 0.0  # Default if no percentage found

def batch_process_files(directory, output_file):
    """
    Process all Excel files in a directory and save to a single CSV.
    
    Args:
        directory: Directory containing Excel files
        output_file: Path to save the output CSV
    """
    # Define media types dictionary for identification
    MEDIA_TYPES = {
        'D8': 'D8',
        'D9': 'D9',
        'ROSIN': 'Rosin',
        'RESIN': 'Resin',
        'LIQUID DIAMONDS': 'Liquid Diamonds',
        'DISTILLATE': 'Distillate'
    }
    
    # Common terpene names to help identify in filenames
    COMMON_TERPENES = [
        'SUNSET SHERBERT', 'SHERBERT', 'BUBBA', 'LEMON', 'PINK RUNTZ',
        'RUSH CAKE', 'BLUE DREAM', 'WATERMELON', 'BLURAZZ', 'TERPENES',
        'APPLE JACK', 'PINK BERRY', 'PINKBERRY', 'SKITTLEZ', 
        'GRAPE', 'GUAVA', 'GASBERRY', 'PINEAPPLE', 'SUPER LEMON',
        'TIGERS BLOOD', 'BLUEBERRY', 'TROPICAL', 'HHC', 'VAPEN', 'GRAPE APE'
    ]
    
    # Find all Excel files in the directory
    excel_files = glob.glob(os.path.join(directory, "*.xlsx")) + glob.glob(os.path.join(directory, "*.xls"))
    
    if not excel_files:
        print(f"No Excel files found in {directory}")
        return
    
    all_data = []
    file_count = 0
    error_count = 0
    
    for file_path in excel_files:
        try:
            # Get the filename without extension
            filename = os.path.basename(file_path)
            filename_no_ext = os.path.splitext(filename)[0]
            filename_upper = filename_no_ext.upper()  # Uppercase for matching
            
            # Initialize variables
            media = "Unknown"
            terpene = "Unknown"
            terpene_pct = 0.0
            
            # Extract media type from filename
            for key, value in MEDIA_TYPES.items():
                if key in filename_upper:
                    media = value
                    break
            
            # Extract percentage from filename
            terpene_pct = extract_percentage(filename_no_ext)
            
            # Extract terpene name
            # First check for common terpene names
            for terp in COMMON_TERPENES:
                if terp in filename_upper:
                    terpene = terp.title()
                    break
            
            # If no known terpene found, try some pattern matching
            if terpene == "Unknown":
                # Try pattern: GER-D8-Apple Jack-93%
                match = re.search(r'GER-[^-]+-([^-]+)-', filename_upper)
                if match:
                    terpene = match.group(1).title()
                
                # Try pattern: Alien Bubba D8
                if terpene == "Unknown":
                    for media_key in MEDIA_TYPES.keys():
                        pattern = f'(.*?){media_key}'
                        match = re.search(pattern, filename_upper)
                        if match:
                            potential_terpene = match.group(1).strip('- _').strip()
                            if potential_terpene and potential_terpene not in ["GER", "BU2500", "CREAM"]:
                                terpene = potential_terpene.title()
                                break
            
            # Extract data from this file
            file_data = extract_viscosity_data(
                file_path, 
                media=media, 
                terpene=terpene, 
                terpene_pct=terpene_pct,
                original_filename=filename
            )
            
            if not file_data.empty:
                all_data.append(file_data)
                file_count += 1
                print(f"Successfully processed file {file_count}: {filename}")
            else:
                error_count += 1
                print(f"Warning: No data extracted from {filename}")
                
        except Exception as e:
            error_count += 1
            print(f"Error processing {file_path}: {str(e)}")
    
    if all_data:
        # Combine all data into a single DataFrame
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Save to CSV
        combined_df.to_csv(output_file, index=False)
        print(f"Successfully processed {file_count} files (with {error_count} errors) and saved {len(combined_df)} data points to {output_file}")
    else:
        print("No data was extracted.")

# Example usage
if __name__ == "__main__":
    # Directory containing your Excel files
    data_dir = "data"
    
    # Create output directory if it doesn't exist
    os.makedirs("data", exist_ok=True)
    
    # Output file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Master Viscosity {timestamp}.csv"
    
    batch_process_files(data_dir, output_file)
import os
import pandas as pd
import glob
from datetime import datetime

def extract_viscosity_data(file_path, media="Unknown", terpene="Unknown", terpene_pct=0.0):
    """
    Extract viscosity data from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        media: Media type (to be filled manually or derived from filename)
        terpene: Terpene type (to be filled manually or derived from filename)
        terpene_pct: Terpene percentage (to be filled manually or derived from filename)
        
    Returns:
        DataFrame with the extracted data in the required format
    """
    try:
        # Read the Excel file
        # Note: skiprows and header might need adjustment based on your Excel structure
        df = pd.read_excel(file_path)
        
        # Look for the header row with "Spindle", "Speed", etc.
        header_row_idx = None
        for i, row in df.iterrows():
            if row.astype(str).str.contains('Spindle').any():
                header_row_idx = i
                break
        
        if header_row_idx is None:
            print(f"Warning: Could not find header row in {file_path}")
            return pd.DataFrame()
        
        # Re-read the file with the correct header row
        df = pd.read_excel(file_path, header=header_row_idx)
        
        # Filter to include only CPA-52Z spindle rows
        df = df[df.iloc[:, 0].astype(str).str.contains('CPA-52Z')]
        
        if df.empty:
            print(f"Warning: No CPA-52Z data found in {file_path}")
            return pd.DataFrame()
        
        # Extract the relevant columns - adjust these indices based on your data structure
        # Assuming "Average (cp-C)" is for temperature and "Average Viscosity" is for viscosity
        temp_col = df.columns[10]  # Average temperature column
        visc_col = df.columns[12]  # Average viscosity column
        
        # Create a new DataFrame with the required structure
        result_df = pd.DataFrame({
            'media': media,
            'terpene': terpene,
            'terpene_pct': terpene_pct,
            'temperature': df[temp_col],
            'viscosity': df[visc_col]
        })
        
        # Extract sample name from filename
        filename = os.path.basename(file_path)
        print(f"Processed {filename}: Found {len(result_df)} data points")
        
        return result_df
        
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")
        return pd.DataFrame()

def batch_process_files(directory, output_file):
    """
    Process all Excel files in a directory and save to a single CSV.
    
    Args:
        directory: Directory containing Excel files
        output_file: Path to save the output CSV
    """
    # Find all Excel files in the directory
    excel_files = glob.glob(os.path.join(directory, "*.xlsx")) + glob.glob(os.path.join(directory, "*.xls"))
    
    if not excel_files:
        print(f"No Excel files found in {directory}")
        return
    
    all_data = []
    
    for file_path in excel_files:
        try:
            # Extract sample info from filename (customize this based on your naming convention)
            filename = os.path.basename(file_path)
            # Example: if filenames are like "D8_BlueDream_5pct.xlsx"
            parts = os.path.splitext(filename)[0].split('_')
            
            media = parts[0] if len(parts) > 0 else "Unknown"
            terpene = parts[1] if len(parts) > 1 else "Unknown"
            
            # Try to extract percentage from the filename if present
            terpene_pct = 0.0
            if len(parts) > 2 and 'pct' in parts[2]:
                try:
                    terpene_pct = float(parts[2].replace('pct', ''))
                except ValueError:
                    pass
            
            # Extract data from this file
            file_data = extract_viscosity_data(file_path, media, terpene, terpene_pct)
            if not file_data.empty:
                all_data.append(file_data)
                
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")
    
    if all_data:
        # Combine all data into a single DataFrame
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Save to CSV
        combined_df.to_csv(output_file, index=False)
        print(f"Successfully saved {len(combined_df)} data points to {output_file}")
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
    output_file = f"data/viscosity_scraped_{timestamp}.csv"
    
    batch_process_files(data_dir, output_file)
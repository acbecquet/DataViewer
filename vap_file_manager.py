"""
vap_file_manager.py
Developed By Charlie Becquet.
Manages the .vap3 file format for the DataViewer Application.

This module implements a custom file format (.vap3) that stores:
- Sheet data as CSVs
- Sheet metadata (is_plotting, is_empty)
- Associated images for each sheet
- Image crop states
- Plot data and settings
"""

import os
import json
import zipfile
import pandas as pd
import pickle
import io
import matplotlib.pyplot as plt
import datetime
import uuid
from typing import Dict, List, Any, Optional, Tuple
from PIL import Image
import tempfile
from utils import plotting_sheet_test

class VapFileManager:
    """Manager for the .vap3 file format, enabling storage and retrieval of test data."""
    
    def __init__(self):
        """Initialize the VapFileManager with version information."""
        self.version = "1.0"
        self.temp_files = []  # Track temporary files for cleanup
    
    def save_to_vap3(self, filepath: str, filtered_sheets: Dict, 
                    sheet_images: Dict, plot_options: List[str], 
                    image_crop_states: Dict = None, 
                    plot_settings: Dict = None) -> bool:
        """
        Save test data to a .vap3 file.
        
        Parameters:
        -----------
        filepath : str
            Path where the .vap3 file will be saved
        filtered_sheets : dict
            Dictionary containing sheet data
        sheet_images : dict
            Dictionary of images associated with sheets
        plot_options : List[str]
            Available plot types
        image_crop_states : dict, optional
            Dictionary tracking crop status of images
        plot_settings : dict, optional
            Dictionary containing plot settings (selected plot type, etc.)
        
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        if not filepath.endswith('.vap3'):
            filepath += '.vap3'
            
        try:
            with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as archive:
                # Create a metadata file
                file_id = str(uuid.uuid4())
                timestamp = datetime.datetime.now().isoformat()
                
                metadata = {
                    'version': self.version,
                    'file_id': file_id,
                    'timestamp': timestamp,
                    'sheet_names': list(filtered_sheets.keys()),
                    'has_images': bool(sheet_images)
                }
                
                # Write metadata to the archive
                archive.writestr('metadata.json', json.dumps(metadata))
                
                # Store plot options
                archive.writestr('plot_options.json', json.dumps(plot_options))
                
                # Store plot settings if provided
                if plot_settings:
                    archive.writestr('plot_settings.json', json.dumps(plot_settings))
                
                # Store each sheet's data
                for sheet_name, sheet_info in filtered_sheets.items():
                    # Check if this is a plotting sheet
                    is_plotting = plotting_sheet_test(sheet_name, sheet_info['data'])
                    is_empty = sheet_info.get('is_empty', False)
                    
                    # Create directory structure for the sheet
                    sheet_dir = f'sheets/{sheet_name}'
                    
                    # Convert DataFrame to CSV for better interoperability
                    buffer = io.StringIO()
                    sheet_info['data'].to_csv(buffer, index=False)
                    archive.writestr(f'{sheet_dir}/data.csv', buffer.getvalue())
                    
                    # Store sheet metadata
                    sheet_metadata = {
                        'is_plotting': is_plotting,
                        'is_empty': is_empty
                    }
                    archive.writestr(f'{sheet_dir}/metadata.json', 
                                    json.dumps(sheet_metadata))
                
                # Store images
                for file_name, file_sheets in sheet_images.items():
                    for sheet_name, images in file_sheets.items():
                        for i, img_path in enumerate(images):
                            if os.path.exists(img_path):
                                # Store the image in the archive
                                img_ext = os.path.splitext(img_path)[1]
                                archive.write(img_path, f'images/{sheet_name}/image_{i}{img_ext}')
                
                # Store image crop states
                if image_crop_states:
                    archive.writestr('image_crop_states.json', json.dumps(image_crop_states))
                
            return True
            
        except Exception as e:
            print(f"Error saving VAP3 file: {e}")
            # Clean up potentially corrupted file
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass
            return False
    
    def load_from_vap3(self, filepath: str) -> Dict[str, Any]:
        """
        Load test data from a .vap3 file.
        
        Parameters:
        -----------
        filepath : str
            Path to the .vap3 file
            
        Returns:
        --------
        dict
            Dictionary containing all extracted data including:
            - filtered_sheets: Dictionary of sheet data
            - sheet_images: Dictionary of images for each sheet
            - image_crop_states: Dictionary of crop states for images
            - plot_options: List of available plot types
            - plot_settings: Dictionary of plot settings
            - metadata: File metadata
        """
        # Initialize the result dictionary
        result = {
            'filtered_sheets': {},
            'sheet_images': {},
            'image_crop_states': {},
            'plot_options': [],
            'plot_settings': {},
            'metadata': {}
        }
        
        try:
            with zipfile.ZipFile(filepath, 'r') as archive:
                # Get the list of all files in the archive
                all_files = archive.namelist()
                
                # Extract metadata
                if 'metadata.json' in all_files:
                    metadata = json.loads(archive.read('metadata.json').decode('utf-8'))
                    result['metadata'] = metadata
                
                # Extract plot options
                if 'plot_options.json' in all_files:
                    plot_options = json.loads(archive.read('plot_options.json').decode('utf-8'))
                    result['plot_options'] = plot_options
                
                # Extract plot settings
                if 'plot_settings.json' in all_files:
                    plot_settings = json.loads(archive.read('plot_settings.json').decode('utf-8'))
                    result['plot_settings'] = plot_settings
                
                # Extract sheet data
                sheet_names = metadata.get('sheet_names', [])
                for sheet_name in sheet_names:
                    # Load sheet metadata
                    sheet_meta_path = f'sheets/{sheet_name}/metadata.json'
                    if sheet_meta_path in all_files:
                        sheet_meta = json.loads(archive.read(sheet_meta_path).decode('utf-8'))
                        
                        # Load sheet data
                        data_path = f'sheets/{sheet_name}/data.csv'
                        if data_path in all_files:
                            csv_data = archive.read(data_path).decode('utf-8')
                            data = pd.read_csv(io.StringIO(csv_data))
                            
                            # Reconstruct sheet info
                            result['filtered_sheets'][sheet_name] = {
                                'data': data,
                                'is_empty': sheet_meta.get('is_empty', False)
                            }
                
                # Extract images if they exist
                current_file = os.path.basename(filepath)
                result['sheet_images'][current_file] = {}
                
                image_files = [f for f in all_files if f.startswith('images/')]
                for img_file in image_files:
                    # Parse the path to get sheet name (e.g., 'images/Sheet1/image_0.png')
                    parts = img_file.split('/')
                    if len(parts) >= 3:
                        sheet_name = parts[1]
                        if sheet_name not in result['sheet_images'][current_file]:
                            result['sheet_images'][current_file][sheet_name] = []
                        
                        # Extract to a temporary file and add to result
                        with tempfile.NamedTemporaryFile(suffix=os.path.splitext(img_file)[1], delete=False) as tmpfile:
                            tmpfile.write(archive.read(img_file))
                            temp_path = tmpfile.name
                            self.temp_files.append(temp_path)  # Track for cleanup
                            result['sheet_images'][current_file][sheet_name].append(temp_path)
                
                # Extract image crop states if they exist
                if 'image_crop_states.json' in all_files:
                    result['image_crop_states'] = json.loads(
                        archive.read('image_crop_states.json').decode('utf-8')
                    )
            
            return result
            
        except Exception as e:
            print(f"Error loading VAP3 file: {e}")
            self.cleanup_temp_files()  # Clean up any temporary files created
            raise
    
    def cleanup_temp_files(self):
        """Clean up any temporary files created during loading."""
        for file_path in self.temp_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"Warning: Could not delete temporary file {file_path}: {e}")
        self.temp_files = []
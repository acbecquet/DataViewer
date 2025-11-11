"""
Excel Image Extractor Module for DataViewer Application

This module handles extraction of embedded images from Excel files and integrates
them into the loaded images display.
"""

import os
import io
import tempfile
from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from utils import debug_print


class ExcelImageExtractor:
    """Extracts embedded images from Excel files and integrates them into the GUI."""
    
    def __init__(self, gui):
        """Initialize the Excel image extractor.
        
        Args:
            gui: Reference to the main DataViewer GUI instance
        """
        self.gui = gui
        self.temp_dir = tempfile.mkdtemp(prefix="excel_images_")
        debug_print(f"DEBUG: Created temp directory for extracted images: {self.temp_dir}")
    
    def extract_images_from_excel(self, excel_path, target_sheet_name=None):
        """Extract all embedded images from an Excel file.
        
        Args:
            excel_path: Path to the Excel file
            target_sheet_name: Optional specific sheet to extract from
            
        Returns:
            Dictionary mapping sheet names to lists of extracted image paths
        """
        debug_print(f"DEBUG: Extracting images from Excel file: {excel_path}")
        
        extracted_images = {}
        
        try:
            # Load workbook with read-only mode for efficiency
            workbook = load_workbook(excel_path, data_only=True)
            
            debug_print(f"DEBUG: Loaded workbook with {len(workbook.sheetnames)} sheets")
            
            # Process each sheet
            for sheet_name in workbook.sheetnames:
                # Skip if target sheet specified and this isn't it
                if target_sheet_name and sheet_name != target_sheet_name:
                    continue
                
                sheet = workbook[sheet_name]
                
                # Check if sheet has images
                if not hasattr(sheet, '_images') or not sheet._images:
                    debug_print(f"DEBUG: No images found in sheet: {sheet_name}")
                    continue
                
                debug_print(f"DEBUG: Found {len(sheet._images)} images in sheet: {sheet_name}")
                
                sheet_images = []
                
                # Extract each image
                for idx, image in enumerate(sheet._images):
                    try:
                        extracted_path = self._extract_single_image(image, sheet_name, idx)
                        if extracted_path:
                            sheet_images.append(extracted_path)
                            debug_print(f"DEBUG: Extracted image {idx+1} from {sheet_name}")
                    
                    except Exception as e:
                        debug_print(f"DEBUG: Failed to extract image {idx+1} from {sheet_name}: {e}")
                        continue
                
                if sheet_images:
                    extracted_images[sheet_name] = sheet_images
                    debug_print(f"DEBUG: Extracted {len(sheet_images)} images from {sheet_name}")
            
            workbook.close()
            
            debug_print(f"DEBUG: Total extraction complete: {sum(len(imgs) for imgs in extracted_images.values())} images from {len(extracted_images)} sheets")
            
            return extracted_images
        
        except Exception as e:
            debug_print(f"ERROR: Failed to extract images from Excel: {e}")
            import traceback
            traceback.print_exc()
            return {}
    
    def _extract_single_image(self, image_obj, sheet_name, index):
        """Extract a single image from Excel to a temporary file.
        
        Args:
            image_obj: Openpyxl image object
            sheet_name: Name of the sheet containing the image
            index: Index of the image in the sheet
            
        Returns:
            Path to the extracted image file, or None if extraction failed
        """
        try:
            # Get image data
            image_data = image_obj._data()
            
            # Open with PIL
            pil_image = Image.open(io.BytesIO(image_data))
            
            # Determine format
            image_format = pil_image.format or 'PNG'
            extension = image_format.lower()
            
            # Create safe filename
            safe_sheet_name = "".join(c if c.isalnum() else "_" for c in sheet_name)
            filename = f"{safe_sheet_name}_image_{index+1}.{extension}"
            filepath = os.path.join(self.temp_dir, filename)
            
            # Save the image
            pil_image.save(filepath)
            
            debug_print(f"DEBUG: Saved extracted image to: {filepath}")
            debug_print(f"DEBUG: Image dimensions: {pil_image.size}, format: {image_format}")
            
            return filepath
        
        except Exception as e:
            debug_print(f"ERROR: Failed to extract single image: {e}")
            return None
    
    def integrate_extracted_images_to_gui(self, extracted_images, current_sheet=None, file_name=None):
        """Integrate extracted images into the GUI's image loader.
    
        Args:
            extracted_images: Dictionary mapping sheet names to image paths
            current_sheet: Optional current sheet name to integrate images for
            file_name: Optional explicit file name to use instead of gui.current_file
        """
        debug_print("DEBUG: Integrating extracted images into GUI")
    
        try:
            # Ensure sheet_images structure exists
            if not hasattr(self.gui, 'sheet_images'):
                self.gui.sheet_images = {}
        
            # Use explicit file_name if provided, otherwise fall back to current_file
            current_file = file_name if file_name else self.gui.current_file
            if not current_file:
                debug_print("DEBUG: No current file set and no file_name provided")
                return
        
            debug_print(f"DEBUG: Using file name: {current_file}")
        
            if current_file not in self.gui.sheet_images:
                self.gui.sheet_images[current_file] = {}
        
            # Ensure image_crop_states exists
            if not hasattr(self.gui, 'image_crop_states'):
                self.gui.image_crop_states = {}
        
            # Add images for each sheet
            for sheet_name, image_paths in extracted_images.items():
                if sheet_name not in self.gui.sheet_images[current_file]:
                    self.gui.sheet_images[current_file][sheet_name] = []
            
                # Add images if not already present
                for img_path in image_paths:
                    if img_path not in self.gui.sheet_images[current_file][sheet_name]:
                        self.gui.sheet_images[current_file][sheet_name].append(img_path)
                    
                        # Set default crop state
                        if img_path not in self.gui.image_crop_states:
                            self.gui.image_crop_states[img_path] = False
            
                debug_print(f"DEBUG: Integrated {len(image_paths)} images for sheet: {sheet_name}")
        
            # If current sheet has extracted images, update the display
            if current_sheet and current_sheet in extracted_images:
                if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
                    debug_print(f"DEBUG: Updating image display for current sheet: {current_sheet}")
                
                    # Load the images into the display
                    self.gui.image_loader.load_images_from_list(
                        self.gui.sheet_images[current_file][current_sheet]
                    )
                
                    # Restore crop states
                    for img_path in self.gui.sheet_images[current_file][current_sheet]:
                        if img_path in self.gui.image_crop_states:
                            self.gui.image_loader.image_crop_states[img_path] = self.gui.image_crop_states[img_path]
                
                    # Force refresh
                    self.gui.image_loader.display_images()
                    self.gui.image_frame.update_idletasks()
        
            debug_print("DEBUG: Image integration complete")
    
        except Exception as e:
            debug_print(f"ERROR: Failed to integrate extracted images: {e}")
            import traceback
            traceback.print_exc()
    
    def cleanup_temp_directory(self):
        """Clean up temporary directory containing extracted images."""
        try:
            if os.path.exists(self.temp_dir):
                import shutil
                shutil.rmtree(self.temp_dir)
                debug_print(f"DEBUG: Cleaned up temp directory: {self.temp_dir}")
        except Exception as e:
            debug_print(f"DEBUG: Error cleaning up temp directory: {e}")


def extract_and_load_excel_images(gui, excel_path, current_sheet=None):
    """Convenience function to extract and load images from Excel file.
    
    Args:
        gui: Main DataViewer GUI instance
        excel_path: Path to Excel file
        current_sheet: Optional current sheet name
        
    Returns:
        Number of images extracted
    """
    debug_print(f"DEBUG: Starting Excel image extraction for: {excel_path}")
    
    try:
        # Create extractor if needed
        if not hasattr(gui, 'excel_image_extractor') or gui.excel_image_extractor is None:
            gui.excel_image_extractor = ExcelImageExtractor(gui)
        
        extractor = gui.excel_image_extractor
        
        # Extract images
        extracted_images = extractor.extract_images_from_excel(excel_path, current_sheet)
        
        if not extracted_images:
            debug_print("DEBUG: No images found in Excel file")
            return 0
        
        # Get file name from path for integration
        file_name = os.path.basename(excel_path)
        debug_print(f"DEBUG: Using file name for integration: {file_name}")
        
        # Integrate into GUI with explicit file name
        extractor.integrate_extracted_images_to_gui(extracted_images, current_sheet, file_name=file_name)
        
        total_images = sum(len(imgs) for imgs in extracted_images.values())
        debug_print(f"DEBUG: Successfully extracted and loaded {total_images} images")
        
        return total_images
    
    except Exception as e:
        debug_print(f"ERROR: Failed to extract and load Excel images: {e}")
        import traceback
        traceback.print_exc()
        return 0
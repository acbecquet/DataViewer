
"""
Improved training extractor that uses OCR to detect actual attribute locations
instead of assuming fixed positions
"""

import cv2
import numpy as np
import os
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pytesseract
import re

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Note: You'll need to install pytesseract: pip install pytesseract
# And download Tesseract: https://github.com/UB-Mannheim/tesseract/wiki

class ImprovedAttributeDetectionExtractor:
    """
    Enhanced extractor that finds actual attribute locations using OCR
    instead of assuming fixed positions.
    """
    
    def __init__(self):
        # High resolution for readable training data
        self.target_size = (600, 140)  # Width x Height - full row capture
        self.display_scale = 2
        
        # Expected attributes (we'll search for these in the image)
        self.expected_attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        # Alternative attribute names to search for (in case of OCR variations)
        self.attribute_variations = {
            "Burnt Taste": ["burnt taste", "burnt", "taste"],
            "Vapor Volume": ["vapor volume", "vapor", "volume"],
            "Overall Flavor": ["overall flavor", "flavor", "flavour"],
            "Smoothness": ["smoothness", "smooth"],
            "Overall Liking": ["overall liking", "liking", "overall", "like"]
        }
        
        self.session_log = []
        
        print("Improved Attribute Detection Extractor initialized")
        print("This version uses OCR to find actual attribute locations")
    
    def preprocess_image_for_extraction(self, image_path):
        """
        Preprocess image for both boundary detection and OCR.
        """
        
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Could not load image: {image_path}")
        
        print(f"Original image size: {img.shape[1]}x{img.shape[0]}")
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Enhance contrast for better OCR
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        return enhanced
    
    def detect_form_cross_lines(self, image):
        """
        Detect the black cross lines that separate the 4 samples.
        """
        
        height, width = image.shape
        print(f"Detecting cross lines in {width}x{height} image")
        
        # Apply edge detection
        edges = cv2.Canny(image, 50, 150, apertureSize=3)
        
        # Detect lines
        lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=int(min(width, height) * 0.3))
        
        horizontal_lines = []
        vertical_lines = []
        
        if lines is not None:
            for line in lines:
                rho, theta = line[0]
                angle_deg = np.degrees(theta)
                
                if abs(angle_deg - 90) < 15:  # Vertical line
                    x = int(rho / np.cos(theta))
                    if width * 0.3 < x < width * 0.7:
                        vertical_lines.append(x)
                        print(f"  Found vertical line at x={x}")
                
                elif abs(angle_deg) < 15 or abs(angle_deg - 180) < 15:  # Horizontal line
                    y = int(rho / np.sin(theta)) if abs(np.sin(theta)) > 0.1 else None
                    if y is not None and height * 0.3 < y < height * 0.8:
                        horizontal_lines.append(y)
                        print(f"  Found horizontal line at y={y}")
        
        # Use detected lines or fallback to estimates
        center_x = int(np.median(vertical_lines)) if vertical_lines else width // 2
        center_y = int(np.median(horizontal_lines)) if horizontal_lines else int(height * 0.6)
        
        print(f"Cross lines detected at: center_x={center_x}, center_y={center_y}")
        return center_x, center_y
    
    def get_sample_regions(self, image):
        """
        Get the 4 sample regions based on cross line detection.
        """
        
        height, width = image.shape
        center_x, center_y = self.detect_form_cross_lines(image)
        
        # Define margins
        top_margin = int(height * 0.15)  # Skip header
        bottom_margin = int(height * 0.95)
        left_margin = int(width * 0.02)
        right_margin = int(width * 0.98)
        
        sample_regions = {
            1: {  # Top-left
                "y_start": top_margin, 
                "y_end": center_y - 5,
                "x_start": left_margin, 
                "x_end": center_x - 5
            },
            2: {  # Top-right
                "y_start": top_margin, 
                "y_end": center_y - 5,
                "x_start": center_x + 5, 
                "x_end": right_margin
            },
            3: {  # Bottom-left
                "y_start": center_y + 5, 
                "y_end": bottom_margin,
                "x_start": left_margin, 
                "x_end": center_x - 5
            },
            4: {  # Bottom-right
                "y_start": center_y + 5, 
                "y_end": bottom_margin,
                "x_start": center_x + 5, 
                "x_end": right_margin
            }
        }
        
        return sample_regions
    
    def find_attributes_in_sample(self, image, sample_region):
        """
        Use OCR to find where each attribute is located within a sample region.
        """
        
        print(f"Searching for attributes in sample region...")
        
        # Extract the sample region
        y_start, y_end = sample_region["y_start"], sample_region["y_end"]
        x_start, x_end = sample_region["x_start"], sample_region["x_end"]
        
        sample_img = image[y_start:y_end, x_start:x_end]
        
        if sample_img.size == 0:
            print("  ERROR: Empty sample region")
            return {}
        
        print(f"  Sample region size: {sample_img.shape[1]}x{sample_img.shape[0]}")
        
        # Enhance image for better OCR
        enhanced = cv2.resize(sample_img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        # Apply additional preprocessing for OCR
        _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        try:
            # Get OCR data with bounding boxes
            ocr_data = pytesseract.image_to_data(binary, output_type=pytesseract.Output.DICT)
            
            found_attributes = {}
            
            # Search for each expected attribute
            for attribute in self.expected_attributes:
                print(f"  Searching for: {attribute}")
                
                # Get all possible variations to search for
                search_terms = [attribute.lower()]
                if attribute in self.attribute_variations:
                    search_terms.extend(self.attribute_variations[attribute])
                
                best_match = None
                best_confidence = 0
                
                # Check each detected text element
                for i, text in enumerate(ocr_data['text']):
                    if not text.strip():
                        continue
                    
                    text_lower = text.lower().strip()
                    confidence = int(ocr_data['conf'][i])
                    
                    # Check if this text matches any of our search terms
                    for search_term in search_terms:
                        if search_term in text_lower or text_lower in search_term:
                            if confidence > best_confidence:
                                # Calculate position (scale back from 2x enhanced image)
                                x = ocr_data['left'][i] // 2
                                y = ocr_data['top'][i] // 2
                                w = ocr_data['width'][i] // 2
                                h = ocr_data['height'][i] // 2
                                
                                best_match = {
                                    'x': x, 'y': y, 'width': w, 'height': h,
                                    'text': text, 'confidence': confidence
                                }
                                best_confidence = confidence
                                print(f"    Found potential match: '{text}' at y={y} (confidence: {confidence})")
                
                if best_match:
                    found_attributes[attribute] = best_match
                    print(f"    ✓ Best match for {attribute}: '{best_match['text']}' at y={best_match['y']}")
                else:
                    print(f"    ✗ No match found for {attribute}")
            
            return found_attributes
            
        except Exception as e:
            print(f"  OCR Error: {e}")
            print("  Falling back to estimated positions...")
            return self.get_fallback_attribute_positions(sample_img)
    
    def get_fallback_attribute_positions(self, sample_img):
        """
        Fallback method if OCR fails - use estimated positions.
        """
        
        height = sample_img.shape[0]
        attr_height = height // len(self.expected_attributes)
        
        fallback_attributes = {}
        for i, attribute in enumerate(self.expected_attributes):
            y_pos = i * attr_height + attr_height // 2
            fallback_attributes[attribute] = {
                'x': 0, 'y': y_pos, 'width': 100, 'height': attr_height,
                'text': 'estimated', 'confidence': 0
            }
            print(f"    Fallback position for {attribute}: y={y_pos}")
        
        return fallback_attributes
    
    def extract_attribute_row(self, image, sample_region, attribute_info):
        """
        Extract the complete row for a specific attribute.
        """
        
        # Get sample region boundaries
        sample_y_start = sample_region["y_start"]
        sample_x_start = sample_region["x_start"]
        sample_x_end = sample_region["x_end"]
        
        # Calculate actual image coordinates
        attr_y_center = sample_y_start + attribute_info['y'] + attribute_info['height'] // 2
        
        # Define row boundaries with buffer
        row_height = max(40, attribute_info['height'] * 2)  # Ensure minimum height
        row_y_start = max(0, attr_y_center - row_height // 2)
        row_y_end = min(image.shape[0], attr_y_center + row_height // 2)
        
        # Use full width of sample region to capture complete rating scale
        row_x_start = sample_x_start
        row_x_end = sample_x_end
        
        # Extract the row
        row_img = image[row_y_start:row_y_end, row_x_start:row_x_end]
        
        print(f"    Extracted row: {row_img.shape[1]}x{row_img.shape[0]} at y={row_y_start}-{row_y_end}")
        
        return row_img
    
    def get_improved_form_regions(self, image):
        """
        Extract regions using OCR-based attribute detection.
        """
        
        height, width = image.shape
        print(f"Processing image dimensions: {width}x{height}")
        
        # Get the 4 sample regions
        sample_regions = self.get_sample_regions(image)
        
        all_regions = {}
        
        # Process each sample region
        for sample_id, sample_region in sample_regions.items():
            print(f"\nProcessing Sample {sample_id}:")
            print(f"  Region: y={sample_region['y_start']}-{sample_region['y_end']}, x={sample_region['x_start']}-{sample_region['x_end']}")
            
            # Find attributes in this sample using OCR
            found_attributes = self.find_attributes_in_sample(image, sample_region)
            
            sample_name = f"Sample {sample_id}"
            all_regions[sample_name] = {}
            
            # Extract row for each found attribute
            for attribute, attr_info in found_attributes.items():
                print(f"  Extracting row for {attribute}...")
                
                try:
                    row_img = self.extract_attribute_row(image, sample_region, attr_info)
                    
                    if row_img.size > 0:
                        # Resize to target size
                        resized = cv2.resize(row_img, self.target_size, interpolation=cv2.INTER_CUBIC)
                        all_regions[sample_name][attribute] = resized
                        
                        # Save debug image
                        debug_info = {
                            'sample_id': sample_id,
                            'attribute': attribute,
                            'ocr_text': attr_info['text'],
                            'confidence': attr_info['confidence'],
                            'y_position': attr_info['y']
                        }
                        self.save_debug_region(resized, f"Sample{sample_id}_{attribute.replace(' ', '_')}", debug_info)
                        
                        print(f"     Successfully extracted and resized to {self.target_size}")
                    else:
                        print(f"     Empty row extracted")
                        
                except Exception as e:
                    print(f"     Error extracting row: {e}")
            
            print(f"  Sample {sample_id} complete: {len(all_regions[sample_name])} attributes extracted")
        
        return all_regions
    
    def save_debug_region(self, region_img, label, debug_info=None):
        """
        Save debug images with enhanced information.
        """
        
        debug_dir = "training_data/debug_regions"
        os.makedirs(debug_dir, exist_ok=True)
        
        # Save the region image
        debug_path = os.path.join(debug_dir, f"debug_{label}.png")
        cv2.imwrite(debug_path, region_img)
        
        # Save debug info
        if debug_info:
            info_path = os.path.join(debug_dir, f"debug_{label}_info.txt")
            with open(info_path, 'w') as f:
                f.write(f"Attribute: {debug_info.get('attribute', 'Unknown')}\n")
                f.write(f"OCR Text Found: '{debug_info.get('ocr_text', 'None')}'\n")
                f.write(f"OCR Confidence: {debug_info.get('confidence', 0)}\n")
                f.write(f"Y Position in Sample: {debug_info.get('y_position', 'Unknown')}\n")
                f.write(f"Final Size: {region_img.shape[1]}x{region_img.shape[0]}\n")
                f.write(f"Target Size: {self.target_size[0]}x{self.target_size[1]}\n")
    
    def extract_training_data_from_images(self, image_folder):
        """
        Main extraction method using OCR-based attribute detection.
        """
        
        print(f"="*60)
        print("IMPROVED ATTRIBUTE DETECTION EXTRACTION")
        print(f"="*60)
        print(f"Source folder: {image_folder}")
        print(f"Method: OCR-based attribute detection")
        print(f"Target resolution: {self.target_size}")
        print()
        
        # Check if Tesseract is available
        try:
            pytesseract.get_tesseract_version()
            print("YES Tesseract OCR is available")
        except:
            print("NO Tesseract OCR not found. Please install it:")
            print("  pip install pytesseract")
            print("  Download Tesseract: https://github.com/UB-Mannheim/tesseract/wiki")
            return
        
        # Find image files
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']
        image_files = []
        
        for filename in os.listdir(image_folder):
            if any(filename.lower().endswith(ext) for ext in image_extensions):
                image_files.append(os.path.join(image_folder, filename))
        
        if not image_files:
            print(f"No image files found in {image_folder}")
            return
        
        print(f"Found {len(image_files)} images to process")
        
        # Create directory structure
        self.create_training_data_structure()
        
        total_regions = 0
        processed_forms = 0
        
        for image_path in image_files:
            print(f"\n{'='*50}")
            print(f"Processing: {os.path.basename(image_path)}")
            print(f"{'='*50}")
            
            try:
                # Preprocess image
                processed_image = self.preprocess_image_for_extraction(image_path)
                
                # Extract regions using OCR-based detection
                regions = self.get_improved_form_regions(processed_image)
                
                # Summary
                total_extracted = sum(len(sample_regions) for sample_regions in regions.values())
                print(f"\nExtraction summary: {total_extracted} total regions found")
                for sample_name, sample_regions in regions.items():
                    print(f"  {sample_name}: {list(sample_regions.keys())}")
                
                processed_forms += 1
                total_regions += total_extracted
                
            except Exception as e:
                print(f"Error processing {os.path.basename(image_path)}: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        print(f"\n{'='*60}")
        print("OCR-BASED EXTRACTION COMPLETE")
        print(f"{'='*60}")
        print(f"Forms processed: {processed_forms}")
        print(f"Total regions extracted: {total_regions}")
        print(f"Debug images saved to: training_data/debug_regions/")
        print()
        print("Next steps:")
        print("1. Check debug images to verify OCR found correct attributes")
        print("2. Look for files ending with '_info.txt' for OCR details")
        print("3. If extraction looks good, proceed with manual rating")
    
    def create_training_data_structure(self):
        """Create training directory structure."""
        
        base_dir = "training_data/sensory_ratings"
        os.makedirs(base_dir, exist_ok=True)
        
        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            os.makedirs(rating_dir, exist_ok=True)
        
        os.makedirs("training_data/logs", exist_ok=True)
        os.makedirs("training_data/debug_regions", exist_ok=True)
        os.makedirs("models", exist_ok=True)
        
        print("Training data structure created")

# Updated test script
if __name__ == "__main__":
    print("Improved Attribute Detection Extractor Test")
    print("="*50)
    
    # Test with your image path
    test_image_path = r"C:\Users\Alexander Becquet\Documents\Python\Python\TPM Data Processing Python Scripts\Standardized Testing GUI\git testing\DataViewer\tests\test_sensory\test.jpg"
    
    print(f"Testing OCR-based attribute detection...")
    
    if os.path.exists(test_image_path):
        try:
            extractor = ImprovedAttributeDetectionExtractor()
            
            # Load and preprocess
            processed_image = extractor.preprocess_image_for_extraction(test_image_path)
            
            # Test the improved extraction
            regions = extractor.get_improved_form_regions(processed_image)
            
            print(f"\n" + "="*50)
            print("EXTRACTION RESULTS:")
            print("="*50)
            
            for sample_name, sample_regions in regions.items():
                print(f"{sample_name}:")
                for attribute, region_img in sample_regions.items():
                    print(f"  ✓ {attribute}: {region_img.shape[1]}x{region_img.shape[0]}")
            
            print(f"\nDebug images and info saved to: training_data/debug_regions/")
            print("Check the debug images to verify that:")
            print("1. Each image shows the correct attribute name")
            print("2. The rating scale 1-9 is visible to the right")
            print("3. OCR correctly identified attribute locations")
            
        except Exception as e:
            print(f"Test failed: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"Test image not found: {test_image_path}")
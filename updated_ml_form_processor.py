# updated_ml_form_processor.py
"""
Updated ML form processor that matches the training extractor boundaries exactly
"""

import cv2
import numpy as np
import os
import tensorflow as tf
from tensorflow import keras
import json
from datetime import datetime

class UpdatedMLFormProcessor:
    """
    ML form processor that uses the exact same boundary detection and region extraction
    as the updated training extractor to ensure consistent results.
    """
    
    def __init__(self, model_path=None):
        # EXACT same parameters as training extractor
        self.target_size = (600, 140)  # Must match training exactly
        
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        # Load trained model if provided
        self.model = None
        if model_path and os.path.exists(model_path):
            try:
                self.model = keras.models.load_model(model_path)
                print(f"Loaded model from {model_path}")
            except Exception as e:
                print(f"Error loading model: {e}")
                print("Will use placeholder predictions")
        else:
            print("No model loaded - using placeholder predictions")
    
    def preprocess_image(self, image_path):
        """
        EXACT same preprocessing as training extractor.
        """
        
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Could not load image: {image_path}")
        
        print(f"Processing image size: {img.shape[1]}x{img.shape[0]}")
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # EXACT same enhancement as training
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # EXACT same denoising as training
        denoised = cv2.bilateralFilter(enhanced, 5, 25, 25)
        
        return denoised
    
    def detect_form_cross_lines(self, image):
        """
        EXACT same cross line detection as training extractor.
        """
        
        height, width = image.shape
        print(f"Detecting cross lines in {width}x{height} image")
        
        # Apply edge detection to find lines
        edges = cv2.Canny(image, 50, 150, apertureSize=3)
        
        # Detect lines using HoughLines
        lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=int(min(width, height) * 0.3))
        
        horizontal_lines = []
        vertical_lines = []
        
        if lines is not None:
            for line in lines:
                rho, theta = line[0]
                
                # Classify as horizontal or vertical
                angle_deg = np.degrees(theta)
                
                if abs(angle_deg - 90) < 15:  # Vertical line (around 90 degrees)
                    x = int(rho / np.cos(theta))
                    if width * 0.3 < x < width * 0.7:  # Center region
                        vertical_lines.append(x)
                        print(f"  Found vertical line at x={x}")
                
                elif abs(angle_deg) < 15 or abs(angle_deg - 180) < 15:  # Horizontal line
                    y = int(rho / np.sin(theta)) if abs(np.sin(theta)) > 0.1 else None
                    if y is not None and height * 0.3 < y < height * 0.8:  # Middle region
                        horizontal_lines.append(y)
                        print(f"  Found horizontal line at y={y}")
        
        # Use the strongest lines or fall back to estimated positions
        if vertical_lines:
            center_x = int(np.median(vertical_lines))
        else:
            center_x = width // 2
            print("  No vertical line detected, using center estimate")
        
        if horizontal_lines:
            center_y = int(np.median(horizontal_lines))
        else:
            center_y = int(height * 0.6)  # Slightly below center for typical forms
            print("  No horizontal line detected, using estimated position")
        
        print(f"Cross lines detected at: center_x={center_x}, center_y={center_y}")
        return center_x, center_y
    
    def extract_regions_for_prediction(self, processed_image):
        """
        EXACT same region extraction as training extractor.
        This ensures training and inference use identical parameters.
        """
        
        height, width = processed_image.shape
        print(f"Extracting regions from {width}x{height} image")
        
        # EXACT same cross line detection as training
        center_x, center_y = self.detect_form_cross_lines(processed_image)
        
        # EXACT same margins as training
        top_margin = int(height * 0.15)  # Skip header area
        bottom_margin = int(height * 0.95)  # Stop before form bottom
        left_margin = int(width * 0.02)  # Small left margin
        right_margin = int(width * 0.98)  # Extend to near edge to capture rating 9
        
        # EXACT same sample regions as training
        sample_regions = {
            1: {  # Top-left sample
                "y_start": top_margin, 
                "y_end": center_y - 5,  # 5 pixel buffer from center line
                "x_start": left_margin, 
                "x_end": center_x - 5   # 5 pixel buffer from center line
            },
            2: {  # Top-right sample  
                "y_start": top_margin, 
                "y_end": center_y - 5,
                "x_start": center_x + 5,  # 5 pixel buffer from center line
                "x_end": right_margin
            },
            3: {  # Bottom-left sample
                "y_start": center_y + 5,  # 5 pixel buffer from center line
                "y_end": bottom_margin,
                "x_start": left_margin, 
                "x_end": center_x - 5
            },
            4: {  # Bottom-right sample
                "y_start": center_y + 5, 
                "y_end": bottom_margin,
                "x_start": center_x + 5, 
                "x_end": right_margin
            }
        }
        
        regions = {}
        
        # Extract regions with EXACT same logic as training
        for sample_id, sample_region in sample_regions.items():
            sample_height = sample_region["y_end"] - sample_region["y_start"]
            attr_height = sample_height // len(self.attributes)
            
            regions[f"Sample {sample_id}"] = {}
            
            for i, attribute in enumerate(self.attributes):
                # Calculate boundaries exactly like training
                y_start = sample_region["y_start"] + i * attr_height
                y_end = sample_region["y_start"] + (i + 1) * attr_height
                x_start = sample_region["x_start"]
                x_end = sample_region["x_end"]
                
                # EXACT same buffer as training extractor
                buffer_y = max(2, int(attr_height * 0.05))  # 5% or minimum 2 pixels
                buffer_x = int((x_end - x_start) * 0.01)    # 1% horizontal buffer
                
                y_start = max(0, y_start - buffer_y)
                y_end = min(height, y_end + buffer_y)
                x_start = max(0, x_start - buffer_x)
                x_end = min(width, x_end + buffer_x)
                
                # Extract region
                region_img = processed_image[y_start:y_end, x_start:x_end]
                
                if region_img.size > 0:
                    # EXACT same resizing as training
                    resized = cv2.resize(region_img, self.target_size, interpolation=cv2.INTER_CUBIC)
                    regions[f"Sample {sample_id}"][attribute] = resized
                    
        return regions
    
    def predict_rating(self, region_image):
        """
        Predict the rating (1-9) for a single region image.
        """
        
        if self.model is None:
            # Placeholder prediction - replace with actual model prediction
            print("  Using placeholder prediction (no model loaded)")
            return np.random.randint(1, 10)  # Random rating 1-9
        
        try:
            # Normalize image for model input
            normalized = region_image.astype(np.float32) / 255.0
            
            # Add batch dimension
            input_data = np.expand_dims(normalized, axis=0)
            
            # Make prediction
            prediction = self.model.predict(input_data, verbose=0)
            
            # Convert to rating (assuming model outputs class probabilities)
            predicted_class = np.argmax(prediction, axis=1)[0]
            predicted_rating = predicted_class + 1  # Convert 0-8 to 1-9
            
            return predicted_rating
            
        except Exception as e:
            print(f"  Prediction error: {e}")
            return 5  # Default middle rating
    
    def process_form_image(self, image_path):
        """
        Process a complete form image and extract all ratings.
        """
        
        print(f"Processing form: {os.path.basename(image_path)}")
        
        try:
            # Preprocess image
            processed_image = self.preprocess_image(image_path)
            
            # Extract regions
            regions = self.extract_regions_for_prediction(processed_image)
            
            # Predict ratings for each region
            results = {}
            for sample_name, sample_regions in regions.items():
                results[sample_name] = {}
                
                print(f"Processing {sample_name}:")
                for attribute, region_img in sample_regions.items():
                    rating = self.predict_rating(region_img)
                    results[sample_name][attribute] = rating
                    print(f"  {attribute}: {rating}")
                
                # Placeholder for comments (can be added later with OCR)
                results[sample_name]['comments'] = ''
            
            print(f"Processing complete: {len(results)} samples processed")
            return results
            
        except Exception as e:
            print(f"Error processing form: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def save_debug_regions(self, regions, output_dir="debug_inference_regions"):
        """
        Save debug images during inference to verify region extraction.
        """
        
        os.makedirs(output_dir, exist_ok=True)
        
        for sample_name, sample_regions in regions.items():
            for attribute, region_img in sample_regions.items():
                safe_sample = sample_name.replace(" ", "_")
                safe_attribute = attribute.replace(" ", "_")
                
                filename = f"inference_{safe_sample}_{safe_attribute}.png"
                filepath = os.path.join(output_dir, filename)
                
                cv2.imwrite(filepath, region_img)
        
        print(f"Debug regions saved to {output_dir}/")

# Usage example
if __name__ == "__main__":
    print("Updated ML Form Processor")
    print("This processor uses exact same boundaries as the updated training extractor")
    print()
    
    # Example usage:
    processor = UpdatedMLFormProcessor()
    
    # Uncomment to test:
    # results = processor.process_form_image('path/to/test/form.jpg')
    # print("Results:", results)
    
    print("To use this processor:")
    print("1. processor = UpdatedMLFormProcessor('path/to/trained/model.h5')")
    print("2. results = processor.process_form_image('path/to/form/image.jpg')")
    print("3. Check debug_inference_regions/ to verify extraction matches training")
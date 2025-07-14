# improved_training_extractor.py
"""
Improved training data extractor with high-resolution regions and better alignment
"""

import cv2
import numpy as np
import os
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import random

class ImprovedTrainingExtractor:
    """
    Enhanced training extractor that captures high-quality, readable training regions
    instead of tiny pixelated images.
    """
    
    def __init__(self):
        # Much higher resolution to ensure readability
        self.target_size = (500, 120)  # Width x Height - even larger for full clarity
        self.display_scale = 2  # Scale factor for UI display
        
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        self.session_log = []
        
        print("Improved Training Extractor initialized")
        print(f"Target region size: {self.target_size} (full row capture)")
        print(f"Display scale: {self.display_scale}x for better viewing")
    
    def preprocess_image_for_extraction(self, image_path):
        """
        Minimal preprocessing that preserves detail while enhancing readability.
        """
        
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Could not load image: {image_path}")
        
        print(f"Original image size: {img.shape[1]}x{img.shape[0]}")
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply minimal enhancement that preserves detail
        # Use adaptive histogram equalization for better contrast
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # Very light denoising to reduce camera noise without blurring
        denoised = cv2.bilateralFilter(enhanced, 3, 20, 20)
        
        return denoised
    
    def get_improved_form_regions(self, image):
        """
        Extract full attribute rows with proper boundaries to capture complete rating scales.
        """
        
        height, width = image.shape
        regions = {}
        
        print(f"Processing image dimensions: {width}x{height}")
        
        # Extended sample regions to capture full width including rating 9
        sample_regions = {
            1: {  # Top-left sample - extended right boundary
                "y_start": int(height * 0.28), "y_end": int(height * 0.56), 
                "x_start": int(width * 0.05), "x_end": int(width * 0.50)  # Extended to 50%
            },
            2: {  # Top-right sample - extended right boundary
                "y_start": int(height * 0.28), "y_end": int(height * 0.56),
                "x_start": int(width * 0.50), "x_end": int(width * 0.98)  # Extended to 98%
            },
            3: {  # Bottom-left sample - extended right boundary
                "y_start": int(height * 0.64), "y_end": int(height * 0.90),
                "x_start": int(width * 0.05), "x_end": int(width * 0.50)  # Extended to 50%
            },
            4: {  # Bottom-right sample - extended right boundary
                "y_start": int(height * 0.64), "y_end": int(height * 0.90),
                "x_start": int(width * 0.50), "x_end": int(width * 0.98)  # Extended to 98%
            }
        }
        
        # Extract individual attribute rows with precise boundaries
        for sample_id, sample_region in sample_regions.items():
            sample_height = sample_region["y_end"] - sample_region["y_start"]
            attr_height = sample_height // len(self.attributes)
            
            regions[f"Sample {sample_id}"] = {}
            
            print(f"Sample {sample_id} region: {sample_region}")
            
            for i, attribute in enumerate(self.attributes):
                # Calculate precise boundaries for each attribute row
                y_start = sample_region["y_start"] + i * attr_height
                y_end = sample_region["y_start"] + (i + 1) * attr_height
                x_start = sample_region["x_start"]
                x_end = sample_region["x_end"]
                
                # Minimal padding to avoid cutting off content
                padding_y = int(attr_height * 0.02)  # Very small Y padding
                padding_x = int((x_end - x_start) * 0.01)  # Minimal X padding
                
                y_start = max(0, y_start - padding_y)
                y_end = min(height, y_end + padding_y)
                x_start = max(0, x_start - padding_x)
                x_end = min(width, x_end + padding_x)
                
                # Extract the full attribute row
                region_img = image[y_start:y_end, x_start:x_end]
                
                if region_img.size > 0:
                    print(f"  {attribute}: extracted {region_img.shape} -> will resize to {self.target_size}")
                    
                    # Resize to target size using high-quality interpolation
                    resized = cv2.resize(region_img, self.target_size, interpolation=cv2.INTER_CUBIC)
                    regions[f"Sample {sample_id}"][attribute] = resized
                    
                    # Save debug image to check extraction quality
                    self.save_debug_region(resized, f"Sample{sample_id}_{attribute}")
                    
        return regions
    
    def save_debug_region(self, region_img, label):
        """
        Save debug images so you can verify extraction quality.
        """
        
        debug_dir = "training_data/debug_regions"
        os.makedirs(debug_dir, exist_ok=True)
        
        debug_path = os.path.join(debug_dir, f"debug_{label}.png")
        cv2.imwrite(debug_path, region_img)
    
    def show_region_and_get_rating_improved(self, region_img, sample_name, attribute, region_num, total_regions):
        """
        Enhanced UI for rating regions with better image display and quality indicators.
        """
        
        # Create main window
        root = tk.Toplevel()
        root.title(f"Rate High-Resolution Region {region_num}/{total_regions}")
        root.geometry("800x600")
        root.configure(bg='white')
        
        # Center and focus the window
        root.transient()
        root.grab_set()
        
        # Convert to PIL for display
        region_pil = Image.fromarray(region_img)
        
        # Scale up for display while maintaining aspect ratio
        display_width = int(self.target_size[0] * self.display_scale)
        display_height = int(self.target_size[1] * self.display_scale)
        region_display = region_pil.resize((display_width, display_height), Image.Resampling.NEAREST)
        region_photo = ImageTk.PhotoImage(region_display)
        
        # Header with progress
        header_frame = tk.Frame(root, bg='white')
        header_frame.pack(pady=10)
        
        title_label = tk.Label(header_frame, text=f"{sample_name} - {attribute}", 
                              font=('Arial', 16, 'bold'), bg='white')
        title_label.pack()
        
        progress_label = tk.Label(header_frame, text=f"Region {region_num} of {total_regions}", 
                                 font=('Arial', 11), bg='white', fg='gray')
        progress_label.pack()
        
        # Quality indicator
        quality_label = tk.Label(header_frame, 
                                text=f"Resolution: {self.target_size[0]}x{self.target_size[1]} pixels (Full Row Capture)", 
                                font=('Arial', 10), bg='white', fg='green')
        quality_label.pack()
        
        # Image display with border
        image_frame = tk.Frame(root, bg='gray', relief='solid', bd=2)
        image_frame.pack(pady=10)
        
        image_label = tk.Label(image_frame, image=region_photo, bg='white')
        image_label.pack(padx=5, pady=5)
        
        # Instructions
        instruction_frame = tk.Frame(root, bg='lightblue', relief='ridge', bd=2)
        instruction_frame.pack(pady=10, padx=20, fill='x')
        
        instruction_text = """
        INSTRUCTIONS:
        • Look at the full rating row in the high-resolution image above
        • Find which number (1-9) has a circle or mark around it  
        • The full row should be visible from attribute name to rating 9
        • Select that number below - if you can't see rating 9, report quality issue
        """
        
        instruction_label = tk.Label(instruction_frame, text=instruction_text, 
                                   font=('Arial', 10), bg='lightblue', 
                                   justify='left', padx=10, pady=5)
        instruction_label.pack()
        
        # Rating selection with larger buttons
        rating_frame = tk.Frame(root, bg='white')
        rating_frame.pack(pady=15)
        
        tk.Label(rating_frame, text="Select the circled rating (1-9):", 
                font=('Arial', 13, 'bold'), bg='white').pack(pady=5)
        
        rating_var = tk.StringVar()
        button_frame = tk.Frame(rating_frame, bg='white')
        button_frame.pack()
        
        # Create larger, more visible rating buttons
        for i in range(1, 10):
            btn = tk.Radiobutton(button_frame, text=str(i), variable=rating_var, 
                                value=str(i), font=('Arial', 14, 'bold'),
                                width=3, height=1, indicatoron=0,
                                selectcolor='lightgreen', bg='lightgray')
            btn.pack(side='left', padx=2, pady=5)
        
        # Control buttons
        control_frame = tk.Frame(root, bg='white')
        control_frame.pack(pady=20)
        
        result = {'rating': None, 'skip': False, 'quality_issue': False}
        
        def submit_rating():
            if rating_var.get():
                result['rating'] = int(rating_var.get())
                root.destroy()
            else:
                messagebox.showwarning("No Rating Selected", "Please select a rating from 1-9")
        
        def skip_region():
            result['skip'] = True
            root.destroy()
        
        def report_quality_issue():
            result['quality_issue'] = True
            root.destroy()
        
        # Enhanced control buttons
        submit_btn = tk.Button(control_frame, text="Submit Rating", 
                              command=submit_rating, bg='lightgreen', 
                              font=('Arial', 12, 'bold'), padx=20, pady=5)
        submit_btn.pack(side='left', padx=10)
        
        skip_btn = tk.Button(control_frame, text="Skip Region", 
                            command=skip_region, bg='lightyellow',
                            font=('Arial', 12), padx=20, pady=5)
        skip_btn.pack(side='left', padx=10)
        
        quality_btn = tk.Button(control_frame, text="Quality Issue", 
                               command=report_quality_issue, bg='lightcoral',
                               font=('Arial', 12), padx=20, pady=5)
        quality_btn.pack(side='left', padx=10)
        
        # Handle window closing
        def on_closing():
            result['skip'] = True
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Set focus and wait
        root.focus_set()
        root.wait_window()
        
        # Handle different result types
        if result['quality_issue']:
            print(f"  Quality issue reported for {sample_name} - {attribute}")
            return 'QUALITY_ISSUE'
        elif result['skip']:
            return None
        else:
            return result['rating']
    
    def save_high_quality_training_region(self, region_img, rating, sample_name, attribute):
        """
        Save high-quality training regions with appropriate augmentation.
        """
        
        base_dir = "training_data/sensory_ratings"
        rating_dir = os.path.join(base_dir, f"rating_{rating}")
        os.makedirs(rating_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        sample_clean = sample_name.replace(" ", "_")
        attribute_clean = attribute.replace(" ", "_")
        
        filenames = []
        
        # Save original high-quality region
        original_filename = f"{sample_clean}_{attribute_clean}_{timestamp}_hq.png"
        original_path = os.path.join(rating_dir, original_filename)
        cv2.imwrite(original_path, region_img)
        filenames.append(original_filename)
        
        # Generate realistic augmentations that preserve readability
        augmentation_configs = [
            {'brightness': 0.15, 'noise': 0.01, 'suffix': 'bright'},
            {'brightness': -0.1, 'noise': 0.015, 'suffix': 'dark'},
            {'brightness': 0.05, 'noise': 0.02, 'suffix': 'noise'},
            {'brightness': 0.1, 'noise': 0.005, 'contrast': 1.1, 'suffix': 'enhanced'},
            {'brightness': -0.05, 'noise': 0.01, 'contrast': 0.9, 'suffix': 'reduced'}
        ]
        
        for config in augmentation_configs:
            augmented = self.apply_quality_preserving_augmentation(region_img, config)
            
            aug_filename = f"{sample_clean}_{attribute_clean}_{timestamp}_{config['suffix']}.png"
            aug_path = os.path.join(rating_dir, aug_filename)
            cv2.imwrite(aug_path, augmented)
            filenames.append(aug_filename)
        
        print(f"    Saved {len(filenames)} high-quality training images for rating {rating}")
        return filenames
    
    def apply_quality_preserving_augmentation(self, image, config):
        """
        Apply augmentation that preserves image quality and readability.
        """
        
        result = image.astype(np.float32)
        
        # Adjust brightness
        if 'brightness' in config:
            result = result * (1 + config['brightness'])
        
        # Adjust contrast
        if 'contrast' in config:
            mean_val = np.mean(result)
            result = (result - mean_val) * config['contrast'] + mean_val
        
        # Add minimal noise
        if 'noise' in config and config['noise'] > 0:
            noise = np.random.normal(0, config['noise'] * 255, result.shape)
            result = result + noise
        
        # Ensure valid range
        result = np.clip(result, 0, 255).astype(np.uint8)
        
        return result
    
    def extract_training_data_from_images(self, image_folder):
        """
        Main method to extract high-quality training data from a folder of form images.
        """
        
        print(f"Starting high-quality training data extraction from: {image_folder}")
        print(f"Target resolution: {self.target_size} (full row capture)")
        
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
        quality_issues = 0
        
        for image_path in image_files:
            print(f"\n{'='*50}")
            print(f"Processing: {os.path.basename(image_path)}")
            print(f"{'='*50}")
            
            try:
                # Preprocess image with quality preservation
                processed_image = self.preprocess_image_for_extraction(image_path)
                
                # Extract high-quality regions
                regions = self.get_improved_form_regions(processed_image)
                
                # Get ratings for each region
                extracted_count = 0
                for sample_name, sample_regions in regions.items():
                    for attribute, region_img in sample_regions.items():
                        total_regions += 1
                        
                        print(f"\nProcessing {sample_name} - {attribute}")
                        
                        rating = self.show_region_and_get_rating_improved(
                            region_img, sample_name, attribute, 
                            total_regions, len(image_files) * 20
                        )
                        
                        if rating == 'QUALITY_ISSUE':
                            quality_issues += 1
                            print(f"  Quality issue reported - skipping this region")
                        elif rating is not None:
                            filenames = self.save_high_quality_training_region(
                                region_img, rating, sample_name, attribute
                            )
                            
                            self.session_log.append({
                                'image': os.path.basename(image_path),
                                'sample': sample_name,
                                'attribute': attribute,
                                'rating': rating,
                                'filenames': filenames,
                                'region_size': self.target_size,
                                'quality': 'high_resolution'
                            })
                            
                            extracted_count += 1
                            print(f"  Saved rating {rating}")
                        else:
                            print(f"  Skipped")
                
                processed_forms += 1
                print(f"\nForm complete: {extracted_count}/20 regions extracted")
                
            except Exception as e:
                print(f"Error processing {os.path.basename(image_path)}: {e}")
                continue
        
        # Save session log
        self.save_session_log(processed_forms, total_regions, quality_issues)
        
        print(f"\n{'='*60}")
        print("HIGH-QUALITY TRAINING EXTRACTION COMPLETE")
        print(f"{'='*60}")
        print(f"Forms processed: {processed_forms}")
        print(f"Total regions extracted: {total_regions}")
        print(f"Quality issues reported: {quality_issues}")
        print(f"Debug images saved to: training_data/debug_regions/")
        print("Next steps:")
        print("1. Review debug images to verify extraction quality")
        print("2. Train your ML model with this high-quality data")
        print("3. Compare accuracy with previous low-resolution training")
    
    def create_training_data_structure(self):
        """Create training directory structure."""
        
        base_dir = "training_data/sensory_ratings"
        os.makedirs(base_dir, exist_ok=True)
        
        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            os.makedirs(rating_dir, exist_ok=True)
        
        # Additional directories for debugging and analysis
        os.makedirs("training_data/logs", exist_ok=True)
        os.makedirs("training_data/debug_regions", exist_ok=True)
        os.makedirs("models", exist_ok=True)
        
        print("High-quality training data structure created")
    
    def save_session_log(self, processed_forms, total_regions, quality_issues):
        """Save detailed session log."""
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"training_data/logs/high_quality_session_{timestamp}.json"
        
        session_data = {
            'extraction_timestamp': timestamp,
            'extraction_method': 'high_quality_regions',
            'target_resolution': f"{self.target_size[0]}x{self.target_size[1]}",
            'forms_processed': processed_forms,
            'total_regions_labeled': total_regions,
            'quality_issues_reported': quality_issues,
            'detailed_extractions': self.session_log,
            'notes': 'Enhanced extraction with readable high-resolution regions'
        }
        
        with open(log_filename, 'w') as f:
            json.dump(session_data, f, indent=2)
        
        print(f"Session log saved to {log_filename}")

# Usage example
if __name__ == "__main__":
    print("High-Quality Training Data Extractor")
    print("This creates readable, high-resolution training regions")
    
    # Example usage:
    # extractor = ImprovedTrainingExtractor()
    # extractor.extract_training_data_from_images('path/to/training/images')
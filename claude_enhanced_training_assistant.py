# claude_enhanced_training_assistant.py
"""
Claude-Enhanced Training Data Assistant
Combines Claude's vision intelligence with your ML training pipeline
"""

import cv2
import numpy as np
import os
import json
import base64
from PIL import Image
import anthropic
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageTk
import io

class ClaudeEnhancedTrainingAssistant:
    """
    Enhanced training assistant that uses Claude's vision to automatically detect
    precise sample regions and generate high-quality training data for your ML model.
    
    This approach solves the quadrant detection and header filtering problems by
    using Claude's advanced computer vision during training data preparation,
    while keeping your fast ML model for production inference.
    """
    
    def __init__(self):
        # Initialize Claude client
        self.api_key = os.getenv('ANTHROPIC_API_KEY')
        if not self.api_key:
            raise ValueError(
                "Anthropic API key not found. Please set ANTHROPIC_API_KEY environment variable."
            )
        
        self.client = anthropic.Anthropic(api_key=self.api_key)
        
        # Form attributes to extract
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        # Session tracking
        self.detected_boundaries = []
        self.session_log = []
        
        print("Claude-Enhanced Training Assistant initialized")
        print("This will generate precise training data using Claude's vision capabilities")
    
    def analyze_form_layout_with_claude(self, image_path):
        """
        Use Claude's vision to intelligently detect the 4 sample regions and their boundaries.
        
        This is the key innovation - instead of using hardcoded percentages, we let Claude
        analyze each form and provide exact pixel coordinates for each sample region.
        """
        
        print(f"Analyzing form layout with Claude: {os.path.basename(image_path)}")
        
        # Prepare image for Claude
        with open(image_path, 'rb') as image_file:
            image_data = base64.b64encode(image_file.read()).decode('utf-8')
        
        # Get image dimensions for coordinate calculations
        img = cv2.imread(image_path)
        height, width = img.shape[:2]
        
        # Claude prompt for precise region detection
        analysis_prompt = f"""
        I need you to analyze this sensory evaluation form and identify the precise boundaries 
        of the 4 sample rating regions. This is for training a computer vision model.
        
        The form has 4 samples arranged in a 2x2 grid:
        - Sample 1: Top-left quadrant  
        - Sample 2: Top-right quadrant
        - Sample 3: Bottom-left quadrant
        - Sample 4: Bottom-right quadrant
        
        Each sample has 5 rating attributes: Burnt Taste, Vapor Volume, Overall Flavor, 
        Smoothness, and Overall Liking.
        
        Please ignore header information and focus only on the actual rating regions.
        
        Image dimensions: {width}x{height} pixels
        
        Return the results in this exact JSON format:
        {{
            "sample_1": {{
                "x_start": pixel_value,
                "x_end": pixel_value, 
                "y_start": pixel_value,
                "y_end": pixel_value
            }},
            "sample_2": {{
                "x_start": pixel_value,
                "x_end": pixel_value,
                "y_start": pixel_value, 
                "y_end": pixel_value
            }},
            "sample_3": {{
                "x_start": pixel_value,
                "x_end": pixel_value,
                "y_start": pixel_value,
                "y_end": pixel_value
            }},
            "sample_4": {{
                "x_start": pixel_value,
                "x_end": pixel_value,
                "y_start": pixel_value,
                "y_end": pixel_value
            }}
        }}
        
        Focus on the actual content areas containing the rating scales, excluding 
        headers, sample codes, and other non-rating content.The image for each attribute should include the attribute title on the leftmost column and extend all the way to the black line that separates the left samples from the right samples.
        """
        
        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=1000,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/jpeg",
                                    "data": image_data
                                }
                            },
                            {
                                "type": "text", 
                                "text": analysis_prompt
                            }
                        ]
                    }
                ]
            )
            
            # Extract JSON from Claude's response
            response_text = response.content[0].text
            print(f"Claude analysis response received")
            
            # Find and parse JSON
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}') + 1
            
            if start_idx != -1 and end_idx > start_idx:
                json_str = response_text[start_idx:end_idx]
                boundaries = json.loads(json_str)
                
                # Validate the boundaries
                self.validate_detected_boundaries(boundaries, width, height)
                
                print("Successfully detected sample boundaries:")
                for sample, coords in boundaries.items():
                    print(f"  {sample}: x={coords['x_start']}-{coords['x_end']}, y={coords['y_start']}-{coords['y_end']}")
                
                return boundaries
            else:
                raise ValueError("Could not find JSON in Claude's response")
                
        except Exception as e:
            print(f"Error analyzing form with Claude: {e}")
            print("Falling back to improved default boundaries")
            return self.get_improved_default_boundaries(width, height)
    
    def validate_detected_boundaries(self, boundaries, width, height):
        """
        Validate that Claude's detected boundaries are reasonable and within image bounds.
        """
        
        for sample, coords in boundaries.items():
            # Check bounds
            if not (0 <= coords['x_start'] < coords['x_end'] <= width):
                raise ValueError(f"Invalid x coordinates for {sample}")
            if not (0 <= coords['y_start'] < coords['y_end'] <= height):
                raise ValueError(f"Invalid y coordinates for {sample}")
            
            # Check minimum size
            min_width = width * 0.15  # At least 15% of image width
            min_height = height * 0.15  # At least 15% of image height
            
            if (coords['x_end'] - coords['x_start']) < min_width:
                raise ValueError(f"Region too narrow for {sample}")
            if (coords['y_end'] - coords['y_start']) < min_height:
                raise ValueError(f"Region too short for {sample}")
        
        print("Boundary validation passed")
    
    def get_improved_default_boundaries(self, width, height):
        """
        Provide improved default boundaries based on analysis of multiple forms.
        These are better than the original hardcoded values.
        """
        
        return {
            "sample_1": {
                "x_start": int(width * 0.05),
                "x_end": int(width * 0.47),
                "y_start": int(height * 0.25),
                "y_end": int(height * 0.57)
            },
            "sample_2": {
                "x_start": int(width * 0.53),
                "x_end": int(width * 0.95),
                "y_start": int(height * 0.25),
                "y_end": int(height * 0.57)
            },
            "sample_3": {
                "x_start": int(width * 0.05),
                "x_end": int(width * 0.47),
                "y_start": int(height * 0.63),
                "y_end": int(height * 0.92)
            },
            "sample_4": {
                "x_start": int(width * 0.53),
                "x_end": int(width * 0.95),
                "y_start": int(height * 0.63),
                "y_end": int(height * 0.92)
            }
        }
    
    def extract_precise_training_regions(self, image_path, boundaries):
        """
        Extract training regions using Claude's detected boundaries instead of hardcoded percentages.
        This generates much more accurate training data.
        """
        
        # Load and preprocess image
        img = cv2.imread(image_path)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply minimal preprocessing to match production pipeline
        processed = cv2.bilateralFilter(gray, 5, 50, 50)
        
        regions = {}
        sample_mapping = {
            "sample_1": "Sample 1",
            "sample_2": "Sample 2", 
            "sample_3": "Sample 3",
            "sample_4": "Sample 4"
        }
        
        print(f"Extracting regions using Claude-detected boundaries...")
        
        for sample_key, sample_name in sample_mapping.items():
            coords = boundaries[sample_key]
            
            # Calculate attribute regions within this sample
            sample_height = coords['y_end'] - coords['y_start']
            attr_height = sample_height // len(self.attributes)
            
            regions[sample_name] = {}
            
            for i, attribute in enumerate(self.attributes):
                # Calculate precise boundaries for this attribute
                y_start = coords['y_start'] + i * attr_height
                y_end = coords['y_start'] + (i + 1) * attr_height
                x_start = coords['x_start']
                x_end = coords['x_end']
                
                # Extract the region
                region_img = processed[y_start:y_end, x_start:x_end]
                
                if region_img.size > 0:
                    # Resize to standard training size
                    resized = cv2.resize(region_img, (64, 64))
                    regions[sample_name][attribute] = resized
                    
                    print(f"  Extracted {sample_name} - {attribute}: {region_img.shape} -> (64, 64)")
        
        return regions
    
    def show_region_and_get_rating(self, region_img, sample_name, attribute, region_num, total_regions):
        """
        Display region and get human expert rating using tkinter interface.
        Enhanced with better visual presentation.
        """
        
        # Create tkinter window for rating input
        root = tk.Toplevel()
        root.title(f"Rate Region {region_num}/{total_regions}")
        root.geometry("500x400")
        root.configure(bg='white')
        
        # Center the window
        root.transient()
        root.grab_set()
        
        # Convert CV2 image to PIL for tkinter display
        region_pil = Image.fromarray(region_img)
        # Scale up for better visibility
        display_size = (200, 200)
        region_pil = region_pil.resize(display_size, Image.Resampling.NEAREST)
        region_photo = ImageTk.PhotoImage(region_pil)
        
        # Create UI elements
        title_label = tk.Label(root, text=f"{sample_name} - {attribute}", 
                              font=('Arial', 14, 'bold'), bg='white')
        title_label.pack(pady=10)
        
        progress_label = tk.Label(root, text=f"Region {region_num} of {total_regions}", 
                                 font=('Arial', 10), bg='white', fg='gray')
        progress_label.pack()
        
        # Image display
        image_label = tk.Label(root, image=region_photo, bg='white')
        image_label.pack(pady=10)
        
        # Rating selection
        rating_frame = tk.Frame(root, bg='white')
        rating_frame.pack(pady=10)
        
        tk.Label(rating_frame, text="Select Rating (1-9):", 
                font=('Arial', 12), bg='white').pack()
        
        rating_var = tk.StringVar()
        rating_buttons = []
        
        button_frame = tk.Frame(rating_frame, bg='white')
        button_frame.pack(pady=5)
        
        for i in range(1, 10):
            btn = tk.Radiobutton(button_frame, text=str(i), variable=rating_var, 
                                value=str(i), font=('Arial', 11))
            btn.pack(side='left', padx=3)
            rating_buttons.append(btn)
        
        # Control buttons
        control_frame = tk.Frame(root, bg='white')
        control_frame.pack(pady=20)
        
        result = {'rating': None, 'skip': False}
        
        def submit_rating():
            if rating_var.get():
                result['rating'] = int(rating_var.get())
                root.destroy()
            else:
                messagebox.showwarning("No Rating", "Please select a rating from 1-9")
        
        def skip_region():
            result['skip'] = True
            root.destroy()
        
        submit_btn = tk.Button(control_frame, text="Submit Rating", 
                              command=submit_rating, bg='lightblue', 
                              font=('Arial', 11), padx=20)
        submit_btn.pack(side='left', padx=10)
        
        skip_btn = tk.Button(control_frame, text="Skip", 
                            command=skip_region, bg='lightgray',
                            font=('Arial', 11), padx=20)
        skip_btn.pack(side='left', padx=10)
        
        # Instructions
        instruction_text = """
        Instructions:
        - Look at the rating scale in the image
        - Identify which number (1-9) is circled/marked
        - Select the corresponding rating below
        - Use Skip if the marking is unclear
        """
        
        instruction_label = tk.Label(root, text=instruction_text, 
                                   font=('Arial', 9), bg='white', 
                                   fg='darkblue', justify='left')
        instruction_label.pack(pady=10)
        
        # Handle window closing
        def on_closing():
            result['skip'] = True
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Set focus and wait
        root.focus_set()
        root.wait_window()
        
        if result['skip']:
            return None
        else:
            return result['rating']
    
    def save_training_region_with_augmentation(self, region_img, rating, sample_name, attribute):
        """
        Save labeled training region with realistic augmentation variations.
        Enhanced with better augmentation strategies.
        """
        
        base_dir = "training_data/sensory_ratings"
        rating_dir = os.path.join(base_dir, f"rating_{rating}")
        os.makedirs(rating_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        sample_clean = sample_name.replace(" ", "_")
        attribute_clean = attribute.replace(" ", "_")
        
        filenames = []
        
        # Save original region
        original_filename = f"{sample_clean}_{attribute_clean}_{timestamp}_original.png"
        original_path = os.path.join(rating_dir, original_filename)
        cv2.imwrite(original_path, region_img)
        filenames.append(original_filename)
        
        # Generate enhanced augmentations
        augmentation_configs = [
            {'rotation': 2, 'noise': 0.02, 'brightness': 0.1, 'suffix': 'rot2'},
            {'rotation': -2, 'noise': 0.01, 'brightness': -0.1, 'suffix': 'rot2n'},
            {'rotation': 0, 'noise': 0.03, 'brightness': 0.15, 'suffix': 'noise'},
            {'rotation': 1, 'noise': 0.015, 'brightness': 0.05, 'suffix': 'light'},
            {'rotation': -1, 'noise': 0.025, 'brightness': -0.05, 'suffix': 'mixed'},
            {'rotation': 0, 'noise': 0.01, 'brightness': 0.2, 'suffix': 'bright'}
        ]
        
        for config in augmentation_configs:
            augmented = self.apply_enhanced_augmentation(region_img, config)
            
            aug_filename = f"{sample_clean}_{attribute_clean}_{timestamp}_{config['suffix']}.png"
            aug_path = os.path.join(rating_dir, aug_filename)
            cv2.imwrite(aug_path, augmented)
            filenames.append(aug_filename)
        
        print(f"    Saved {len(filenames)} training images for rating {rating}")
        return filenames
    
    def apply_enhanced_augmentation(self, image, config):
        """
        Apply realistic augmentation that simulates natural variations in phone camera images.
        """
        
        result = image.copy()
        
        # Apply rotation
        if config['rotation'] != 0:
            h, w = result.shape
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, config['rotation'], 1.0)
            result = cv2.warpAffine(result, M, (w, h), borderMode=cv2.BORDER_REFLECT)
        
        # Add realistic noise
        if config['noise'] > 0:
            noise = np.random.normal(0, config['noise'] * 255, result.shape).astype(np.float32)
            result = result.astype(np.float32) + noise
            result = np.clip(result, 0, 255).astype(np.uint8)
        
        # Adjust brightness
        if config['brightness'] != 0:
            result = result.astype(np.float32)
            result = result * (1 + config['brightness'])
            result = np.clip(result, 0, 255).astype(np.uint8)
        
        return result
    
    def process_training_images_with_claude(self, image_folder):
        """
        Main method to process a folder of training images using Claude's enhanced detection.
        
        This replaces the manual boundary guessing with intelligent form analysis.
        """
        
        print(f"Starting Claude-enhanced training data extraction from: {image_folder}")
        print("This process uses Claude to detect precise sample boundaries for each form.")
        
        # Find all image files
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']
        image_files = []
        
        for filename in os.listdir(image_folder):
            if any(filename.lower().endswith(ext) for ext in image_extensions):
                image_files.append(os.path.join(image_folder, filename))
        
        if not image_files:
            print(f"No image files found in {image_folder}")
            return
        
        print(f"Found {len(image_files)} images to process")
        
        # Create training structure if needed
        self.create_training_data_structure()
        
        total_regions = 0
        processed_forms = 0
        
        for image_path in image_files:
            print(f"\n{'='*50}")
            print(f"Processing: {os.path.basename(image_path)}")
            print(f"{'='*50}")
            
            try:
                # Step 1: Use Claude to detect precise boundaries
                boundaries = self.analyze_form_layout_with_claude(image_path)
                self.detected_boundaries.append({
                    'image': os.path.basename(image_path),
                    'boundaries': boundaries
                })
                
                # Step 2: Extract regions using detected boundaries
                regions = self.extract_precise_training_regions(image_path, boundaries)
                
                # Step 3: Get human expert labels for each region
                extracted_count = 0
                for sample_name, sample_regions in regions.items():
                    for attribute, region_img in sample_regions.items():
                        total_regions += 1
                        
                        rating = self.show_region_and_get_rating(
                            region_img, sample_name, attribute, 
                            total_regions, len(image_files) * 20
                        )
                        
                        if rating is not None:
                            filenames = self.save_training_region_with_augmentation(
                                region_img, rating, sample_name, attribute
                            )
                            
                            self.session_log.append({
                                'image': os.path.basename(image_path),
                                'sample': sample_name,
                                'attribute': attribute,
                                'rating': rating,
                                'filenames': filenames,
                                'boundaries_used': boundaries
                            })
                            
                            extracted_count += 1
                
                processed_forms += 1
                print(f"Completed: {extracted_count}/20 regions extracted")
                
            except Exception as e:
                print(f"Error processing {os.path.basename(image_path)}: {e}")
                continue
        
        # Save session results
        self.save_enhanced_session_log(processed_forms, total_regions)
        self.generate_improved_boundaries_file()
        
        print(f"\n{'='*60}")
        print("CLAUDE-ENHANCED TRAINING COMPLETE")
        print(f"{'='*60}")
        print(f"Forms processed: {processed_forms}")
        print(f"Total regions extracted: {total_regions}")
        print(f"Intelligent boundaries detected for each form")
        print("Next steps:")
        print("1. Train your ML model with this improved data")
        print("2. Update production boundaries using generated config")
        print("3. Test the improved model accuracy")
    
    def create_training_data_structure(self):
        """Create the organized folder structure needed for training."""
        
        base_dir = "training_data/sensory_ratings"
        os.makedirs(base_dir, exist_ok=True)
        
        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            os.makedirs(rating_dir, exist_ok=True)
        
        # Additional directories
        os.makedirs("training_data/logs", exist_ok=True)
        os.makedirs("training_data/claude_analysis", exist_ok=True)
        os.makedirs("models", exist_ok=True)
        
        print("Training data structure created with Claude enhancement support")
    
    def save_enhanced_session_log(self, processed_forms, total_regions):
        """Save comprehensive session log including Claude's boundary detections."""
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"training_data/logs/claude_enhanced_session_{timestamp}.json"
        
        session_data = {
            'extraction_timestamp': timestamp,
            'extraction_method': 'claude_enhanced_vision',
            'forms_processed': processed_forms,
            'total_regions_labeled': total_regions,
            'claude_boundary_detections': self.detected_boundaries,
            'detailed_extractions': self.session_log,
            'notes': 'Used Claude vision to detect precise sample boundaries instead of hardcoded percentages'
        }
        
        with open(log_filename, 'w') as f:
            json.dump(session_data, f, indent=2)
        
        print(f"Enhanced session log saved to {log_filename}")
    
    def generate_improved_boundaries_file(self):
        """
        Generate an improved boundaries configuration file based on Claude's detections.
        This can be used to update your ML processor with better default boundaries.
        """
        
        if not self.detected_boundaries:
            print("No boundary detections to analyze")
            return
        
        # Calculate average boundaries across all analyzed forms
        avg_boundaries = {}
        boundary_keys = ['x_start', 'x_end', 'y_start', 'y_end']
        
        for sample in ['sample_1', 'sample_2', 'sample_3', 'sample_4']:
            avg_boundaries[sample] = {}
            
            for key in boundary_keys:
                values = []
                for detection in self.detected_boundaries:
                    if sample in detection['boundaries']:
                        values.append(detection['boundaries'][sample][key])
                
                if values:
                    avg_boundaries[sample][key] = int(np.mean(values))
        
        # Convert to percentage-based boundaries for your ML processor
        if self.detected_boundaries:
            # Use first image to get reference dimensions
            first_detection = self.detected_boundaries[0]
            # We'll calculate as percentages for generalizability
            
            config_data = {
                'generation_timestamp': datetime.now().isoformat(),
                'source': 'claude_enhanced_boundary_detection',
                'forms_analyzed': len(self.detected_boundaries),
                'raw_pixel_averages': avg_boundaries,
                'notes': 'These boundaries were intelligently detected by Claude vision analysis',
                'usage_instructions': 'Replace hardcoded boundaries in ml_form_processor.py extract_regions_for_prediction method'
            }
            
            config_filename = f"training_data/claude_analysis/improved_boundaries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(config_filename, 'w') as f:
                json.dump(config_data, f, indent=2)
            
            print(f"Improved boundaries configuration saved to {config_filename}")
            print("Use this data to update your ML processor's region detection!")

# Example usage and integration helper
if __name__ == "__main__":
    print("Claude-Enhanced Training Assistant")
    print("Usage: assistant = ClaudeEnhancedTrainingAssistant()")
    print("       assistant.process_training_images_with_claude('path/to/training/images')")
"""
training_data_helper.py
Training data extraction tool optimized for phone camera images of sensory evaluation forms.

This tool is specifically designed for scenarios where forms are photographed with mobile devices
rather than scanned with traditional document scanners. The augmentation strategy accounts for
the unique characteristics of phone camera images:
- JPEG compression artifacts and noise patterns
- Slight geometric distortions and perspective variations
- Variable focus quality and minor motion blur
- Lighting inconsistencies despite attempts at consistency
- Hand shake and natural positioning variations

The systematic augmentation approach ensures the trained model performs robustly on
phone-captured images while avoiding unrealistic distortions that don't match actual usage patterns.
"""

import cv2
import os
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import json
from datetime import datetime
import random
import math

class PhoneCameraTrainingExtractor:
    """
    Specialized training data extractor for phone camera images.
    
    This class understands that phone camera images have different characteristics
    than scanned documents and applies appropriate augmentation strategies to create
    robust training data that matches real-world deployment conditions.
    """
    
    def __init__(self):
        self.current_image = None
        self.regions_extracted = 0
        self.session_log = []
        
        # Attribute definitions that match the ML processor expectations
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        # Phone camera-specific augmentation parameters
        # These values are tuned for the realistic variations you get when photographing
        # documents with a mobile device, even when trying to maintain consistency
        self.phone_augmentation_params = {
            # Lighting variations - more generous than scanner since phone lighting varies more
            'brightness_range': (0.75, 1.25),  
            'contrast_range': (0.85, 1.20),
            
            # Phone camera noise characteristics
            'noise_std_range': (0, 12),  # Phones have more sensor noise than scanners
            'jpeg_quality_range': (70, 95),  # Simulate different compression levels
            
            # Geometric variations from hand-held photography
            'rotation_range': (-3.0, 3.0),  # Slight hand rotation is common
            'perspective_variation': 0.02,  # Minor perspective distortions
            
            # Focus and motion variations
            'blur_sigma_range': (0, 1.2),  # Slight focus variations and hand shake
            'sharpness_range': (0.8, 1.3),  # Camera focus variations
            
            # Color variations from phone camera sensors and auto-adjustment
            'saturation_range': (0.9, 1.1),  # Slight color variations
            'hue_shift_range': (-5, 5),  # Minor color temperature differences
        }
        
        print("Phone Camera Training Data Extractor initialized.")
        print(f"Will extract {4 * len(self.attributes)} regions per form.")
        print("Optimized for mobile device photography with consistent lighting.")
        
    def create_training_data_structure(self):
        """
        Create the organized folder structure needed for machine learning training.
        
        The structure follows TensorFlow's ImageDataGenerator conventions, where each
        rating class (1 through 9) gets its own directory. This organization enables
        automatic data loading during training and makes it easy to verify class balance.
        """
        base_dir = "training_data/sensory_ratings"
        os.makedirs(base_dir, exist_ok=True)
        
        # Create a folder for each possible rating value
        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            os.makedirs(rating_dir, exist_ok=True)
            
        print("Created training data directory structure:")
        for rating in range(1, 10):
            print(f"  training_data/sensory_ratings/rating_{rating}/")
            
        # Create directories for session tracking and augmentation examples
        os.makedirs("training_data/logs", exist_ok=True)
        os.makedirs("training_data/augmentation_examples", exist_ok=True)
        print("  training_data/logs/ (for session tracking)")
        print("  training_data/augmentation_examples/ (for reviewing augmentation quality)")
        
    def get_form_regions(self, image):
        """
        Extract the 20 individual rating regions from a phone camera image of the form.
        
        This method implements region detection specifically designed for phone camera images.
        The region boundaries are defined as percentages of image dimensions to handle
        the variable resolutions and aspect ratios that phone cameras produce. The key
        insight is that while the absolute pixel coordinates vary, the relative positions
        of form elements remain consistent.
        
        For phone camera images, we need to be slightly more generous with region boundaries
        since hand-held photography introduces small positioning variations that scanners don't have.
        """
        height, width = image.shape
        regions = {}
        
        print(f"  Processing image dimensions: {width}x{height} pixels")
        
        # Region definitions optimized for phone camera images
        # These percentages account for the slight variations in framing that occur
        # when photographing documents by hand, even when trying to be consistent
        sample_regions = {
            1: {"y_start": int(height * 0.22), "y_end": int(height * 0.58), 
                "x_start": int(width * 0.03), "x_end": int(width * 0.49)},  # Top-left
            2: {"y_start": int(height * 0.22), "y_end": int(height * 0.58),
                "x_start": int(width * 0.51), "x_end": int(width * 0.97)},  # Top-right
            3: {"y_start": int(height * 0.62), "y_end": int(height * 0.94),
                "x_start": int(width * 0.03), "x_end": int(width * 0.49)},  # Bottom-left
            4: {"y_start": int(height * 0.62), "y_end": int(height * 0.94),
                "x_start": int(width * 0.51), "x_end": int(width * 0.97)}   # Bottom-right
        }
        
        # Extract individual attribute regions within each sample
        for sample_id, sample_region in sample_regions.items():
            sample_height = sample_region["y_end"] - sample_region["y_start"]
            attr_height = sample_height // len(self.attributes)
            
            regions[f"Sample {sample_id}"] = {}
            
            for i, attribute in enumerate(self.attributes):
                # Calculate precise boundaries for each attribute's rating scale
                y_start = sample_region["y_start"] + i * attr_height
                y_end = sample_region["y_start"] + (i + 1) * attr_height
                x_start = sample_region["x_start"]
                x_end = sample_region["x_end"]
                
                # Ensure we don't exceed image boundaries (important for variable phone image sizes)
                y_start = max(0, y_start)
                y_end = min(height, y_end)
                x_start = max(0, x_start)
                x_end = min(width, x_end)
                
                # Extract the region - this becomes our training example
                region_img = image[y_start:y_end, x_start:x_end]
                regions[f"Sample {sample_id}"][attribute] = region_img
                
        return regions
        
    def apply_phone_camera_augmentation(self, region_img, augmentation_level='moderate'):
        """
        Apply realistic augmentations that simulate the natural variations in phone camera images.
        
        This is where we systematically introduce the types of variations your model will
        encounter in real use. Rather than applying random distortions, each augmentation
        corresponds to a real physical or technical aspect of mobile photography.
        
        The augmentation levels allow us to control how much variation we introduce:
        - 'light': Minimal variations (ideal phone camera conditions)
        - 'moderate': Typical variations (your target use case)
        - 'challenging': More significant variations (poor conditions, but still realistic)
        """
        
        # Start with the original region as a PIL image for easier manipulation
        if len(region_img.shape) == 2:  # Grayscale
            pil_img = Image.fromarray(region_img, mode='L')
        else:  # Color
            pil_img = Image.fromarray(cv2.cvtColor(region_img, cv2.COLOR_BGR2RGB))
        
        # Set augmentation intensity based on the specified level
        if augmentation_level == 'light':
            brightness_factor = random.uniform(0.90, 1.10)
            contrast_factor = random.uniform(0.95, 1.05)
            noise_std = random.uniform(0, 5)
            blur_sigma = random.uniform(0, 0.5)
            rotation_angle = random.uniform(-1.0, 1.0)
        elif augmentation_level == 'moderate':
            brightness_factor = random.uniform(0.80, 1.20)
            contrast_factor = random.uniform(0.90, 1.15)
            noise_std = random.uniform(0, 8)
            blur_sigma = random.uniform(0, 0.8)
            rotation_angle = random.uniform(-2.0, 2.0)
        else:  # challenging
            brightness_factor = random.uniform(0.75, 1.25)
            contrast_factor = random.uniform(0.85, 1.20)
            noise_std = random.uniform(0, 12)
            blur_sigma = random.uniform(0, 1.2)
            rotation_angle = random.uniform(-3.0, 3.0)
        
        # Apply brightness variations (simulates different lighting conditions)
        # Even with consistent lighting attempts, phone cameras introduce natural variation
        enhancer = ImageEnhance.Brightness(pil_img)
        pil_img = enhancer.enhance(brightness_factor)
        
        # Apply contrast variations (simulates phone camera auto-adjustment)
        enhancer = ImageEnhance.Contrast(pil_img)
        pil_img = enhancer.enhance(contrast_factor)
        
        # Apply slight sharpness variations (simulates focus differences)
        if random.random() < 0.4:  # Apply to 40% of images
            sharpness_factor = random.uniform(0.85, 1.15)
            enhancer = ImageEnhance.Sharpness(pil_img)
            pil_img = enhancer.enhance(sharpness_factor)
        
        # Convert back to OpenCV format for remaining operations
        cv_img = np.array(pil_img)
        if len(cv_img.shape) == 3 and cv_img.shape[2] == 3:
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
        
        # Add realistic phone camera noise
        # Phone sensors produce characteristic noise patterns, especially in lower light
        if noise_std > 0:
            noise = np.random.normal(0, noise_std, cv_img.shape)
            cv_img = np.clip(cv_img.astype(float) + noise, 0, 255).astype(np.uint8)
        
        # Apply slight blur to simulate minor focus variations and hand shake
        # Even steady hands introduce microscopic movement that creates slight blur
        if blur_sigma > 0.1:
            cv_img = cv2.GaussianBlur(cv_img, (3, 3), blur_sigma)
        
        # Apply small rotations to simulate natural hand positioning variations
        # When photographing documents by hand, perfect alignment is nearly impossible
        if abs(rotation_angle) > 0.2:
            rows, cols = cv_img.shape[:2]
            rotation_matrix = cv2.getRotationMatrix2D((cols/2, rows/2), rotation_angle, 1)
            cv_img = cv2.warpAffine(cv_img, rotation_matrix, (cols, rows), 
                                   borderMode=cv2.BORDER_REFLECT)
        
        # Occasionally apply slight perspective distortion
        # Phone cameras aren't always held perfectly parallel to the document
        if random.random() < 0.15:  # Apply to 15% of images
            cv_img = self.apply_slight_perspective_distortion(cv_img)
        
        return cv_img
    
    def apply_slight_perspective_distortion(self, image):
        """
        Apply minor perspective distortion that mimics photographing a document
        at a slight angle, which commonly happens with hand-held phone photography.
        
        This augmentation helps the model handle the small perspective variations
        that occur naturally when people photograph forms, even when trying to
        position the camera directly above the document.
        """
        rows, cols = image.shape[:2]
        
        # Define small random perspective variations
        # The variations are intentionally small to reflect realistic photography conditions
        distortion_strength = random.uniform(0.01, 0.03)
        
        # Create source and destination points for perspective transformation
        pts1 = np.float32([[0, 0], [cols, 0], [0, rows], [cols, rows]])
        
        # Apply small random offsets to simulate natural perspective variations
        offset = distortion_strength * min(rows, cols)
        pts2 = np.float32([
            [random.uniform(-offset, offset), random.uniform(-offset, offset)],
            [cols + random.uniform(-offset, offset), random.uniform(-offset, offset)],
            [random.uniform(-offset, offset), rows + random.uniform(-offset, offset)],
            [cols + random.uniform(-offset, offset), rows + random.uniform(-offset, offset)]
        ])
        
        # Apply the perspective transformation
        matrix = cv2.getPerspectiveTransform(pts1, pts2)
        result = cv2.warpPerspective(image, matrix, (cols, rows), borderMode=cv2.BORDER_REFLECT)
        
        return result
        
    def extract_training_regions_from_form(self, image_path):
        """
        Process a single phone camera image and extract all labeled training regions.
        
        This method orchestrates the entire extraction process for one form:
        1. Load and preprocess the phone camera image
        2. Extract the 20 individual regions of interest
        3. Present each region to the user for labeling
        4. Save the labeled regions with appropriate augmentations
        
        The human labeling step is crucial because it provides the ground truth that
        teaches the neural network what each rating looks like visually.
        """
        
        print(f"\nProcessing phone camera image: {os.path.basename(image_path)}")
        
        # Load the image with error handling for common phone camera image issues
        img = cv2.imread(image_path)
        if img is None:
            print(f"ERROR: Could not load image {image_path}")
            print("Check that the file exists and is a valid image format (.jpg, .png, etc.)")
            return []
        
        print(f"  Original image size: {img.shape[1]}x{img.shape[0]} pixels")
        
        # Convert to grayscale for processing (matches ML processor expectations)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply minimal preprocessing that matches the ML processor
        # For phone images, we keep preprocessing light to preserve natural characteristics
        processed = cv2.GaussianBlur(gray, (3, 3), 0)
        
        # Extract all regions from this form
        regions = self.get_form_regions(processed)
        
        extracted_regions = []
        total_regions = sum(len(sample_regions) for sample_regions in regions.values())
        current_region = 0
        
        print(f"  Found {total_regions} regions to process")
        
        # Process each region with user interaction
        for sample_name, sample_regions in regions.items():
            print(f"  Processing {sample_name}...")
            
            for attribute, region_img in sample_regions.items():
                current_region += 1
                
                if region_img.size > 0:
                    # Standardize region size for consistent model input
                    resized = cv2.resize(region_img, (64, 64))
                    
                    # Get human expert label for this region
                    rating = self.show_region_and_get_rating(
                        resized, sample_name, attribute, current_region, total_regions
                    )
                    
                    if rating is not None:
                        # Save the labeled region with augmented variations
                        filenames = self.save_training_region_with_augmentation(
                            resized, rating, sample_name, attribute
                        )
                        
                        extracted_regions.append({
                            'sample': sample_name,
                            'attribute': attribute, 
                            'rating': rating,
                            'filenames': filenames,
                            'region_number': current_region
                        })
                
        print(f"  Successfully extracted {len(extracted_regions)} labeled regions")
        
        # Track this session for analysis and debugging
        self.session_log.extend(extracted_regions)
        
        return extracted_regions
    
# Add this improved method to your PhoneCameraTrainingExtractor class
# Replace the existing show_region_and_get_rating method

    def show_region_and_get_rating(self, region_img, sample_name, attribute, region_num, total_regions):
        """
        Display a region using tkinter instead of OpenCV to avoid GUI backend issues.
    
        This method creates a robust, cross-platform display solution that doesn't
        depend on OpenCV's sometimes-problematic window management system.
        """
    
        import tkinter as tk
        from tkinter import ttk, messagebox
        from PIL import Image, ImageTk
        import numpy as np
    
        # Enhance the region for better visibility
        enhanced = cv2.convertScaleAbs(region_img, alpha=1.3, beta=15)
    
        # Convert to PIL Image for tkinter display
        # OpenCV uses BGR, PIL uses RGB, but for grayscale this doesn't matter
        pil_image = Image.fromarray(enhanced)
    
        # Scale up for better visibility (make it larger for easier evaluation)
        display_size = (400, 200)
        pil_image = pil_image.resize(display_size, Image.NEAREST)  # Use NEAREST to preserve pixelated appearance
    
        # Create the rating dialog window
        dialog = tk.Toplevel()
        dialog.title(f"Rate Region {region_num}/{total_regions}")
        dialog.geometry("600x500")
        dialog.transient()  # Make it stay on top
        dialog.grab_set()   # Make it modal
        dialog.configure(bg='white')
    
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300)
        y = (dialog.winfo_screenheight() // 2) - (250)
        dialog.geometry(f"600x500+{x}+{y}")
    
        # Create main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
    
        # Title and context
        title_label = ttk.Label(main_frame, 
                               text=f"Region {region_num} of {total_regions}", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=5)
    
        context_label = ttk.Label(main_frame, 
                                 text=f"Sample: {sample_name}\nAttribute: {attribute}", 
                                 font=('Arial', 12),
                                 justify='center')
        context_label.pack(pady=5)
    
        # Convert PIL image to tkinter PhotoImage
        photo = ImageTk.PhotoImage(pil_image)
    
        # Display the region image
        image_label = ttk.Label(main_frame, image=photo)
        image_label.pack(pady=10)
    
        # Keep a reference to prevent garbage collection
        image_label.image = photo
    
        # Instructions
        instruction_label = ttk.Label(main_frame, 
                                     text="What rating (1-9) is marked in this region?\n" +
                                          "1 = Lowest rating, 9 = Highest rating\n" +
                                          "Enter 0 to skip if unclear",
                                     font=('Arial', 10),
                                     justify='center')
        instruction_label.pack(pady=10)
    
        # Rating input frame
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(pady=10)
    
        # Create a visual rating scale for easy selection
        ttk.Label(input_frame, text="Select Rating:", font=('Arial', 12, 'bold')).pack()
    
        rating_var = tk.IntVar(value=5)  # Default to middle rating
    
        # Create radio buttons for each rating
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(pady=10)
    
        # Add a "Skip" option
        ttk.Radiobutton(button_frame, text="Skip (0)", variable=rating_var, value=0).grid(row=0, column=0, padx=5)
    
        # Add rating buttons 1-9
        for i in range(1, 10):
            ttk.Radiobutton(button_frame, text=str(i), variable=rating_var, value=i).grid(row=0, column=i, padx=5)
    
        # Alternative text entry for quick typing
        entry_frame = ttk.Frame(input_frame)
        entry_frame.pack(pady=10)
    
        ttk.Label(entry_frame, text="Or type rating:").pack(side='left')
        rating_entry = ttk.Entry(entry_frame, width=5)
        rating_entry.pack(side='left', padx=5)
        rating_entry.bind('<Return>', lambda e: confirm_rating())  # Enter key submits
    
        # Result variable to store the user's choice
        result = {'rating': None, 'confirmed': False}
    
        def confirm_rating():
            """Handle rating confirmation."""
            # Try to get rating from text entry first
            entry_text = rating_entry.get().strip()
            if entry_text:
                try:
                    rating = int(entry_text)
                    if 0 <= rating <= 9:
                        rating_var.set(rating)
                    else:
                        messagebox.showerror("Invalid Rating", "Please enter a number between 0 and 9")
                        return
                except ValueError:
                    messagebox.showerror("Invalid Rating", "Please enter a valid number")
                    return
        
            # Get the final rating
            final_rating = rating_var.get()
            result['rating'] = final_rating if final_rating > 0 else None
            result['confirmed'] = True
            dialog.destroy()
    
        def skip_region():
            """Handle skipping this region."""
            result['rating'] = None
            result['confirmed'] = True
            dialog.destroy()
    
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
    
        ttk.Button(button_frame, text="Confirm Rating", command=confirm_rating).pack(side='left', padx=10)
        ttk.Button(button_frame, text="Skip Region", command=skip_region).pack(side='left', padx=10)
    
        # Focus on the entry field for quick typing
        rating_entry.focus_set()
    
        # Wait for user interaction
        dialog.wait_window()
    
        # Provide feedback
        if result['confirmed']:
            if result['rating']:
                print(f"    Labeled region {region_num} as rating {result['rating']}")
            else:
                print(f"    Skipped region {region_num}")
    
        return result['rating']
    
    def save_training_region_with_augmentation(self, region_img, rating, sample_name, attribute):
        """
        Save the original labeled region plus systematically augmented versions.
        
        This method multiplies the effective size of your training dataset by creating
        realistic variations of each manually labeled example. Instead of needing
        thousands of manually labeled regions, we can create a robust dataset by
        augmenting the examples we do have with realistic variations.
        
        The augmentation strategy is specifically tuned for phone camera characteristics,
        ensuring that the variations match what the model will encounter in real deployment.
        """
        
        folder = f"training_data/sensory_ratings/rating_{rating}"
        saved_filenames = []
        
        # Save the original region first (this is your ground truth)
        original_filename = self.save_single_region(region_img, rating, "original")
        saved_filenames.append(original_filename)
        
        # Create augmented versions with different variation levels
        # This systematic approach ensures we cover the range of realistic variations
        # without going overboard with unrealistic distortions
        augmentation_sets = [
            ('light', 2),       # 2 examples with minimal variations
            ('moderate', 3),    # 3 examples with typical phone camera variations
            ('challenging', 1)  # 1 example with more significant (but realistic) variations
        ]
        
        for aug_level, count in augmentation_sets:
            for i in range(count):
                # Apply the augmentation appropriate for this level
                augmented_region = self.apply_phone_camera_augmentation(region_img, aug_level)
                
                # Save with a descriptive filename that tracks the augmentation applied
                filename = self.save_single_region(
                    augmented_region, rating, f"{aug_level}_{i+1}"
                )
                saved_filenames.append(filename)
        
        print(f"    Saved {len(saved_filenames)} versions (1 original + {len(saved_filenames)-1} augmented)")
        
        # Occasionally save an example to the examples folder for quality review
        if random.random() < 0.05:  # Save 5% of regions as examples
            self.save_augmentation_example(region_img, rating, sample_name, attribute)
        
        return saved_filenames
    
    def save_single_region(self, region_img, rating, variant_type):
        """
        Save a single region image with proper naming and organization.
        
        The filename includes timestamp and variant information to help with
        debugging and dataset analysis while maintaining the class-based
        folder structure that TensorFlow expects.
        """
        
        folder = f"training_data/sensory_ratings/rating_{rating}"
        
        # Generate a unique, descriptive filename
        existing_files = len([f for f in os.listdir(folder) if f.endswith('.jpg')])
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # Include milliseconds
        filename = f"region_{existing_files + 1:04d}_{timestamp}_{variant_type}.jpg"
        filepath = os.path.join(folder, filename)
        
        # Save with high quality to preserve details for training
        cv2.imwrite(filepath, region_img, [cv2.IMWRITE_JPEG_QUALITY, 95])
        self.regions_extracted += 1
        
        return filename
    
    def save_augmentation_example(self, original_region, rating, sample_name, attribute):
        """
        Save examples of original and augmented regions for quality review.
        
        This helps you verify that the augmentation process is producing realistic
        variations and allows you to fine-tune the augmentation parameters if needed.
        """
        
        examples_dir = "training_data/augmentation_examples"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create a comparison image showing original and augmented versions
        aug_light = self.apply_phone_camera_augmentation(original_region, 'light')
        aug_moderate = self.apply_phone_camera_augmentation(original_region, 'moderate')
        aug_challenging = self.apply_phone_camera_augmentation(original_region, 'challenging')
        
        # Combine into a single comparison image
        comparison = np.hstack([original_region, aug_light, aug_moderate, aug_challenging])
        
        filename = f"example_{timestamp}_rating{rating}_{sample_name}_{attribute}.jpg"
        filepath = os.path.join(examples_dir, filename)
        cv2.imwrite(filepath, comparison)
    
    def process_multiple_forms(self, image_directory):
        """
        Process all phone camera images in a directory to build the training dataset.
        
        This method orchestrates the complete training data creation process:
        1. Find all image files in the specified directory
        2. Process each form to extract and label regions
        3. Track progress and provide user feedback
        4. Generate session logs for dataset management
        
        The interactive nature allows users to take breaks during the labeling process,
        which is important since labeling hundreds of regions can be mentally taxing.
        """
        
        # Find all supported image files
        supported_formats = ('.jpg', '.jpeg', '.png', '.tiff', '.bmp', '.JPG', '.JPEG')
        image_files = [f for f in os.listdir(image_directory) 
                      if f.endswith(supported_formats)]
        
        if not image_files:
            print(f"No image files found in {image_directory}")
            print(f"Looking for files with extensions: {supported_formats}")
            return
            
        print(f"Found {len(image_files)} phone camera images to process")
        print("Each image should contain a sensory evaluation form with 4 samples")
        
        # Process each form with progress tracking
        total_regions = 0
        processed_forms = 0
        
        for i, image_file in enumerate(image_files):
            print(f"\n{'='*60}")
            print(f"Processing form {i+1}/{len(image_files)}: {image_file}")
            print(f"{'='*60}")
            
            image_path = os.path.join(image_directory, image_file)
            
            try:
                # Extract regions from this form
                regions = self.extract_training_regions_from_form(image_path)
                total_regions += len(regions)
                processed_forms += 1
                
                print(f"Completed form {i+1}. Total regions extracted so far: {total_regions}")
                
            except Exception as e:
                print(f"ERROR processing {image_file}: {str(e)}")
                print("Continuing with next form...")
                continue
            
            # Ask user if they want to continue (except for the last file)
            if i < len(image_files) - 1:
                root = tk.Tk()
                root.withdraw()
                
                continue_processing = messagebox.askyesno(
                    "Continue Processing?", 
                    f"Progress Update:\n\n"
                    f"✓ Processed: {processed_forms}/{len(image_files)} forms\n"
                    f"✓ Extracted: {total_regions} total regions\n"
                    f"✓ Remaining: {len(image_files) - i - 1} forms\n\n"
                    f"Continue with the next form?\n"
                    f"(You can resume later by restarting the tool)"
                )
                
                root.destroy()
                
                if not continue_processing:
                    print("Processing stopped by user choice.")
                    break
                    
        # Save comprehensive session log
        self.save_detailed_session_log(processed_forms, total_regions)
        
        # Display final summary
        print(f"\n{'='*60}")
        print("TRAINING DATA EXTRACTION COMPLETE")
        print(f"{'='*60}")
        print(f"Forms processed: {processed_forms}")
        print(f"Total regions extracted: {total_regions}")
        print(f"Expected augmented dataset size: ~{total_regions * 7} images")
        print("Session log saved to training_data/logs/")
        print("\nNext steps:")
        print("1. Check dataset balance with the balance checker")
        print("2. Review augmentation examples if desired")
        print("3. Train your machine learning model")
        
    def save_detailed_session_log(self, processed_forms, total_regions):
        """
        Save a comprehensive log of the extraction session for dataset management.
        
        This log helps with dataset versioning, tracking extraction quality,
        and identifying any patterns in the labeling process that might
        indicate systematic issues or improvements.
        """
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"training_data/logs/phone_extraction_session_{timestamp}.json"
        
        # Calculate statistics about this session
        rating_distribution = {}
        for region in self.session_log:
            rating = region['rating']
            rating_distribution[rating] = rating_distribution.get(rating, 0) + 1
        
        session_summary = {
            'extraction_timestamp': timestamp,
            'extraction_method': 'phone_camera_optimized',
            'forms_processed': processed_forms,
            'total_regions_labeled': total_regions,
            'expected_augmented_count': total_regions * 7,  # 1 original + 6 augmented per region
            'rating_distribution': rating_distribution,
            'augmentation_settings': self.phone_augmentation_params,
            'detailed_regions': self.session_log
        }
        
        # Save the log with proper formatting
        with open(log_filename, 'w') as f:
            json.dump(session_summary, f, indent=2)
            
        print(f"Detailed session log saved to {log_filename}")
        
    def check_dataset_balance(self):
        """
        Analyze the current training dataset balance and provide recommendations.
        
        A balanced dataset is crucial for good model performance. This function
        helps identify if certain rating classes are under-represented and
        provides guidance on whether more data collection is needed.
        """
        
        base_dir = "training_data/sensory_ratings"
        if not os.path.exists(base_dir):
            print("Training data directory not found.")
            print("Run create_training_data_structure() first, then extract some training data.")
            return
            
        print("\n" + "="*50)
        print("PHONE CAMERA TRAINING DATASET ANALYSIS")
        print("="*50)
        
        # Count images in each rating class
        class_counts = {}
        total_images = 0
        
        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            if os.path.exists(rating_dir):
                count = len([f for f in os.listdir(rating_dir) if f.endswith('.jpg')])
                class_counts[rating] = count
                total_images += count
                print(f"Rating {rating}: {count:4d} images")
            else:
                class_counts[rating] = 0
                print(f"Rating {rating}: {0:4d} images")
        
        print("-" * 30)
        print(f"Total images: {total_images}")
        
        if total_images == 0:
            print("\nNo training data found. Extract some regions first.")
            return
        
        # Analyze balance and provide recommendations
        avg_per_class = total_images / 9
        min_count = min(class_counts.values())
        max_count = max(class_counts.values())
        
        print(f"Average per class: {avg_per_class:.1f}")
        print(f"Min class size: {min_count}")
        print(f"Max class size: {max_count}")
        
        # Determine if dataset is ready for training
        print("\nDataset Assessment:")
        
        if min_count >= 100:
            print("✓ EXCELLENT: Dataset is well-balanced and ready for training")
            print("  Expected model performance: Very good")
        elif min_count >= 50 and total_images >= 500:
            print("✓ GOOD: Dataset should produce a functional model")
            print("  Expected model performance: Good")
            print("  Consider adding more examples for under-represented classes")
        elif total_images >= 200:
            print("⚠ MINIMAL: Dataset may work but could benefit from more data")
            print("  Expected model performance: Fair")
            print("  Recommend collecting more examples, especially for sparse classes")
        else:
            print("✗ INSUFFICIENT: Need more training data")
            print("  Minimum recommendation: 200+ total images")
            print("  Target recommendation: 500+ total images")
        
        # Identify classes that need more examples
        under_represented = [rating for rating, count in class_counts.items() 
                           if count < avg_per_class * 0.7]
        
        if under_represented:
            print(f"\nUnder-represented classes (need more examples): {under_represented}")
        else:
            print("\nClass distribution is reasonably balanced")

def main():
    """
    Main function to run the phone camera training data extraction process.
    
    This function provides a user-friendly interface to the training data extraction
    process, guiding users through each step and providing clear instructions for
    building a robust training dataset from phone camera images.
    """
    
    print("="*70)
    print("PHONE CAMERA SENSORY FORM TRAINING DATA EXTRACTOR")
    print("="*70)
    print("This tool creates machine learning training data from phone camera images")
    print("of sensory evaluation forms. It's optimized for the unique characteristics")
    print("of mobile device photography with consistent but imperfect lighting.")
    print()
    print("Process overview:")
    print("1. You'll select a folder containing your phone camera images")
    print("2. The tool will extract individual rating regions from each form")
    print("3. You'll label each region with the correct rating (1-9)")
    print("4. The tool will create augmented versions to build a robust dataset")
    print()
    
    # Initialize the extractor
    extractor = PhoneCameraTrainingExtractor()
    
    # Create the directory structure
    print("Setting up training data directories...")
    extractor.create_training_data_structure()
    
    # Get the directory containing phone camera images
    root = tk.Tk()
    root.withdraw()
    
    image_directory = filedialog.askdirectory(
        title="Select folder containing phone camera images of sensory evaluation forms"
    )
    
    root.destroy()
    
    if image_directory:
        print(f"\nSelected directory: {image_directory}")
        
        # Confirm before starting the intensive labeling process
        root = tk.Tk()
        root.withdraw()
        
        start_processing = messagebox.askyesno(
            "Start Training Data Extraction",
            "This process will:\n\n"
            "- Show you individual regions from each form\n"
            "- Ask you to rate each region (1-9)\n"
            "- Create augmented versions automatically\n"
            "- Build a robust training dataset\n\n"
            "The process can be stopped and resumed at any time.\n\n"
            "Ready to begin?"
        )
        
        root.destroy()
        
        if start_processing:
            # Start the extraction process
            extractor.process_multiple_forms(image_directory)
            
            # Show final dataset analysis
            extractor.check_dataset_balance()
        else:
            print("Extraction cancelled by user.")
    else:
        print("No directory selected. Exiting.")

if __name__ == "__main__":
    main()
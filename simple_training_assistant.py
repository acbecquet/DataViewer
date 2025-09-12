# simple_training_assistant.py
"""
Simple Training Data Assistant
Uses your ImprovedAttributeDetectionExtractor directly to handle all attribute separation
"""

import cv2
import numpy as np
import os
import json
from PIL import Image
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageTk

# Import your existing extractor class
from updated_training_extractor import ImprovedAttributeDetectionExtractor

class SimpleTrainingAssistant:
    """
    Simple training assistant that uses your ImprovedAttributeDetectionExtractor
    to handle all the attribute separation automatically.
    """

    def __init__(self):
        print("DEBUG: Initializing SimpleTrainingAssistant")

        # Initialize your extractor - let it handle everything
        try:
            self.extractor = ImprovedAttributeDetectionExtractor()
            print("DEBUG: Successfully initialized ImprovedAttributeDetectionExtractor")
        except Exception as e:
            print(f"ERROR: Failed to initialize ImprovedAttributeDetectionExtractor: {e}")
            raise

        # Session tracking
        self.session_log = []

        print("DEBUG: SimpleTrainingAssistant initialized")
        print("DEBUG: Will use your extractor to handle all attribute separation")

    def show_region_and_get_rating(self, region_img, sample_name, attribute, region_num, total_regions):
        """
        Display region and get human expert rating using tkinter interface.
        """

        print(f"DEBUG: Showing region {region_num}/{total_regions}: {sample_name} - {attribute}")

        # Create tkinter window for rating input
        root = tk.Toplevel()
        root.title(f"Rate Region {region_num}/{total_regions}")
        root.geometry("700x600")
        root.configure(bg='white')

        # Center the window
        root.transient()
        root.grab_set()

        # Convert CV2 image to PIL for tkinter display
        region_pil = Image.fromarray(region_img)
        # Scale up for better visibility while maintaining aspect ratio
        original_width, original_height = region_pil.size
        scale_factor = min(500/original_width, 300/original_height)
        new_width = int(original_width * scale_factor)
        new_height = int(original_height * scale_factor)
        region_pil = region_pil.resize((new_width, new_height), Image.Resampling.NEAREST)
        region_photo = ImageTk.PhotoImage(region_pil)

        # Create UI elements
        title_label = tk.Label(root, text=f"{sample_name} - {attribute}",
                              font=('Arial', 16, 'bold'), bg='white')
        title_label.pack(pady=10)

        progress_label = tk.Label(root, text=f"Region {region_num} of {total_regions}",
                                 font=('Arial', 12), bg='white', fg='gray')
        progress_label.pack()

        # Debug info
        debug_label = tk.Label(root, text=f"Original: {original_width}x{original_height}, Display: {new_width}x{new_height}",
                              font=('Arial', 9), bg='white', fg='blue')
        debug_label.pack()

        # Image display with border
        image_frame = tk.Frame(root, bg='black', padx=2, pady=2)
        image_frame.pack(pady=15)

        image_label = tk.Label(image_frame, image=region_photo, bg='white')
        image_label.pack()

        # Rating selection
        rating_frame = tk.Frame(root, bg='white')
        rating_frame.pack(pady=20)

        tk.Label(rating_frame, text="Select Rating (1-9):",
                font=('Arial', 14, 'bold'), bg='white').pack()

        rating_var = tk.StringVar()

        button_frame = tk.Frame(rating_frame, bg='white')
        button_frame.pack(pady=10)

        for i in range(1, 10):
            btn = tk.Radiobutton(button_frame, text=str(i), variable=rating_var,
                                value=str(i), font=('Arial', 12, 'bold'),
                                bg='lightblue', selectcolor='yellow')
            btn.pack(side='left', padx=5)

        # Control buttons
        control_frame = tk.Frame(root, bg='white')
        control_frame.pack(pady=20)

        result = {'rating': None, 'skip': False}

        def submit_rating():
            if rating_var.get():
                result['rating'] = int(rating_var.get())
                print(f"DEBUG: User selected rating {result['rating']} for {sample_name} - {attribute}")
                root.destroy()
            else:
                messagebox.showwarning("No Rating", "Please select a rating from 1-9")

        def skip_region():
            result['skip'] = True
            print(f"DEBUG: User skipped {sample_name} - {attribute}")
            root.destroy()

        submit_btn = tk.Button(control_frame, text="Submit Rating",
                              command=submit_rating, bg='lightgreen',
                              font=('Arial', 12, 'bold'), padx=25, pady=5)
        submit_btn.pack(side='left', padx=15)

        skip_btn = tk.Button(control_frame, text="Skip Region",
                            command=skip_region, bg='lightcoral',
                            font=('Arial', 12, 'bold'), padx=25, pady=5)
        skip_btn.pack(side='left', padx=15)

        # Instructions
        instruction_text = """
        Instructions:
        - Look at the rating scale in the image carefully
        - Identify which number (1-9) is circled/marked/selected
        - Select the corresponding rating using the buttons above
        - Use 'Skip Region' if the marking is unclear or damaged
        """

        instruction_label = tk.Label(root, text=instruction_text,
                                   font=('Arial', 10), bg='white',
                                   fg='darkblue', justify='left')
        instruction_label.pack(pady=15)

        # Handle window closing
        def on_closing():
            result['skip'] = True
            print(f"DEBUG: Window closed, skipping {sample_name} - {attribute}")
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
        """

        print(f"DEBUG: Saving training data for rating {rating}: {sample_name} - {attribute}")

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
        print(f"DEBUG:   Saved original: {original_filename}")

        # Generate enhanced augmentations
        augmentation_configs = [
            {'rotation': 2, 'noise': 0.02, 'brightness': 0.1, 'suffix': 'rot2'},
            {'rotation': -2, 'noise': 0.01, 'brightness': -0.1, 'suffix': 'rot2n'},
            {'rotation': 0, 'noise': 0.03, 'brightness': 0.15, 'suffix': 'noise'},
            {'rotation': 1, 'noise': 0.015, 'brightness': 0.05, 'suffix': 'light'},
            {'rotation': -1, 'noise': 0.025, 'brightness': -0.05, 'suffix': 'mixed'},
            {'rotation': 0, 'noise': 0.01, 'brightness': 0.2, 'suffix': 'bright'},
            {'rotation': 0.5, 'noise': 0.02, 'brightness': -0.15, 'suffix': 'dim'},
            {'rotation': -0.5, 'noise': 0.018, 'brightness': 0.08, 'suffix': 'subtle'}
        ]

        for config in augmentation_configs:
            try:
                augmented = self.apply_augmentation(region_img, config)

                aug_filename = f"{sample_clean}_{attribute_clean}_{timestamp}_{config['suffix']}.png"
                aug_path = os.path.join(rating_dir, aug_filename)
                cv2.imwrite(aug_path, augmented)
                filenames.append(aug_filename)

            except Exception as e:
                print(f"DEBUG:   Failed to create augmentation {config['suffix']}: {e}")

        print(f"DEBUG:   Saved {len(filenames)} training images for rating {rating}")
        return filenames

    def process_training_images(self, image_folder):
        """
        Main method to process a folder of training images.
        Just loads images and passes them to your extractor!
        """

        print(f"DEBUG: Starting training data extraction from: {image_folder}")
        print("DEBUG: Using your ImprovedAttributeDetectionExtractor for all processing")

        # Find all image files
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']
        image_files = []

        for filename in os.listdir(image_folder):
            if any(filename.lower().endswith(ext) for ext in image_extensions):
                image_files.append(os.path.join(image_folder, filename))

        if not image_files:
            print(f"ERROR: No image files found in {image_folder}")
            return

        print(f"DEBUG: Found {len(image_files)} images to process")

        # Create training structure
        self.create_training_data_structure()

        total_regions_processed = 0
        total_regions_labeled = 0
        processed_forms = 0

        for image_path in image_files:
            print(f"\n{'='*60}")
            print(f"DEBUG: Processing: {os.path.basename(image_path)}")
            print(f"{'='*60}")

            try:
                # Load the image
                image = cv2.imread(image_path)
                image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                if image is None:
                    print(f"ERROR: Could not load image: {image_path}")
                    continue

                print(f"DEBUG: Loaded image: {image.shape}")

                # Pass directly to your extractor - let it handle everything!
                regions = self.extractor.get_improved_form_regions(image)
                print(f"DEBUG: Your extractor returned {len(regions)} sample regions")

                # Process each region returned by your extractor
                extracted_count = 0
                for sample_name, sample_attributes in regions.items():
                    print(f"DEBUG: Processing {sample_name} with {len(sample_attributes)} attributes")

                    for attribute, region_img in sample_attributes.items():
                        total_regions_processed += 1

                        # Show the region and get rating
                        rating = self.show_region_and_get_rating(
                            region_img, sample_name, attribute,
                            total_regions_processed, len(image_files) * 20
                        )

                        if rating is not None:
                            # Save the training data
                            filenames = self.save_training_region_with_augmentation(
                                region_img, rating, sample_name, attribute
                            )

                            # Log the extraction
                            self.session_log.append({
                                'image': os.path.basename(image_path),
                                'sample': sample_name,
                                'attribute': attribute,
                                'rating': rating,
                                'filenames': filenames,
                                'region_shape': region_img.shape
                            })

                            extracted_count += 1
                            total_regions_labeled += 1
                            print(f"DEBUG: Successfully labeled region {total_regions_processed}")
                        else:
                            print(f"DEBUG: Skipped region {total_regions_processed}")

                processed_forms += 1
                print(f"DEBUG: Completed form: {extracted_count} regions labeled")

            except Exception as e:
                print(f"ERROR: Error processing {os.path.basename(image_path)}: {e}")
                import traceback
                print(f"DEBUG: Full traceback: {traceback.format_exc()}")
                continue

        # Save session results
        self.save_session_log(processed_forms, total_regions_processed, total_regions_labeled)

        print(f"\n{'='*70}")
        print("TRAINING DATA EXTRACTION COMPLETE")
        print(f"{'='*70}")
        print(f"DEBUG: Forms processed: {processed_forms}")
        print(f"DEBUG: Total regions processed: {total_regions_processed}")
        print(f"DEBUG: Total regions labeled: {total_regions_labeled}")
        if total_regions_processed > 0:
            print(f"DEBUG: Success rate: {(total_regions_labeled/total_regions_processed)*100:.1f}%")
        print("DEBUG: Your ImprovedAttributeDetectionExtractor handled all the attribute separation!")

    def create_training_data_structure(self):
        """Create the organized folder structure needed for training."""

        print("DEBUG: Creating training data structure")

        base_dir = "training_data/sensory_ratings"
        os.makedirs(base_dir, exist_ok=True)

        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            os.makedirs(rating_dir, exist_ok=True)

        # Additional directories
        os.makedirs("training_data/logs", exist_ok=True)
        os.makedirs("models", exist_ok=True)

        print("DEBUG: Training data structure created")

    def save_session_log(self, processed_forms, total_processed, total_labeled):
        """Save session log."""

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"training_data/logs/training_session_{timestamp}.json"

        session_data = {
            'extraction_timestamp': timestamp,
            'extraction_method': 'ImprovedAttributeDetectionExtractor',
            'forms_processed': processed_forms,
            'total_regions_processed': total_processed,
            'total_regions_labeled': total_labeled,
            'success_rate_percent': (total_labeled/total_processed)*100 if total_processed > 0 else 0,
            'detailed_extractions': self.session_log
        }

        with open(log_filename, 'w') as f:
            json.dump(session_data, f, indent=2)

        print(f"DEBUG: Session log saved to {log_filename}")

    def apply_augmentation(self, image, config):
            """
            Apply augmentation based on configuration.
            """
            augmented = image.copy().astype(np.float32)

            # Apply rotation
            if config.get('rotation', 0) != 0:
                angle = config['rotation']
                h, w = augmented.shape
                center = (w // 2, h // 2)
                rotation_matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
                augmented = cv2.warpAffine(augmented, rotation_matrix, (w, h),
                                         borderMode=cv2.BORDER_REFLECT)

            # Apply noise
            if config.get('noise', 0) > 0:
                noise_level = config['noise']
                noise = np.random.normal(0, noise_level * 255, augmented.shape)
                augmented = augmented + noise
                augmented = np.clip(augmented, 0, 255)

            # Apply brightness adjustment
            if config.get('brightness', 0) != 0:
                brightness_factor = 1.0 + config['brightness']
                augmented = augmented * brightness_factor
                augmented = np.clip(augmented, 0, 255)

            return augmented.astype(np.uint8)

# Example usage
if __name__ == "__main__":
    print("Simple Training Assistant")
    print("Uses your ImprovedAttributeDetectionExtractor directly!")
    print()

    try:
        assistant = SimpleTrainingAssistant()
        print("Assistant initialized successfully!")
        print("Ready to process training images.")
        print()
        print("Usage:")
        print("  assistant.process_training_images('path/to/your/images')")
        assistant.process_training_images(r"C:\Users\Alexander Becquet\Documents\Python\Python\TPM Data Processing Python Scripts\Standardized Testing GUI\git testing\DataViewer\scanned_forms\July 15 Training")
    except Exception as e:
        print(f"Error initializing assistant: {e}")

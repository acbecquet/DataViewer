# enhanced_training_workflow.py
"""
Enhanced training workflow that directly utilizes the shadow removal functionality
in find_attribute_boundaries_with_ocr from updated_training_extractor.py
"""

import os
import cv2
import numpy as np
from datetime import datetime
import traceback

# Direct import of the extractor with shadow removal capabilities
from updated_training_extractor import ImprovedAttributeDetectionExtractor

def test_shadow_removal_on_image(image_path):
    """
    Test the shadow removal functionality on a single image
    """
    print(f"Testing shadow removal on: {os.path.basename(image_path)}")
    print("="*60)
    
    # Initialize the extractor
    extractor = ImprovedAttributeDetectionExtractor()
    
    try:
        # Load and preprocess the image
        print("DEBUG: Step 1 - Loading and preprocessing image...")
        processed_image = extractor.preprocess_image_for_extraction(image_path)
        print(f"DEBUG: Image loaded: {processed_image.shape[1]}x{processed_image.shape[0]} pixels")
        
        # Detect form cross lines for center coordinates
        print("DEBUG: Step 2 - Detecting form structure...")
        center_x, center_y = extractor.detect_form_cross_lines(processed_image)
        print(f"DEBUG: Form center detected at: ({center_x}, {center_y})")
        
        # Use the shadow removal function directly
        print("DEBUG: Step 3 - Applying shadow removal with find_attribute_boundaries_with_ocr...")
        print("DEBUG: This will apply:")
        print("DEBUG:   - 26% search width (leftmost column)")
        print("DEBUG:   - 2x scaling preprocessing")  
        print("DEBUG:   - Shadow removal: dilate(7x7) -> medianBlur(21x21) -> absdiff -> normalize")
        print("DEBUG:   - OCR on shadow-free image")
        print("DEBUG:   - Tighter grouping (10px tolerance)")
        
        # Call the shadow removal function directly
        attribute_boundaries = extractor.find_attribute_boundaries_with_ocr(
            processed_image, center_x, center_y
        )
        
        print("DEBUG: Step 4 - Shadow removal OCR successful!")
        print(f"DEBUG: Boundaries found: {list(attribute_boundaries.keys())}")
        for key, value in attribute_boundaries.items():
            print(f"DEBUG:   {key}: {value}")
        
        # Now get sample regions using the shadow-removal detected boundaries
        print("DEBUG: Step 5 - Getting sample regions from shadow-removal boundaries...")
        sample_regions = extractor.get_sample_regions(processed_image)
        
        print(f"DEBUG: Sample regions calculated: {len(sample_regions)} regions")
        for sample_id, region in sample_regions.items():
            height = region['y_end'] - region['y_start']
            width = region['x_end'] - region['x_start']
            print(f"DEBUG:   Sample {sample_id}: {width}x{height} at ({region['x_start']},{region['y_start']})")
        
        # Extract training data using the complete workflow
        print("DEBUG: Step 6 - Extracting complete training data...")
        extracted_regions = extractor.get_improved_form_regions(processed_image)
        
        # Summary
        print("\nDEBUG: SHADOW REMOVAL EXTRACTION COMPLETE")
        print("="*60)
        total_extracted = sum(len(sample_data) for sample_data in extracted_regions.values())
        print(f"DEBUG: Total samples: {len(extracted_regions)}")
        print(f"DEBUG: Total attributes: {total_extracted}")
        
        for sample_name, sample_data in extracted_regions.items():
            print(f"DEBUG: {sample_name}: {len(sample_data)} attributes")
            for attribute in sample_data.keys():
                print(f"DEBUG:   - {attribute}")
        
        print(f"\nDEBUG: Debug images saved to: training_data/debug_regions/")
        print(f"DEBUG: Shadow removal images saved to: training_data/debug_regions/debug_minimal_1x_shadows_removed.png")
        print(f"DEBUG: Check debug files ending with '_info.txt' for detailed extraction info")
        
        return True, extracted_regions
        
    except Exception as e:
        print(f"DEBUG: Shadow removal failed: {e}")
        print("DEBUG: Full error traceback:")
        traceback.print_exc()
        print("\nDEBUG: Error analysis:")
        error_str = str(e)
        if "OCR failed to find required attributes" in error_str:
            print("DEBUG: Issue: OCR could not detect all required attribute names")
            print("DEBUG: Possible causes:")
            print("DEBUG:   1. Image quality too poor for OCR")
            print("DEBUG:   2. Attribute names are different than expected")
            print("DEBUG:   3. Shadow removal preprocessing didn't work well enough")
            print("DEBUG: Check the shadow removal debug image: debug_minimal_1x_shadows_removed.png")
        elif "Tesseract" in error_str:
            print("DEBUG: Issue: Tesseract OCR not properly installed")
            print("DEBUG: Install with: pip install pytesseract")
            print("DEBUG: Download from: https://github.com/UB-Mannheim/tesseract/wiki")
        return False, None

def process_training_folder_with_shadow_removal(training_folder):
    """
    Process entire folder using shadow removal
    """
    print("ENHANCED TRAINING WORKFLOW - SHADOW REMOVAL")
    print("="*80)
    print(f"Training folder: {training_folder}")
    print(f"Shadow removal method: find_attribute_boundaries_with_ocr")
    print()
    
    if not os.path.exists(training_folder):
        print(f"ERROR: Training folder not found: {training_folder}")
        return False
    
    # Find image files
    image_files = [f for f in os.listdir(training_folder) 
                  if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff'))]
    
    if not image_files:
        print(f"ERROR: No image files found in {training_folder}")
        return False
    
    print(f"Found {len(image_files)} images to process with shadow removal")
    
    # Initialize extractor once
    extractor = ImprovedAttributeDetectionExtractor()
    
    # Process each image
    successful_count = 0
    failed_count = 0
    
    for i, image_file in enumerate(image_files, 1):
        image_path = os.path.join(training_folder, image_file)
        print(f"\n[{i}/{len(image_files)}] Processing: {image_file}")
        print("-" * 50)
        
        try:
            # Direct call to the complete extraction with shadow removal
            processed_image = extractor.preprocess_image_for_extraction(image_path)
            extracted_regions = extractor.get_improved_form_regions(processed_image)
            
            if extracted_regions:
                total_attributes = sum(len(sample_data) for sample_data in extracted_regions.values())
                print(f"SUCCESS: {len(extracted_regions)} samples, {total_attributes} attributes extracted")
                successful_count += 1
            else:
                print("FAILED: No regions extracted")
                failed_count += 1
                
        except Exception as e:
            print(f"FAILED: {e}")
            failed_count += 1
            continue
    
    # Final summary
    print(f"\n{'='*80}")
    print("SHADOW REMOVAL PROCESSING COMPLETE")
    print(f"{'='*80}")
    print(f"Total images: {len(image_files)}")
    print(f"Successful: {successful_count}")
    print(f"Failed: {failed_count}")
    print(f"Success rate: {(successful_count/len(image_files)*100):.1f}%")
    
    if successful_count > 0:
        print(f"\nResults available in:")
        print(f"  - training_data/debug_regions/ (extracted training images)")
        print(f"  - Look for debug_minimal_1x_shadows_removed.png (shadow removal results)")
        print(f"  - Check *_info.txt files for detailed extraction information")
        return True
    else:
        print(f"\nNo successful extractions. Check error messages above.")
        return False

def main():
    """
    Main function - choose single image test or full folder processing
    """
    print("Enhanced Training Workflow - Direct Shadow Removal Usage")
    print("="*60)
    print("Options:")
    print("1. Test shadow removal on single image")
    print("2. Process entire folder with shadow removal")
    print()
    
    choice = input("Enter choice (1 or 2): ").strip()
    
    if choice == "1":
        image_path = input("Enter path to test image: ").strip()
        if os.path.exists(image_path):
            success, extracted_data = test_shadow_removal_on_image(image_path)
            if success:
                print("\nShadow removal test completed successfully!")
            else:
                print("\nShadow removal test failed. Check debug output above.")
        else:
            print(f"Image not found: {image_path}")
    
    elif choice == "2":
        training_folder = input("Enter path to training folder: ").strip()
        success = process_training_folder_with_shadow_removal(training_folder)
        if success:
            print("\nFolder processing completed!")
        else:
            print("\nFolder processing failed. Check debug output above.")
    
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()
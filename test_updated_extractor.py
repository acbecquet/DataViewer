# test_improved_extractor.py
"""
Test script for the improved OCR-based extractor
"""

import os
import cv2
from updated_training_extractor import ImprovedAttributeDetectionExtractor

def test_ocr_attribute_detection(image_path):
    """
    Test the OCR-based attribute detection.
    """
    
    print(f"Testing OCR attribute detection on: {os.path.basename(image_path)}")
    
    extractor = ImprovedAttributeDetectionExtractor()
    
    # Load and preprocess image
    processed_image = extractor.preprocess_image_for_extraction(image_path)
    print(f"Preprocessed image size: {processed_image.shape[1]}x{processed_image.shape[0]}")
    
    # Test the OCR-based extraction
    regions = extractor.get_improved_form_regions(processed_image)
    
    print(f"\nOCR Extraction Results:")
    for sample_name, sample_regions in regions.items():
        print(f"{sample_name}: {len(sample_regions)} attributes found")
        for attribute, region_img in sample_regions.items():
            print(f"  ✓ {attribute}: {region_img.shape[1]}x{region_img.shape[0]} -> {extractor.target_size}")
    
    print(f"\nDebug images and OCR info saved to: training_data/debug_regions/")
    print("Check these files to verify:")
    print("1. debug_Sample1_Burnt_Taste.png shows the Burnt Taste row")
    print("2. debug_Sample1_Burnt_Taste_info.txt shows OCR details")
    print("3. Each row contains the attribute name + rating scale 1-9")
    
    return regions

if __name__ == "__main__":
    print("Improved OCR-Based Extractor Test")
    print("="*50)
    
    # Your test image path
    test_image_path = r"C:\Users\Alexander Becquet\Documents\Python\Python\TPM Data Processing Python Scripts\Standardized Testing GUI\git testing\DataViewer\tests\test_sensory\james 1.jpg"
    
    if os.path.exists(test_image_path):
        print(f"✓ Test image found!")
        
        try:
            regions = test_ocr_attribute_detection(test_image_path)
            
            print("\n" + "="*50)
            print("✓ OCR Test completed!")
            print("="*50)
            print("Next steps:")
            print("1. Check training_data/debug_regions/ for extracted images")
            print("2. Look at the _info.txt files to see what OCR detected")
            print("3. Verify each image shows the correct attribute + rating scale")
            print("4. If good, run full extraction on your training folder")
            
        except Exception as e:
            print(f"✗ Test failed: {e}")
            import traceback
            traceback.print_exc()
            
            if "pytesseract" in str(e):
                print("\nTo fix OCR issues:")
                print("1. pip install pytesseract")
                print("2. Download Tesseract: https://github.com/UB-Mannheim/tesseract/wiki")
                print("3. Add Tesseract to your PATH")
    else:
        print(f"✗ Test image not found: {test_image_path}")
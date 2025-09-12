
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
        Get the 4 sample regions using OCR to find actual attribute boundaries.
        """

        height, width = image.shape
        print(f"Finding sample regions using OCR boundary detection...")

        # First detect cross lines for left/right separation
        center_x, center_y = self.detect_form_cross_lines(image)

        # Define margins
        left_margin = int(width * 0.02)
        right_margin = int(width * 0.98)

        # STEP 1: Find attribute boundaries using OCR in leftmost column
        attribute_boundaries = self.find_attribute_boundaries_with_ocr(image, center_x, center_y)

        # STEP 2: Create regions based on found boundaries
        sample_regions = {}

        # Top samples (1 and 2) - same Y coordinates
        if 'top_start' in attribute_boundaries and 'top_end' in attribute_boundaries:
            top_start_y = attribute_boundaries['top_start']
            top_end_y = attribute_boundaries['top_end']

            sample_regions[1] = {  # Top-left
                "y_start": top_start_y,
                "y_end": top_end_y,
                "x_start": left_margin,
                "x_end": center_x - 5,
                "content_y_start": top_start_y,
                "content_y_end": top_end_y
            }

            sample_regions[2] = {  # Top-right - SAME Y coordinates
                "y_start": top_start_y,
                "y_end": top_end_y,
                "x_start": center_x + 5,
                "x_end": right_margin,
                "content_y_start": top_start_y,
                "content_y_end": top_end_y
            }

            print(f"  Top samples: y={top_start_y}-{top_end_y} (height={top_end_y - top_start_y})")

        # Bottom samples (3 and 4) - same Y coordinates but different from top
        if 'bottom_start' in attribute_boundaries and 'bottom_end' in attribute_boundaries:
            bottom_start_y = attribute_boundaries['bottom_start']
            bottom_end_y = attribute_boundaries['bottom_end']

            sample_regions[3] = {  # Bottom-left
                "y_start": bottom_start_y,
                "y_end": bottom_end_y,
                "x_start": left_margin,
                "x_end": center_x - 5,
                "content_y_start": bottom_start_y,
                "content_y_end": bottom_end_y
            }

            sample_regions[4] = {  # Bottom-right - SAME Y coordinates
                "y_start": bottom_start_y,
                "y_end": bottom_end_y,
                "x_start": center_x + 5,
                "x_end": right_margin,
                "content_y_start": bottom_start_y,
                "content_y_end": bottom_end_y
            }

            print(f"  Bottom samples: y={bottom_start_y}-{bottom_end_y} (height={bottom_end_y - bottom_start_y})")

        print(f"DEBUG: OCR-based sample regions:")
        for sample_id, region in sample_regions.items():
            region_height = region['y_end'] - region['y_start']
            print(f"  Sample {sample_id}: y={region['y_start']}-{region['y_end']} (height={region_height}), x={region['x_start']}-{region['x_end']}")

        return sample_regions

    def find_attribute_boundaries_with_ocr(self, image, center_x, center_y):
        """
        Find burnt_taste and overall_liking from the SAME sample region,
        then calculate region height from their difference.
        """

        height, width = image.shape
        print(f"  Using SAME-REGION approach: find burnt_taste + overall_liking from same sample...")
        print(f"  Dynamic threshold: center_y={center_y}")

        search_width = int(width * 0.30)
        left_column = image[:,0:search_width]

        # Create debug directory
        debug_dir = "training_data/debug_regions"
        os.makedirs(debug_dir, exist_ok=True)

        # Preprocess image with shadow removal
        preprocessed = cv2.resize(left_column, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        search_region = preprocessed[:, :search_width*2]

        # Shadow removal pipeline
        region_planes = cv2.split(search_region)
        result_planes = []
        for plane in region_planes:
            dilated_img = cv2.dilate(plane, np.ones((7,7), np.uint8))
            bg_img = cv2.medianBlur(dilated_img, 21)
            diff_img = 255 - cv2.absdiff(plane, bg_img)
            result_planes.append(diff_img)

        search_region = cv2.merge(result_planes)
        cv2.imwrite(os.path.join(debug_dir, "debug_minimal_1x_shadows_removed.png"), search_region)

        try:
            # Get OCR data and group lines
            ocr_data = pytesseract.image_to_data(search_region, output_type=pytesseract.Output.DICT)

            text_items = []
            for i in range(len(ocr_data['text'])):
                text = ocr_data['text'][i].strip().lower()
                if text and len(text) > 1 and ocr_data['conf'][i] > 20:
                    y = ocr_data['top'][i] // 2
                    text_items.append({'text': text, 'y': y, 'confidence': ocr_data['conf'][i]})

            # Group nearby text items
            grouped_lines = []
            text_items.sort(key=lambda x: x['y'])
            current_line = []
            for item in text_items:
                if not current_line or abs(item['y'] - current_line[-1]['y']) <= 10:
                    current_line.append(item)
                else:
                    if current_line:
                        combined_text = ' '.join([t['text'] for t in current_line])
                        avg_y = sum([t['y'] for t in current_line]) // len(current_line)
                        grouped_lines.append({'text': combined_text, 'y': avg_y})
                    current_line = [item]
            if current_line:
                combined_text = ' '.join([t['text'] for t in current_line])
                avg_y = sum([t['y'] for t in current_line]) // len(current_line)
                grouped_lines.append({'text': combined_text, 'y': avg_y})

            print(f"  Grouped into {len(grouped_lines)} text lines")

            # FIND ALL burnt_taste and overall_liking positions
            all_burnt_taste = []
            all_overall_liking = []

            for line in grouped_lines:
                text = line['text']
                y = line['y']

                # Find ALL burnt_taste occurrences
                if any(variant in text for variant in [
                    'burnt taste', 'surnt taste', 'gumt taste', 'burnt tase', 'surmt taste'
                ]):
                    sample_region = "TOP" if y < center_y else "BOTTOM"
                    all_burnt_taste.append({'y': y, 'text': text, 'region': sample_region})
                    print(f"  Found BURNT_TASTE in {sample_region}: y={y} '{text}'")

                # Find ALL overall_liking occurrences
                elif (any(liking_word in text for liking_word in ['liking', 'iking', 'linking']) and
                      not any(flavor_word in text for flavor_word in ['flavor', 'flavour'])) or \
                     (any(word in text for word in ['uldne', 'qyeral']) and
                      not any(flavor_word in text for flavor_word in ['flavor', 'flavour'])):
                    sample_region = "TOP" if y < center_y else "BOTTOM"
                    all_overall_liking.append({'y': y, 'text': text, 'region': sample_region})
                    print(f"  Found OVERALL_LIKING in {sample_region}: y={y} '{text}'")

            # STRATEGY: Try to find a complete pair from the same region
            print(f"  Found {len(all_burnt_taste)} burnt_taste and {len(all_overall_liking)} overall_liking")

            region_height = None
            source_region = None
            burnt_taste_y = None
            overall_liking_y = None

            # Try TOP region first
            top_burnt = [bt for bt in all_burnt_taste if bt['region'] == 'TOP']
            top_liking = [ol for ol in all_overall_liking if ol['region'] == 'TOP']

            if top_burnt and top_liking:
                burnt_taste_y = top_burnt[0]['y']  # Take first occurrence
                overall_liking_y = top_liking[0]['y']  # Take first occurrence
                region_height = abs(overall_liking_y - burnt_taste_y)
                source_region = "TOP"
                print(f"  ✓ Using TOP region: burnt_taste={burnt_taste_y}, overall_liking={overall_liking_y}")
                print(f"  ✓ Calculated region height: {region_height}px from TOP sample")

            # Try BOTTOM region if TOP didn't work
            if region_height is None:
                bottom_burnt = [bt for bt in all_burnt_taste if bt['region'] == 'BOTTOM']
                bottom_liking = [ol for ol in all_overall_liking if ol['region'] == 'BOTTOM']

                if bottom_burnt and bottom_liking:
                    burnt_taste_y = bottom_burnt[0]['y']  # Take first occurrence
                    overall_liking_y = bottom_liking[0]['y']  # Take first occurrence
                    region_height = abs(overall_liking_y - burnt_taste_y)
                    source_region = "BOTTOM"
                    print(f"  ✓ Using BOTTOM region: burnt_taste={burnt_taste_y}, overall_liking={overall_liking_y}")
                    print(f"  ✓ Calculated region height: {region_height}px from BOTTOM sample")

            # VALIDATION: Check if region height makes sense
            if region_height is None:
                raise Exception("Could not find burnt_taste and overall_liking in the same region")

            max_reasonable_height = height * 0.4  # No more than 40% of image height
            if region_height > max_reasonable_height:
                print(f"  ⚠️ WARNING: Calculated region height {region_height}px seems too large (>{max_reasonable_height:.0f}px)")
                print(f"  This might indicate attributes from different regions were matched")
                raise Exception(f"Calculated region height {region_height}px is too large, probably mixed regions")

            print(f"  📏 VALIDATED REGION HEIGHT: {region_height}px from {source_region} sample")

            # Calculate boundaries using the validated region height
            buffer = 20

            # Apply the region height but USE ACTUAL POSITIONS when available
            if source_region == "TOP":
                # Use the actual top positions we found
                top_start_y = max(0, burnt_taste_y - buffer)
                top_end_y = min(height, overall_liking_y + buffer + 30)

                # Estimate bottom region with same height, but check if we have bottom positions
                bottom_burnt = [bt for bt in all_burnt_taste if bt['region'] == 'BOTTOM']
                if bottom_burnt:
                    # Use actual bottom burnt_taste position
                    bottom_start_y = max(0, bottom_burnt[0]['y'] - buffer)
                    bottom_end_y = min(height, bottom_start_y + region_height + 2*buffer)
                else:
                    # Estimate bottom region
                    bottom_start_y = max(center_y + 50, 0)
                    bottom_end_y = min(height, bottom_start_y + region_height + 2*buffer)

            else:  # source_region == "BOTTOM"
                # Use the actual bottom positions we found
                bottom_start_y = max(0, burnt_taste_y - buffer)
                bottom_end_y = min(height, overall_liking_y + buffer + 30)

                # Use actual TOP positions if available (don't just estimate!)
                top_burnt = [bt for bt in all_burnt_taste if bt['region'] == 'TOP']
                if top_burnt:
                    # Use actual top burnt_taste position with calculated height
                    top_start_y = max(0, top_burnt[0]['y'] - buffer)
                    top_end_y = min(height, top_start_y + region_height + 2*buffer)
                    print(f"  Using actual TOP burnt_taste at y={top_burnt[0]['y']} for TOP region")
                else:
                    # Estimate top region only if no TOP positions found
                    top_end_y = min(center_y - 50, height)
                    top_start_y = max(0, top_end_y - region_height - 2*buffer)
                    print(f"  Estimating TOP region (no TOP attributes found)")

            boundaries = {
                'top_start': top_start_y,
                'top_end': top_end_y,
                'bottom_start': bottom_start_y,
                'bottom_end': bottom_end_y
            }

            print(f"  ✓ CONSISTENT BOUNDARIES using {source_region}-derived height {region_height}px:")
            print(f"    TOP: y={top_start_y}-{top_end_y} (height={top_end_y - top_start_y})")
            print(f"    BOTTOM: y={bottom_start_y}-{bottom_end_y} (height={bottom_end_y - bottom_start_y})")

            return boundaries

        except Exception as e:
            print(f"  ❌ SAME-REGION OCR BOUNDARY DETECTION FAILED: {e}")
            raise Exception(f"OCR boundary detection failed: {e}")

    def find_attributes_in_sample(self, image, sample_region):
        """
        Find attributes using precise 5-section division of OCR-detected boundaries.
        """

        print(f"Finding attributes in sample region with 5-section division...")

        # Extract the sample region
        y_start, y_end = sample_region["y_start"], sample_region["y_end"]
        x_start, x_end = sample_region["x_start"], sample_region["x_end"]

        sample_img = image[y_start:y_end, x_start:x_end]

        if sample_img.size == 0:
            print("  ERROR: Empty sample region")
            return {}

        region_height = y_end - y_start
        print(f"  Sample region: {sample_img.shape[1]}x{sample_img.shape[0]}, region_height={region_height}")

        # DIVIDE INTO 5 EQUAL SECTIONS for the 5 attributes
        section_height = region_height // 5

        print(f"  Dividing into 5 sections of {section_height} pixels each:")

        found_attributes = {}

        # Assign each attribute to its section
        for i, attribute in enumerate(self.expected_attributes):
            section_y_start = i * section_height
            section_y_end = (i + 1) * section_height
            section_y_center = section_y_start + section_height // 2

            # Try OCR in this specific section first
            section_img = sample_img[section_y_start:section_y_end, :]
            ocr_result = self.try_ocr_in_section(section_img, attribute)

            if ocr_result:
                # Adjust coordinates to sample region
                ocr_result['y'] = section_y_start + ocr_result['y']
                found_attributes[attribute] = ocr_result
                print(f"    Section {i+1} ({attribute}): ✓ OCR found '{ocr_result['text']}' at y={ocr_result['y']}")
            else:
                # Use section center as fallback
                found_attributes[attribute] = {
                    'x': 0,
                    'y': section_y_center,
                    'width': 100,
                    'height': section_height,
                    'text': f'section_{i+1}',
                    'confidence': 0,
                    'absolute_y': y_start + section_y_center
                }
                print(f"    Section {i+1} ({attribute}): Using center position y={section_y_center}")

        return found_attributes

    def try_ocr_in_section(self, section_img, target_attribute):
        """
        Try OCR within a specific section to find the target attribute.
        """

        if section_img.shape[0] < 10 or section_img.shape[1] < 10:
            return None

        try:
            # Enhance section for OCR
            enhanced = cv2.resize(section_img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
            _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # Get OCR data
            ocr_data = pytesseract.image_to_data(binary, output_type=pytesseract.Output.DICT,config='--psm 6')

            best_match = None
            best_score = 0

            for i, text in enumerate(ocr_data['text']):
                if int(ocr_data['conf'][i]) < 25:
                    continue

                match_score = self.calculate_attribute_match_score(text.strip(), target_attribute)

                if match_score > best_score and match_score > 0.5:
                    best_match = {
                        'x': int(ocr_data['left'][i] / 2),
                        'y': int(ocr_data['top'][i] / 2),
                        'width': int(ocr_data['width'][i] / 2),
                        'height': int(ocr_data['height'][i] / 2),
                        'text': text.strip(),
                        'confidence': int(ocr_data['conf'][i]),
                        'match_score': match_score
                    }
                    best_score = match_score

            return best_match

        except Exception as e:
            return None

    # Add the preprocessing methods
    def preprocess_for_ocr_enhanced(self, img):
        """Enhanced preprocessing with scaling and contrast."""
        enhanced = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
        enhanced = clahe.apply(enhanced)
        _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return binary

    def preprocess_for_ocr_binary(self, img):
        """Simple binary preprocessing."""
        _, binary = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return cv2.resize(binary, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    def preprocess_for_ocr_original(self, img):
        """Minimal preprocessing."""
        return cv2.resize(img, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)

    def get_fallback_attributes(self, sample_img, content_y_start_relative=None, content_y_end_relative=None):
        """
        Generate fallback attribute positions with content area awareness.
        """

        height = sample_img.shape[0]

        # Use content area if provided, otherwise use full sample
        if content_y_start_relative is not None and content_y_end_relative is not None:
            effective_start = max(0, content_y_start_relative)
            effective_end = min(height, content_y_end_relative)
            effective_height = effective_end - effective_start
            print(f"    Using content-area fallback: y={effective_start}-{effective_end} (height={effective_height})")
        else:
            effective_start = 0
            effective_end = height
            effective_height = height
            print(f"    Using full-sample fallback: height={effective_height}")

        attr_height = effective_height // len(self.expected_attributes)

        fallback_attributes = {}
        for i, attribute in enumerate(self.expected_attributes):
            y_pos = effective_start + i * attr_height + attr_height // 2
            fallback_attributes[attribute] = {
                'x': 0, 'y': y_pos, 'width': 100, 'height': attr_height,
                'text': 'estimated', 'confidence': 0,
                'absolute_y': y_pos  # For consistency with OCR results
            }
            print(f"    Fallback position for {attribute}: y={y_pos}")

        return fallback_attributes

    def extract_attribute_row(self, image, sample_region, attribute_info, sample_id=None):
        """
        Extract the complete row with text CENTERED in the extracted image.
        Adjust buffer for right samples due to aspect ratio scaling.
        """

        # Get sample region boundaries
        sample_y_start = sample_region["y_start"]
        sample_x_start = sample_region["x_start"]
        sample_x_end = sample_region["x_end"]

        # Calculate actual image coordinates - this is where the text starts
        text_y_start = sample_y_start + attribute_info['y']
        text_height = attribute_info.get('height', 30)

        # Calculate the CENTER of the text
        text_y_center = text_y_start + text_height // 2

        print(f"DEBUG: Text starts at y={text_y_start}, height={text_height}, center={text_y_center}")

        # ASPECT RATIO ADJUSTED ROW HEIGHT CALCULATION
        base_row_height = max(80, int(text_height * 3))

        if sample_id in [2, 4]:  # Right-side samples with full width
            # Divide buffer by 4 for aspect ratio scaling compensation
            row_height = base_row_height // 2
            print(f"DEBUG: Sample {sample_id} (RIGHT): Reduced row height {base_row_height} -> {row_height} (÷4 for aspect ratio)")
        else:  # Left-side samples with partial width
            row_height = base_row_height
            print(f"DEBUG: Sample {sample_id} (LEFT): Standard row height {row_height}")

        # CENTER the row around the text center
        row_y_start = max(0, text_y_center - row_height // 2)
        row_y_end = min(image.shape[0], text_y_center + row_height // 2)

        # Adjust if we hit boundaries
        actual_row_height = row_y_end - row_y_start
        if actual_row_height < row_height:
            # If we hit a boundary, try to expand the other direction
            if row_y_start == 0:
                # Hit top boundary, expand down more
                row_y_end = min(image.shape[0], row_y_start + row_height)
            elif row_y_end == image.shape[0]:
                # Hit bottom boundary, expand up more
                row_y_start = max(0, row_y_end - row_height)

        # WIDTH HANDLING
        image_width = image.shape[1]

        if sample_id in [2, 4]:  # Right-side samples - show FULL WIDTH
            row_x_start = int(image_width * 0.02)  # Start from left margin
            row_x_end = int(image_width * 0.98)    # Go to right margin
            print(f"DEBUG: Sample {sample_id} (RIGHT): Using FULL WIDTH x={row_x_start}-{row_x_end}")
            print(f"DEBUG: This will show attribute name on left + full rating scale on right")
            print(f"DEBUG: Aspect ratio compensation: {base_row_height} -> {row_height} height")
        else:  # Left-side samples - use extended width
            row_x_start = sample_x_start
            extension_pixels = int(image_width * 0.15)
            row_x_end = min(image_width, sample_x_end + extension_pixels)
            print(f"DEBUG: Sample {sample_id} (LEFT): Using extended width x={row_x_start}-{row_x_end}")

        print(f"DEBUG: CENTERED Row boundaries: x={row_x_start}-{row_x_end}, y={row_y_start}-{row_y_end}")
        print(f"DEBUG: Text center y={text_y_center} should be at middle of {row_y_start}-{row_y_end}")
        print(f"DEBUG: Y span: {row_y_end - row_y_start} pixels, X span: {row_x_end - row_x_start} pixels")

        # Extract the row
        row_img = image[row_y_start:row_y_end, row_x_start:row_x_end]

        print(f"    Extracted CENTERED row: {row_img.shape[1]}x{row_img.shape[0]} at y={row_y_start}-{row_y_end}")

        return row_img

    def calculate_attribute_match_score(self, ocr_text, target_attribute):
        """
        Calculate how well OCR text matches a target attribute with improved disambiguation.
        """

        ocr_words = ocr_text.lower().strip().split()
        target_words = target_attribute.lower().split()

        # Special handling for similar attributes
        if target_attribute == "Overall Flavor":
            # Must contain "flavor" or "flavour" - be strict about this
            flavor_words = ["flavor", "flavour", "flav"]
            has_flavor = any(any(fw in word for fw in flavor_words) for word in ocr_words)
            has_overall = any("overall" in word for word in ocr_words)

            if has_flavor and has_overall:
                print(f"      MATCH: '{ocr_text}' -> Overall Flavor (has both overall + flavor)")
                return 0.95
            elif has_flavor and not any("lik" in word for word in ocr_words):
                # Has flavor but no "liking" words
                print(f"      PARTIAL: '{ocr_text}' -> Overall Flavor (has flavor, no liking)")
                return 0.7
            else:
                return 0.0

        elif target_attribute == "Overall Liking":
            # Must contain "liking" or "like" - be strict about this
            liking_words = ["liking", "like", "lik"]
            has_liking = any(any(lw in word for lw in liking_words) for word in ocr_words)
            has_overall = any("overall" in word for word in ocr_words)

            if has_liking and has_overall:
                print(f"      MATCH: '{ocr_text}' -> Overall Liking (has both overall + liking)")
                return 0.95
            elif has_liking and not any("flav" in word for word in ocr_words):
                # Has liking but no "flavor" words
                print(f"      PARTIAL: '{ocr_text}' -> Overall Liking (has liking, no flavor)")
                return 0.7
            else:
                return 0.0

        elif target_attribute == "Vapor Volume":
            # Be more flexible with Vapor Volume detection
            vapor_words = ["vapor", "vapour", "vap"]
            volume_words = ["volume", "vol", "olume"]

            has_vapor = any(any(vw in word for vw in vapor_words) for word in ocr_words)
            has_volume = any(any(vow in word for vow in volume_words) for word in ocr_words)

            if has_vapor and has_volume:
                print(f"      MATCH: '{ocr_text}' -> Vapor Volume (has both vapor + volume)")
                return 0.95
            elif has_vapor or has_volume:
                # Accept partial matches for Vapor Volume since it's tricky for OCR
                print(f"      PARTIAL: '{ocr_text}' -> Vapor Volume (has vapor or volume)")
                return 0.6
            else:
                return 0.0

        elif target_attribute == "Burnt Taste":
            burnt_words = ["burnt", "burn", "urnt"]
            taste_words = ["taste", "tast", "aste"]

            has_burnt = any(any(bw in word for bw in burnt_words) for word in ocr_words)
            has_taste = any(any(tw in word for tw in taste_words) for word in ocr_words)

            if has_burnt and has_taste:
                return 0.95
            elif has_burnt or has_taste:
                return 0.6
            else:
                return 0.0

        elif target_attribute == "Smoothness":
            smooth_words = ["smoothness", "smooth", "mooth"]
            has_smooth = any(any(sw in word for sw in smooth_words) for word in ocr_words)

            if has_smooth:
                return 0.95
            else:
                return 0.0

        # Fallback: general word matching
        matches = 0
        for target_word in target_words:
            for ocr_word in ocr_words:
                if target_word in ocr_word or ocr_word in target_word:
                    matches += 1
                    break

        score = matches / len(target_words) if target_words else 0
        print(f"      GENERAL: '{ocr_text}' -> {target_attribute} (score={score:.2f})")
        return score

    def get_improved_form_regions(self, image):
        """
        Extract regions using OCR-based attribute detection with improved boundaries.
        """

        height, width = image.shape
        print(f"Processing image dimensions: {width}x{height}")

        # Get the 4 sample regions with improved content boundaries
        sample_regions = self.get_sample_regions(image)

        all_regions = {}

        # Process each sample region
        for sample_id, sample_region in sample_regions.items():
            print(f"\nProcessing Sample {sample_id}:")
            print(f"  Region: y={sample_region['y_start']}-{sample_region['y_end']}, x={sample_region['x_start']}-{sample_region['x_end']}")
            print(f"  Content: y={sample_region['content_y_start']}-{sample_region['content_y_end']} (excludes comments)")

            # Find attributes in this sample using OCR
            found_attributes = self.find_attributes_in_sample(image, sample_region)

            sample_name = f"Sample {sample_id}"
            all_regions[sample_name] = {}

            # Extract row for each found attribute
            for attribute, attr_info in found_attributes.items():
                print(f"  Extracting row for {attribute}...")

                try:
                    # PASS SAMPLE_ID to extraction method for width handling
                    row_img = self.extract_attribute_row(image, sample_region, attr_info, sample_id)

                    if row_img.size > 0:
                        # Resize to target size
                        resized = cv2.resize(row_img, self.target_size, interpolation=cv2.INTER_CUBIC)
                        all_regions[sample_name][attribute] = resized

                        # Save debug image with sample info
                        debug_info = {
                            'sample_id': sample_id,
                            'attribute': attribute,
                            'ocr_text': attr_info['text'],
                            'confidence': attr_info['confidence'],
                            'y_position': attr_info['y'],
                            'width_mode': 'FULL_WIDTH' if sample_id in [2, 4] else 'EXTENDED',
                            'comments_excluded': True
                        }
                        self.save_debug_region(resized, f"Sample{sample_id}_{attribute.replace(' ', '_')}", debug_info)

                        print(f"     ✓ Successfully extracted and resized to {self.target_size}")
                        if sample_id in [2, 4]:
                            print(f"     ✓ Full width mode - attribute name should be visible on left")
                    else:
                        print(f"     ✗ Empty row extracted")

                except Exception as e:
                    print(f"     ✗ Error extracting row: {e}")

            print(f"  Sample {sample_id} complete: {len(all_regions[sample_name])} attributes extracted")

        return all_regions

    def save_debug_region(self, region_img, label, debug_info=None):
        """
        Save debug images with OCR boundary information.
        """

        debug_dir = "training_data/debug_regions"
        os.makedirs(debug_dir, exist_ok=True)

        # Save the region image
        debug_path = os.path.join(debug_dir, f"debug_{label}.png")
        cv2.imwrite(debug_path, region_img)

        # Save enhanced debug info
        if debug_info:
            info_path = os.path.join(debug_dir, f"debug_{label}_info.txt")
            with open(info_path, 'w') as f:
                f.write(f"Attribute: {debug_info.get('attribute', 'Unknown')}\n")
                f.write(f"Sample ID: {debug_info.get('sample_id', 'Unknown')}\n")
                f.write(f"Width Mode: {debug_info.get('width_mode', 'Unknown')}\n")
                f.write(f"Boundary Detection: OCR-based (Burnt Taste -> Overall Liking)\n")
                f.write(f"Section Division: 5 equal sections within OCR boundaries\n")
                f.write(f"Comments Excluded: {debug_info.get('comments_excluded', True)}\n")
                f.write(f"OCR Text Found: '{debug_info.get('ocr_text', 'None')}'\n")
                f.write(f"OCR Confidence: {debug_info.get('confidence', 0)}\n")
                f.write(f"Y Position in Sample: {debug_info.get('y_position', 'Unknown')}\n")
                f.write(f"Final Size: {region_img.shape[1]}x{region_img.shape[0]}\n")
                f.write(f"Target Size: {self.target_size[0]}x{self.target_size[1]}\n")
                f.write(f"\nBoundary Strategy:\n")
                f.write(f"1. OCR scans leftmost column for 'Burnt Taste' and 'Overall Liking'\n")
                f.write(f"2. Creates precise boundaries excluding headers and comments\n")
                f.write(f"3. Divides content area into 5 equal sections for attributes\n")
                f.write(f"4. Top samples (1,2) share same Y coordinates\n")
                f.write(f"5. Bottom samples (3,4) share same Y coordinates\n")

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
    test_image_path = r"C:\Users\Alexander Becquet\Documents\Python\Python\TPM Data Processing Python Scripts\Standardized Testing GUI\git testing\DataViewer\tests\test_sensory\james 1.jpg"

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

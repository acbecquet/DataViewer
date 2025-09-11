# controllers/image_controller.py
"""
controllers/image_controller.py
Image handling controller that coordinates image operations.
This consolidates functionality from image_loader.py and related image handling code.
"""

import os
import re
import copy
import shutil
import time
import uuid
import tempfile
import threading
import base64
import io
import json
from typing import Optional, Dict, Any, List, Tuple

import tkinter as tk
from tkinter import filedialog, Canvas, Scrollbar, Frame, Label, PhotoImage, messagebox, ttk

# Local imports
from utils import debug_print, show_success_message, APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT
from models.data_model import DataModel
from models.image_model import ImageModel, ImageMetadata, CropState


class ImageController:
    """Controller for image handling and display operations."""
    
    def __init__(self, data_model: DataModel, image_service: Any):
        """Initialize the image controller."""
        self.data_model = data_model
        self.image_service = image_service
        self.image_model = ImageModel()
        
        # GUI reference will be set when connected
        self.gui = None
        
        # Image processing state
        self.processed_image_cache = {}
        self.processing_status = {}
        self.temp_files = []  # Track temporary files for cleanup
        
        # Current image loader instance
        self.current_image_loader = None
        
        # Initialize Claude API client if available
        self.claude_client = None
        self._init_claude_api()
        
        print("DEBUG: ImageController initialized")
        print(f"DEBUG: Connected to DataModel and ImageService")
        print("DEBUG: Image cache and processing systems initialized")
    
    def set_gui_reference(self, gui):
        """Set reference to main GUI for UI operations."""
        self.gui = gui
        print("DEBUG: ImageController connected to GUI")
    
    def _init_claude_api(self):
        """Initialize Claude API client for intelligent cropping."""
        try:
            import anthropic
            api_key = os.getenv('ANTHROPIC_API_KEY')
            if api_key:
                self.claude_client = anthropic.Anthropic(api_key=api_key)
                print("DEBUG: Claude API client initialized successfully")
            else:
                print("DEBUG: ANTHROPIC_API_KEY not found, Claude cropping disabled")
        except ImportError:
            print("DEBUG: anthropic package not available, Claude cropping disabled")
        except Exception as e:
            print(f"DEBUG: Claude API initialization failed: {e}")
    
    def _lazy_import_pil(self):
        """Lazy import PIL modules."""
        try:
            from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
            print("TIMING: Lazy loaded PIL modules")
            return Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
        except ImportError as e:
            print(f"ERROR: PIL not available: {e}")
            return None, None, None, None, None, None
    
    def _lazy_import_cv2(self):
        """Lazy import cv2."""
        try:
            import cv2
            print("TIMING: Lazy loaded cv2")
            return cv2
        except ImportError as e:
            print(f"ERROR: cv2 not available: {e}")
            return None
    
    def _lazy_import_numpy(self):
        """Lazy import numpy."""
        try:
            import numpy as np
            print("TIMING: Lazy loaded numpy for ImageController")
            return np
        except ImportError as e:
            print(f"ERROR: numpy not available: {e}")
            return None
    
    def create_image_loader(self, parent_frame, is_plotting_sheet: bool, 
                           on_images_selected_callback=None) -> 'ImageLoader':
        """Create and configure an ImageLoader instance."""
        debug_print(f"DEBUG: ImageController creating ImageLoader for parent frame")
        debug_print(f"DEBUG: is_plotting_sheet={is_plotting_sheet}")
        
        try:
            # Create new ImageLoader instance
            image_loader = ImageLoader(
                parent=parent_frame,
                is_plotting_sheet=is_plotting_sheet,
                on_images_selected=on_images_selected_callback,
                image_controller=self
            )
            
            # Store reference for management
            self.current_image_loader = image_loader
            
            debug_print("DEBUG: ImageLoader created successfully")
            return image_loader
            
        except Exception as e:
            debug_print(f"ERROR: Failed to create ImageLoader: {e}")
            return None
    
    def load_sheet_images(self, sheet_name: str, image_paths: List[str]) -> bool:
        """Load images for a specific sheet."""
        debug_print(f"DEBUG: ImageController loading {len(image_paths)} images for sheet: {sheet_name}")
        
        try:
            # Clear existing images for the sheet
            self.image_model.clear_sheet_images(sheet_name)
            
            # Add new images
            for img_path in image_paths:
                if os.path.exists(img_path):
                    # Create image metadata
                    metadata = ImageMetadata(
                        image_path=img_path,
                        sheet_name=sheet_name,
                        image_type="sheet",
                        display_name=os.path.basename(img_path)
                    )
                    
                    # Add to model
                    self.image_model.add_sheet_image(sheet_name, metadata)
                else:
                    debug_print(f"WARNING: Image file not found: {img_path}")
            
            debug_print(f"DEBUG: Successfully loaded images for sheet {sheet_name}")
            return True
            
        except Exception as e:
            debug_print(f"ERROR: Failed to load sheet images: {e}")
            return False
    
    def add_sample_images(self, sample_images: Dict[str, List[str]], 
                         sample_crop_states: Dict[str, bool] = None,
                         sample_header_data: Dict[str, Any] = None) -> bool:
        """Add sample images to the controller."""
        debug_print(f"DEBUG: ImageController adding sample images for {len(sample_images)} samples")
        
        try:
            if sample_crop_states is None:
                sample_crop_states = {}
            if sample_header_data is None:
                sample_header_data = {}
            
            # Store in GUI if available
            if self.gui:
                if not hasattr(self.gui, 'pending_sample_images'):
                    self.gui.pending_sample_images = {}
                if not hasattr(self.gui, 'pending_sample_image_crop_states'):
                    self.gui.pending_sample_image_crop_states = {}
                if not hasattr(self.gui, 'pending_sample_header_data'):
                    self.gui.pending_sample_header_data = {}
                
                # Merge with existing
                self.gui.pending_sample_images.update(sample_images)
                self.gui.pending_sample_image_crop_states.update(sample_crop_states)
                self.gui.pending_sample_header_data.update(sample_header_data)
                
                debug_print(f"DEBUG: Updated GUI with {len(sample_images)} sample images")
            
            # Add to image model
            for sample_id, image_paths in sample_images.items():
                for img_path in image_paths:
                    if os.path.exists(img_path):
                        # Create metadata
                        metadata = ImageMetadata(
                            image_path=img_path,
                            sheet_name="sample",
                            image_type="sample",
                            display_name=f"{sample_id}_{os.path.basename(img_path)}"
                        )
                        
                        # Set crop state if available
                        if img_path in sample_crop_states:
                            crop_state = CropState()
                            crop_state.crop_enabled = sample_crop_states[img_path]
                            metadata.crop_state = crop_state
                        
                        # Add to model
                        self.image_model.add_sample_image(sample_id, metadata)
            
            debug_print("DEBUG: Sample images added to image model")
            return True
            
        except Exception as e:
            debug_print(f"ERROR: Failed to add sample images: {e}")
            return False
    
    def store_images_for_sheet(self, sheet_name: str, image_paths: List[str], 
                              crop_states: Dict[str, bool] = None) -> bool:
        """Store image paths and crop states for a specific sheet."""
        debug_print(f"DEBUG: ImageController storing {len(image_paths)} images for sheet: {sheet_name}")
        
        try:
            if not sheet_name:
                debug_print("ERROR: No sheet name provided")
                return False
            
            if crop_states is None:
                crop_states = {}
            
            # Store in GUI if available
            if self.gui:
                # Initialize storage structures
                if not hasattr(self.gui, 'sheet_images'):
                    self.gui.sheet_images = {}
                if not hasattr(self.gui, 'image_crop_states'):
                    self.gui.image_crop_states = {}
                
                current_file = getattr(self.gui, 'current_file', 'default')
                
                # Ensure file key exists
                if current_file not in self.gui.sheet_images:
                    self.gui.sheet_images[current_file] = {}
                
                # Store images for the sheet
                self.gui.sheet_images[current_file][sheet_name] = image_paths
                
                # Store crop states
                for img_path in image_paths:
                    if img_path in crop_states:
                        self.gui.image_crop_states[img_path] = crop_states[img_path]
                    elif img_path not in self.gui.image_crop_states:
                        # Default to crop enabled if GUI has crop setting
                        crop_enabled = getattr(self.gui, 'crop_enabled', None)
                        if hasattr(crop_enabled, 'get'):
                            self.gui.image_crop_states[img_path] = crop_enabled.get()
                        else:
                            self.gui.image_crop_states[img_path] = False
                
                debug_print(f"DEBUG: Stored image paths in GUI: {len(image_paths)} images")
                debug_print(f"DEBUG: Stored crop states in GUI: {len(self.gui.image_crop_states)} states")
            
            # Also store in image model
            self.load_sheet_images(sheet_name, image_paths)
            
            return True
            
        except Exception as e:
            debug_print(f"ERROR: Failed to store images for sheet: {e}")
            return False
    
    def process_image_with_crop(self, img_path: str, should_crop: bool = False) -> Any:
        """Process image with optional cropping and caching."""
        debug_print(f"DEBUG: Processing image with crop: {os.path.basename(img_path)}, should_crop={should_crop}")
        
        try:
            Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance = self._lazy_import_pil()
            if Image is None:
                debug_print("ERROR: PIL not available for image processing")
                return None
            
            # Create cache key
            cache_key = f"{img_path}_{should_crop}"
            
            # Check cache first
            if cache_key in self.processed_image_cache:
                debug_print(f"DEBUG: Using cached processed image for {os.path.basename(img_path)}")
                cached_img = self.processed_image_cache[cache_key]
                return cached_img.copy()
            
            # Load original image
            with Image.open(img_path) as original_img:
                processed_img = original_img.copy()
                
                # Apply cropping if requested
                if should_crop:
                    processed_img = self._apply_intelligent_crop(processed_img, img_path)
                
                # Cache the processed image
                self.processed_image_cache[cache_key] = processed_img.copy()
                debug_print(f"DEBUG: Cached processed image for {os.path.basename(img_path)}")
                
                return processed_img
                
        except Exception as e:
            debug_print(f"ERROR: Failed to process image {img_path}: {e}")
            return None
    
    def _apply_intelligent_crop(self, img, img_path: str) -> Any:
        """Apply intelligent cropping to an image."""
        debug_print(f"DEBUG: Applying intelligent crop to {os.path.basename(img_path)}")
        
        try:
            # Try Claude-based intelligent cropping first
            if self.claude_client:
                cropped_img = self._claude_intelligent_crop(img, img_path)
                if cropped_img:
                    debug_print("DEBUG: Claude intelligent crop successful")
                    return cropped_img
            
            # Fallback to computer vision cropping
            debug_print("DEBUG: Using fallback smart crop")
            return self._smart_crop_fallback(img, img_path)
            
        except Exception as e:
            debug_print(f"ERROR: Intelligent crop failed: {e}")
            return img  # Return original if cropping fails
    
    def _claude_intelligent_crop(self, img, img_path: str) -> Any:
        """Use Claude API for intelligent image cropping."""
        debug_print(f"DEBUG: Attempting Claude intelligent crop for {os.path.basename(img_path)}")
        
        try:
            if not self.claude_client:
                debug_print("DEBUG: Claude client not available")
                return None
            
            # Resize image for API if too large
            max_dimension = 1568
            scale_factor = 1.0
            
            if max(img.size) > max_dimension:
                scale_factor = max_dimension / max(img.size)
                new_size = (int(img.size[0] * scale_factor), int(img.size[1] * scale_factor))
                resized_img = img.resize(new_size, img.Resampling.LANCZOS if hasattr(img, 'Resampling') else 3)
                debug_print(f"DEBUG: Resized image for API: {img.size} -> {resized_img.size}")
            else:
                resized_img = img
            
            # Convert to base64
            buffer = io.BytesIO()
            resized_img.save(buffer, format='PNG')
            image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
            # Call Claude API
            response = self.claude_client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=1024,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": image_data
                            }
                        },
                        {
                            "type": "text",
                            "content": "Analyze this image and identify the main subject or object. Return ONLY a JSON object with 'left', 'top', 'right', 'bottom' coordinates (as integers) that would crop to show just the main subject with minimal background. The coordinates should be relative to the image dimensions."
                        }
                    ]
                }]
            )
            
            # Parse response
            response_text = response.content[0].text
            debug_print(f"DEBUG: Claude response: {response_text}")
            
            # Extract JSON from response
            json_match = re.search(r'\{[^}]*\}', response_text)
            if json_match:
                crop_box = json.loads(json_match.group())
                debug_print(f"DEBUG: Parsed crop box: {crop_box}")
                
                # Apply coordinates (scale back if needed)
                with img as original_img:
                    if scale_factor != 1.0:
                        left = int(crop_box['left'] / scale_factor)
                        top = int(crop_box['top'] / scale_factor)
                        right = int(crop_box['right'] / scale_factor)
                        bottom = int(crop_box['bottom'] / scale_factor)
                    else:
                        left = crop_box['left']
                        top = crop_box['top']
                        right = crop_box['right']
                        bottom = crop_box['bottom']
                    
                    # Validate coordinates
                    width, height = original_img.size
                    left = max(0, min(left, width - 1))
                    top = max(0, min(top, height - 1))
                    right = max(left + 1, min(right, width))
                    bottom = max(top + 1, min(bottom, height))
                    
                    debug_print(f"DEBUG: Final crop coordinates: ({left}, {top}, {right}, {bottom})")
                    
                    cropped_img = original_img.crop((left, top, right, bottom))
                    debug_print(f"DEBUG: Cropped image size: {cropped_img.size}")
                    return cropped_img
            
        except Exception as e:
            debug_print(f"ERROR: Claude intelligent crop failed: {e}")
            return None
    
    def _smart_crop_fallback(self, img, img_path: str) -> Any:
        """Fallback cropping using computer vision techniques."""
        debug_print(f"DEBUG: Applying smart crop fallback to {os.path.basename(img_path)}")
        
        try:
            cv2 = self._lazy_import_cv2()
            np = self._lazy_import_numpy()
            
            if cv2 is None or np is None:
                debug_print("DEBUG: OpenCV or numpy not available, returning original image")
                return img
            
            # Convert PIL to OpenCV format
            img_array = np.array(img)
            if len(img_array.shape) == 3:
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            else:
                img_cv = img_array
            
            # Convert to grayscale for edge detection
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY) if len(img_cv.shape) == 3 else img_cv
            
            # Apply Gaussian blur
            blurred = cv2.GaussianBlur(gray, (5, 5), 0)
            
            # Edge detection
            edges = cv2.Canny(blurred, 50, 150)
            
            # Find contours
            contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            if contours:
                # Find largest contour
                largest_contour = max(contours, key=cv2.contourArea)
                
                # Get bounding rectangle
                x, y, w, h = cv2.boundingRect(largest_contour)
                
                # Add padding
                padding = 20
                height, width = gray.shape
                x = max(0, x - padding)
                y = max(0, y - padding)
                w = min(width - x, w + 2 * padding)
                h = min(height - y, h + 2 * padding)
                
                # Crop using PIL
                cropped_img = img.crop((x, y, x + w, y + h))
                debug_print(f"DEBUG: Smart crop applied, new size: {cropped_img.size}")
                return cropped_img
            else:
                debug_print("DEBUG: No contours found, returning original image")
                return img
                
        except Exception as e:
            debug_print(f"ERROR: Smart crop fallback failed: {e}")
            return img
    
    def get_image_for_report(self, img_path: str, target_width: int = None, 
                           target_height: int = None) -> Any:
        """Get processed image for report generation."""
        debug_print(f"DEBUG: Getting image for report: {os.path.basename(img_path)}")
        
        try:
            Image, _, _, _, _, _ = self._lazy_import_pil()
            if Image is None:
                debug_print("ERROR: PIL not available for get_image_for_report")
                return None
            
            # Check for processed version in cache
            should_crop = False
            if self.gui and hasattr(self.gui, 'image_crop_states'):
                should_crop = self.gui.image_crop_states.get(img_path, False)
            
            cache_key = f"{img_path}_{should_crop}"
            
            debug_print(f"DEBUG: Image crop state for {os.path.basename(img_path)}: {should_crop}")
            
            if cache_key in self.processed_image_cache:
                debug_print(f"DEBUG: Using cached processed image for report")
                img = self.processed_image_cache[cache_key].copy()
            else:
                debug_print(f"DEBUG: Using original image for report")
                img = Image.open(img_path)
            
            # Resize if target dimensions specified
            if target_width or target_height:
                original_width, original_height = img.size
                
                if target_width and target_height:
                    img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
                elif target_width:
                    ratio = target_width / original_width
                    new_height = int(original_height * ratio)
                    img = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
                elif target_height:
                    ratio = target_height / original_height
                    new_width = int(original_width * ratio)
                    img = img.resize((new_width, target_height), Image.Resampling.LANCZOS)
                
                debug_print(f"DEBUG: Resized image for report to: {img.size}")
            
            return img
            
        except Exception as e:
            debug_print(f"ERROR: Failed to get image for report {img_path}: {e}")
            try:
                # Fallback to original image
                Image, _, _, _, _, _ = self._lazy_import_pil()
                if Image:
                    return Image.open(img_path)
            except:
                pass
            return None
    
    def set_crop_state(self, image_path: str, crop_enabled: bool) -> None:
        """Set crop state for an image."""
        debug_print(f"DEBUG: Setting crop state for {os.path.basename(image_path)}: {crop_enabled}")
        
        try:
            # Update image model
            crop_state = CropState()
            crop_state.crop_enabled = crop_enabled
            self.image_model.set_crop_state(image_path, crop_state)
            
            # Update GUI if available
            if self.gui and hasattr(self.gui, 'image_crop_states'):
                self.gui.image_crop_states[image_path] = crop_enabled
            
            # Clear processed cache for this image to force reprocessing
            cache_keys_to_remove = [key for key in self.processed_image_cache.keys() if image_path in key]
            for key in cache_keys_to_remove:
                del self.processed_image_cache[key]
                debug_print(f"DEBUG: Removed cached image: {key}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to set crop state: {e}")
    
    def toggle_global_crop_mode(self) -> bool:
        """Toggle global crop mode."""
        try:
            self.image_model.toggle_crop_mode()
            current_state = self.image_model.crop_enabled
            
            debug_print(f"DEBUG: Global crop mode toggled to: {current_state}")
            
            # Update GUI if available
            if self.gui and hasattr(self.gui, 'crop_enabled'):
                if hasattr(self.gui.crop_enabled, 'set'):
                    self.gui.crop_enabled.set(current_state)
            
            # Clear all processed image cache to force reprocessing
            self.clear_image_cache()
            
            return current_state
            
        except Exception as e:
            debug_print(f"ERROR: Failed to toggle crop mode: {e}")
            return False
    
    def clear_image_cache(self) -> None:
        """Clear the processed image cache."""
        debug_print("DEBUG: Clearing processed image cache")
        
        try:
            cache_count = len(self.processed_image_cache)
            self.processed_image_cache.clear()
            self.processing_status.clear()
            
            debug_print(f"DEBUG: Cleared {cache_count} cached images")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to clear image cache: {e}")
    
    def get_cache_info(self) -> Dict[str, Any]:
        """Get information about current cache state."""
        try:
            cache_info = {
                'cached_images': len(self.processed_image_cache),
                'cache_keys': list(self.processed_image_cache.keys()),
                'processing_status': self.processing_status.copy(),
                'temp_files': len(self.temp_files)
            }
            
            debug_print(f"DEBUG: Cache info - {cache_info['cached_images']} cached images, {cache_info['temp_files']} temp files")
            return cache_info
            
        except Exception as e:
            debug_print(f"ERROR: Failed to get cache info: {e}")
            return {}
    
    def load_sample_images_from_vap3(self, vap_data: Dict[str, Any]) -> bool:
        """Load sample images from VAP3 data."""
        debug_print("DEBUG: Loading sample images from VAP3 data")
        
        try:
            sample_images = vap_data.get('sample_images', {})
            sample_crop_states = vap_data.get('sample_image_crop_states', {})
            sample_header_data = vap_data.get('sample_images_metadata', {}).get('header_data', {})
            
            debug_print(f"DEBUG: Found {len(sample_images)} samples in VAP3 data")
            
            if sample_images:
                # Process sample images
                success = self.add_sample_images(sample_images, sample_crop_states, sample_header_data)
                
                if success and self.gui:
                    # Trigger processing of sample images in GUI
                    if hasattr(self.gui, 'process_pending_sample_images'):
                        self.gui.process_pending_sample_images()
                
                debug_print("DEBUG: Sample images loaded from VAP3 successfully")
                return True
            else:
                debug_print("DEBUG: No sample images found in VAP3 data")
                return True
                
        except Exception as e:
            debug_print(f"ERROR: Failed to load sample images from VAP3: {e}")
            return False
    
    def cleanup_temp_files(self) -> None:
        """Clean up temporary files created during processing."""
        debug_print("DEBUG: Cleaning up temporary image files")
        
        try:
            cleaned_count = 0
            for temp_file in self.temp_files[:]:  # Copy list to avoid modification during iteration
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        cleaned_count += 1
                    self.temp_files.remove(temp_file)
                except Exception as e:
                    debug_print(f"WARNING: Failed to remove temp file {temp_file}: {e}")
            
            debug_print(f"DEBUG: Cleaned up {cleaned_count} temporary files")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to cleanup temp files: {e}")
    
    def update_image_display_for_sheet(self, sheet_name: str) -> None:
        """Update image display when sheet changes."""
        debug_print(f"DEBUG: Updating image display for sheet: {sheet_name}")
        
        try:
            if self.current_image_loader and self.gui:
                # Get images for the sheet
                current_file = getattr(self.gui, 'current_file', None)
                
                if (current_file and 
                    hasattr(self.gui, 'sheet_images') and 
                    current_file in self.gui.sheet_images and 
                    sheet_name in self.gui.sheet_images[current_file]):
                    
                    sheet_images = self.gui.sheet_images[current_file][sheet_name]
                    
                    # Load images into the loader
                    self.current_image_loader.load_images_from_list(sheet_images)
                    
                    # Restore crop states
                    if hasattr(self.gui, 'image_crop_states'):
                        for img_path in sheet_images:
                            if img_path in self.gui.image_crop_states:
                                self.current_image_loader.image_crop_states[img_path] = self.gui.image_crop_states[img_path]
                    
                    # Refresh display
                    self.current_image_loader.display_images()
                    
                    debug_print(f"DEBUG: Updated image display with {len(sheet_images)} images")
                else:
                    debug_print(f"DEBUG: No images found for sheet {sheet_name}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to update image display for sheet: {e}")
    
    # Getter methods for external access
    def get_sheet_images(self, sheet_name: str) -> List[str]:
        """Get image paths for a specific sheet."""
        try:
            if self.gui and hasattr(self.gui, 'sheet_images'):
                current_file = getattr(self.gui, 'current_file', None)
                if (current_file and 
                    current_file in self.gui.sheet_images and 
                    sheet_name in self.gui.sheet_images[current_file]):
                    return self.gui.sheet_images[current_file][sheet_name]
            return []
        except Exception as e:
            debug_print(f"ERROR: Failed to get sheet images: {e}")
            return []
    
    def get_sample_images(self) -> Dict[str, Any]:
        """Get current sample images."""
        try:
            if self.gui and hasattr(self.gui, 'pending_sample_images'):
                return self.gui.pending_sample_images
            return {}
        except Exception as e:
            debug_print(f"ERROR: Failed to get sample images: {e}")
            return {}
    
    def get_crop_states(self) -> Dict[str, bool]:
        """Get current crop states."""
        try:
            if self.gui and hasattr(self.gui, 'image_crop_states'):
                return self.gui.image_crop_states.copy()
            return {}
        except Exception as e:
            debug_print(f"ERROR: Failed to get crop states: {e}")
            return {}


class ImageLoader:
    """Image loading and display component managed by ImageController."""
    
    def __init__(self, parent, is_plotting_sheet: bool, on_images_selected=None, image_controller=None):
        """Initialize the ImageLoader."""
        debug_print(f"DEBUG: Initializing ImageLoader with parent={parent}")
        debug_print(f"DEBUG: is_plotting_sheet={is_plotting_sheet}")
        
        self.parent = parent
        self.is_plotting_sheet = is_plotting_sheet
        self.on_images_selected = on_images_selected
        self.image_controller = image_controller
        
        # Image management
        self.image_files = []
        self.image_widgets = []
        self.close_buttons = []
        self.image_crop_states = {}
        self.image_references = []  # Prevent garbage collection
        
        # UI components
        self.frame = None
        self.canvas = None
        self.scrollbar = None
        self.scrollable_frame = None
        self.scrollable_frame_id = None
        
        # Check parent validity
        if not self.parent or not self.parent.winfo_exists():
            debug_print("ERROR: Parent frame does not exist! Skipping UI setup.")
            return
        
        debug_print("DEBUG: Setting up ImageLoader UI...")
        self.setup_ui()
    
    def setup_ui(self):
        """Set up the UI elements."""
        debug_print("DEBUG: ImageLoader setting up UI")
        
        try:
            if not self.parent.winfo_exists():
                debug_print("ERROR: Parent destroyed before UI setup")
                return
            
            # Clear any existing widgets
            for widget in self.parent.winfo_children():
                widget.destroy()
            
            # Create main frame
            self.frame = Frame(self.parent)
            self.frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            # Create canvas and scrollbar
            self.canvas = Canvas(self.frame, height=140, bg="white")
            self.scrollbar = Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
            self.scrollable_frame = Frame(self.canvas, bg="white")
            
            # Create scrollable window
            self.scrollable_frame_id = self.canvas.create_window(
                (0, 0),
                window=self.scrollable_frame,
                anchor="nw"
            )
            
            # Configure scrolling
            def on_canvas_configure(event):
                self.canvas.itemconfig(self.scrollable_frame_id, width=event.width)
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            self.canvas.bind("<Configure>", on_canvas_configure)
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            
            # Pack canvas and scrollbar
            self.canvas.pack(side="left", fill="both", expand=True)
            self.scrollbar.pack(side="right", fill="y")
            
            debug_print("DEBUG: ImageLoader UI setup completed")
            
        except Exception as e:
            debug_print(f"ERROR: ImageLoader UI setup failed: {e}")
    
    def add_images(self):
        """Open file dialog to select and add images."""
        debug_print("DEBUG: ImageLoader opening file dialog for image selection")
        
        try:
            if not (self.parent and self.parent.winfo_exists()):
                debug_print("ERROR: Parent frame no longer exists. Cannot load images.")
                return
            
            # Open file dialog
            file_paths = filedialog.askopenfilenames(
                title="Select Images",
                filetypes=[
                    ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.pdf"),
                    ("All files", "*.*")
                ]
            )
            
            if file_paths:
                debug_print(f"DEBUG: Selected {len(file_paths)} image files")
                
                # Add to image files list
                for file_path in file_paths:
                    if file_path not in self.image_files:
                        self.image_files.append(file_path)
                
                # Display updated images
                self.display_images()
                
                # Notify callback if provided
                if self.on_images_selected:
                    self.on_images_selected(self.image_files)
                    debug_print("DEBUG: Notified callback of selected images")
            else:
                debug_print("DEBUG: No images selected")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to add images: {e}")
    
    def load_images_from_list(self, image_paths: List[str]):
        """Load images from a provided list."""
        debug_print(f"DEBUG: ImageLoader loading {len(image_paths)} images from list")
        
        try:
            self.image_files = list(image_paths)  # Create copy
            debug_print(f"DEBUG: Loaded image files: {[os.path.basename(p) for p in image_paths]}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to load images from list: {e}")
    
    def display_images(self):
        """Display loaded images in the scrollable frame."""
        debug_print("DEBUG: ImageLoader displaying images")
        
        try:
            if not self.canvas or not self.scrollable_frame:
                debug_print("DEBUG: Canvas or scrollable_frame not initialized")
                return
            
            # Store current scroll position
            prev_pos = (0, 0)
            try:
                prev_pos = self.canvas.yview()
            except:
                pass
            
            # Clear existing widgets
            debug_print("DEBUG: Clearing previous image widgets")
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            
            self.image_widgets.clear()
            self.close_buttons.clear()
            self.image_references = []
            
            if not self.image_files:
                debug_print("DEBUG: No image files to display")
                return
            
            # Get PIL modules
            if self.image_controller:
                Image, ImageTk, _, _, _, _ = self.image_controller._lazy_import_pil()
            else:
                Image, ImageTk, _, _, _, _ = None, None, None, None, None, None
            
            if Image is None or ImageTk is None:
                debug_print("ERROR: PIL modules not available for display")
                return
            
            # Display parameters
            max_height = 135
            padding = 5
            current_x = padding
            current_y = padding
            max_row_height = 0
            
            # Get frame width
            try:
                frame_width = self.scrollable_frame.winfo_width()
                if frame_width <= 1:
                    frame_width = 800  # Default width
            except:
                frame_width = 800
            
            debug_print(f"DEBUG: Displaying {len(self.image_files)} images")
            
            # Process each image
            for img_index, img_path in enumerate(self.image_files):
                try:
                    debug_print(f"DEBUG: Processing image {img_index + 1}/{len(self.image_files)}: {os.path.basename(img_path)}")
                    
                    # Check if image should be cropped
                    should_crop = self.image_crop_states.get(img_path, False)
                    
                    # Get processed image from controller
                    if self.image_controller:
                        processed_img = self.image_controller.process_image_with_crop(img_path, should_crop)
                    else:
                        processed_img = Image.open(img_path)
                    
                    if processed_img is None:
                        debug_print(f"WARNING: Failed to process image {img_path}")
                        continue
                    
                    # Resize for display
                    scaling_factor = max_height / processed_img.height
                    new_width = int(processed_img.width * scaling_factor)
                    display_img = processed_img.resize((new_width, max_height), Image.Resampling.LANCZOS)
                    
                    # Convert to PhotoImage
                    photo = ImageTk.PhotoImage(display_img)
                    self.image_references.append(photo)  # Prevent garbage collection
                    
                    # Check if we need to wrap to next row
                    if current_x + new_width + padding > frame_width and current_x > padding:
                        current_y += max_row_height + padding
                        current_x = padding
                        max_row_height = 0
                    
                    # Create image container frame
                    img_frame = Frame(self.scrollable_frame, bg="white", relief="solid", bd=1)
                    img_frame.place(x=current_x, y=current_y, width=new_width + 4, height=max_height + 25)
                    
                    # Create image label
                    img_label = Label(img_frame, image=photo, bg="white")
                    img_label.pack(pady=2)
                    
                    # Create close button
                    close_btn = ttk.Button(
                        img_frame,
                        text="×",
                        width=3,
                        command=lambda idx=img_index: self.remove_image(idx)
                    )
                    close_btn.pack()
                    
                    # Store references
                    self.image_widgets.append(img_frame)
                    self.close_buttons.append(close_btn)
                    
                    # Update position tracking
                    current_x += new_width + padding + 4
                    max_row_height = max(max_row_height, max_height + 25)
                    
                    debug_print(f"DEBUG: Displayed image {os.path.basename(img_path)} at position ({current_x - new_width - padding - 4}, {current_y})")
                    
                except Exception as e:
                    debug_print(f"ERROR: Failed to display image {img_path}: {e}")
                    continue
            
            # Update canvas scroll region
            total_height = current_y + max_row_height + padding
            self.scrollable_frame.configure(width=frame_width, height=total_height)
            self.canvas.configure(scrollregion=(0, 0, frame_width, total_height))
            
            # Restore scroll position
            try:
                self.canvas.yview_moveto(prev_pos[0])
            except:
                pass
            
            debug_print(f"DEBUG: Image display completed - {len(self.image_widgets)} images displayed")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to display images: {e}")
    
    def remove_image(self, index: int):
        """Remove an image from the display and file list."""
        debug_print(f"DEBUG: Removing image at index {index}")
        
        try:
            if 0 <= index < len(self.image_files):
                removed_file = self.image_files.pop(index)
                debug_print(f"DEBUG: Removed image: {os.path.basename(removed_file)}")
                
                # Remove from crop states
                if removed_file in self.image_crop_states:
                    del self.image_crop_states[removed_file]
                
                # Refresh display
                self.display_images()
                
                # Notify callback
                if self.on_images_selected:
                    self.on_images_selected(self.image_files)
            else:
                debug_print(f"ERROR: Invalid image index: {index}")
                
        except Exception as e:
            debug_print(f"ERROR: Failed to remove image: {e}")
    
    def clear_images(self):
        """Clear all loaded images."""
        debug_print("DEBUG: Clearing all images from ImageLoader")
        
        try:
            self.image_files.clear()
            self.image_crop_states.clear()
            self.display_images()  # Refresh to show empty state
            
            if self.on_images_selected:
                self.on_images_selected(self.image_files)
                
        except Exception as e:
            debug_print(f"ERROR: Failed to clear images: {e}")
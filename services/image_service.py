"""
services/image_service.py
Consolidated image processing and management service.
This consolidates all image operations from image_loader.py and related modules.
"""

# Standard library imports
import os
import base64
import tempfile
import threading
import json
import time
from typing import Optional, Dict, Any, List, Tuple, Union, Callable
from pathlib import Path
from dataclasses import dataclass

# Third party imports (lazy loaded)
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def debug_print(message: str):
    """Debug print function for image operations."""
    print(f"DEBUG: ImageService - {message}")


@dataclass
class ImageMetadata:
    """Metadata for an image."""
    image_path: str
    sheet_name: str = ""
    image_type: str = "standard"  # standard, sample, plot, etc.
    display_name: str = ""
    creation_time: Optional[str] = None
    file_size: Optional[int] = None
    dimensions: Optional[Tuple[int, int]] = None
    crop_enabled: bool = False


@dataclass
class CropState:
    """Crop state information for an image."""
    crop_enabled: bool = False
    crop_coordinates: Optional[Tuple[int, int, int, int]] = None
    crop_method: str = "none"  # none, intelligent, manual, smart_fallback
    crop_timestamp: Optional[str] = None


class ImageService:
    """
    Consolidated service for image processing and management.
    Handles image loading, processing, cropping, caching, and sample image management.
    """
    
    def __init__(self, claude_client=None):
        """Initialize the image service."""
        debug_print("Initializing ImageService")
        
        # Core configuration
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.pdf']
        self.max_display_height = 135
        self.cache_size_limit = 100  # Maximum cached images
        
        # AI integration
        self.claude_client = claude_client
        self.enable_intelligent_crop = True
        
        # Caching system
        self.processed_image_cache = {}  # Cache for processed images
        self.image_metadata_cache = {}   # Cache for image metadata
        self.processing_status = {}      # Track processing status
        
        # Sample image management
        self.sample_images = {}          # {sample_id: [image_paths]}
        self.sample_image_crop_states = {}  # {image_path: crop_state}
        self.sample_header_data = {}     # {sample_id: header_data}
        
        # Sheet image management
        self.sheet_images = {}           # {file_name: {sheet_name: [image_paths]}}
        self.image_crop_states = {}      # {image_path: crop_enabled}
        
        # Pending operations
        self.pending_sample_images = {}
        self.pending_sample_image_crop_states = {}
        self.pending_sample_header_data = {}
        
        # Threading locks
        self.cache_lock = threading.Lock()
        self.processing_lock = threading.Lock()
        
        # Lazy-loaded dependencies
        self._pil_modules = None
        self._cv2 = None
        self._numpy = None
        
        debug_print("ImageService initialized successfully")
        debug_print(f"Supported formats: {', '.join(self.supported_formats)}")
        debug_print(f"Intelligent crop enabled: {self.enable_intelligent_crop}")
    
    # ===================== DEPENDENCY MANAGEMENT =====================
    
    def _lazy_import_pil(self):
        """Lazy import PIL modules."""
        if self._pil_modules is None:
            try:
                from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
                self._pil_modules = (Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance)
                debug_print("PIL modules loaded successfully")
            except ImportError:
                debug_print("WARNING: PIL not available")
                self._pil_modules = (None, None, None, None, None, None)
        return self._pil_modules
    
    def _lazy_import_cv2(self):
        """Lazy import OpenCV."""
        if self._cv2 is None:
            try:
                import cv2
                self._cv2 = cv2
                debug_print("OpenCV loaded successfully")
            except ImportError:
                debug_print("WARNING: OpenCV not available")
                self._cv2 = None
        return self._cv2
    
    def _lazy_import_numpy(self):
        """Lazy import NumPy."""
        if self._numpy is None:
            try:
                import numpy as np
                self._numpy = np
                debug_print("NumPy loaded successfully")
            except ImportError:
                debug_print("WARNING: NumPy not available")
                self._numpy = None
        return self._numpy
    
    # ===================== IMAGE LOADING AND PROCESSING =====================
    
    def load_image(self, image_path: str, target_size: Optional[Tuple[int, int]] = None) -> Optional[Any]:
        """Load an image with optional resizing."""
        debug_print(f"Loading image: {os.path.basename(image_path)}")
        
        try:
            Image, _, _, _, _, _ = self._lazy_import_pil()
            if Image is None:
                debug_print("ERROR: PIL not available for image loading")
                return None
            
            if not os.path.exists(image_path):
                debug_print(f"ERROR: Image file not found: {image_path}")
                return None
            
            # Load image
            with Image.open(image_path) as img:
                # Convert to RGB if necessary
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Resize if target size specified
                if target_size:
                    img = img.resize(target_size, Image.Resampling.LANCZOS)
                
                # Store metadata
                self._store_image_metadata(image_path, img)
                
                return img.copy()
                
        except Exception as e:
            debug_print(f"ERROR: Failed to load image {image_path}: {e}")
            return None
    
    def process_image(self, image_path: str, crop_enabled: bool = False, 
                     target_size: Optional[Tuple[int, int]] = None) -> Optional[Any]:
        """Process an image with optional cropping and resizing."""
        debug_print(f"Processing image: {os.path.basename(image_path)}, crop={crop_enabled}")
        
        try:
            # Create cache key
            cache_key = f"{image_path}_{crop_enabled}_{target_size}"
            
            # Check cache first
            with self.cache_lock:
                if cache_key in self.processed_image_cache:
                    debug_print(f"Using cached processed image for {os.path.basename(image_path)}")
                    return self.processed_image_cache[cache_key].copy()
            
            # Load original image
            original_img = self.load_image(image_path)
            if original_img is None:
                return None
            
            processed_img = original_img.copy()
            
            # Apply cropping if requested
            if crop_enabled:
                cropped_img = self.apply_intelligent_crop(processed_img, image_path)
                if cropped_img:
                    processed_img = cropped_img
            
            # Apply target sizing
            if target_size:
                processed_img = processed_img.resize(target_size, processed_img.Resampling.LANCZOS)
            elif not target_size and crop_enabled:
                # Default display resizing for cropped images
                processed_img = self._resize_for_display(processed_img)
            
            # Cache the processed image
            with self.cache_lock:
                self.processed_image_cache[cache_key] = processed_img.copy()
                self._manage_cache_size()
            
            # Update processing status
            self.processing_status[image_path] = crop_enabled
            
            debug_print(f"Processed image {os.path.basename(image_path)}: {processed_img.size}")
            return processed_img
            
        except Exception as e:
            debug_print(f"ERROR: Failed to process image {image_path}: {e}")
            return None
    
    def _resize_for_display(self, img) -> Any:
        """Resize image for display with consistent height."""
        try:
            scaling_factor = self.max_display_height / img.height
            new_width = int(img.width * scaling_factor)
            return img.resize((new_width, self.max_display_height), img.Resampling.LANCZOS)
        except Exception as e:
            debug_print(f"ERROR: Failed to resize image for display: {e}")
            return img
    
    def _store_image_metadata(self, image_path: str, img) -> None:
        """Store metadata for an image."""
        try:
            metadata = ImageMetadata(
                image_path=image_path,
                display_name=os.path.basename(image_path),
                creation_time=time.strftime('%Y-%m-%d %H:%M:%S'),
                file_size=os.path.getsize(image_path),
                dimensions=(img.width, img.height)
            )
            self.image_metadata_cache[image_path] = metadata
        except Exception as e:
            debug_print(f"WARNING: Failed to store metadata for {image_path}: {e}")
    
    # ===================== INTELLIGENT CROPPING =====================
    
    def apply_intelligent_crop(self, img, image_path: str) -> Optional[Any]:
        """Apply intelligent cropping to an image."""
        debug_print(f"Applying intelligent crop to {os.path.basename(image_path)}")
        
        try:
            # Try Claude-based intelligent cropping first
            if self.claude_client and self.enable_intelligent_crop:
                cropped_img = self._claude_intelligent_crop(img, image_path)
                if cropped_img:
                    debug_print("Claude intelligent crop successful")
                    return cropped_img
            
            # Fallback to computer vision cropping
            debug_print("Using fallback smart crop")
            return self._smart_crop_fallback(img, image_path)
            
        except Exception as e:
            debug_print(f"ERROR: Intelligent crop failed: {e}")
            return img  # Return original if cropping fails
    
    def _claude_intelligent_crop(self, img, image_path: str) -> Optional[Any]:
        """Use Claude API for intelligent image cropping."""
        debug_print(f"Claude intelligent crop for {os.path.basename(image_path)}")
        
        try:
            if not self.claude_client:
                debug_print("No Claude client available")
                return None
            
            Image, _, _, _, _, _ = self._lazy_import_pil()
            if Image is None:
                return None
            
            # Convert image to base64 for API
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                temp_path = temp_file.name
                
                # Scale image for API if too large
                api_img = img.copy()
                scale_factor = 1.0
                max_dimension = 1024
                
                if max(api_img.width, api_img.height) > max_dimension:
                    if api_img.width > api_img.height:
                        scale_factor = max_dimension / api_img.width
                        new_size = (max_dimension, int(api_img.height * scale_factor))
                    else:
                        scale_factor = max_dimension / api_img.height
                        new_size = (int(api_img.width * scale_factor), max_dimension)
                    
                    api_img = api_img.resize(new_size, Image.Resampling.LANCZOS)
                    debug_print(f"Scaled image for API: {api_img.size}, scale_factor: {scale_factor}")
                
                api_img.save(temp_path, 'PNG')
                
                # Encode for API
                with open(temp_path, 'rb') as f:
                    image_data = base64.b64encode(f.read()).decode('utf-8')
                
                # Clean up temp file
                os.unlink(temp_path)
            
            # Call Claude API
            prompt = """
            Please analyze this image and determine the optimal crop boundaries to focus on the main subject.
            The image appears to be from a scientific/testing context.
            
            Return ONLY a JSON object with the crop boundaries in this exact format:
            {"left": x1, "top": y1, "right": x2, "bottom": y2}
            
            The coordinates should be pixel values relative to the image dimensions.
            """
            
            # This would be the actual Claude API call
            # For now, using a placeholder response structure
            response = self._call_claude_api(prompt, image_data)
            
            if response and 'crop_box' in response:
                crop_box = response['crop_box']
                
                # Scale coordinates back to original image size
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
                width, height = img.size
                left = max(0, min(left, width - 1))
                top = max(0, min(top, height - 1))
                right = max(left + 1, min(right, width))
                bottom = max(top + 1, min(bottom, height))
                
                debug_print(f"Claude crop box: ({left}, {top}, {right}, {bottom})")
                cropped_img = img.crop((left, top, right, bottom))
                debug_print(f"Claude cropped image size: {cropped_img.size}")
                return cropped_img
            
        except Exception as e:
            debug_print(f"ERROR: Claude intelligent crop failed: {e}")
        
        return None
    
    def _call_claude_api(self, prompt: str, image_data: str) -> Optional[Dict]:
        """Call Claude API for image analysis. Placeholder implementation."""
        debug_print("Calling Claude API for image analysis")
        
        try:
            # This is a placeholder for the actual Claude API call
            # In the real implementation, you would use your Claude client here
            # 
            # response = self.claude_client.analyze_image(prompt, image_data)
            # 
            # For now, returning None to trigger fallback
            return None
            
        except Exception as e:
            debug_print(f"ERROR: Claude API call failed: {e}")
            return None
    
    def _smart_crop_fallback(self, img, image_path: str) -> Any:
        """Fallback cropping using computer vision techniques."""
        debug_print(f"Smart crop fallback for {os.path.basename(image_path)}")
        
        try:
            cv2 = self._lazy_import_cv2()
            np = self._lazy_import_numpy()
            
            if cv2 is None or np is None:
                debug_print("OpenCV/NumPy not available, using center crop")
                return self._center_crop(img)
            
            # Convert to numpy array for OpenCV processing
            img_array = np.array(img)
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            
            # Apply multiple edge detection techniques
            edges1 = cv2.Canny(gray, 50, 150)
            edges2 = cv2.Canny(gray, 100, 200)
            edges = cv2.bitwise_or(edges1, edges2)
            
            # Find contours
            contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            if not contours:
                debug_print("No contours found, using center crop")
                return self._center_crop(img)
            
            # Find the most significant contour (largest area)
            largest_contour = max(contours, key=cv2.contourArea)
            x, y, w, h = cv2.boundingRect(largest_contour)
            
            # Add margin
            margin_percent = 5
            width, height = img.size
            margin_x = int(width * margin_percent / 100)
            margin_y = int(height * margin_percent / 100)
            
            crop_box = (
                max(0, x - margin_x),
                max(0, y - margin_y),
                min(width, x + w + margin_x),
                min(height, y + h + margin_y)
            )
            
            debug_print(f"Smart crop box: {crop_box}")
            return img.crop(crop_box)
            
        except Exception as e:
            debug_print(f"ERROR: Smart crop fallback failed: {e}")
            return self._center_crop(img)
    
    def _center_crop(self, img, crop_factor: float = 0.8) -> Any:
        """Simple center crop as last resort."""
        try:
            width, height = img.size
            crop_size_w = int(width * crop_factor)
            crop_size_h = int(height * crop_factor)
            left = (width - crop_size_w) // 2
            top = (height - crop_size_h) // 2
            
            crop_box = (left, top, left + crop_size_w, top + crop_size_h)
            debug_print(f"Center crop box: {crop_box}")
            return img.crop(crop_box)
        except Exception as e:
            debug_print(f"ERROR: Center crop failed: {e}")
            return img
    
    # ===================== SAMPLE IMAGE MANAGEMENT =====================
    
    def add_sample_images(self, sample_images: Dict[str, List[str]], 
                         sample_crop_states: Dict[str, bool] = None,
                         sample_header_data: Dict[str, Any] = None) -> bool:
        """Add sample images with metadata."""
        debug_print(f"Adding sample images for {len(sample_images)} samples")
        
        try:
            if sample_crop_states is None:
                sample_crop_states = {}
            if sample_header_data is None:
                sample_header_data = {}
            
            # Store in pending operations
            self.pending_sample_images.update(sample_images)
            self.pending_sample_image_crop_states.update(sample_crop_states)
            self.pending_sample_header_data.update(sample_header_data)
            
            # Process each sample's images
            for sample_id, image_paths in sample_images.items():
                debug_print(f"Processing {len(image_paths)} images for {sample_id}")
                
                # Validate image paths
                valid_paths = []
                for img_path in image_paths:
                    if os.path.exists(img_path):
                        valid_paths.append(img_path)
                        
                        # Set crop state if provided
                        if img_path in sample_crop_states:
                            self.sample_image_crop_states[img_path] = sample_crop_states[img_path]
                    else:
                        debug_print(f"WARNING: Image not found: {img_path}")
                
                if valid_paths:
                    self.sample_images[sample_id] = valid_paths
                    
                    # Store header data if provided
                    if sample_id in sample_header_data:
                        self.sample_header_data[sample_id] = sample_header_data[sample_id]
            
            debug_print(f"Successfully added sample images for {len(self.sample_images)} samples")
            return True
            
        except Exception as e:
            debug_print(f"ERROR: Failed to add sample images: {e}")
            return False
    
    def get_sample_images(self, sample_id: str) -> List[str]:
        """Get image paths for a specific sample."""
        return self.sample_images.get(sample_id, [])
    
    def get_sample_images_summary(self) -> Tuple[Dict[str, int], int]:
        """Get a summary of all loaded sample images."""
        summary = {}
        total_images = 0
        
        for sample_id, images in self.sample_images.items():
            summary[sample_id] = len(images)
            total_images += len(images)
        
        debug_print(f"Sample images summary: {summary} (Total: {total_images})")
        return summary, total_images
    
    def load_sample_images_from_vap3(self, vap_data: Dict[str, Any]) -> bool:
        """Load sample images from VAP3 data."""
        debug_print("Loading sample images from VAP3 data")
        
        try:
            sample_images = vap_data.get('sample_images', {})
            sample_crop_states = vap_data.get('sample_image_crop_states', {})
            sample_header_data = vap_data.get('sample_header_data', {})
            
            if sample_images:
                debug_print(f"Found {len(sample_images)} samples with images in VAP3")
                return self.add_sample_images(sample_images, sample_crop_states, sample_header_data)
            else:
                debug_print("No sample images found in VAP3 data")
                return True
                
        except Exception as e:
            debug_print(f"ERROR: Failed to load sample images from VAP3: {e}")
            return False
    
    # ===================== SHEET IMAGE MANAGEMENT =====================
    
    def store_images_for_sheet(self, file_name: str, sheet_name: str, 
                              image_paths: List[str], crop_states: Dict[str, bool] = None) -> bool:
        """Store image paths and crop states for a specific sheet."""
        debug_print(f"Storing {len(image_paths)} images for sheet {sheet_name}")
        
        try:
            if crop_states is None:
                crop_states = {}
            
            # Initialize file structure if needed
            if file_name not in self.sheet_images:
                self.sheet_images[file_name] = {}
            
            # Store image paths
            self.sheet_images[file_name][sheet_name] = image_paths.copy()
            
            # Store crop states
            for img_path in image_paths:
                if img_path in crop_states:
                    self.image_crop_states[img_path] = crop_states[img_path]
                elif img_path not in self.image_crop_states:
                    self.image_crop_states[img_path] = False  # Default state
            
            debug_print(f"Stored images for {file_name}/{sheet_name}: {len(image_paths)} images")
            return True
            
        except Exception as e:
            debug_print(f"ERROR: Failed to store images for sheet: {e}")
            return False
    
    def get_sheet_images(self, file_name: str, sheet_name: str) -> List[str]:
        """Get image paths for a specific sheet."""
        try:
            return self.sheet_images.get(file_name, {}).get(sheet_name, [])
        except Exception as e:
            debug_print(f"ERROR: Failed to get sheet images: {e}")
            return []
    
    def get_all_sheet_images(self, file_name: str) -> Dict[str, List[str]]:
        """Get all sheet images for a file."""
        return self.sheet_images.get(file_name, {})
    
    # ===================== IMAGE SELECTION AND UI INTEGRATION =====================
    
    def select_images_dialog(self, title: str = "Select Images") -> List[str]:
        """Show file dialog for image selection."""
        debug_print("Opening image selection dialog")
        
        try:
            file_types = [
                ("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
                ("PDF Files", "*.pdf"),
                ("All Files", "*.*")
            ]
            
            selected_files = filedialog.askopenfilenames(title=title, filetypes=file_types)
            
            if selected_files:
                debug_print(f"Selected {len(selected_files)} images")
                return list(selected_files)
            else:
                debug_print("No images selected")
                return []
                
        except Exception as e:
            debug_print(f"ERROR: Image selection dialog failed: {e}")
            return []
    
    def get_image_for_report(self, image_path: str, target_width: int = None, 
                           target_height: int = None) -> Optional[Any]:
        """Get processed image for report generation."""
        debug_print(f"Getting image for report: {os.path.basename(image_path)}")
        
        try:
            # Check if we have a processed version in cache
            crop_enabled = self.image_crop_states.get(image_path, False)
            
            # Load processed image
            processed_img = self.process_image(image_path, crop_enabled)
            if processed_img is None:
                debug_print(f"Failed to process image for report: {image_path}")
                return None
            
            # Resize for report if dimensions specified
            if target_width or target_height:
                original_width, original_height = processed_img.size
                
                if target_width and target_height:
                    processed_img = processed_img.resize((target_width, target_height), 
                                                       processed_img.Resampling.LANCZOS)
                elif target_width:
                    ratio = target_width / original_width
                    new_height = int(original_height * ratio)
                    processed_img = processed_img.resize((target_width, new_height), 
                                                       processed_img.Resampling.LANCZOS)
                elif target_height:
                    ratio = target_height / original_height
                    new_width = int(original_width * ratio)
                    processed_img = processed_img.resize((new_width, target_height), 
                                                       processed_img.Resampling.LANCZOS)
                
                debug_print(f"Resized image for report to: {processed_img.size}")
            
            return processed_img
            
        except Exception as e:
            debug_print(f"ERROR: Failed to get image for report: {e}")
            # Fallback to original image
            try:
                return self.load_image(image_path)
            except:
                return None
    
    # ===================== CACHE MANAGEMENT =====================
    
    def _manage_cache_size(self) -> None:
        """Manage cache size to prevent memory issues."""
        if len(self.processed_image_cache) > self.cache_size_limit:
            # Remove oldest entries (simple FIFO)
            items_to_remove = len(self.processed_image_cache) - self.cache_size_limit + 10
            keys_to_remove = list(self.processed_image_cache.keys())[:items_to_remove]
            
            for key in keys_to_remove:
                del self.processed_image_cache[key]
            
            debug_print(f"Cache cleanup: removed {len(keys_to_remove)} entries")
    
    def clear_cache(self) -> None:
        """Clear all cached data."""
        debug_print("Clearing image cache")
        
        with self.cache_lock:
            self.processed_image_cache.clear()
            self.image_metadata_cache.clear()
            self.processing_status.clear()
        
        debug_print("Image cache cleared")
    
    def remove_image(self, image_path: str) -> bool:
        """Remove an image from all storage and cache."""
        debug_print(f"Removing image: {os.path.basename(image_path)}")
        
        try:
            # Remove from crop states
            if image_path in self.image_crop_states:
                del self.image_crop_states[image_path]
            
            if image_path in self.sample_image_crop_states:
                del self.sample_image_crop_states[image_path]
            
            # Remove from cache
            with self.cache_lock:
                cache_keys_to_remove = [key for key in self.processed_image_cache.keys() 
                                      if key.startswith(image_path)]
                for cache_key in cache_keys_to_remove:
                    del self.processed_image_cache[cache_key]
            
            # Remove from processing status
            if image_path in self.processing_status:
                del self.processing_status[image_path]
            
            # Remove from metadata cache
            if image_path in self.image_metadata_cache:
                del self.image_metadata_cache[image_path]
            
            # Remove from sample images
            for sample_id, image_paths in self.sample_images.items():
                if image_path in image_paths:
                    image_paths.remove(image_path)
            
            # Remove from sheet images
            for file_name, sheets in self.sheet_images.items():
                for sheet_name, image_paths in sheets.items():
                    if image_path in image_paths:
                        image_paths.remove(image_path)
            
            debug_print(f"Successfully removed image: {os.path.basename(image_path)}")
            return True
            
        except Exception as e:
            debug_print(f"ERROR: Failed to remove image: {e}")
            return False
    
    # ===================== VALIDATION AND UTILITIES =====================
    
    def validate_image_path(self, image_path: str) -> Tuple[bool, str]:
        """Validate if an image path is accessible and supported."""
        if not os.path.exists(image_path):
            return False, f"Image file does not exist: {image_path}"
        
        file_ext = Path(image_path).suffix.lower()
        if file_ext not in self.supported_formats:
            return False, f"Unsupported image format: {file_ext}"
        
        try:
            # Try to open the image to validate format
            Image, _, _, _, _, _ = self._lazy_import_pil()
            if Image:
                with Image.open(image_path) as img:
                    # Just opening it validates the format
                    pass
            return True, "Valid image"
        except Exception as e:
            return False, f"Invalid image file: {e}"
    
    def get_image_info(self, image_path: str) -> Dict[str, Any]:
        """Get comprehensive information about an image."""
        try:
            # Check cache first
            if image_path in self.image_metadata_cache:
                metadata = self.image_metadata_cache[image_path]
                return {
                    'filename': metadata.display_name,
                    'path': metadata.image_path,
                    'size': metadata.file_size,
                    'dimensions': metadata.dimensions,
                    'creation_time': metadata.creation_time,
                    'crop_enabled': self.image_crop_states.get(image_path, False),
                    'processed': image_path in self.processing_status
                }
            
            # Get info from file system
            path = Path(image_path)
            stat = path.stat()
            
            info = {
                'filename': path.name,
                'path': str(path),
                'size': stat.st_size,
                'modified_time': stat.st_mtime,
                'extension': path.suffix.lower(),
                'exists': path.exists(),
                'crop_enabled': self.image_crop_states.get(image_path, False),
                'processed': image_path in self.processing_status
            }
            
            # Try to get image dimensions
            try:
                Image, _, _, _, _, _ = self._lazy_import_pil()
                if Image:
                    with Image.open(image_path) as img:
                        info['dimensions'] = img.size
                        info['mode'] = img.mode
            except:
                info['dimensions'] = None
                info['mode'] = 'unknown'
            
            return info
            
        except Exception as e:
            debug_print(f"Failed to get image info: {e}")
            return {
                'filename': Path(image_path).name if image_path else 'Unknown',
                'exists': False,
                'error': str(e)
            }
    
    def get_service_status(self) -> Dict[str, Any]:
        """Get comprehensive service status information."""
        return {
            'cache_size': len(self.processed_image_cache),
            'cache_limit': self.cache_size_limit,
            'metadata_cache_size': len(self.image_metadata_cache),
            'sample_images_count': len(self.sample_images),
            'sheet_images_count': sum(len(sheets) for sheets in self.sheet_images.values()),
            'total_crop_states': len(self.image_crop_states),
            'intelligent_crop_enabled': self.enable_intelligent_crop,
            'claude_client_available': self.claude_client is not None,
            'pil_available': self._lazy_import_pil()[0] is not None,
            'opencv_available': self._lazy_import_cv2() is not None,
            'numpy_available': self._lazy_import_numpy() is not None,
            'supported_formats': self.supported_formats,
            'pending_sample_images': len(self.pending_sample_images)
        }
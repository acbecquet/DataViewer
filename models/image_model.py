# models/image_model.py
"""
models/image_model.py
Image model for managing images, processing, cropping, and metadata.
Consolidated from image_loader.py, data_collection_window.py image functionality, and image processing operations.
"""

import os
import base64
import io
import json
import re
import tempfile
import threading
import time
import uuid
from typing import Optional, Dict, Any, List, Tuple, Union, Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, Canvas, Scrollbar, Frame, Label, PhotoImage, messagebox, ttk


def debug_print(message: str):
    """Debug print function for image operations."""
    print(f"DEBUG: ImageModel - {message}")


@dataclass
class CropState:
    """Model for image crop state and coordinates."""
    x1: int = 0
    y1: int = 0
    x2: int = 0
    y2: int = 0
    is_cropped: bool = False
    crop_enabled: bool = False
    crop_method: str = "none"  # none, intelligent, manual, smart_fallback, center
    crop_timestamp: Optional[datetime] = None
    
    def __post_init__(self):
        """Post-initialization processing."""
        if self.is_cropped:
            debug_print(f"Created CropState with crop ({self.x1},{self.y1}) to ({self.x2},{self.y2}) using {self.crop_method}")
        else:
            debug_print("Created CropState with no crop")
    
    def set_crop(self, x1: int, y1: int, x2: int, y2: int, method: str = "manual"):
        """Set crop coordinates and method."""
        self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2
        self.is_cropped = True
        self.crop_method = method
        self.crop_timestamp = datetime.now()
        debug_print(f"Set crop coordinates ({x1},{y1}) to ({x2},{y2}) using {method}")
    
    def clear_crop(self):
        """Clear crop coordinates."""
        self.x1 = self.y1 = self.x2 = self.y2 = 0
        self.is_cropped = False
        self.crop_method = "none"
        self.crop_timestamp = None
        debug_print("Cleared crop coordinates")
    
    def get_crop_box(self) -> Tuple[int, int, int, int]:
        """Get crop coordinates as tuple."""
        return (self.x1, self.y1, self.x2, self.y2)


@dataclass
class ImageMetadata:
    """Enhanced metadata for images associated with sheets and samples."""
    image_path: str
    sheet_name: str
    image_type: str = "sample"  # sample, plot, reference, standard
    display_name: Optional[str] = None
    crop_state: CropState = field(default_factory=CropState)
    thumbnail_path: Optional[str] = None
    file_size: int = 0
    dimensions: Optional[Tuple[int, int]] = None
    created_at: datetime = field(default_factory=datetime.now)
    last_processed: Optional[datetime] = None
    processing_status: str = "none"  # none, processing, processed, failed
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    # Sample-specific metadata
    sample_id: Optional[str] = None
    test_name: Optional[str] = None
    media_type: Optional[str] = None
    viscosity: Optional[float] = None
    header_data: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Post-initialization processing."""
        if self.display_name is None:
            self.display_name = Path(self.image_path).name
        
        # Get file size if file exists
        if Path(self.image_path).exists():
            try:
                self.file_size = Path(self.image_path).stat().st_size
            except:
                self.file_size = 0
        
        debug_print(f"Created ImageMetadata for '{self.display_name}' in sheet '{self.sheet_name}'")
        debug_print(f"Image path: {self.image_path}, Size: {self.file_size} bytes, Type: {self.image_type}")
    
    def update_processing_status(self, status: str):
        """Update processing status and timestamp."""
        self.processing_status = status
        self.last_processed = datetime.now()
        debug_print(f"Updated processing status to '{status}' for {self.display_name}")


@dataclass
class ImageCache:
    """Cache entry for processed images."""
    original_path: str
    processed_image: Any  # PIL Image object
    cache_key: str
    created_at: datetime = field(default_factory=datetime.now)
    access_count: int = 0
    last_accessed: Optional[datetime] = None
    file_size_bytes: int = 0
    crop_applied: bool = False
    crop_method: str = "none"
    
    def access_cache(self):
        """Update access statistics."""
        self.access_count += 1
        self.last_accessed = datetime.now()
    
    def get_cache_info(self) -> Dict[str, Any]:
        """Get cache information."""
        return {
            'original_path': self.original_path,
            'cache_key': self.cache_key,
            'created_at': self.created_at,
            'access_count': self.access_count,
            'last_accessed': self.last_accessed,
            'file_size_bytes': self.file_size_bytes,
            'crop_applied': self.crop_applied,
            'crop_method': self.crop_method
        }


class ImageModel:
    """
    Main image model for managing images, processing, cropping, and metadata.
    Consolidated from image_loader.py, data_collection_window.py, and image processing functionality.
    """
    
    def __init__(self):
        """Initialize the image model."""
        # Core image storage
        self.sheet_images: Dict[str, Dict[str, List[ImageMetadata]]] = {}  # file -> sheet -> images
        self.sample_images: Dict[str, List[ImageMetadata]] = {}  # sample_id -> images
        self.all_images: Dict[str, ImageMetadata] = {}  # image_path -> metadata
        
        # Crop state management
        self.image_crop_states: Dict[str, CropState] = {}  # image_path -> crop_state
        self.global_crop_enabled: bool = False
        
        # Image processing and caching
        self.processed_image_cache: Dict[str, ImageCache] = {}  # cache_key -> cached_image
        self.processing_status: Dict[str, str] = {}  # image_path -> status
        
        # Sample image management (from data_collection_window.py)
        self.pending_sample_images: Dict[str, List[str]] = {}  # sample_id -> image_paths
        self.pending_sample_image_crop_states: Dict[str, bool] = {}  # image_path -> crop_enabled
        self.pending_sample_header_data: Dict[str, Any] = {}  # sample_id -> header_data
        self.sample_image_metadata: Dict[str, Dict[str, Any]] = {}  # file -> sheet -> metadata
        
        # UI state management
        self.image_files: List[str] = []  # Current list of image files
        self.image_widgets: List[Any] = []  # UI widgets for images
        self.close_buttons: List[Any] = []  # Close button widgets
        self.image_references: List[Any] = []  # Keep references to prevent garbage collection
        
        # Configuration
        self.max_display_height: int = 135
        self.cache_size_limit: int = 100
        self.cache_duration_hours: int = 24
        self.supported_formats: List[str] = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.pdf']
        
        # Claude API integration
        self.claude_client: Optional[Any] = None
        self.enable_intelligent_crop: bool = True
        self._init_claude_api()
        
        # Lazy loading infrastructure
        self._pil_loaded = False
        self._cv2_loaded = False
        self._numpy_loaded = False
        
        # PIL components (loaded lazily)
        self.Image = None
        self.ImageTk = None
        self.ImageFilter = None
        self.ImageOps = None
        self.ImageChops = None
        self.ImageEnhance = None
        
        # Threading support
        self.cache_lock = threading.Lock()
        self.processing_lock = threading.Lock()
        
        debug_print("ImageModel initialized")
        debug_print(f"Supported formats: {', '.join(self.supported_formats)}")
        debug_print(f"Intelligent crop enabled: {self.enable_intelligent_crop}")
    
    # ===================== INITIALIZATION AND DEPENDENCIES =====================
    
    def _init_claude_api(self):
        """Initialize Claude API client for intelligent cropping."""
        try:
            import anthropic
            api_key = os.getenv('ANTHROPIC_API_KEY')
            if api_key:
                self.claude_client = anthropic.Anthropic(api_key=api_key)
                debug_print("Claude API client initialized successfully")
            else:
                debug_print("ANTHROPIC_API_KEY not found, Claude cropping disabled")
        except ImportError:
            debug_print("anthropic package not available, Claude cropping disabled")
        except Exception as e:
            debug_print(f"Claude API initialization failed: {e}")
    
    def _lazy_import_pil(self):
        """Lazy import PIL modules."""
        if self._pil_loaded:
            return self.Image, self.ImageTk, self.ImageFilter, self.ImageOps, self.ImageChops, self.ImageEnhance
        
        try:
            from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
            
            self.Image = Image
            self.ImageTk = ImageTk
            self.ImageFilter = ImageFilter
            self.ImageOps = ImageOps
            self.ImageChops = ImageChops
            self.ImageEnhance = ImageEnhance
            self._pil_loaded = True
            
            debug_print("PIL modules loaded successfully")
            return Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
            
        except ImportError as e:
            debug_print(f"Error importing PIL: {e}")
            return None, None, None, None, None, None
    
    def _lazy_import_cv2(self):
        """Lazy import OpenCV."""
        if self._cv2_loaded:
            return self.cv2
        
        try:
            import cv2
            self.cv2 = cv2
            self._cv2_loaded = True
            debug_print("OpenCV loaded successfully")
            return cv2
        except ImportError:
            debug_print("OpenCV not available")
            return None
    
    def _lazy_import_numpy(self):
        """Lazy import NumPy."""
        if self._numpy_loaded:
            return self.np
        
        try:
            import numpy as np
            self.np = np
            self._numpy_loaded = True
            debug_print("NumPy loaded successfully")
            return np
        except ImportError:
            debug_print("NumPy not available")
            return None
    
    # ===================== IMAGE METADATA MANAGEMENT =====================
    
    def add_image_metadata(self, image_path: str, sheet_name: str, image_type: str = "sample",
                          sample_id: str = None, test_name: str = None, 
                          header_data: Dict[str, Any] = None) -> ImageMetadata:
        """Add image metadata to the model."""
        # Create metadata
        metadata = ImageMetadata(
            image_path=image_path,
            sheet_name=sheet_name,
            image_type=image_type,
            sample_id=sample_id,
            test_name=test_name,
            header_data=header_data or {}
        )
        
        # Store in all_images registry
        self.all_images[image_path] = metadata
        
        # Initialize crop state if not exists
        if image_path not in self.image_crop_states:
            self.image_crop_states[image_path] = CropState(crop_enabled=self.global_crop_enabled)
        
        metadata.crop_state = self.image_crop_states[image_path]
        
        debug_print(f"Added image metadata: {metadata.display_name} ({image_type}) to sheet {sheet_name}")
        return metadata
    
    def get_image_metadata(self, image_path: str) -> Optional[ImageMetadata]:
        """Get image metadata by path."""
        return self.all_images.get(image_path)
    
    def update_image_metadata(self, image_path: str, **kwargs):
        """Update image metadata fields."""
        if image_path in self.all_images:
            metadata = self.all_images[image_path]
            for key, value in kwargs.items():
                if hasattr(metadata, key):
                    setattr(metadata, key, value)
                    debug_print(f"Updated metadata {key} = {value} for {metadata.display_name}")
                else:
                    # Store in metadata dict if not a direct attribute
                    metadata.metadata[key] = value
    
    # ===================== SHEET IMAGE MANAGEMENT =====================
    
    def add_sheet_image(self, file_name: str, sheet_name: str, image_path: str, 
                       image_type: str = "standard") -> ImageMetadata:
        """Add an image to a specific sheet."""
        # Initialize nested dictionaries
        if file_name not in self.sheet_images:
            self.sheet_images[file_name] = {}
        if sheet_name not in self.sheet_images[file_name]:
            self.sheet_images[file_name][sheet_name] = []
        
        # Create metadata
        metadata = self.add_image_metadata(image_path, sheet_name, image_type)
        
        # Add to sheet images
        self.sheet_images[file_name][sheet_name].append(metadata)
        
        debug_print(f"Added image to file '{file_name}' sheet '{sheet_name}' - total: {len(self.sheet_images[file_name][sheet_name])}")
        return metadata
    
    def get_sheet_images(self, file_name: str, sheet_name: str) -> List[ImageMetadata]:
        """Get all images for a specific sheet."""
        return self.sheet_images.get(file_name, {}).get(sheet_name, [])
    
    def remove_sheet_image(self, file_name: str, sheet_name: str, image_path: str) -> bool:
        """Remove an image from a sheet."""
        if file_name in self.sheet_images and sheet_name in self.sheet_images[file_name]:
            images = self.sheet_images[file_name][sheet_name]
            for i, metadata in enumerate(images):
                if metadata.image_path == image_path:
                    del images[i]
                    self._cleanup_image_data(image_path)
                    debug_print(f"Removed image from sheet {sheet_name}")
                    return True
        return False
    
    def clear_sheet_images(self, file_name: str, sheet_name: str):
        """Clear all images for a specific sheet."""
        if file_name in self.sheet_images and sheet_name in self.sheet_images[file_name]:
            images = self.sheet_images[file_name][sheet_name]
            for metadata in images:
                self._cleanup_image_data(metadata.image_path)
            
            image_count = len(images)
            del self.sheet_images[file_name][sheet_name]
            debug_print(f"Cleared {image_count} images from sheet '{sheet_name}' in file '{file_name}'")
    
    # ===================== SAMPLE IMAGE MANAGEMENT =====================
    
    def add_sample_image(self, sample_id: str, image_path: str, test_name: str = None,
                        header_data: Dict[str, Any] = None) -> ImageMetadata:
        """Add a sample image with metadata."""
        # Create metadata
        metadata = self.add_image_metadata(
            image_path=image_path,
            sheet_name=test_name or "sample",
            image_type="sample",
            sample_id=sample_id,
            test_name=test_name,
            header_data=header_data
        )
        
        # Add to sample images
        if sample_id not in self.sample_images:
            self.sample_images[sample_id] = []
        self.sample_images[sample_id].append(metadata)
        
        debug_print(f"Added sample image for sample '{sample_id}' - total: {len(self.sample_images[sample_id])}")
        return metadata
    
    def get_sample_images(self, sample_id: str) -> List[ImageMetadata]:
        """Get all images for a specific sample."""
        return self.sample_images.get(sample_id, [])
    
    def add_pending_sample_images(self, sample_images: Dict[str, List[str]], 
                                 sample_crop_states: Dict[str, bool] = None,
                                 sample_header_data: Dict[str, Any] = None):
        """Add pending sample images from data collection."""
        debug_print(f"Adding pending sample images for {len(sample_images)} samples")
        
        # Store pending data
        self.pending_sample_images.update(sample_images)
        
        if sample_crop_states:
            self.pending_sample_image_crop_states.update(sample_crop_states)
        
        if sample_header_data:
            self.pending_sample_header_data.update(sample_header_data)
        
        # Process immediately to create proper metadata
        self._process_pending_sample_images()
    
    def _process_pending_sample_images(self):
        """Process pending sample images into proper metadata structures."""
        for sample_id, image_paths in self.pending_sample_images.items():
            # Get header data for this sample
            header_data = self.pending_sample_header_data.get(sample_id, {})
            test_name = header_data.get('test', 'Unknown Test')
            
            for img_path in image_paths:
                if os.path.exists(img_path):
                    # Create sample image metadata
                    metadata = self.add_sample_image(
                        sample_id=sample_id,
                        image_path=img_path,
                        test_name=test_name,
                        header_data=header_data
                    )
                    
                    # Set crop state if available
                    if img_path in self.pending_sample_image_crop_states:
                        crop_enabled = self.pending_sample_image_crop_states[img_path]
                        metadata.crop_state.crop_enabled = crop_enabled
                        self.image_crop_states[img_path].crop_enabled = crop_enabled
        
        debug_print(f"Processed {len(self.pending_sample_images)} pending sample image groups")
    
    # ===================== CROP STATE MANAGEMENT =====================
    
    def set_crop_state(self, image_path: str, crop_state: CropState):
        """Set crop state for an image."""
        self.image_crop_states[image_path] = crop_state
        
        # Update metadata if exists
        if image_path in self.all_images:
            self.all_images[image_path].crop_state = crop_state
        
        debug_print(f"Set crop state for image: {Path(image_path).name}")
    
    def get_crop_state(self, image_path: str) -> Optional[CropState]:
        """Get crop state for an image."""
        return self.image_crop_states.get(image_path)
    
    def toggle_global_crop_mode(self) -> bool:
        """Toggle global crop mode and return new state."""
        self.global_crop_enabled = not self.global_crop_enabled
        debug_print(f"Toggled global crop mode to {self.global_crop_enabled}")
        return self.global_crop_enabled
    
    def set_image_crop_enabled(self, image_path: str, enabled: bool):
        """Enable or disable cropping for a specific image."""
        if image_path not in self.image_crop_states:
            self.image_crop_states[image_path] = CropState()
        
        self.image_crop_states[image_path].crop_enabled = enabled
        
        # Clear cache for this image so it gets reprocessed
        self._clear_image_cache(image_path)
        
        debug_print(f"Set crop enabled={enabled} for {Path(image_path).name}")
    
    # ===================== IMAGE PROCESSING AND CROPPING =====================
    
    def process_image(self, image_path: str, force_reprocess: bool = False) -> Optional[Any]:
        """Process an image with optional cropping based on crop state."""
        try:
            Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance = self._lazy_import_pil()
            if Image is None:
                debug_print("ERROR: PIL not available for image processing")
                return None
            
            debug_print(f"Processing image: {Path(image_path).name}")
            
            # Get crop state
            crop_state = self.get_crop_state(image_path)
            should_crop = crop_state.crop_enabled if crop_state else False
            
            # Create cache key
            cache_key = f"{image_path}_{should_crop}"
            
            # Check cache first (unless forcing reprocess)
            if not force_reprocess and cache_key in self.processed_image_cache:
                cached_entry = self.processed_image_cache[cache_key]
                cached_entry.access_cache()
                debug_print(f"Using cached processed image for {Path(image_path).name}")
                return cached_entry.processed_image.copy()
            
            # Load and process image
            with Image.open(image_path) as original_img:
                processed_img = original_img.copy()
                crop_method = "none"
                
                # Apply cropping if requested
                if should_crop:
                    cropped_img = self._apply_intelligent_crop(processed_img, image_path)
                    if cropped_img != processed_img:
                        processed_img = cropped_img
                        crop_method = crop_state.crop_method if crop_state else "intelligent"
                
                # Cache the processed image
                cache_entry = ImageCache(
                    original_path=image_path,
                    processed_image=processed_img.copy(),
                    cache_key=cache_key,
                    crop_applied=should_crop,
                    crop_method=crop_method
                )
                
                self.processed_image_cache[cache_key] = cache_entry
                debug_print(f"Cached processed image for {Path(image_path).name} (crop={should_crop})")
                
                # Update processing status
                self.processing_status[image_path] = "processed"
                if image_path in self.all_images:
                    self.all_images[image_path].update_processing_status("processed")
                
                return processed_img
                
        except Exception as e:
            debug_print(f"ERROR: Failed to process image {image_path}: {e}")
            self.processing_status[image_path] = "failed"
            return None
    
    def _apply_intelligent_crop(self, img: Any, image_path: str) -> Any:
        """Apply intelligent cropping to an image."""
        debug_print(f"Applying intelligent crop to {Path(image_path).name}")
        
        try:
            # Try Claude-based intelligent cropping first
            if self.claude_client and self.enable_intelligent_crop:
                cropped_img = self._claude_intelligent_crop(img, image_path)
                if cropped_img:
                    debug_print("Claude intelligent crop successful")
                    # Update crop state
                    if image_path in self.image_crop_states:
                        self.image_crop_states[image_path].crop_method = "intelligent"
                    return cropped_img
            
            # Fallback to computer vision cropping
            debug_print("Using fallback smart crop")
            cropped_img = self._smart_crop_fallback(img, image_path)
            if image_path in self.image_crop_states:
                self.image_crop_states[image_path].crop_method = "smart_fallback"
            return cropped_img
            
        except Exception as e:
            debug_print(f"ERROR: Intelligent crop failed: {e}")
            return img  # Return original if cropping fails
    
    def _claude_intelligent_crop(self, img: Any, image_path: str) -> Optional[Any]:
        """Use Claude API for intelligent image cropping."""
        if not self.claude_client:
            debug_print("Claude client not available")
            return None
        
        try:
            debug_print(f"Using Claude intelligent crop for {Path(image_path).name}")
            
            # Convert image to base64
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr = img_byte_arr.getvalue()
            image_data = base64.b64encode(img_byte_arr).decode('utf-8')
            
            # Resize for API if image is very large
            scale_factor = 1.0
            if max(img.size) > 1000:
                scale_factor = 1000 / max(img.size)
                new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
                resized_img = img.resize(new_size)
                
                # Re-encode resized image
                img_byte_arr = io.BytesIO()
                resized_img.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                image_data = base64.b64encode(img_byte_arr).decode('utf-8')
                debug_print(f"Resized image for API by factor {scale_factor}")
            
            # Call Claude API
            response = self.claude_client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=1000,
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
                            "text": "Analyze this image and identify the main subject that should be preserved when cropping. Return ONLY a JSON object with 'left', 'top', 'right', 'bottom' coordinates (as integers) that would crop to show just the main subject with minimal background. The coordinates should be relative to the image dimensions."
                        }
                    ]
                }]
            )
            
            # Parse response
            response_text = response.content[0].text
            debug_print(f"Claude response: {response_text}")
            
            # Extract JSON from response
            json_match = re.search(r'\{[^}]*\}', response_text)
            if json_match:
                crop_box = json.loads(json_match.group())
                debug_print(f"Parsed crop box: {crop_box}")
                
                # Scale coordinates back if needed
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
                
                debug_print(f"Final crop coordinates: ({left}, {top}, {right}, {bottom})")
                
                # Apply crop and update crop state
                cropped_img = img.crop((left, top, right, bottom))
                
                # Update crop state with coordinates
                if image_path in self.image_crop_states:
                    crop_state = self.image_crop_states[image_path]
                    crop_state.set_crop(left, top, right, bottom, "intelligent")
                
                debug_print(f"Claude cropped image size: {cropped_img.size}")
                return cropped_img
            
        except Exception as e:
            debug_print(f"ERROR: Claude intelligent crop failed: {e}")
        
        return None
    
    def _smart_crop_fallback(self, img: Any, image_path: str) -> Any:
        """Fallback cropping using computer vision techniques."""
        debug_print(f"Smart crop fallback for {Path(image_path).name}")
        
        try:
            cv2 = self._lazy_import_cv2()
            np = self._lazy_import_numpy()
            
            if cv2 is None or np is None:
                debug_print("OpenCV/NumPy not available, using center crop")
                return self._center_crop(img)
            
            # Convert PIL to numpy array for OpenCV
            img_array = np.array(img)
            if len(img_array.shape) == 3:
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            else:
                gray = img_array
            
            # Apply Gaussian blur
            blurred = cv2.GaussianBlur(gray, (5, 5), 0)
            
            # Apply multiple edge detection techniques
            edges1 = cv2.Canny(blurred, 50, 150)
            edges2 = cv2.Canny(blurred, 100, 200)
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
            
            # Update crop state
            if image_path in self.image_crop_states:
                crop_state = self.image_crop_states[image_path]
                crop_state.set_crop(*crop_box, "smart_fallback")
            
            return img.crop(crop_box)
            
        except Exception as e:
            debug_print(f"ERROR: Smart crop fallback failed: {e}")
            return self._center_crop(img)
    
    def _center_crop(self, img: Any, crop_factor: float = 0.8) -> Any:
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
    
    def resize_for_display(self, img: Any, max_height: int = None) -> Any:
        """Resize image for display purposes."""
        try:
            max_height = max_height or self.max_display_height
            if img.height <= max_height:
                return img
            
            scaling_factor = max_height / img.height
            new_width = int(img.width * scaling_factor)
            
            Image, _, _, _, _, _ = self._lazy_import_pil()
            if Image:
                resized_img = img.resize((new_width, max_height), Image.Resampling.LANCZOS)
                debug_print(f"Resized image to ({new_width}, {max_height}) for display")
                return resized_img
            else:
                return img
        except Exception as e:
            debug_print(f"ERROR: Failed to resize image for display: {e}")
            return img
    
    # ===================== IMAGE CACHE MANAGEMENT =====================
    
    def _clear_image_cache(self, image_path: str):
        """Clear cache entries for a specific image."""
        keys_to_remove = [key for key in self.processed_image_cache.keys() if key.startswith(image_path)]
        for key in keys_to_remove:
            del self.processed_image_cache[key]
            debug_print(f"Removed cache entry: {key}")
    
    def clear_all_cache(self):
        """Clear all processed image cache."""
        cache_count = len(self.processed_image_cache)
        self.processed_image_cache.clear()
        debug_print(f"Cleared all image cache ({cache_count} entries)")
    
    def cleanup_expired_cache(self):
        """Remove expired cache entries."""
        current_time = datetime.now()
        expired_keys = []
        
        for cache_key, cache_entry in self.processed_image_cache.items():
            age_hours = (current_time - cache_entry.created_at).total_seconds() / 3600
            if age_hours > self.cache_duration_hours:
                expired_keys.append(cache_key)
        
        for key in expired_keys:
            del self.processed_image_cache[key]
        
        if expired_keys:
            debug_print(f"Cleaned up {len(expired_keys)} expired cache entries")
    
    def limit_cache_size(self):
        """Limit cache size by removing least recently used entries."""
        if len(self.processed_image_cache) <= self.cache_size_limit:
            return
        
        # Sort by last accessed time
        cache_items = [(key, entry) for key, entry in self.processed_image_cache.items()]
        cache_items.sort(key=lambda x: x[1].last_accessed or datetime.min)
        
        # Remove oldest entries
        entries_to_remove = len(cache_items) - self.cache_size_limit
        for i in range(entries_to_remove):
            key = cache_items[i][0]
            del self.processed_image_cache[key]
        
        debug_print(f"Limited cache size by removing {entries_to_remove} least recently used entries")
    
    # ===================== UI SUPPORT METHODS =====================
    
    def add_image_files(self, new_files: List[str], current_sheet: str = None) -> List[str]:
        """Add new image files to the current list."""
        added_files = []
        
        for file_path in new_files:
            if file_path not in self.image_files:
                # Validate file format
                file_ext = Path(file_path).suffix.lower()
                if file_ext in self.supported_formats:
                    self.image_files.append(file_path)
                    added_files.append(file_path)
                    
                    # Set initial crop state
                    if file_path not in self.image_crop_states:
                        self.image_crop_states[file_path] = CropState(crop_enabled=self.global_crop_enabled)
                    
                    # Create metadata if sheet provided
                    if current_sheet:
                        self.add_image_metadata(file_path, current_sheet, "standard")
                else:
                    debug_print(f"Skipping unsupported file format: {file_ext}")
        
        debug_print(f"Added {len(added_files)} new image files")
        return added_files
    
    def remove_image_file(self, image_path: str) -> bool:
        """Remove an image file from the current list."""
        if image_path in self.image_files:
            self.image_files.remove(image_path)
            self._cleanup_image_data(image_path)
            debug_print(f"Removed image file: {Path(image_path).name}")
            return True
        return False
    
    def _cleanup_image_data(self, image_path: str):
        """Clean up all data associated with an image."""
        # Remove from crop states
        if image_path in self.image_crop_states:
            del self.image_crop_states[image_path]
        
        # Remove from processing status
        if image_path in self.processing_status:
            del self.processing_status[image_path]
        
        # Remove from all_images registry
        if image_path in self.all_images:
            del self.all_images[image_path]
        
        # Clear cache entries
        self._clear_image_cache(image_path)
        
        debug_print(f"Cleaned up all data for image: {Path(image_path).name}")
    
    def get_image_files(self) -> List[str]:
        """Get current list of image files."""
        return self.image_files.copy()
    
    def clear_image_files(self):
        """Clear all image files and associated data."""
        file_count = len(self.image_files)
        
        # Clean up data for all files
        for image_path in self.image_files:
            self._cleanup_image_data(image_path)
        
        # Clear lists
        self.image_files.clear()
        self.image_widgets.clear()
        self.close_buttons.clear()
        self.image_references.clear()
        
        debug_print(f"Cleared {file_count} image files and associated data")
    
    # ===================== UTILITY METHODS =====================
    
    def get_model_stats(self) -> Dict[str, Any]:
        """Get comprehensive statistics about the image model."""
        return {
            'total_images': len(self.all_images),
            'sheet_images_count': sum(len(sheets) for sheets in self.sheet_images.values()),
            'sample_images_count': len(self.sample_images),
            'current_image_files': len(self.image_files),
            'cached_images': len(self.processed_image_cache),
            'crop_states': len(self.image_crop_states),
            'processing_status': len(self.processing_status),
            'global_crop_enabled': self.global_crop_enabled,
            'claude_api_available': self.claude_client is not None,
            'pil_loaded': self._pil_loaded,
            'cv2_loaded': self._cv2_loaded,
            'numpy_loaded': self._numpy_loaded,
            'supported_formats': self.supported_formats,
            'cache_settings': {
                'max_display_height': self.max_display_height,
                'cache_size_limit': self.cache_size_limit,
                'cache_duration_hours': self.cache_duration_hours
            }
        }
    
    def validate_image_path(self, image_path: str) -> bool:
        """Validate that an image path exists and has supported format."""
        if not os.path.exists(image_path):
            return False
        
        file_ext = Path(image_path).suffix.lower()
        return file_ext in self.supported_formats
    
    def get_image_info(self, image_path: str) -> Dict[str, Any]:
        """Get comprehensive information about a specific image."""
        info = {
            'path': image_path,
            'exists': os.path.exists(image_path),
            'metadata': None,
            'crop_state': None,
            'processing_status': self.processing_status.get(image_path, 'none'),
            'cached': False,
            'cache_entries': []
        }
        
        # Add metadata if available
        if image_path in self.all_images:
            metadata = self.all_images[image_path]
            info['metadata'] = {
                'display_name': metadata.display_name,
                'sheet_name': metadata.sheet_name,
                'image_type': metadata.image_type,
                'sample_id': metadata.sample_id,
                'test_name': metadata.test_name,
                'file_size': metadata.file_size,
                'dimensions': metadata.dimensions,
                'created_at': metadata.created_at,
                'processing_status': metadata.processing_status
            }
        
        # Add crop state if available
        if image_path in self.image_crop_states:
            crop_state = self.image_crop_states[image_path]
            info['crop_state'] = {
                'crop_enabled': crop_state.crop_enabled,
                'is_cropped': crop_state.is_cropped,
                'crop_method': crop_state.crop_method,
                'crop_coordinates': crop_state.get_crop_box(),
                'crop_timestamp': crop_state.crop_timestamp
            }
        
        # Check cache entries
        cache_keys = [key for key in self.processed_image_cache.keys() if key.startswith(image_path)]
        info['cached'] = len(cache_keys) > 0
        info['cache_entries'] = cache_keys
        
        return info
    
    def export_image_metadata(self, file_name: str = None) -> Dict[str, Any]:
        """Export image metadata for saving to database or file."""
        export_data = {
            'sheet_images': {},
            'sample_images': {},
            'image_crop_states': {},
            'sample_image_metadata': self.sample_image_metadata.copy(),
            'export_timestamp': datetime.now().isoformat()
        }
        
        # Export sheet images
        if file_name and file_name in self.sheet_images:
            for sheet_name, images in self.sheet_images[file_name].items():
                export_data['sheet_images'][sheet_name] = [img.image_path for img in images]
        else:
            # Export all sheet images
            for file_key, sheets in self.sheet_images.items():
                export_data['sheet_images'][file_key] = {}
                for sheet_name, images in sheets.items():
                    export_data['sheet_images'][file_key][sheet_name] = [img.image_path for img in images]
        
        # Export sample images
        for sample_id, images in self.sample_images.items():
            export_data['sample_images'][sample_id] = [img.image_path for img in images]
        
        # Export crop states
        for image_path, crop_state in self.image_crop_states.items():
            export_data['image_crop_states'][image_path] = {
                'crop_enabled': crop_state.crop_enabled,
                'is_cropped': crop_state.is_cropped,
                'crop_method': crop_state.crop_method,
                'crop_coordinates': crop_state.get_crop_box(),
                'crop_timestamp': crop_state.crop_timestamp.isoformat() if crop_state.crop_timestamp else None
            }
        
        debug_print(f"Exported image metadata: {len(export_data['sheet_images'])} sheet groups, {len(export_data['sample_images'])} sample groups")
        return export_data
    
    def import_image_metadata(self, import_data: Dict[str, Any]):
        """Import image metadata from database or file."""
        try:
            # Import sheet images
            if 'sheet_images' in import_data:
                for file_or_sheet, data in import_data['sheet_images'].items():
                    if isinstance(data, dict):  # New format: file -> sheet -> images
                        for sheet_name, image_paths in data.items():
                            for img_path in image_paths:
                                if self.validate_image_path(img_path):
                                    self.add_sheet_image(file_or_sheet, sheet_name, img_path)
                    else:  # Old format: sheet -> images
                        for img_path in data:
                            if self.validate_image_path(img_path):
                                self.add_sheet_image("imported", file_or_sheet, img_path)
            
            # Import sample images
            if 'sample_images' in import_data:
                for sample_id, image_paths in import_data['sample_images'].items():
                    for img_path in image_paths:
                        if self.validate_image_path(img_path):
                            self.add_sample_image(sample_id, img_path)
            
            # Import crop states
            if 'image_crop_states' in import_data:
                for image_path, crop_data in import_data['image_crop_states'].items():
                    crop_state = CropState()
                    crop_state.crop_enabled = crop_data.get('crop_enabled', False)
                    crop_state.is_cropped = crop_data.get('is_cropped', False)
                    crop_state.crop_method = crop_data.get('crop_method', 'none')
                    
                    coords = crop_data.get('crop_coordinates')
                    if coords and len(coords) == 4:
                        crop_state.set_crop(*coords, crop_state.crop_method)
                    
                    self.set_crop_state(image_path, crop_state)
            
            # Import sample image metadata
            if 'sample_image_metadata' in import_data:
                self.sample_image_metadata.update(import_data['sample_image_metadata'])
            
            debug_print("Successfully imported image metadata")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to import image metadata: {e}")


# Export the main classes
__all__ = ['ImageModel', 'ImageMetadata', 'CropState', 'ImageCache']

# Debug output for model initialization
debug_print("ImageModel module loaded successfully")
debug_print("Available classes: ImageModel, ImageMetadata, CropState, ImageCache")
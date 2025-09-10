# services/image_service.py
"""
services/image_service.py
Image processing service.
This will contain the image handling logic from image_loader.py.
"""

from typing import Optional, Dict, Any, List, Tuple
from PIL import Image, ImageTk
import tkinter as tk
from pathlib import Path
import os


class ImageService:
    """Service for image processing and management."""
    
    def __init__(self):
        """Initialize the image service."""
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']
        self.thumbnail_size = (150, 150)
        self.image_cache: Dict[str, Any] = {}
        
        print("DEBUG: ImageService initialized")
        print(f"DEBUG: Supported formats: {', '.join(self.supported_formats)}")
    
    def load_image(self, image_path: str) -> Tuple[bool, Optional[Image.Image], str]:
        """Load an image from file path."""
        print(f"DEBUG: ImageService loading image: {image_path}")
        
        try:
            if not Path(image_path).exists():
                return False, None, f"Image file not found: {image_path}"
            
            file_ext = Path(image_path).suffix.lower()
            if file_ext not in self.supported_formats:
                return False, None, f"Unsupported image format: {file_ext}"
            
            # Check cache first
            if image_path in self.image_cache:
                print(f"DEBUG: ImageService loaded from cache: {image_path}")
                return True, self.image_cache[image_path], "Loaded from cache"
            
            # Load image
            image = Image.open(image_path)
            
            # Cache the image
            self.image_cache[image_path] = image
            
            print(f"DEBUG: ImageService loaded image: {image.size[0]}x{image.size[1]}")
            return True, image, "Success"
            
        except Exception as e:
            error_msg = f"Failed to load image: {e}"
            print(f"ERROR: ImageService - {error_msg}")
            return False, None, error_msg
    
    def create_thumbnail(self, image: Image.Image, size: Tuple[int, int] = None) -> Tuple[bool, Optional[Image.Image], str]:
        """Create a thumbnail from an image."""
        print(f"DEBUG: ImageService creating thumbnail")
        
        try:
            if not image:
                return False, None, "No image provided"
            
            thumbnail_size = size or self.thumbnail_size
            
            # Create thumbnail
            thumbnail = image.copy()
            thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)
            
            print(f"DEBUG: ImageService created thumbnail: {thumbnail.size[0]}x{thumbnail.size[1]}")
            return True, thumbnail, "Success"
            
        except Exception as e:
            error_msg = f"Failed to create thumbnail: {e}"
            print(f"ERROR: ImageService - {error_msg}")
            return False, None, error_msg
    
    def crop_image(self, image: Image.Image, x1: int, y1: int, x2: int, y2: int) -> Tuple[bool, Optional[Image.Image], str]:
        """Crop an image to specified coordinates."""
        print(f"DEBUG: ImageService cropping image: ({x1},{y1}) to ({x2},{y2})")
        
        try:
            if not image:
                return False, None, "No image provided"
            
            # Validate crop coordinates
            width, height = image.size
            x1 = max(0, min(x1, width))
            y1 = max(0, min(y1, height))
            x2 = max(x1, min(x2, width))
            y2 = max(y1, min(y2, height))
            
            if x1 >= x2 or y1 >= y2:
                return False, None, "Invalid crop coordinates"
            
            # Crop image
            cropped = image.crop((x1, y1, x2, y2))
            
            print(f"DEBUG: ImageService cropped to: {cropped.size[0]}x{cropped.size[1]}")
            return True, cropped, "Success"
            
        except Exception as e:
            error_msg = f"Failed to crop image: {e}"
            print(f"ERROR: ImageService - {error_msg}")
            return False, None, error_msg
    
    def save_image(self, image: Image.Image, output_path: str, format: str = "PNG") -> Tuple[bool, str]:
        """Save an image to file."""
        print(f"DEBUG: ImageService saving image to: {output_path}")
        
        try:
            if not image:
                return False, "No image provided"
            
            # Create output directory
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            
            # Save image
            image.save(output_path, format=format)
            
            print(f"DEBUG: ImageService saved image: {output_path}")
            return True, "Success"
            
        except Exception as e:
            error_msg = f"Failed to save image: {e}"
            print(f"ERROR: ImageService - {error_msg}")
            return False, error_msg
    
    def convert_for_tkinter(self, image: Image.Image) -> Tuple[bool, Optional[ImageTk.PhotoImage], str]:
        """Convert PIL image to Tkinter-compatible format."""
        print("DEBUG: ImageService converting image for Tkinter")
        
        try:
            if not image:
                return False, None, "No image provided"
            
            # Convert to PhotoImage
            photo_image = ImageTk.PhotoImage(image)
            
            print("DEBUG: ImageService converted image for Tkinter")
            return True, photo_image, "Success"
            
        except Exception as e:
            error_msg = f"Failed to convert image for Tkinter: {e}"
            print(f"ERROR: ImageService - {error_msg}")
            return False, None, error_msg
    
    def find_images_in_directory(self, directory: str, pattern: str = "*") -> List[str]:
        """Find images in a directory matching a pattern."""
        print(f"DEBUG: ImageService searching for images in: {directory}")
        
        try:
            dir_path = Path(directory)
            if not dir_path.exists():
                print(f"WARNING: Directory not found: {directory}")
                return []
            
            image_files = []
            for ext in self.supported_formats:
                image_files.extend(dir_path.glob(f"{pattern}{ext}"))
                image_files.extend(dir_path.glob(f"{pattern}{ext.upper()}"))
            
            image_paths = [str(f) for f in image_files]
            
            print(f"DEBUG: ImageService found {len(image_paths)} images")
            return image_paths
            
        except Exception as e:
            print(f"ERROR: ImageService failed to find images: {e}")
            return []
    
    def clear_cache(self):
        """Clear the image cache."""
        cache_size = len(self.image_cache)
        self.image_cache.clear()
        print(f"DEBUG: ImageService cleared cache ({cache_size} images)")
    
    def get_image_info(self, image_path: str) -> Dict[str, Any]:
        """Get information about an image file."""
        try:
            if not Path(image_path).exists():
                return {}
            
            with Image.open(image_path) as img:
                info = {
                    'filename': Path(image_path).name,
                    'format': img.format,
                    'mode': img.mode,
                    'size': img.size,
                    'width': img.size[0],
                    'height': img.size[1],
                    'file_size': Path(image_path).stat().st_size
                }
            
            return info
            
        except Exception as e:
            print(f"ERROR: ImageService failed to get image info: {e}")
            return {}
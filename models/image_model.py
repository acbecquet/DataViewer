# models/image_model.py
"""
models/image_model.py
Image handling models and metadata structures.
These models will replace image-related structures from image_loader.py.
"""

from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


@dataclass
class CropState:
    """Model for image crop state and coordinates."""
    x1: int = 0
    y1: int = 0
    x2: int = 0
    y2: int = 0
    is_cropped: bool = False
    crop_enabled: bool = False
    
    def __post_init__(self):
        """Post-initialization processing."""
        if self.is_cropped:
            print(f"DEBUG: Created CropState with crop ({self.x1},{self.y1}) to ({self.x2},{self.y2})")
        else:
            print("DEBUG: Created CropState with no crop")
    
    def set_crop(self, x1: int, y1: int, x2: int, y2: int):
        """Set crop coordinates."""
        self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2
        self.is_cropped = True
        print(f"DEBUG: Set crop coordinates ({x1},{y1}) to ({x2},{y2})")
    
    def clear_crop(self):
        """Clear crop coordinates."""
        self.x1 = self.y1 = self.x2 = self.y2 = 0
        self.is_cropped = False
        print("DEBUG: Cleared crop coordinates")


@dataclass
class ImageMetadata:
    """Metadata for images associated with sheets and samples."""
    image_path: str
    sheet_name: str
    image_type: str = "sample"  # sample, plot, reference
    display_name: Optional[str] = None
    crop_state: CropState = field(default_factory=CropState)
    thumbnail_path: Optional[str] = None
    file_size: int = 0
    created_at: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Post-initialization processing."""
        if Path(self.image_path).exists():
            self.file_size = Path(self.image_path).stat().st_size
        print(f"DEBUG: Created ImageMetadata for '{self.display_name}' in sheet '{self.sheet_name}'")
        print(f"DEBUG: Image path: {self.image_path}, Size: {self.file_size} bytes")


class ImageModel:
    """Main image model for managing images and their states."""
    
    def __init__(self):
        """Initialize the image model."""
        self.sheet_images: Dict[str, List[ImageMetadata]] = {}  # sheet_name -> images
        self.image_crop_states: Dict[str, CropState] = {}  # image_path -> crop_state
        self.sample_images: Dict[str, ImageMetadata] = {}  # sample_id -> image
        self.pending_images: Dict[str, Any] = {}  # For batch operations
        self.crop_enabled = False
        
        print("DEBUG: ImageModel initialized")
        print("DEBUG: Ready to manage sheet images, crop states, and sample images")
    
    def add_sheet_image(self, sheet_name: str, image_metadata: ImageMetadata):
        """Add an image to a specific sheet."""
        if sheet_name not in self.sheet_images:
            self.sheet_images[sheet_name] = []
        
        self.sheet_images[sheet_name].append(image_metadata)
        self.image_crop_states[image_metadata.image_path] = image_metadata.crop_state
        
        print(f"DEBUG: Added image to sheet '{sheet_name}' - total: {len(self.sheet_images[sheet_name])}")
    
    def get_sheet_images(self, sheet_name: str) -> List[ImageMetadata]:
        """Get all images for a specific sheet."""
        return self.sheet_images.get(sheet_name, [])
    
    def set_crop_state(self, image_path: str, crop_state: CropState):
        """Set crop state for an image."""
        self.image_crop_states[image_path] = crop_state
        print(f"DEBUG: Set crop state for image: {image_path}")
    
    def get_crop_state(self, image_path: str) -> Optional[CropState]:
        """Get crop state for an image."""
        return self.image_crop_states.get(image_path)
    
    def add_sample_image(self, sample_id: str, image_metadata: ImageMetadata):
        """Add a sample image."""
        self.sample_images[sample_id] = image_metadata
        print(f"DEBUG: Added sample image for sample '{sample_id}'")
    
    def clear_sheet_images(self, sheet_name: str):
        """Clear all images for a specific sheet."""
        if sheet_name in self.sheet_images:
            image_count = len(self.sheet_images[sheet_name])
            del self.sheet_images[sheet_name]
            print(f"DEBUG: Cleared {image_count} images from sheet '{sheet_name}'")
    
    def toggle_crop_mode(self):
        """Toggle global crop mode."""
        self.crop_enabled = not self.crop_enabled
        print(f"DEBUG: Toggled crop mode to {self.crop_enabled}")
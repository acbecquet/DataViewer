# controllers/image_controller.py
"""
controllers/image_controller.py
Image handling controller that coordinates image operations.
This replaces the coordination logic currently in image_loader.py.
"""

from typing import Optional, Dict, Any, List
from models.data_model import DataModel
from models.image_model import ImageModel, ImageMetadata, CropState


class ImageController:
    """Controller for image handling and display."""
    
    def __init__(self, data_model: DataModel, image_service: Any):
        """Initialize the image controller."""
        self.data_model = data_model
        self.image_service = image_service
        self.image_model = ImageModel()
        
        print("DEBUG: ImageController initialized")
        print(f"DEBUG: Connected to DataModel and ImageService")
    
    def load_sheet_images(self, sheet_name: str) -> bool:
        """Load images for a specific sheet."""
        print(f"DEBUG: ImageController loading images for sheet: {sheet_name}")
        
        try:
            # Load images through service (placeholder)
            # image_paths = self.image_service.find_sheet_images(sheet_name)
            
            # For now, create placeholder image metadata
            # In real implementation, would iterate through found images
            print(f"DEBUG: ImageController loaded images for {sheet_name}")
            return True
            
        except Exception as e:
            print(f"ERROR: ImageController failed to load images for {sheet_name}: {e}")
            return False
    
    def add_sample_image(self, sample_id: str, image_path: str) -> bool:
        """Add a sample image."""
        print(f"DEBUG: ImageController adding sample image: {sample_id}")
        
        try:
            # Create image metadata
            metadata = ImageMetadata(
                image_path=image_path,
                sheet_name="sample",
                image_type="sample",
                display_name=sample_id
            )
            
            # Add to model
            self.image_model.add_sample_image(sample_id, metadata)
            
            print(f"DEBUG: ImageController added sample image for {sample_id}")
            return True
            
        except Exception as e:
            print(f"ERROR: ImageController failed to add sample image: {e}")
            return False
    
    def set_crop_state(self, image_path: str, x1: int, y1: int, x2: int, y2: int):
        """Set crop state for an image."""
        crop_state = CropState()
        crop_state.set_crop(x1, y1, x2, y2)
        self.image_model.set_crop_state(image_path, crop_state)
        print(f"DEBUG: ImageController set crop for {image_path}")
    
    def toggle_crop_mode(self):
        """Toggle global crop mode."""
        self.image_model.toggle_crop_mode()
        print(f"DEBUG: ImageController crop mode: {self.image_model.crop_enabled}")
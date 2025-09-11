# views/widgets/image_widget.py
"""
views/widgets/image_widget.py
Image widget for displaying and managing images.
This contains image display functionality from image_loader.py.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Frame, Label, Button, Canvas, Scrollbar
from typing import Optional, Dict, Any, List, Callable, Tuple
import os
import tempfile


class ImageWidget:
    """Widget for displaying and managing images."""
    
    def __init__(self, parent: tk.Widget, controller: Optional[Any] = None):
        """Initialize the image widget."""
        self.parent = parent
        self.controller = controller
        
        # UI Components
        self.frame: Optional[Frame] = None
        self.canvas: Optional[Canvas] = None
        self.scrollbar: Optional[Scrollbar] = None
        self.scrollable_frame: Optional[Frame] = None
        self.scrollable_frame_id: Optional[int] = None
        
        # Image management
        self.image_files: List[str] = []
        self.image_widgets: List[tk.Widget] = []
        self.close_buttons: List[tk.Widget] = []
        self.image_references: List[Any] = []  # Keep references to prevent garbage collection
        self.image_crop_states: Dict[str, bool] = {}
        
        # Processing and caching
        self.processed_image_cache: Dict[str, Any] = {}
        self.processing_status: Dict[str, str] = {}
        self.temp_files: List[str] = []  # Track temporary files for cleanup
        
        # Configuration
        self.max_image_height = 135
        self.is_plotting_sheet = False
        self.auto_crop_enabled = False
        
        # Callbacks
        self.on_images_selected: Optional[Callable] = None
        self.on_image_removed: Optional[Callable] = None
        self.on_image_clicked: Optional[Callable] = None
        
        print("DEBUG: ImageWidget initialized")
    
    def setup_ui(self):
        """Set up the UI elements."""
        try:
            if not self.parent.winfo_exists():
                print("ERROR: ImageWidget - parent destroyed before UI setup")
                return False
            
            # Clear existing widgets
            for widget in self.parent.winfo_children():
                widget.destroy()
            
            # Create main frame
            self.frame = Frame(self.parent)
            self.frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            # Create canvas and scrollbar
            self.canvas = Canvas(self.frame, height=140, bg="white")
            self.scrollbar = Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
            self.scrollable_frame = Frame(self.canvas, bg="white")
            
            # Create window in canvas
            self.scrollable_frame_id = self.canvas.create_window(
                (0, 0), window=self.scrollable_frame, anchor="nw"
            )
            
            # Configure scrolling
            def on_canvas_configure(event):
                self.canvas.itemconfig(self.scrollable_frame_id, width=event.width)
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            self.canvas.bind("<Configure>", on_canvas_configure)
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            
            # Pack components
            self.canvas.pack(side="left", fill="both", expand=True)
            self.scrollbar.pack(side="right", fill="y")
            
            print("DEBUG: ImageWidget - UI setup completed")
            return True
            
        except Exception as e:
            print(f"ERROR: ImageWidget - UI setup failed: {e}")
            return False
    
    def add_images(self):
        """Open file dialog to select and add images."""
        try:
            if not (self.parent and self.parent.winfo_exists()):
                print("ERROR: ImageWidget - parent frame no longer exists")
                return
            
            # Setup UI if not already done
            if not self.frame:
                if not self.setup_ui():
                    return
            
            # Open file dialog
            filetypes = [
                ("Image files", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.pdf"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("PNG files", "*.png"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
            
            file_paths = filedialog.askopenfilenames(
                title="Select Images",
                filetypes=filetypes
            )
            
            if file_paths:
                # Add new files to existing list
                new_files = [f for f in file_paths if f not in self.image_files]
                self.image_files.extend(new_files)
                
                # Initialize crop states for new files
                for file_path in new_files:
                    if file_path not in self.image_crop_states:
                        self.image_crop_states[file_path] = self.auto_crop_enabled
                
                # Display all images
                self.display_images()
                
                # Notify callback
                if self.on_images_selected:
                    self.on_images_selected(self.image_files.copy())
                
                print(f"DEBUG: ImageWidget - added {len(new_files)} new images, total: {len(self.image_files)}")
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error adding images: {e}")
    
    def load_images_from_list(self, image_paths: List[str]):
        """Load images from a provided list of paths."""
        try:
            if not image_paths:
                self.image_files = []
                self.clear_display()
                return
            
            # Setup UI if needed
            if not self.frame:
                if not self.setup_ui():
                    return
            
            # Filter valid files
            valid_paths = []
            for path in image_paths:
                if os.path.exists(path):
                    valid_paths.append(path)
                else:
                    print(f"DEBUG: ImageWidget - file not found: {path}")
            
            self.image_files = valid_paths
            
            # Initialize crop states
            for file_path in self.image_files:
                if file_path not in self.image_crop_states:
                    self.image_crop_states[file_path] = self.auto_crop_enabled
            
            # Display images
            self.display_images()
            
            print(f"DEBUG: ImageWidget - loaded {len(valid_paths)} images from list")
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error loading images from list: {e}")
    
    def display_images(self):
        """Display all images in the scrollable frame."""
        try:
            if not self.canvas or not self.scrollable_frame:
                print("DEBUG: ImageWidget - canvas or scrollable_frame not initialized")
                return
            
            # Save scroll position
            prev_pos = (0, 0)
            try:
                prev_pos = self.canvas.yview()
            except:
                pass
            
            # Clear previous widgets
            self.clear_display()
            
            if not self.image_files:
                print("DEBUG: ImageWidget - no image files to display")
                return
            
            # Get PIL components
            Image, ImageTk = self._get_pil_components()
            if not Image or not ImageTk:
                print("ERROR: ImageWidget - PIL not available")
                return
            
            # Update parent size
            self.parent.update_idletasks()
            
            # Calculate layout
            padding = 5
            current_x = padding
            current_y = padding
            max_row_height = 0
            
            # Get frame width
            try:
                frame_width = self.scrollable_frame.winfo_width()
                if frame_width <= 1:
                    frame_width = 800  # Default fallback
            except:
                frame_width = 800
            
            print(f"DEBUG: ImageWidget - displaying {len(self.image_files)} images")
            
            # Process each image
            for img_index, img_path in enumerate(self.image_files):
                try:
                    # Get processed image
                    pil_image = self._get_processed_image(img_path, Image)
                    if not pil_image:
                        continue
                    
                    # Calculate dimensions
                    original_width, original_height = pil_image.size
                    aspect_ratio = original_width / original_height
                    
                    if original_height > self.max_image_height:
                        new_height = self.max_image_height
                        new_width = int(new_height * aspect_ratio)
                    else:
                        new_width = original_width
                        new_height = original_height
                    
                    # Resize image
                    display_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(display_image)
                    
                    # Calculate item dimensions
                    item_width = new_width + 10
                    item_height = new_height + 30
                    
                    # Check for row wrap
                    if current_x + item_width > frame_width - padding and current_x > padding:
                        current_x = padding
                        current_y += max_row_height + padding
                        max_row_height = 0
                    
                    # Create container
                    container = self._create_image_container(img_path, photo, item_width, item_height,
                                                           current_x, current_y)
                    
                    # Update positions
                    current_x += item_width + padding
                    max_row_height = max(max_row_height, item_height)
                    
                    # Keep reference to prevent garbage collection
                    self.image_references.append(photo)
                    
                except Exception as e:
                    print(f"DEBUG: ImageWidget - error processing image {img_index}: {e}")
                    continue
            
            # Update scroll region
            total_height = current_y + max_row_height + padding
            self.scrollable_frame.configure(height=total_height)
            
            try:
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
                self.parent.update_idletasks()
                
                # Restore scroll position
                if prev_pos and len(prev_pos) >= 2:
                    self.canvas.yview_moveto(prev_pos[0])
            except Exception as e:
                print(f"DEBUG: ImageWidget - error updating scroll region: {e}")
            
            print(f"DEBUG: ImageWidget - display completed. Total height: {total_height}")
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error in display_images: {e}")
    
    def _get_pil_components(self):
        """Get PIL components with lazy loading."""
        try:
            from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
            return Image, ImageTk
        except ImportError as e:
            print(f"ERROR: ImageWidget - PIL import failed: {e}")
            return None, None
    
    def _get_processed_image(self, img_path: str, Image: Any) -> Optional[Any]:
        """Get processed image from cache or process it."""
        try:
            # Check cache first
            if img_path in self.processed_image_cache:
                return self.processed_image_cache[img_path]
            
            # Process new image
            if img_path.lower().endswith('.pdf'):
                pil_image = self._process_pdf_image(img_path, Image)
            else:
                pil_image = Image.open(img_path)
                
                # Apply auto-crop if enabled
                if self.image_crop_states.get(img_path, False):
                    pil_image = self._auto_crop_image(pil_image, Image)
            
            # Cache processed image
            if pil_image:
                self.processed_image_cache[img_path] = pil_image
                self.processing_status[img_path] = "processed"
            
            return pil_image
            
        except Exception as e:
            print(f"DEBUG: ImageWidget - error processing image {img_path}: {e}")
            self.processing_status[img_path] = f"error: {str(e)}"
            return None
    
    def _process_pdf_image(self, pdf_path: str, Image: Any) -> Optional[Any]:
        """Process PDF file to extract first page as image."""
        try:
            # Try using pdf2image
            try:
                from pdf2image import convert_from_path
                pages = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=150)
                if pages:
                    return pages[0]
            except ImportError:
                print("DEBUG: ImageWidget - pdf2image not available")
            
            # Fallback: try using PyMuPDF (fitz)
            try:
                import fitz
                pdf_doc = fitz.open(pdf_path)
                if len(pdf_doc) > 0:
                    page = pdf_doc[0]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x scaling
                    img_data = pix.tobytes("ppm")
                    
                    # Create temporary file
                    with tempfile.NamedTemporaryFile(suffix='.ppm', delete=False) as tmp:
                        tmp.write(img_data)
                        tmp_path = tmp.name
                        self.temp_files.append(tmp_path)
                    
                    return Image.open(tmp_path)
                pdf_doc.close()
            except ImportError:
                print("DEBUG: ImageWidget - PyMuPDF not available")
            
            return None
            
        except Exception as e:
            print(f"DEBUG: ImageWidget - error processing PDF {pdf_path}: {e}")
            return None
    
    def _auto_crop_image(self, image: Any, Image: Any) -> Any:
        """Apply automatic cropping to remove white/transparent borders."""
        try:
            # Simple auto-crop implementation
            # Convert to RGB if needed
            if image.mode in ('RGBA', 'LA'):
                background = Image.new('RGB', image.size, (255, 255, 255))
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                image = background
            elif image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Get bounding box of non-white content
            bbox = image.getbbox()
            if bbox:
                # Add small margin
                margin = 10
                left, top, right, bottom = bbox
                left = max(0, left - margin)
                top = max(0, top - margin)
                right = min(image.size[0], right + margin)
                bottom = min(image.size[1], bottom + margin)
                
                return image.crop((left, top, right, bottom))
            
            return image
            
        except Exception as e:
            print(f"DEBUG: ImageWidget - error in auto-crop: {e}")
            return image
    
    def _create_image_container(self, img_path: str, photo: Any, width: int, height: int,
                              x: int, y: int) -> tk.Widget:
        """Create container widget for image display."""
        try:
            # Create container frame
            container = Frame(self.scrollable_frame, bg="white", relief="solid", bd=1)
            container.place(x=x, y=y, width=width, height=height)
            
            # Create image label
            img_label = Label(container, image=photo, bg="white", cursor="hand2")
            img_label.pack(expand=True)
            
            # Bind click events
            img_label.bind("<Button-1>", lambda e: self._on_image_click(img_path))
            img_label.bind("<Double-Button-1>", lambda e: self._on_image_double_click(img_path))
            
            # Create close button
            close_btn = Button(container, text="×", font=("Arial", 10, "bold"),
                             fg="red", bg="white", cursor="hand2", relief="raised", bd=1)
            close_btn.place(relx=1.0, rely=0.0, anchor="ne")
            close_btn.bind("<Button-1>", lambda e: self.remove_image(img_path))
            
            # Store references
            self.image_widgets.append(container)
            self.close_buttons.append(close_btn)
            
            return container
            
        except Exception as e:
            print(f"DEBUG: ImageWidget - error creating image container: {e}")
            return None
    
    def remove_image(self, img_path: str):
        """Remove an image from the display."""
        try:
            if img_path in self.image_files:
                self.image_files.remove(img_path)
                
                # Clear from cache
                if img_path in self.processed_image_cache:
                    del self.processed_image_cache[img_path]
                if img_path in self.processing_status:
                    del self.processing_status[img_path]
                if img_path in self.image_crop_states:
                    del self.image_crop_states[img_path]
                
                # Refresh display
                self.display_images()
                
                # Notify callback
                if self.on_image_removed:
                    self.on_image_removed(img_path)
                if self.on_images_selected:
                    self.on_images_selected(self.image_files.copy())
                
                print(f"DEBUG: ImageWidget - removed image: {os.path.basename(img_path)}")
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error removing image: {e}")
    
    def clear_display(self):
        """Clear all image widgets from display."""
        try:
            # Destroy widgets
            for widget in self.image_widgets:
                try:
                    widget.destroy()
                except:
                    pass
            
            for button in self.close_buttons:
                try:
                    button.destroy()
                except:
                    pass
            
            # Clear references
            self.image_widgets.clear()
            self.close_buttons.clear()
            self.image_references.clear()
            
            print("DEBUG: ImageWidget - display cleared")
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error clearing display: {e}")
    
    def clear_all(self):
        """Clear all images and reset the widget."""
        try:
            self.clear_display()
            self.image_files.clear()
            self.processed_image_cache.clear()
            self.processing_status.clear()
            self.image_crop_states.clear()
            
            # Clean up temp files
            for temp_file in self.temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.unlink(temp_file)
                except:
                    pass
            self.temp_files.clear()
            
            print("DEBUG: ImageWidget - all images cleared")
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error clearing all: {e}")
    
    # Event handlers
    def _on_image_click(self, img_path: str):
        """Handle single click on image."""
        try:
            if self.on_image_clicked:
                self.on_image_clicked(img_path, "single")
            print(f"DEBUG: ImageWidget - image clicked: {os.path.basename(img_path)}")
        except Exception as e:
            print(f"ERROR: ImageWidget - error in image click: {e}")
    
    def _on_image_double_click(self, img_path: str):
        """Handle double click on image."""
        try:
            if self.on_image_clicked:
                self.on_image_clicked(img_path, "double")
            else:
                # Default action: open image in system viewer
                self._open_image_externally(img_path)
            print(f"DEBUG: ImageWidget - image double-clicked: {os.path.basename(img_path)}")
        except Exception as e:
            print(f"ERROR: ImageWidget - error in image double-click: {e}")
    
    def _open_image_externally(self, img_path: str):
        """Open image in external application."""
        try:
            import subprocess
            import platform
            
            if platform.system() == 'Windows':
                os.startfile(img_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', img_path])
            else:  # Linux
                subprocess.run(['xdg-open', img_path])
                
        except Exception as e:
            print(f"DEBUG: ImageWidget - error opening image externally: {e}")
    
    # Public interface methods
    def set_auto_crop_enabled(self, enabled: bool):
        """Set auto-crop enabled state."""
        self.auto_crop_enabled = enabled
        print(f"DEBUG: ImageWidget - auto-crop set to: {enabled}")
    
    def set_plotting_sheet_mode(self, is_plotting_sheet: bool):
        """Set plotting sheet mode."""
        self.is_plotting_sheet = is_plotting_sheet
        print(f"DEBUG: ImageWidget - plotting sheet mode: {is_plotting_sheet}")
    
    def get_image_files(self) -> List[str]:
        """Get list of current image files."""
        return self.image_files.copy()
    
    def get_image_count(self) -> int:
        """Get count of current images."""
        return len(self.image_files)
    
    def has_images(self) -> bool:
        """Check if widget has any images."""
        return len(self.image_files) > 0
    
    def get_image_for_report(self, img_path: str, target_width: int = None, 
                           target_height: int = None) -> Optional[Any]:
        """Get processed image suitable for report generation."""
        try:
            Image, _ = self._get_pil_components()
            if not Image:
                return None
            
            # Get processed image
            processed_img = self._get_processed_image(img_path, Image)
            if not processed_img:
                return None
            
            # Resize for report if dimensions specified
            if target_width or target_height:
                if target_width and target_height:
                    processed_img = processed_img.resize((target_width, target_height), Image.Resampling.LANCZOS)
                elif target_width:
                    ratio = target_width / processed_img.size[0]
                    new_height = int(processed_img.size[1] * ratio)
                    processed_img = processed_img.resize((target_width, new_height), Image.Resampling.LANCZOS)
                elif target_height:
                    ratio = target_height / processed_img.size[1]
                    new_width = int(processed_img.size[0] * ratio)
                    processed_img = processed_img.resize((new_width, target_height), Image.Resampling.LANCZOS)
            
            return processed_img
            
        except Exception as e:
            print(f"ERROR: ImageWidget - error getting image for report: {e}")
            return None
    
    def clear_cache(self):
        """Clear the image processing cache."""
        self.processed_image_cache.clear()
        self.processing_status.clear()
        print("DEBUG: ImageWidget - cache cleared")
    
    def get_cache_info(self) -> Dict[str, Any]:
        """Get cache information for debugging."""
        return {
            'cached_images': len(self.processed_image_cache),
            'cache_keys': list(self.processed_image_cache.keys()),
            'processing_status': self.processing_status.copy()
        }
    
    def refresh_display(self):
        """Refresh the image display."""
        try:
            self.display_images()
            print("DEBUG: ImageWidget - display refreshed")
        except Exception as e:
            print(f"ERROR: ImageWidget - error refreshing display: {e}")
    
    def __del__(self):
        """Cleanup when widget is destroyed."""
        try:
            # Clean up temporary files
            for temp_file in self.temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.unlink(temp_file)
                except:
                    pass
        except:
            pass
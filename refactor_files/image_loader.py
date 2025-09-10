import os
import base64
import io
from tkinter import filedialog, Canvas, Scrollbar, Frame, Label, PhotoImage, messagebox
from utils import debug_print

class ImageLoader:
    def __init__(self, parent, is_plotting_sheet, on_images_selected=None, main_gui=None):
        """
        Initializes the ImageLoader class.
        :param parent: The parent Tkinter frame where images will be displayed.
        :param is_plotting_sheet: Boolean indicating whether the current sheet is a plotting sheet.
        :param on_images_selected: Callback function when images are selected
        :param main_gui: Reference to main GUI for crop settings
        """
        debug_print(f"DEBUG: Initializing ImageLoader with parent={parent} (is_plotting_sheet={is_plotting_sheet})")
        self.on_images_selected = on_images_selected
        self.parent = parent
        self.is_plotting_sheet = is_plotting_sheet
        self.main_gui = main_gui
        self.image_files = []
        self.image_widgets = []
        self.close_buttons = []
        self.image_crop_states = {}
        self.image_references = []  # Keep references to prevent garbage collection
        self.scrollable_frame_id = None
        self.frame = None
        self.canvas = None
        self.scrollbar = None
        self.scrollable_frame = None
        
        self.processed_image_cache = {}
        self.processing_status = {}

        # Initialize Claude API client if available
        self.claude_client = None
        self._init_claude_api()
        
        if not self.parent or not self.parent.winfo_exists():
            print("ERROR: Parent frame does not exist! Skipping UI setup.")
            return

        debug_print("DEBUG: Calling setup_ui()...")
        self.setup_ui()

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

    def _lazy_import_pil():
        """Lazy import PIL modules."""
        try:
            from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
            print("TIMING: Lazy loaded PIL modules")
            return Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
        except ImportError as e:
            print(f"Error importing PIL: {e}")
            return None, None, None, None, None, None

    def _lazy_import_cv2():
        """Lazy import cv2."""
        try:
            import cv2
            print("TIMING: Lazy loaded cv2")
            return cv2
        except ImportError as e:
            print(f"Error importing cv2: {e}")
            return None

    def _lazy_import_numpy():
        """Lazy import numpy."""
        try:
            import numpy as np
            print("TIMING: Lazy loaded numpy for ImageLoader")
            return np
        except ImportError as e:
            print(f"Error importing numpy: {e}")
            return None

    def setup_ui(self):
        """Sets up the UI elements only if parent exists"""
        debug_print("DEBUG: Attempting UI setup")
        if not self.parent.winfo_exists():
            print("ERROR: Parent destroyed before UI setup")
            return

        try:
            # Clear any existing widgets first
            for widget in self.parent.winfo_children():
                widget.destroy()
                
            self.frame = Frame(self.parent)
            self.frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            self.canvas = Canvas(self.frame, height=140, bg="white")
            self.scrollbar = Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
            self.scrollable_frame = Frame(self.canvas, bg="white")

            # Create window and store its ID
            self.scrollable_frame_id = self.canvas.create_window(
                (0, 0), 
                window=self.scrollable_frame, 
                anchor="nw"
            )
            
            # Update canvas scroll region and frame width on resize
            def on_canvas_configure(event):
                self.canvas.itemconfig(self.scrollable_frame_id, width=event.width)
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            self.canvas.bind("<Configure>", on_canvas_configure)
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            
            self.canvas.pack(side="left", fill="both", expand=True)
            self.scrollbar.pack(side="right", fill="y")
            
            debug_print("DEBUG: UI setup completed")
        except Exception as e:
            print(f"CRITICAL: UI setup failed - {str(e)}")

    def add_images(self):
        """
        Opens a file dialog to select images and displays them in the scrollable frame.
        """
        if not (self.parent and self.parent.winfo_exists()):
            print("ERROR: Parent frame no longer exists. Cannot load images.")
            return
        
        if not self.frame:  # Reinitialize UI if frame is missing
            print("WARNING: Frame missing, reinitializing UI")
            self.setup_ui()
            if not self.frame:  # Still missing after setup
                print("CRITICAL: Failed to create frame")
                return

        debug_print("DEBUG: Opening file dialog for image selection...")
        file_types = [
            ("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
            ("PDF Files", "*.pdf"),
            ("All Files", "*.*")
        ]
        
        new_files = filedialog.askopenfilenames(title="Select Images", filetypes=file_types)
        if not new_files:
            debug_print("DEBUG: No new images selected")
            return

        # Capture crop state at time of loading
        crop_state = self.main_gui.crop_enabled.get() if self.main_gui else False
        
        # Add new files and set their crop states
        for file in new_files:
            if file not in self.image_files:
                self.image_files.append(file)
                self.image_crop_states[file] = crop_state

        # Store images in main GUI
        if self.main_gui:
            self.main_gui.store_images(self.main_gui.selected_sheet.get(), self.image_files)
        
        debug_print(f"DEBUG: Selected image files: {self.image_files}")
        debug_print(f"DEBUG: Crop states: {self.image_crop_states}")

        if self.on_images_selected:
            self.on_images_selected(self.image_files)

        debug_print("DEBUG: Calling display_images()...")
        self.display_images()

    def load_images_from_list(self, image_paths):
        """Load images from a provided list"""
        self.image_files = image_paths
        self.display_images()

    def claude_intelligent_crop(self, img_path):
        """Use Claude API to analyze image and determine optimal crop boundaries."""
        if not self.claude_client:
            debug_print("DEBUG: Claude API not available, falling back to basic crop")
            return self.smart_crop_fallback(img_path)
    
        try:
            debug_print(f"DEBUG: Using Claude API for intelligent cropping: {img_path}")
        
            # Prepare image for Claude API
            Image, _, _, _, _, _ = _lazy_import_pil()
            if Image is None:
                return self.smart_crop_fallback(img_path)
        
            with Image.open(img_path) as img:
                # Convert to RGB if necessary
                if img.mode != 'RGB':
                    img = img.convert('RGB')
            
                # Resize if too large for API
                max_size = 1568
                original_size = img.size
                scale_factor = 1.0
                if max(img.width, img.height) > max_size:
                    scale_factor = max_size / max(img.width, img.height)
                    new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
                    img = img.resize(new_size, Image.Resampling.LANCZOS)
                    debug_print(f"DEBUG: Resized image to {new_size} for API processing (scale: {scale_factor})")
            
                # Convert to base64
                buffer = io.BytesIO()
                img.save(buffer, format='JPEG', quality=90)
                image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
                # Create prompt for Claude - VERY specific about bottoms
                prompt = f"""
                Analyze this image and find the COMPLETE boundaries of the entire device/object for tight cropping.

                CRITICAL - PAY SPECIAL ATTENTION TO THE BOTTOM:
                1. Electronic devices often have CURVED, ROUNDED, or EXTENDED bottoms below the main rectangular body
                2. The bottom boundary is NOT where the label ends - it's where the physical device ends
                3. Look for curved bases, rounded bottoms, protruding connectors, or any physical extensions
                4. Many vape devices, cartridges, and electronic components have rounded bottoms that extend well below the label area
                5. Follow the device contours all the way to the absolute bottom edge

                COMPLETE MEASUREMENT REQUIREMENTS:
                - TOP: Include mouthpieces, tips, caps, anything above the main body
                - BOTTOM: Include curved bases, rounded bottoms, connectors - follow the device shape to its absolute lowest point
                - SIDES: Include any side extensions, buttons, connectors
                - The device likely has a distinctive shape - capture ALL of it

                IMAGE DETAILS:
                - Image size: {img.width} x {img.height} pixels
                - Look beyond rectangular boundaries - follow the actual device shape
                - Electronic devices typically have rounded/curved bottoms

                MEASUREMENT VALIDATION:
                - If the detected height seems too short compared to width, extend the bottom boundary
                - Electronic devices are typically taller than they are wide
                - Double-check that you've included the complete bottom curve/extension

                Return JSON with exact measurements:
                {{
                    "device_analysis": {{
                        "has_curved_bottom": <true/false>,
                        "has_extended_bottom": <true/false>,
                        "bottom_shape_description": "<describe the bottom shape>",
                        "estimated_device_type": "<what type of device this appears to be>"
                    }},
                    "complete_boundaries": {{
                        "top_pixel": <absolute topmost pixel including any extensions>,
                        "bottom_pixel": <absolute bottommost pixel following device contour>,
                        "left_pixel": <absolute leftmost pixel>,
                        "right_pixel": <absolute rightmost pixel>
                    }},
                    "crop_box": {{
                        "left": <left_pixel - 15>,
                        "top": <top_pixel - 15>,
                        "right": <right_pixel + 15>,
                        "bottom": <bottom_pixel + 15>
                    }},
                    "measurements": {{
                        "total_device_height": <bottom_pixel - top_pixel>,
                        "total_device_width": <right_pixel - left_pixel>,
                        "height_to_width_ratio": <height/width>
                    }},
                    "confidence": <0-100>,
                    "description": "<complete device with all parts>"
                }}

                CRITICAL: Make absolutely sure you've followed the device shape to its absolute bottom edge!
                """
            
                # Send request to Claude
                response = self.claude_client.messages.create(
                    model="claude-3-5-sonnet-20241022",
                    max_tokens=2000,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "text",
                                    "text": prompt
                                },
                                {
                                    "type": "image",
                                    "source": {
                                        "type": "base64",
                                        "media_type": "image/jpeg",
                                        "data": image_data
                                    }
                                }
                            ]
                        }
                    ]
                )
            
                # Parse response
                response_text = response.content[0].text
                debug_print(f"DEBUG: Claude response: {response_text}")
            
                # Extract JSON from response
                import json
                start_idx = response_text.find('{')
                end_idx = response_text.rfind('}') + 1
            
                if start_idx == -1 or end_idx == 0:
                    debug_print("DEBUG: No JSON found in Claude response, using fallback")
                    return self.smart_crop_fallback(img_path)
            
                crop_data = json.loads(response_text[start_idx:end_idx])
                confidence = crop_data.get('confidence', 0)
            
                if confidence < 60:  # Lowered threshold since we want to try harder
                    debug_print(f"DEBUG: Claude confidence too low ({confidence}), using fallback")
                    return self.smart_crop_fallback(img_path)
            
                # Extract detailed analysis
                device_analysis = crop_data.get('device_analysis', {})
                boundaries = crop_data.get('complete_boundaries', {})
                measurements = crop_data.get('measurements', {})
                crop_box = crop_data['crop_box']
            
                debug_print(f"DEBUG: Device analysis: {device_analysis}")
                debug_print(f"DEBUG: Boundaries detected: {boundaries}")
                debug_print(f"DEBUG: Measurements: {measurements}")
            
                # VALIDATION: Check if the height seems reasonable
                height_width_ratio = measurements.get('height_to_width_ratio', 0)
                device_height = measurements.get('total_device_height', 0)
                device_width = measurements.get('total_device_width', 0)
            
                debug_print(f"DEBUG: Height/Width ratio: {height_width_ratio}")
            
                # If the device seems too short (likely missing bottom), extend it
                if height_width_ratio < 1.2:  # Most electronic devices are taller than wide
                    debug_print("WARNING: Device appears too short - likely missing bottom portion")
                    debug_print("DEBUG: Attempting to extend bottom boundary")
                
                    # Extend bottom by 20% of current height or at least 50 pixels
                    extension = max(int(device_height * 0.2), 50)
                    original_bottom = boundaries.get('bottom_pixel', crop_box.get('bottom', 0) + 15)
                    extended_bottom = original_bottom + extension
                
                    # Update the crop box
                    crop_box['bottom'] = min(extended_bottom + 15, img.height - 1)
                    debug_print(f"DEBUG: Extended bottom from {original_bottom} to {extended_bottom}")
                    debug_print(f"DEBUG: New crop bottom: {crop_box['bottom']}")
            
                # Convert coordinates, accounting for any scaling done for API
                with Image.open(img_path) as original_img:
                    if scale_factor != 1.0:
                        # Scale coordinates back up to original image size
                        left = int(crop_box['left'] / scale_factor)
                        top = int(crop_box['top'] / scale_factor)
                        right = int(crop_box['right'] / scale_factor)
                        bottom = int(crop_box['bottom'] / scale_factor)
                        debug_print(f"DEBUG: Scaled coordinates back by factor {1/scale_factor}")
                    else:
                        left = crop_box['left']
                        top = crop_box['top']
                        right = crop_box['right']
                        bottom = crop_box['bottom']
                
                    # Ensure coordinates are valid
                    width, height = original_img.size
                    left = max(0, min(left, width - 1))
                    top = max(0, min(top, height - 1))
                    right = max(left + 1, min(right, width))
                    bottom = max(top + 1, min(bottom, height))
                
                    debug_print(f"DEBUG: Final validated crop box: ({left}, {top}, {right}, {bottom})")
                    debug_print(f"DEBUG: Final crop dimensions: {right - left} x {bottom - top} pixels")
                
                    cropped_img = original_img.crop((left, top, right, bottom))
                    debug_print(f"DEBUG: Cropped image created with size: {cropped_img.size}")
                    return cropped_img
            
        except Exception as e:
            print(f"ERROR: Claude intelligent crop failed: {e}")
            return self.smart_crop_fallback(img_path)

    def smart_crop_fallback(self, img_path):
        """Fallback cropping using improved computer vision techniques."""
        try:
            Image, _, _, _, _, _ = _lazy_import_pil()
            cv2 = _lazy_import_cv2()
            np = _lazy_import_numpy()
            
            if Image is None:
                debug_print("DEBUG: PIL not available for fallback crop")
                return Image.open(img_path)
                
            with Image.open(img_path) as img:
                debug_print(f"DEBUG: Using fallback crop for {img_path}")
                
                if cv2 is None or np is None:
                    debug_print("DEBUG: OpenCV/numpy not available, returning center crop")
                    # Simple center crop as last resort
                    width, height = img.size
                    crop_size = min(width, height) * 0.8
                    left = (width - crop_size) // 2
                    top = (height - crop_size) // 2
                    return img.crop((left, top, left + crop_size, top + crop_size))
                
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
                    debug_print("DEBUG: No contours found, returning center crop")
                    width, height = img.size
                    margin = min(width, height) * 0.1
                    return img.crop((margin, margin, width - margin, height - margin))
                
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
                
                debug_print(f"DEBUG: Fallback crop box: {crop_box}")
                return img.crop(crop_box)
                
        except Exception as e:
            print(f"ERROR: Fallback crop failed: {e}")
            try:
                return Image.open(img_path)
            except:
                return None

    def smart_crop(self, img, margin_percent=10):
        """Enhanced smart crop using Claude API when available."""
        # This method now serves as a wrapper for the new intelligent cropping
        # The actual cropping logic has been moved to claude_intelligent_crop and smart_crop_fallback
        debug_print("DEBUG: smart_crop called - this should now use claude_intelligent_crop")
        return img  # Return unchanged for now, as processing happens in process_image

    def process_image(self, img_path):
        """Load and process image with optional cropping based on toggle. Cache results to avoid re-processing."""
        try:
            Image, ImageTk, _, _, _, _ = _lazy_import_pil()
            if Image is None:
                print("ERROR: PIL not available")
                return None
                
            debug_print(f"DEBUG: Processing image: {img_path}")
            
            should_crop = self.image_crop_states.get(img_path, False)
            
            # Create a cache key based on image path and crop state
            cache_key = f"{img_path}_{should_crop}"
            
            # Check if we already have this processed image in cache
            if cache_key in self.processed_image_cache:
                debug_print(f"DEBUG: Using cached processed image for {os.path.basename(img_path)}")
                cached_img = self.processed_image_cache[cache_key]
                
                # Resize for display (this is cheap so we can do it every time)
                max_height = 135
                scaling_factor = max_height / cached_img.height
                new_width = int(cached_img.width * scaling_factor)
                
                debug_print(f"DEBUG: Resizing cached image to: ({new_width}, {max_height})")
                return cached_img.resize((new_width, max_height), Image.Resampling.LANCZOS)
            
            # Image not in cache - process it
            if should_crop:
                debug_print(f"DEBUG: Auto-Crop ENABLED - Processing with Claude API (FIRST TIME): {os.path.basename(img_path)}")
                processed_img = self.claude_intelligent_crop(img_path)
                if processed_img is None:
                    debug_print("DEBUG: Intelligent crop failed, loading original")
                    processed_img = Image.open(img_path)
            else:
                debug_print(f"DEBUG: Auto-Crop DISABLED - Loading original image: {os.path.basename(img_path)}")
                processed_img = Image.open(img_path)
            
            # Store the processed image in cache (at full resolution before display resizing)
            self.processed_image_cache[cache_key] = processed_img.copy()
            self.processing_status[img_path] = should_crop
            debug_print(f"DEBUG: Cached processed image for {os.path.basename(img_path)} (crop={should_crop})")
            
            debug_print(f"DEBUG: Final processed image size: {processed_img.size}")
            
            # Resize for display
            max_height = 135
            scaling_factor = max_height / processed_img.height
            new_width = int(processed_img.width * scaling_factor)
            
            debug_print(f"DEBUG: Resizing processed image to: ({new_width}, {max_height})")
            return processed_img.resize((new_width, max_height), Image.Resampling.LANCZOS)
            
        except Exception as e:
            print(f"ERROR: Failed to process image {img_path}: {e}")
            return None

    def remove_image(self, path_to_remove):
        """Remove an image from the list and refresh display"""
        debug_print(f"DEBUG: Removing image: {path_to_remove}")
        
        # Filter out the removed path
        self.image_files = [p for p in self.image_files if p != path_to_remove]
        
        # Remove from crop states
        if path_to_remove in self.image_crop_states:
            del self.image_crop_states[path_to_remove]
        
        # IMPORTANT: Clean up cache entries for this image
        cache_keys_to_remove = [key for key in self.processed_image_cache.keys() if key.startswith(path_to_remove)]
        for cache_key in cache_keys_to_remove:
            del self.processed_image_cache[cache_key]
            debug_print(f"DEBUG: Removed cache entry: {cache_key}")
        
        if path_to_remove in self.processing_status:
            del self.processing_status[path_to_remove]
    
        # Update parent storage if callback exists
        if self.on_images_selected:
            self.on_images_selected(self.image_files)
        
        # Redraw images with updated list (will use cached images for remaining ones)
        self.display_images()
        
        debug_print(f"DEBUG: Image removed. Remaining: {len(self.image_files)} images")

    def clear_cache(self):
        """Clear the processed image cache - useful for debugging or memory management"""
        self.processed_image_cache.clear()
        self.processing_status.clear()
        debug_print("DEBUG: Cleared processed image cache")

    def get_cache_info(self):
        """Get information about the current cache state - useful for debugging"""
        cache_info = {
            'cached_images': len(self.processed_image_cache),
            'cache_keys': list(self.processed_image_cache.keys()),
            'processing_status': self.processing_status
        }
        debug_print(f"DEBUG: Cache info: {cache_info}")
        return cache_info

    def display_images(self):
        """
        Loads and displays the selected images in the GUI with additional debugging.
        """
        if not self.canvas or not self.scrollable_frame:
            debug_print("DEBUG: Canvas or scrollable_frame not initialized")
            return
            
        prev_pos = (0, 0)
        try:
            prev_pos = self.canvas.yview()
        except:
            pass
            
        debug_print("DEBUG: Clearing previous image widgets...")
        for widget in self.scrollable_frame.winfo_children():
            debug_print(f"DEBUG: Destroying widget {widget}")
            widget.destroy()
        
        self.image_widgets.clear()
        self.close_buttons.clear()
        self.image_references = []

        if not self.image_files:
            debug_print("DEBUG: No image files to display")
            return

        Image, ImageTk, _,_,_,_ = _lazy_import_pil()
        if Image is None or ImageTk is None:
            print("ERROR: PIL modules not available for display")
            return


        # Ensure parent has updated size before calculating max width
        self.parent.update_idletasks()
        max_height = 135
        
        debug_print(f"DEBUG: Total images to process: {len(self.image_files)}")

        # Calculate positions
        padding = 5
        current_x = padding
        current_y = padding
        max_row_height = 0
        
        # Get frame width, with fallback
        try:
            frame_width = self.scrollable_frame.winfo_width()
            if frame_width <= 1:  # Frame not properly initialized
                frame_width = 800  # Default width
        except:
            frame_width = 800

        for img_index, img_path in enumerate(self.image_files):
            debug_print(f"DEBUG: Processing image {img_index + 1}/{len(self.image_files)}: {os.path.basename(img_path)}")
            
            # Create container frame for image + close button
            container = Frame(self.scrollable_frame, bg="white", relief="solid", bd=1)
            container.pack_propagate(False)

            if img_path.lower().endswith(".pdf"):
                label = Label(container, text=f"PDF: {os.path.basename(img_path)}", bg="white")
                label.pack(padx=5, pady=5)
                item_width = 200  # Fixed width for PDF labels
                item_height = 50
            else:
                try:
                    img = self.process_image(img_path)
                    if img is None:
                        debug_print(f"DEBUG: Skipping failed image: {img_path}")
                        continue

                    img_tk = ImageTk.PhotoImage(img)
                    label = Label(container, image=img_tk, bg="white")
                    label.pack(padx=2, pady=2)
                    
                    self.image_references.append(img_tk)
                    item_width = img.width + 4  # Add padding
                    item_height = img.height + 4
                    
                    debug_print(f"DEBUG: Image {img_index + 1} processed successfully: {item_width}x{item_height}")
                    
                except Exception as e:
                    print(f"ERROR: Failed to process image {img_path}: {str(e)}")
                    continue

            # Create close button
            close_btn = Label(container,
                text="✕",
                font=("Arial", 8, "bold"),
                fg="red",
                bg="white",
                cursor="hand2",
                relief="raised",
                bd=1
            )
            close_btn.place(relx=1.0, rely=0.0, anchor="ne")
            close_btn.bind("<Button-1>", lambda e, path=img_path: self.remove_image(path))

            self.image_widgets.append(container)
            self.close_buttons.append(close_btn)

            # Set container dimensions
            container.config(width=item_width, height=item_height)

            # Wrap to new row if needed
            if current_x + item_width > frame_width - padding:
                current_x = padding
                current_y += max_row_height + padding
                max_row_height = 0
                debug_print(f"DEBUG: Wrapping to new row at y={current_y}")

            # Position container
            container.place(x=current_x, y=current_y)
            debug_print(f"DEBUG: Placed container at ({current_x}, {current_y})")

            # Update tracking variables
            current_x += item_width + padding
            if item_height > max_row_height:
                max_row_height = item_height

        # Update scrolling area
        total_height = current_y + max_row_height + padding
        self.scrollable_frame.configure(height=total_height)
        
        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self.parent.update_idletasks()
            if prev_pos and len(prev_pos) >= 2:
                self.canvas.yview_moveto(prev_pos[0])
        except Exception as e:
            debug_print(f"DEBUG: Error updating scroll region: {e}")
        
        debug_print(f"DEBUG: Display completed. Total height: {total_height}")

    # Add this method to your ImageLoader class

    def get_image_for_report(self, img_path, target_width=None, target_height=None):
        """
        Get the best available version of an image for report generation.
        Returns the processed (cropped) version if available, otherwise the original.
    
        Args:
            img_path: Path to the original image
            target_width: Optional target width for the returned image
            target_height: Optional target height for the returned image
    
        Returns:
            PIL Image object (processed version if available, otherwise original)
        """
        try:
            Image, _, _, _, _, _ = _lazy_import_pil()
            if Image is None:
                return None
            
            debug_print(f"DEBUG: Getting image for report: {os.path.basename(img_path)}")
        
            # Check if we have a processed version in cache
            should_crop = self.image_crop_states.get(img_path, False)
            cache_key = f"{img_path}_{should_crop}"
        
            if cache_key in self.processed_image_cache:
                debug_print(f"DEBUG: Using processed version for report: {os.path.basename(img_path)} (cropped={should_crop})")
                img = self.processed_image_cache[cache_key].copy()
            else:
                debug_print(f"DEBUG: Using original image for report: {os.path.basename(img_path)}")
                img = Image.open(img_path)
        
            # Resize if target dimensions are specified
            if target_width or target_height:
                original_width, original_height = img.size
            
                if target_width and target_height:
                    # Both dimensions specified - use them directly
                    img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
                elif target_width:
                    # Only width specified - maintain aspect ratio
                    ratio = target_width / original_width
                    new_height = int(original_height * ratio)
                    img = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
                elif target_height:
                    # Only height specified - maintain aspect ratio
                    ratio = target_height / original_height
                    new_width = int(original_width * ratio)
                    img = img.resize((new_width, target_height), Image.Resampling.LANCZOS)
            
                debug_print(f"DEBUG: Resized image for report to: {img.size}")
        
            return img
        
        except Exception as e:
            print(f"ERROR: Failed to get image for report {img_path}: {e}")
            try:
                # Fallback to original image
                return Image.open(img_path)
            except:
                return None

# Update the ImageLoader methods with better error handling and debugging

    def save_image_for_report(self, img_path, output_path, target_width=None, target_height=None):
        """
        Save the best available version of an image to a file for report use.
        """
        try:
            debug_print(f"DEBUG: save_image_for_report called for {os.path.basename(img_path)}")
        
            img = self.get_image_for_report(img_path, target_width, target_height)
            if img is None:
                debug_print(f"DEBUG: Failed to get image for {img_path}")
                return False
            
            # Ensure output directory exists
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                debug_print(f"DEBUG: Created output directory: {output_dir}")
        
            # Save as JPEG for reports (good compression, universal support)
            if img.mode != 'RGB':
                img = img.convert('RGB')
        
            img.save(output_path, 'JPEG', quality=95)
            debug_print(f"DEBUG: Successfully saved report image to: {output_path}")
            debug_print(f"DEBUG: Saved image size: {img.size}")
        
            # Verify the file was created
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                debug_print(f"DEBUG: Verified saved file exists, size: {file_size} bytes")
                return True
            else:
                debug_print(f"ERROR: File was not created: {output_path}")
                return False
        
        except Exception as e:
            print(f"ERROR: Failed to save image for report {img_path}: {e}")
            import traceback
            traceback.print_exc()
            return False

    def get_image_for_report(self, img_path, target_width=None, target_height=None):
        """
        Get the best available version of an image for report generation.
        """
        try:
            Image, _, _, _, _, _ = _lazy_import_pil()
            if Image is None:
                debug_print("ERROR: PIL not available for get_image_for_report")
                return None
            
            debug_print(f"DEBUG: Getting image for report: {os.path.basename(img_path)}")
        
            # Check if we have a processed version in cache
            should_crop = self.image_crop_states.get(img_path, False)
            cache_key = f"{img_path}_{should_crop}"
        
            debug_print(f"DEBUG: Image crop state for {os.path.basename(img_path)}: {should_crop}")
            debug_print(f"DEBUG: Looking for cache key: {cache_key}")
            debug_print(f"DEBUG: Available cache keys: {list(self.processed_image_cache.keys())}")
        
            if cache_key in self.processed_image_cache:
                debug_print(f"DEBUG: FOUND processed version in cache for {os.path.basename(img_path)} (cropped={should_crop})")
                img = self.processed_image_cache[cache_key].copy()
                debug_print(f"DEBUG: Retrieved processed image size: {img.size}")
            else:
                debug_print(f"DEBUG: NO processed version found, using original for {os.path.basename(img_path)}")
                img = Image.open(img_path)
                debug_print(f"DEBUG: Original image size: {img.size}")
        
            # Resize if target dimensions are specified
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
            print(f"ERROR: Failed to get image for report {img_path}: {e}")
            import traceback
            traceback.print_exc()
            try:
                # Fallback to original image
                debug_print(f"DEBUG: Attempting fallback to original image")
                return Image.open(img_path)
            except:
                debug_print(f"ERROR: Fallback to original also failed")
                return None

# Add these functions at module level for lazy imports
def _lazy_import_pil():
    """Lazy import PIL modules."""
    try:
        from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
        print("TIMING: Lazy loaded PIL modules")
        return Image, ImageTk, ImageFilter, ImageOps, ImageChops, ImageEnhance
    except ImportError as e:
        print(f"Error importing PIL: {e}")
        return None, None, None, None, None, None

def _lazy_import_cv2():
    """Lazy import cv2."""
    try:
        import cv2
        print("TIMING: Lazy loaded cv2")
        return cv2
    except ImportError as e:
        print(f"Error importing cv2: {e}")
        return None

def _lazy_import_numpy():
    """Lazy import numpy."""
    try:
        import numpy as np
        print("TIMING: Lazy loaded numpy for ImageLoader")
        return np
    except ImportError as e:
        print(f"Error importing numpy: {e}")
        return None
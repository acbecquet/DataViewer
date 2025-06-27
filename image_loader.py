import os
from tkinter import filedialog, Canvas, Scrollbar, Frame, Label, PhotoImage
from PIL import Image, ImageTk
from PIL import ImageFilter, ImageOps, ImageChops, ImageEnhance
from utils import debug_print
import cv2
import numpy as np
class ImageLoader:
    def __init__(self, parent, is_plotting_sheet, on_images_selected = None, main_gui = None):
        """
        Initializes the ImageLoader class.
        :param parent: The parent Tkinter frame where images will be displayed.
        :param is_plotting_sheet: Boolean indicating whether the current sheet is a plotting sheet.
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
        self.scrollable_frame_id = None
        if not self.parent or not self.parent.winfo_exists():
            print("ERROR: Parent frame does not exist! Skipping UI setup.")
            return  # Prevents crashes if parent is invalid

        debug_print("DEBUG: Calling setup_ui()...")
        self.setup_ui()

    def setup_ui(self):
        """Sets up the UI elements only if parent exists"""
        debug_print("DEBUG: Attempting UI setup")
        if not self.parent.winfo_exists():
            print("ERROR: Parent destroyed before UI setup")
            return

        try:
            self.frame = Frame(self.parent)
            self.frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            self.canvas = Canvas(self.frame, height = 140)
            self.scrollbar = Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
            self.scrollable_frame = Frame(self.canvas)

            # Create window and store its ID
            self.scrollable_frame_id = self.canvas.create_window(
                (0, 0), 
                window=self.scrollable_frame, 
                anchor="nw",
                width=self.canvas.winfo_width()  # Initial width
            )
            
            # Update canvas scroll region and frame width on resize
            self.canvas.bind("<Configure>", lambda e: (
                self.canvas.itemconfig(self.scrollable_frame_id, width=e.width),
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            ))
            
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            
            self.canvas.pack(side="left", fill="both", expand=True)
            self.scrollbar.pack(side="right", fill="y")
            debug_print("DEBUG: UI setup completed")
        except Exception as e:
            print(f"CRITICAL: UI setup failed - {str(e)}")

        self.frame.pack_propagate(True) # allow frame to resize to fit images
        self.scrollable_frame.pack_propagate(True)
        self.canvas.update_idletasks()

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
        
        new_files = filedialog.askopenfilenames(title="Select Images", filetypes = file_types)
        if not new_files:
            debug_print("Debug: No new images selected")

        crop_state = self.main_gui.crop_enabled.get() # capture crop state at time of loading
        for file in new_files:
            if file not in self.image_files:
                self.image_crop_states[file] = crop_state # store crop state for each file after extending list

        
        self.image_files.extend(new_files)        
                
        self.main_gui.store_images(self.main_gui.selected_sheet.get(), self.image_files)
                
        
        
        debug_print(f"DEBUG: Selected image files: {self.image_files}")
        debug_print(f"Debug: Crop states: {self.image_crop_states}")

        if not self.image_files:
            debug_print("DEBUG: No images selected.")
            return

        if self.on_images_selected:
            self.on_images_selected(self.image_files)

        debug_print("DEBUG: Calling display_images()...")
        self.display_images()

    def load_images_from_list(self, image_paths):
        """Load images from a provided list"""
        self.image_files = image_paths
        self.display_images()

    def display_images(self):
        """
        Loads and displays the selected images in the GUI with additional debugging.
        """
        prev_pos = self.canvas.yview()
        debug_print("DEBUG: Clearing previous image widgets...")
        for widget in self.scrollable_frame.winfo_children():
            debug_print(f"DEBUG: Destroying widget {widget}")
            widget.destroy()
        
        self.image_widgets.clear()
        self.close_buttons.clear()
        self.image_references = []  # Ensure we retain image references

        if not self.image_files:
            return

        # Ensure parent has updated size before calculating max width
        self.parent.update_idletasks()
        max_height = 135 # fixed max height for images
        

        debug_print(f"DEBUG: Total images to process: {len(self.image_files)}")

        # Calculate positions
        padding = 5
        current_x = padding
        current_y = padding
        max_row_height = 0
        frame_width = self.scrollable_frame.winfo_width()
        row_images = []

        for img_index, img_path in enumerate(self.image_files):
            # Handle both images and PDFs with place()
            # Create container frame for image + close button
            container = Frame(self.scrollable_frame)
            container.pack_propagate(False)

            if img_path.lower().endswith(".pdf"):
                label = Label(container, text=f"PDF: {os.path.basename(img_path)}")
                label.pack()
                item_width = label.winfo_reqwidth()
                item_height = label.winfo_reqheight()
            else:
                try:
                    img = self.process_image(img_path) # auto applies smart crop
                    if img is None:
                        continue

                    img_tk = ImageTk.PhotoImage(img)
                    label = Label(container, image = img_tk)
                    label.pack()
                    
                    self.image_references.append(img_tk)
                    item_width = img.width
                    item_height = img.height
                except Exception as e:
                    print(f"ERROR: {str(e)}")
                    continue

            # Create close button
            close_btn = Label(container,
                text="x",  # Unicode multiplication sign
                font=("Arial", 10, "bold"),
                fg="black",
                bg="white",
                cursor="hand2"
            )
            close_btn.place(relx=1.0, rely=0.0, x=0, y=0, anchor="ne")  # Position in top-right
            close_btn.bind("<Button-1>", lambda e, path=img_path: self.remove_image(path))

            self.image_widgets.append(container)
            self.close_buttons.append(close_btn)

            # Track container dimensions
            container.config(width=item_width, height=item_height)
            

            # Wrap to new row if needed
            if current_x + item_width > frame_width - padding:
                current_x = padding
                current_y += max_row_height + padding
                max_row_height = 0

            # Position container
            container.place(x=current_x, y=current_y)

            # Update tracking variables
            current_x += item_width + padding
            if item_height > max_row_height:
                max_row_height = item_height

        # Update scrolling area
        total_height = current_y + max_row_height + padding
        self.scrollable_frame.configure(height=total_height)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.parent.update_idletasks()
        self.canvas.yview_moveto(prev_pos[0])

    def remove_image(self, path_to_remove):
        """Remove an image from the list and refresh display"""
        # Filter out the removed path (removes all occurrences)
        self.image_files = [p for p in self.image_files if p != path_to_remove]
    
        # Update parent storage if callback exists
        if self.on_images_selected:
            self.on_images_selected(self.image_files)
        
        # Redraw images with updated list
        self.display_images()

    # Add to ImageLoader class
    def smart_crop(self, img, margin_percent=10):
        """Auto-crop image using adaptive thresholding and contour detection."""
        debug_print(f"DEBUG: Running smart_crop on image. Original size: {img.size}")

        # Convert to grayscale
        grayscale = img.convert("L")
        np_img = np.array(grayscale)
        blurred = cv2.GaussianBlur(np_img, (5,5), 0) # apply blur
        edges = cv2.Canny(blurred, 100, 200) # adjust thresholds if needed

        # Find contours (detects all edges)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        if not contours:
            print("WARNING: No significant contours detected. Returning original image.")
            return img  # Return original if no object is found

        # Find the tallest, narrow object by filtering contours
        selected_contour = None
        max_height = 0
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            aspect_ratio = h / w  # Tall objects should have a high aspect ratio
            if aspect_ratio > 2.5 and h > max_height:  # Adjust aspect ratio threshold as needed
                selected_contour = contour
                max_height = h

        if selected_contour is None:
            print("WARNING: No valid tall object detected. Returning original image.")
            return img

        # Get refined bounding box
        x, y, w, h = cv2.boundingRect(selected_contour)
        debug_print(f"DEBUG: Selected Bounding Box: (x={x}, y={y}, w={w}, h={h})")

        # Apply margin but ensure we stay within image bounds
        width, height = img.size
        margin_x = min(int(width * margin_percent / 100), x)  # Prevent negative margins
        margin_y = min(int(height * margin_percent / 100), y)

        crop_box = (
            max(0, x - margin_x),
            max(0, y - margin_y),
            min(width, x + w + margin_x),
            min(height, y + h + margin_y)
        )

        debug_print(f"DEBUG: Cropping image with box: {crop_box}")

        cropped_img = img.crop(crop_box)
        debug_print(f"DEBUG: Cropped image size: {cropped_img.size}")

        return cropped_img
    
    def process_image(self, img_path):
        """Load and process image with optional cropping based on toggle."""
        try:
            debug_print(f"DEBUG: Opening image: {img_path}")
            img = Image.open(img_path)
            debug_print(f"DEBUG: Loaded image size: {img.size}")

            should_crop = self.image_crop_states.get(img_path, False)

            if should_crop:
                debug_print("DEBUG: Auto-Crop was ENABLED at load time. Applying smart_crop.")
                img = self.smart_crop(img, margin_percent=10)
            else:
                debug_print("DEBUG: Auto-Crop was DISABLED at load time. Skipping cropping.")

            # Resize for display
            max_height = 135
            scaling_factor = max_height / img.height
            new_width = int(img.width * scaling_factor)

            debug_print(f"DEBUG: Resizing image to: ({new_width}, {max_height})")
            return img.resize((new_width, max_height), Image.Resampling.LANCZOS)

        except Exception as e:
            print(f"ERROR: Failed to process image {img_path}: {e}")
            return None


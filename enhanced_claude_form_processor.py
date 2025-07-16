"""
enhanced_claude_form_processor.py
Enhanced Claude AI integration for batch processing sensory evaluation forms
with shadow removal preprocessing and interactive review interface
"""

import anthropic
import base64
import json
import os
from PIL import Image, ImageTk
import io
import cv2
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import glob

class EnhancedClaudeFormProcessor:
    """
    Enhanced processor that uses Claude AI for batch processing of sensory evaluation forms
    with shadow removal preprocessing and interactive review capabilities.
    """
    
    def __init__(self):
        # Try to get API key from environment variable
        self.api_key = os.getenv('ANTHROPIC_API_KEY')
        
        if not self.api_key:
            raise ValueError(
                "Anthropic API key not found. Please set the ANTHROPIC_API_KEY environment variable.\n\n"
                "You can get an API key from: https://console.anthropic.com/\n\n"
                "Set it in your environment:\n"
                "Windows: set ANTHROPIC_API_KEY=your_key_here\n"
                "Mac/Linux: export ANTHROPIC_API_KEY=your_key_here"
            )
        
        self.client = anthropic.Anthropic(api_key=self.api_key)
        
        # Expected metrics from your sensory evaluation form
        self.metrics = [
            "Burnt Taste",
            "Vapor Volume", 
            "Overall Flavor",
            "Smoothness",
            "Overall Liking"
        ]
        
        # Batch processing results
        self.batch_results = {}
        self.current_review_index = 0

        # NEW: Add callback mechanism and review completion tracking
        self.final_reviewed_results = {}
        self.review_complete = False
        self.main_app_callback = None
    
        print("DEBUG: EnhancedClaudeFormProcessor initialized with callback support")
    def set_main_app_callback(self, callback_function):
        """Set callback function to communicate with main application."""
        self.main_app_callback = callback_function
        print("DEBUG: Main application callback set")

    def get_reviewed_results(self):
        """Return the final reviewed results."""
        return self.final_reviewed_results

    def is_review_complete(self):
        """Check if review process is complete."""
        return self.review_complete

    def apply_shadow_removal(self, image):
        """
        Apply shadow removal preprocessing to improve form readability.
        Uses the same techniques from your existing shadow removal code.
        """
        
        # Shadow removal pipeline
        region_planes = cv2.split(image)
        result_planes = []
        for plane in region_planes:
            dilated_img = cv2.dilate(plane, np.ones((7,7), np.uint8))
            bg_img = cv2.medianBlur(dilated_img, 21)
            diff_img = 255 - cv2.absdiff(plane, bg_img)
            result_planes.append(diff_img)
        normalized = cv2.merge(result_planes)
        return normalized
    
    def prepare_image_with_preprocessing(self, image_path):
        """
        Prepare image for Claude API with shadow removal preprocessing.
        """
        
        try:
            print(f"  Preprocessing: {os.path.basename(image_path)}")
            
            # Load image
            image = cv2.imread(image_path)
            if image is None:
                raise ValueError(f"Could not load image: {image_path}")
            
            # Apply shadow removal
            processed = self.apply_shadow_removal(image)
            
            # Convert back to PIL Image for API processing
            pil_image = Image.fromarray(processed)
            
            # Convert to RGB 
            if pil_image.mode != 'RGB':
                pil_image = pil_image.convert('RGB')
            
            # Resize if too large (Claude has size limits)
            max_size = 1568  # Claude's recommended max dimension
            if max(pil_image.width, pil_image.height) > max_size:
                ratio = max_size / max(pil_image.width, pil_image.height)
                new_size = (int(pil_image.width * ratio), int(pil_image.height * ratio))
                pil_image = pil_image.resize(new_size, Image.Resampling.LANCZOS)
                print(f"    Resized to {new_size} for API processing")
            
            # Convert to base64
            buffer = io.BytesIO()
            pil_image.save(buffer, format='JPEG', quality=95)
            image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
            return image_data, processed  # Return both API data and processed image for review
            
        except Exception as e:
            raise Exception(f"Failed to prepare image {image_path}: {e}")
    
    def process_batch_images(self, image_folder):
        """
        Process a batch of images from a folder.
        """
        
        print("="*80)
        print("CLAUDE AI BATCH FORM PROCESSING")
        print("="*80)
        
        # Find all image files (avoiding duplicates)
        image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.tiff']
        image_files = set()  # Use set to avoid duplicates
        
        for ext in image_extensions:
            # Add both lowercase and uppercase patterns
            image_files.update(glob.glob(os.path.join(image_folder, ext)))
            image_files.update(glob.glob(os.path.join(image_folder, ext.upper())))
        
        # Convert back to sorted list
        image_files = sorted(list(image_files))
        
        if not image_files:
            raise ValueError(f"No image files found in {image_folder}")
        
        print(f"Found {len(image_files)} images to process")
        
        # Process each image
        self.batch_results = {}
        failed_images = []
        
        for i, image_path in enumerate(image_files, 1):
            try:
                print(f"\nProcessing {i}/{len(image_files)}: {os.path.basename(image_path)}")
                
                # Preprocess and prepare image
                image_data, processed_image = self.prepare_image_with_preprocessing(image_path)
                
                # Process with Claude
                extracted_data = self.process_single_image_with_claude(image_data, image_path)
                
                # Store results
                self.batch_results[image_path] = {
                    'extracted_data': extracted_data,
                    'processed_image': processed_image,
                    'status': 'success'
                }
                
                print(f"  ✓ Successfully processed - found {len(extracted_data)} samples")
                
            except Exception as e:
                print(f"  ✗ Failed to process: {e}")
                failed_images.append(image_path)
                self.batch_results[image_path] = {
                    'extracted_data': None,
                    'processed_image': None,
                    'status': 'failed',
                    'error': str(e)
                }
        
        print(f"\n" + "="*80)
        print("BATCH PROCESSING COMPLETE")
        print("="*80)
        print(f"Successfully processed: {len(image_files) - len(failed_images)}/{len(image_files)}")
        if failed_images:
            print(f"Failed images: {len(failed_images)}")
            for failed in failed_images:
                print(f"  - {os.path.basename(failed)}")
        
        return self.batch_results
    
    def process_single_image_with_claude(self, image_data, image_path):
        """
        Process a single preprocessed image with Claude AI.
        """
        
        # Create the enhanced prompt
        prompt = self.create_enhanced_extraction_prompt()
        
        try:
            # Send request to Claude
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=2000,  # Increased for sample names
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
            
            # Extract and parse the response
            response_text = response.content[0].text
            extracted_data = self.parse_claude_response(response_text)
            
            return extracted_data
            
        except anthropic.APIError as e:
            raise Exception(f"Claude API error: {e}")
        except Exception as e:
            raise Exception(f"AI processing failed: {e}")
    
    def create_enhanced_extraction_prompt(self):
        """
        Create an enhanced prompt that includes sample name extraction.
        """
        
        prompt = f"""
Please analyze this sensory evaluation form image and extract the ratings data along with sample names.

FORM STRUCTURE:
- The form contains 4 samples arranged in a 2x2 grid
- Each sample has a NAME/ID above the "Burnt Taste" row (look for text labels)
- Each sample has ratings for 5 attributes: {', '.join(self.metrics)}
- Ratings are on a 1-9 scale where participants circle their chosen number
- Look for circled numbers, checkmarks, or other clear markings indicating the selected rating

EXTRACTION INSTRUCTIONS:
1. Identify each sample section (4 total in a 2x2 grid)
2. For each sample, find the SAMPLE NAME/ID that appears above the "Burnt Taste" row
3. For each sample, find the rating for each of the 5 attributes
4. Look for circled numbers, heavy marks, checkmarks, or other clear indicators
5. If a rating is unclear or unmarked, use null
6. If a sample name is unclear, use a descriptive name like "Top Left Sample"

RESPONSE FORMAT:
Return the data as a JSON object with this structure (using actual sample names from the image):

{{
    "Sample A": {{
        "sample_name": "Sample A",
        "Burnt Taste": 5,
        "Vapor Volume": 7,
        "Overall Flavor": 6,
        "Smoothness": 8,
        "Overall Liking": 7,
        "comments": ""
    }},
    "Sample B": {{
        "sample_name": "Sample B", 
        "Burnt Taste": 4,
        "Vapor Volume": 6,
        "Overall Flavor": 5,
        "Smoothness": 7,
        "Overall Liking": 6,
        "comments": ""
    }},
    "Sample C": {{
        "sample_name": "Sample C",
        "Burnt Taste": 6,
        "Vapor Volume": 8,
        "Overall Flavor": 7,
        "Smoothness": 6,
        "Overall Liking": 8,
        "comments": ""
    }},
    "Sample D": {{
        "sample_name": "Sample D",
        "Burnt Taste": 3,
        "Vapor Volume": 5,
        "Overall Flavor": 4,
        "Smoothness": 5,
        "Overall Liking": 4,
        "comments": ""
    }}
}}

IMPORTANT:
- Use the actual sample names/IDs from the image as both key and sample_name value
- Use only numbers 1-9 for ratings, or null if unclear
- Include all 5 attributes for each sample
- Use the exact attribute names I provided
- Return only the JSON object, no additional text
- If you see handwritten comments, include them in the comments field
"""
        
        return prompt
    
    def parse_claude_response(self, response_text):
        """
        Parse Claude's response and convert it to the expected format.
        """
        
        try:
            # Find JSON in the response
            response_text = response_text.strip()
            
            # Look for JSON object boundaries
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}') + 1
            
            if start_idx == -1 or end_idx == 0:
                raise ValueError("No JSON object found in response")
            
            json_text = response_text[start_idx:end_idx]
            
            # Parse the JSON
            data = json.loads(json_text)
            
            # Validate the structure
            validated_data = {}
            
            for sample_key, sample_data in data.items():
                if not isinstance(sample_data, dict):
                    continue
                    
                validated_sample = {}
                
                # Extract sample name
                sample_name = sample_data.get('sample_name', sample_key)
                validated_sample['sample_name'] = str(sample_name)
                
                # Validate each metric
                for metric in self.metrics:
                    rating = sample_data.get(metric)
                    
                    # Convert and validate rating
                    if rating is None or rating == "null":
                        validated_sample[metric] = None  # Keep as None for review
                    elif isinstance(rating, (int, float)):
                        # Ensure rating is in valid range
                        rating = max(1, min(9, int(rating)))
                        validated_sample[metric] = rating
                    else:
                        print(f"Warning: Invalid rating '{rating}' for {sample_key} - {metric}")
                        validated_sample[metric] = None
                
                # Include comments
                validated_sample['comments'] = str(sample_data.get('comments', ''))
                
                validated_data[sample_key] = validated_sample
            
            if not validated_data:
                raise ValueError("No valid sample data found in response")
            
            return validated_data
            
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in Claude response: {e}\n\nResponse: {response_text[:200]}...")
        except Exception as e:
            raise ValueError(f"Failed to parse Claude response: {e}")
    
    def launch_review_interface(self):
        """
        Launch the interactive review interface to verify and edit results.
        """
        
        if not self.batch_results:
            messagebox.showwarning("No Data", "No batch processing results to review. Process some images first.")
            return
        
        # Filter successful results for review
        self.review_items = [(path, data) for path, data in self.batch_results.items() 
                            if data['status'] == 'success']
        
        if not self.review_items:
            messagebox.showerror("No Valid Results", "No successfully processed images to review.")
            return
        
        self.current_review_index = 0
        self.create_review_window()
    
    def create_review_window(self):
        """
        Create the main review interface window with improved layout.
        """
        
        self.review_window = tk.Toplevel()
        self.review_window.title("Claude AI Form Processing Review")
        self.review_window.geometry("1600x1000")  # CHANGED: Larger window
        self.review_window.configure(bg='white')
        
        # Create main frame with scrollbar
        main_frame = tk.Frame(self.review_window, bg='white')
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Header frame
        header_frame = tk.Frame(main_frame, bg='lightblue', height=60)
        header_frame.pack(fill='x', pady=(0, 10))
        header_frame.pack_propagate(False)
        
        # Progress info
        self.progress_label = tk.Label(header_frame, text="", font=('Arial', 14, 'bold'), bg='lightblue')
        self.progress_label.pack(side='left', padx=10, pady=15)
        
        # Navigation buttons
        nav_frame = tk.Frame(header_frame, bg='lightblue')
        nav_frame.pack(side='right', padx=10, pady=10)
        
        self.prev_btn = tk.Button(nav_frame, text="◀ Previous", command=self.previous_image,
                                 bg='lightcoral', font=('Arial', 10, 'bold'))
        self.prev_btn.pack(side='left', padx=5)
        
        self.next_btn = tk.Button(nav_frame, text="Next ▶", command=self.next_image,
                                 bg='lightgreen', font=('Arial', 10, 'bold'))
        self.next_btn.pack(side='left', padx=5)
        
        save_btn = tk.Button(nav_frame, text="💾 Save All", command=self.save_all_results,
                            bg='gold', font=('Arial', 10, 'bold'))
        save_btn.pack(side='left', padx=10)
        
        # Content frame with new proportions
        content_frame = tk.Frame(main_frame, bg='white')
        content_frame.pack(fill='both', expand=True)
        
        # CHANGED: Left side - Image display (larger proportion)
        image_frame = tk.Frame(content_frame, bg='white', width=1000)  # Increased width
        image_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        self.image_label = tk.Label(image_frame, bg='white', text="Loading...")
        self.image_label.pack(pady=10)
        
        # CHANGED: Right side - Data editing (smaller proportion for 2x2 grid)
        edit_frame = tk.Frame(content_frame, bg='lightgray', width=600)  # Fixed smaller width
        edit_frame.pack(side='right', fill='both')
        edit_frame.pack_propagate(False)
        
        # Scrollable frame for data editing
        canvas = tk.Canvas(edit_frame, bg='white')
        scrollbar = ttk.Scrollbar(edit_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg='white')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Initialize display
        self.update_review_display()
    
    def update_review_display(self):
        """
        Update the review display with current image and data.
        """
        
        if not self.review_items:
            return
        
        image_path, data = self.review_items[self.current_review_index]
        
        # Update progress
        progress_text = f"Image {self.current_review_index + 1} of {len(self.review_items)}: {os.path.basename(image_path)}"
        self.progress_label.config(text=progress_text)
        
        # Update navigation buttons
        self.prev_btn.config(state='normal' if self.current_review_index > 0 else 'disabled')
        self.next_btn.config(state='normal' if self.current_review_index < len(self.review_items) - 1 else 'disabled')
        
        # Display processed image
        self.display_processed_image(data['processed_image'])
        
        # Display editable data
        self.display_editable_data(data['extracted_data'])
    
    def display_processed_image(self, processed_image):
        """
        Display the shadow-removed processed image with 2x larger size.
        """
        
        try:
            # Convert to PIL and resize for display - INCREASED SIZE
            pil_image = Image.fromarray(processed_image)
            
            # CHANGED: Larger display size (2x increase)
            display_width = 1000  # Was 580, now 1000
            aspect_ratio = pil_image.height / pil_image.width
            display_height = int(display_width * aspect_ratio)
            
            # CHANGED: Higher max height
            if display_height > 800:  # Was 400, now 800
                display_height = 800
                display_width = int(display_height / aspect_ratio)
            
            pil_image = pil_image.resize((display_width, display_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(pil_image)
            
            # Update label
            self.image_label.config(image=photo, text="")
            self.image_label.image = photo  # Keep a reference
            
        except Exception as e:
            self.image_label.config(image="", text=f"Error displaying image: {e}")
    
    def display_editable_data(self, extracted_data):
        """
        Display editable form data in a 2x2 grid layout to match the form structure.
        """
        
        # Clear previous widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # Title
        title_label = tk.Label(self.scrollable_frame, text="Extracted Data (Click to Edit)", 
                              font=('Arial', 14, 'bold'), bg='white')
        title_label.pack(pady=10)
        
        # Store entry widgets for saving
        self.entry_widgets = {}
        
        # Convert to list for easier indexing
        sample_items = list(extracted_data.items())
        
        # CHANGED: Create 2x2 grid layout to match form structure
        grid_frame = tk.Frame(self.scrollable_frame, bg='white')
        grid_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Configure grid weights for even distribution
        grid_frame.grid_columnconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(1, weight=1)
        grid_frame.grid_rowconfigure(0, weight=1)  # ✅ CORRECT
        grid_frame.grid_rowconfigure(1, weight=1)
        
        for i, (sample_key, sample_data) in enumerate(sample_items):
            # Calculate grid position (2x2 layout)
            row = i // 2  # 0 for first two samples, 1 for last two
            col = i % 2   # 0 for left column, 1 for right column
            
            # CHANGED: Smaller sample frames (60% width) arranged in grid
            sample_frame = tk.LabelFrame(grid_frame, text=f"{sample_key}", 
                                       font=('Arial', 11, 'bold'), bg='white', 
                                       padx=8, pady=8, relief='raised', bd=2)
            sample_frame.grid(row=row, column=col, sticky='nsew', 
                            padx=8, pady=8)  # Even spacing
            
            self.entry_widgets[sample_key] = {}
            
            # CHANGED: Compact sample name section
            name_frame = tk.Frame(sample_frame, bg='white')
            name_frame.pack(fill='x', pady=(0, 8))
            
            tk.Label(name_frame, text="Name:", font=('Arial', 9, 'bold'), 
                    bg='white', width=8, anchor='w').pack(side='left')
            name_entry = tk.Entry(name_frame, font=('Arial', 9), width=15)  # Shorter width
            name_entry.pack(side='left', padx=(3, 0), fill='x', expand=True)
            name_entry.insert(0, str(sample_data.get('sample_name', '')))
            self.entry_widgets[sample_key]['sample_name'] = name_entry
            
            # CHANGED: Compact metrics section
            metrics_frame = tk.Frame(sample_frame, bg='white')
            metrics_frame.pack(fill='both', expand=True)
            
            for metric in self.metrics:
                metric_frame = tk.Frame(metrics_frame, bg='white')
                metric_frame.pack(fill='x', pady=1)
                
                # CHANGED: Shorter labels and entries
                label_text = metric.replace(' ', '\n') if len(metric) > 12 else metric  # Break long labels
                metric_label = tk.Label(metric_frame, text=f"{label_text}:", 
                                      font=('Arial', 8), bg='white', 
                                      width=12, anchor='w', justify='left')
                metric_label.pack(side='left')
                
                metric_entry = tk.Entry(metric_frame, font=('Arial', 9), width=8)  # Shorter width
                metric_entry.pack(side='left', padx=(3, 0))
                
                value = sample_data.get(metric)
                if value is not None:
                    metric_entry.insert(0, str(value))
                
                self.entry_widgets[sample_key][metric] = metric_entry
            
            # CHANGED: Compact comments section
            comments_frame = tk.Frame(sample_frame, bg='white')
            comments_frame.pack(fill='x', pady=(8, 0))
            
            tk.Label(comments_frame, text="Comments:", font=('Arial', 8, 'bold'), 
                    bg='white', width=12, anchor='w').pack(side='top', anchor='w')
            
            # CHANGED: Smaller text widget
            comments_text = tk.Text(comments_frame, font=('Arial', 8), 
                                  width=25, height=3, wrap='word')  # Smaller dimensions
            comments_text.pack(fill='x', expand=True, pady=(2, 0))
            comments_text.insert('1.0', str(sample_data.get('comments', '')))
            self.entry_widgets[sample_key]['comments'] = comments_text
        
        # If we have fewer than 4 samples, add placeholder frames to maintain layout
        total_samples = len(sample_items)
        if total_samples < 4:
            for i in range(total_samples, 4):
                row = i // 2
                col = i % 2
                placeholder_frame = tk.Frame(grid_frame, bg='lightgray', relief='sunken', bd=1)
                placeholder_frame.grid(row=row, column=col, sticky='nsew', padx=8, pady=8)
                tk.Label(placeholder_frame, text="No Sample", 
                        font=('Arial', 10), bg='lightgray', fg='gray').pack(expand=True)
    
    def previous_image(self):
        """Navigate to previous image."""
        if self.current_review_index > 0:
            self.save_current_edits()
            self.current_review_index -= 1
            self.update_review_display()
    
    def next_image(self):
        """Navigate to next image."""
        if self.current_review_index < len(self.review_items) - 1:
            self.save_current_edits()
            self.current_review_index += 1
            self.update_review_display()
    
    def save_current_edits(self):
        """Save the current edits back to the data structure."""
        if not self.review_items or not hasattr(self, 'entry_widgets'):
            return
        
        image_path, data = self.review_items[self.current_review_index]
        
        for sample_key, widgets in self.entry_widgets.items():
            if sample_key in data['extracted_data']:
                # Update sample name
                data['extracted_data'][sample_key]['sample_name'] = widgets['sample_name'].get()
                
                # Update metrics
                for metric in self.metrics:
                    value = widgets[metric].get().strip()
                    if value:
                        try:
                            data['extracted_data'][sample_key][metric] = int(value)
                        except ValueError:
                            data['extracted_data'][sample_key][metric] = None
                    else:
                        data['extracted_data'][sample_key][metric] = None
                
                # CHANGED: Update comments (now Text widget instead of Entry)
                if isinstance(widgets['comments'], tk.Text):
                    data['extracted_data'][sample_key]['comments'] = widgets['comments'].get('1.0', tk.END).strip()
                else:
                    data['extracted_data'][sample_key]['comments'] = widgets['comments'].get()
    
    def save_all_results(self):
        """Save all results and return to main application with data loaded."""
        print("DEBUG: Starting save_all_results process")
    
        # Save current edits first
        self.save_current_edits()
        print("DEBUG: Current edits saved")
    
        # Prepare the reviewed data for loading into main application
        reviewed_batch_results = {}
    
        for image_path, data in self.batch_results.items():
            if data['status'] == 'success':
                reviewed_batch_results[image_path] = data
                print(f"DEBUG: Prepared reviewed data for {image_path}")
    
        print(f"DEBUG: Total reviewed results: {len(reviewed_batch_results)}")
    
        # Store the reviewed results for the main application to access
        self.final_reviewed_results = reviewed_batch_results
    
        # Signal that review is complete
        self.review_complete = True
    
        # Close the review window
        if hasattr(self, 'review_window') and self.review_window:
            print("DEBUG: Closing review window")
            self.review_window.destroy()
            self.review_window = None
    
        print("DEBUG: Review process completed successfully")
    
        # Show success message
        total_samples = sum(len(data['extracted_data']) for data in reviewed_batch_results.values() 
                           if data['status'] == 'success')
    
        print(f"DEBUG: Total samples to be loaded: {total_samples}")
    
        # Don't show the messagebox here since we want seamless transition
        # The main application will handle the loading and show appropriate messages


# Usage example and testing
if __name__ == "__main__":
    try:
        processor = EnhancedClaudeFormProcessor()
        print("Enhanced Claude Form Processor initialized successfully!")
        print("\nUsage:")
        print("  # Process a batch of images")
        print("  results = processor.process_batch_images('path/to/image/folder')")
        print("  ")
        print("  # Launch review interface")
        print("  processor.launch_review_interface()")
        
    except Exception as e:
        print(f"Initialization error: {e}")
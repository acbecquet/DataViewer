"""
VAP3 File Handler Module for DataViewer Application

This module handles loading and saving of .vap3 files (custom application format).
"""

# Standard library imports
import os
import copy
import traceback

# Third party imports
import tkinter as tk
from tkinter import filedialog, messagebox

# Local imports
from utils import debug_print, show_success_message


class Vap3FileHandler:
    """Handles VAP3 file loading and saving operations."""
    
    def __init__(self, file_manager):
        """Initialize with reference to parent FileManager."""
        self.file_manager = file_manager
        self.gui = file_manager.gui
        self.root = file_manager.root
        
    def save_as_vap3(self, filepath=None) -> None:
        """Save the current session to a .vap3 file."""
        from vap_file_manager import VapFileManager

        if not filepath:
            filepath = filedialog.asksaveasfilename(
                title="Save As VAP3",
                defaultextension=".vap3",
                filetypes=[("VAP3 files", "*.vap3")]
            )

        if not filepath:
            return

        try:
            self.gui.progress_dialog.show_progress_bar("Saving VAP3 file...")
            self.root.update_idletasks()

            vap_manager = VapFileManager()

            # Collect plot settings
            plot_settings = {
                'selected_plot_type': self.gui.selected_plot_type.get() if hasattr(self.gui, 'selected_plot_type') else None
            }

            # Collect image sample mapping
            if hasattr(self.gui, 'image_sample_mapping') and self.gui.image_sample_mapping:
                plot_settings['image_sample_mapping'] = self.gui.image_sample_mapping.copy()
                debug_print(f"DEBUG: Collected image sample mapping for save: {len(self.gui.image_sample_mapping)} entries")

            # Get image crop states
            image_crop_states = getattr(self.gui, 'image_crop_states', {})
            if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
                image_crop_states.update(self.gui.image_loader.image_crop_states)

            success = vap_manager.save_to_vap3(
                filepath,
                self.gui.filtered_sheets,
                self.gui.sheet_images,
                self.gui.plot_options,
                image_crop_states,
                plot_settings
            )

            self.gui.progress_dialog.update_progress_bar(100)
            self.root.update_idletasks()

            if success:
                show_success_message("Success", f"Data saved successfully to {filepath}", self.gui.root)
            else:
                messagebox.showerror("Error", "Failed to save data")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the file: {e}")
        finally:
            self.gui.progress_dialog.hide_progress_bar()

    def load_vap3_file(self, filepath=None, display_name=None, append_to_existing=False) -> bool:
        """Load a .vap3 file and update the application state."""
        from vap_file_manager import VapFileManager

        if not filepath:
            filepath = filedialog.askopenfilename(
                title="Open VAP3 File",
                filetypes=[("VAP3 files", "*.vap3")]
            )

        if not filepath:
            return False

        try:
            self.gui.progress_dialog.show_progress_bar("Loading VAP3 file...")
            self.root.update_idletasks()

            vap_manager = VapFileManager()
            result = vap_manager.load_from_vap3(filepath)

            # Use display_name if provided, otherwise use the actual filename
            if display_name:
                current_file_name = display_name
                debug_print(f"DEBUG: Using provided display name: {display_name}")
            else:
                current_file_name = os.path.basename(filepath)
                debug_print(f"DEBUG: Using actual filepath basename: {os.path.basename(filepath)}")

            # Check if this file is already loaded (avoid duplicates)
            existing_file = None
            for file_data in self.gui.all_filtered_sheets:
                if file_data["file_name"] == current_file_name or file_data["file_path"] == filepath:
                    existing_file = file_data
                    debug_print(f"DEBUG: File already loaded: {current_file_name}")
                    break

            vap_data = vap_manager.load_from_vap3(filepath)
            self.gui.load_sample_images_from_vap3(vap_data)

            if existing_file:
                # File already loaded, just make it active
                self.file_manager.set_active_file(existing_file["file_name"])
                self.file_manager.update_ui_for_current_file()
                debug_print(f"DEBUG: Switched to existing file: {existing_file['file_name']}")
            else:
                # New file, add it to the collection
                debug_print(f"DEBUG: Adding new file: {current_file_name}")

                # Update current session data (this will be the active file)
                self.gui.filtered_sheets = result['filtered_sheets']
                self.gui.sheets = {name: sheet_info['data'] for name, sheet_info in result['filtered_sheets'].items()}

                # Store old image sample mapping temporarily
                old_image_sample_mapping = {}
                if 'plot_settings' in result and result['plot_settings']:
                    if 'image_sample_mapping' in result['plot_settings']:
                        old_image_sample_mapping = result['plot_settings']['image_sample_mapping']
                        debug_print(f"DEBUG: Found image sample mapping in VAP3: {len(old_image_sample_mapping)} entries")

                # Handle sheet images - need to be careful about file-specific images
                if 'sheet_images' in result and result['sheet_images']:
                    if not hasattr(self.gui, 'sheet_images'):
                        self.gui.sheet_images = {}
                    
                    if result['sheet_images']:
                        vap3_key = list(result['sheet_images'].keys())[0]
                        self.gui.sheet_images[current_file_name] = result['sheet_images'][vap3_key]
                        debug_print(f"DEBUG: Loaded sheet images for {current_file_name} (from VAP3 key: {vap3_key})")
                        debug_print(f"DEBUG: Sheets with images: {list(self.gui.sheet_images[current_file_name].keys())}")
                        for sheet_name, imgs in self.gui.sheet_images[current_file_name].items():
                            debug_print(f"DEBUG: Sheet {sheet_name}: {len(imgs)} images")

                        # Initialize image_sample_mapping if needed
                        if not hasattr(self.gui, 'image_sample_mapping'):
                            self.gui.image_sample_mapping = {}

                        # Rebuild image sample mapping with new temporary paths
                        if old_image_sample_mapping:
                            # Create mapping of old basenames to sample numbers
                            basename_to_sample = {}
                            for old_path, sample_num in old_image_sample_mapping.items():
                                basename = os.path.basename(old_path)
                                basename_to_sample[basename] = sample_num
                                debug_print(f"DEBUG: Old mapping: {basename} -> Sample {sample_num}")
                            
                            # Map new temporary paths to sample numbers using basenames
                            for sheet_name, new_image_paths in self.gui.sheet_images[current_file_name].items():
                                for new_path in new_image_paths:
                                    new_basename = os.path.basename(new_path)
                                    if new_basename in basename_to_sample:
                                        sample_num = basename_to_sample[new_basename]
                                        self.gui.image_sample_mapping[new_path] = sample_num
                                        debug_print(f"DEBUG: Mapped new path to Sample {sample_num}: {new_basename}")
                            
                            debug_print(f"DEBUG: Rebuilt image sample mapping with {len(self.gui.image_sample_mapping)} entries")
                        else:
                            # FALLBACK: Try to reconstruct mapping from image filenames
                            # Excel embedded images often have predictable names
                            debug_print(f"DEBUG: No saved mapping found, attempting reconstruction from filenames")
                            
                            import re
                            for sheet_name, new_image_paths in self.gui.sheet_images[current_file_name].items():
                                # Sort images to ensure consistent ordering
                                sorted_paths = sorted(new_image_paths)
                                for idx, img_path in enumerate(sorted_paths):
                                    basename = os.path.basename(img_path)
                                    
                                    # Try to extract sample number from filename patterns
                                    # Pattern 1: image_0.jpeg, image_1.jpeg, etc.
                                    match = re.search(r'image[_\s-]*(\d+)', basename, re.IGNORECASE)
                                    if match:
                                        sample_num = int(match.group(1)) + 1  # Convert 0-based to 1-based
                                        self.gui.image_sample_mapping[img_path] = sample_num
                                        debug_print(f"DEBUG: Reconstructed mapping from filename: {basename} -> Sample {sample_num}")
                                    else:
                                        # Pattern 2: If no number found, use order-based assignment
                                        sample_num = idx + 1
                                        self.gui.image_sample_mapping[img_path] = sample_num
                                        debug_print(f"DEBUG: Assigned by order: {basename} -> Sample {sample_num}")
                            
                            if self.gui.image_sample_mapping:
                                debug_print(f"DEBUG: Reconstructed {len(self.gui.image_sample_mapping)} mappings from filenames")

                # Handle image crop states
                if 'image_crop_states' in result and result['image_crop_states']:
                    if not hasattr(self.gui, 'image_crop_states'):
                        self.gui.image_crop_states = {}
                    self.gui.image_crop_states.update(result['image_crop_states'])

                # Handle plot options and settings
                if 'plot_options' in result and result['plot_options']:
                    self.gui.plot_options = result['plot_options']

                if 'plot_settings' in result and result['plot_settings']:
                    if 'selected_plot_type' in result['plot_settings'] and result['plot_settings']['selected_plot_type']:
                        self.gui.selected_plot_type.set(result['plot_settings']['selected_plot_type'])

                # Set current file info
                self.gui.current_file = current_file_name
                self.gui.file_path = filepath

                # Add to all_filtered_sheets (append instead of replace if append_to_existing is True)
                new_file_data = {
                    "file_name": current_file_name,
                    "file_path": filepath,
                    "filtered_sheets": copy.deepcopy(result['filtered_sheets'])
                }

                if append_to_existing:
                    # Append to existing files
                    self.gui.all_filtered_sheets.append(new_file_data)
                    debug_print(f"DEBUG: Appended file to existing collection. Total files: {len(self.gui.all_filtered_sheets)}")
                else:
                    # Replace existing files (original behavior for single file loading)
                    self.gui.all_filtered_sheets = [new_file_data]
                    debug_print(f"DEBUG: Replaced file collection with single file")

                # Update UI
                self.file_manager.update_file_dropdown()
                self.file_manager.update_ui_for_current_file()

            self.gui.progress_dialog.update_progress_bar(100)
            self.root.update_idletasks()

            # Only show success message if this wasn't called from load_from_database
            if not display_name:
                show_success_message("Success", f"VAP3 file loaded successfully: {current_file_name}", self.gui.root)
            else:
                debug_print(f"DEBUG: VAP3 file loaded with display name: {current_file_name}")

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load VAP3 file: {str(e)}")
            return False
        finally:
            self.gui.progress_dialog.hide_progress_bar()
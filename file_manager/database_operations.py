"""
Database Operations Module for DataViewer Application

This module handles all database-related operations including storing, loading, 
and browsing files in the application database.
"""

# Standard library imports
import os
import re
import copy
import time
import tempfile
import traceback
from typing import Optional

# Third party imports
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, Button, ttk, Frame, Listbox, Scrollbar

# Local imports
from database_manager import DatabaseManager
from utils import debug_print, show_success_message, FONT, APP_BACKGROUND_COLOR, plotting_sheet_test

class DatabaseOperations:
    """Handles database storage, loading, and browsing operations."""
    
    def __init__(self, file_manager):
        """Initialize with reference to parent FileManager."""
        self.file_manager = file_manager
        self.gui = file_manager.gui
        self.root = file_manager.root
        self.db_manager = DatabaseManager()
        
    def _store_file_in_database(self, original_file_path, display_filename=None):
        """
        Store the current file in the database.
        Enhanced with duplicate checking, better error handling, and sample image support.
        """
        debug_print(f"DEBUG: Checking if file {original_file_path} needs database storage")

        # Check if already stored
        if original_file_path in self.file_manager.stored_files_cache:
            debug_print("DEBUG: File already stored in database, skipping")
            return

        try:
            debug_print("DEBUG: Storing new file in database...")
            # Show progress dialog
            self.gui.progress_dialog.show_progress_bar("Storing file in database...")
            self.gui.root.update_idletasks()

            # CREATE: Collect sample notes from current filtered_sheets (ADD THIS SECTION)
            sample_notes_data = {}
            for sheet_name, sheet_info in self.gui.filtered_sheets.items():
                header_data = sheet_info.get('header_data')
                if header_data and 'samples' in header_data:
                    sheet_notes = {}
                    for i, sample_data in enumerate(header_data['samples']):
                        sample_notes = sample_data.get('sample_notes', '')
                        if sample_notes.strip():
                            sheet_notes[f"Sample {i+1}"] = sample_notes

                    if sheet_notes:
                        sample_notes_data[sheet_name] = sheet_notes
                        debug_print(f"DEBUG: Collected notes for sheet {sheet_name}: {len(sheet_notes)} samples")

            # Create a temporary VAP3 file
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_vap3_path = temp_file.name
                debug_print(f"DEBUG: Created temporary VAP3 file: {temp_vap3_path}")

            # Save current state as VAP3
            from vap_file_manager import VapFileManager
            vap_manager = VapFileManager()

            # Collect plot settings safely
            plot_settings = {}
            if hasattr(self.gui, 'selected_plot_type'):
                plot_settings['selected_plot_type'] = self.gui.selected_plot_type.get()

            debug_print(f"DEBUG: Plot settings: {plot_settings}")

            # Get image crop states safely
            image_crop_states = getattr(self.gui, 'image_crop_states', {})
            if hasattr(self.gui, 'image_loader') and self.gui.image_loader:
                image_crop_states.update(getattr(self.gui.image_loader, 'image_crop_states', {}))

            debug_print(f"DEBUG: Image crop states: {len(image_crop_states)} items")

            # Collect sample images for database storage
            sample_images = {}
            sample_image_crop_states = {}
            sample_header_data = {}

            # Check if we have pending sample images
            if hasattr(self.gui, 'pending_sample_images'):
                sample_images = getattr(self.gui, 'pending_sample_images', {})
                sample_image_crop_states = getattr(self.gui, 'pending_sample_image_crop_states', {})
                sample_header_data = getattr(self.gui, 'pending_sample_header_data', {})
                debug_print(f"DEBUG: Found pending sample images: {len(sample_images)} samples")

            # Also check sample_image_metadata structure for all files/sheets
            if hasattr(self.gui, 'sample_image_metadata'):
                current_file = getattr(self.gui, 'current_file', None)
                if current_file and current_file in self.gui.sample_image_metadata:
                    for sheet_name, metadata in self.gui.sample_image_metadata[current_file].items():
                        sheet_sample_images = metadata.get('sample_images', {})
                        if sheet_sample_images:
                            # Merge with existing sample images
                            sample_images.update(sheet_sample_images)
                            sample_image_crop_states.update(metadata.get('sample_image_crop_states', {}))
                            if not sample_header_data:  # Use first header data found
                                sample_header_data = metadata.get('header_data', {})
                            debug_print(f"DEBUG: Found sample images in metadata for {sheet_name}: {len(sheet_sample_images)} samples")

            debug_print(f"DEBUG: Total sample images for database storage: {len(sample_images)} samples")

            if hasattr(self.gui, 'sheet_images'):
                current_file = getattr(self.gui, 'current_file', None) or original_file_path
                if current_file in self.gui.sheet_images:
                    debug_print(f"DEBUG: Checking for embedded Excel images in sheet_images")
                    # Sheet images contain the embedded images extracted from Excel
                    # We need to preserve these when storing to database
                    for sheet_name, image_paths in self.gui.sheet_images[current_file].items():
                        if image_paths:
                            debug_print(f"DEBUG: Found {len(image_paths)} embedded images for sheet: {sheet_name}")
                            # These images are already in the correct format and will be
                            # included via the sheet_images parameter in save_to_vap3

            debug_print(f"DEBUG: Total sample images for database storage: {len(sample_images)} samples")
            if sample_images:
                debug_print(f"DEBUG: Sample image groups: {list(sample_images.keys())}")

            # Extract and construct the display filename
            if display_filename is None:
                if original_file_path.endswith('.vap3'):
                    display_filename = os.path.basename(original_file_path)
                else:
                    # For Excel files, create .vap3 extension
                    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
                    display_filename = f"{base_name}.vap3"

            debug_print(f"DEBUG: Using display filename: {display_filename}")

            # Save the VAP3 file with sample images included
            success = vap_manager.save_to_vap3(
                temp_vap3_path,
                self.gui.filtered_sheets,
                getattr(self.gui, 'sheet_images', {}),
                getattr(self.gui, 'plot_options', []),
                image_crop_states,
                plot_settings,
                sample_images,
                sample_image_crop_states,
                sample_header_data
            )

            if not success:
                raise Exception("Failed to create temporary VAP3 file")

            debug_print("DEBUG: VAP3 file created successfully with sample images")

            # Store metadata about the original file (ADD sample_notes to metadata)
            meta_data = {
                'display_filename': display_filename,
                'original_filename': os.path.basename(original_file_path),
                'original_path': original_file_path,
                'creation_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                'sheet_count': len(self.gui.filtered_sheets),
                'plot_options': getattr(self.gui, 'plot_options', []),
                'plot_settings': plot_settings,
                'has_sample_images': bool(sample_images),
                'sample_count': len(sample_images),
                'sample_notes': sample_notes_data 
            }

            debug_print(f"DEBUG: Metadata to store: {meta_data}")

            # Store the VAP3 file in the database with the proper display filename
            file_id = self.db_manager.store_vap3_file(temp_vap3_path, meta_data)
            debug_print(f"DEBUG: File stored with ID: {file_id}")

            # Store sheet metadata
            for sheet_name, sheet_info in self.gui.filtered_sheets.items():
                is_plotting = plotting_sheet_test(sheet_name, sheet_info["data"])
                is_empty = sheet_info.get("is_empty", False)

                sheet_id = self.db_manager.store_sheet_info(
                    file_id,
                    sheet_name,
                    is_plotting,
                    is_empty
                )
                debug_print(f"DEBUG: Stored sheet '{sheet_name}' with ID: {sheet_id}")

            # Store associated images
            if hasattr(self.gui, 'sheet_images') and hasattr(self.gui, 'current_file') and self.gui.current_file in self.gui.sheet_images:
                current_file = self.gui.current_file
                for sheet_name, image_paths in self.gui.sheet_images[current_file].items():
                    for image_path in image_paths:
                        if os.path.exists(image_path):
                            crop_enabled = image_crop_states.get(image_path, False)
                            self.db_manager.store_image(file_id, sheet_name, image_path, crop_enabled)
                            debug_print(f"DEBUG: Stored image for sheet {sheet_name}: {os.path.basename(image_path)}")

            # Clean up temporary file
            try:
                os.unlink(temp_vap3_path)
            except:
                pass

            # Add to cache to prevent re-storing
            self.file_manager.stored_files_cache.add(original_file_path)

            debug_print("DEBUG: File stored in database successfully")

        except Exception as e:
            debug_print(f"ERROR: Failed to store file in database: {e}")
            import traceback
            traceback.print_exc()
            raise e
        finally:
            self.gui.progress_dialog.hide_progress_bar()

    def load_from_database(self, file_id=None, show_success_msg=True, batch_operation=False):
        """Load a file from the database."""
        debug_print(f"DEBUG: load_from_database called with file_id={file_id}, show_success_message={show_success_msg}, batch_operation={batch_operation}")
        try:
            # Only show progress dialog if not part of a batch operation
            if not batch_operation:
                self.gui.progress_dialog.show_progress_bar("Loading from database...")
                self.gui.root.update_idletasks()

            if file_id is None:
                # Show a dialog to select from available files
                file_list = self.db_manager.list_files()
                if not file_list:
                    show_success_message("Info", "No files found in the database.", self.gui.root)
                    return False

                # Create file selection dialog
                selection_dialog = FileSelectionDialog(self.gui.root, file_list)
                file_id = selection_dialog.show()

                if file_id is None:
                    debug_print("DEBUG: No file selected from dialog")
                    return False

            # Get file data from database - CORRECTED METHOD NAME
            file_data = self.db_manager.get_file_by_id(file_id)
            if not file_data:
                if show_success_msg:
                    messagebox.showerror("Error", "File not found in database.")
                return False

            raw_database_filename = file_data['filename']
            debug_print(f"DEBUG: Raw database filename: {raw_database_filename}")
            created_at = file_data.get('created_at')

            # Get the proper display filename - prioritize metadata, then fallback to database filename
            display_filename = None

            # First, try to get display_filename from metadata
            if 'meta_data' in file_data and file_data['meta_data']:
                display_filename = file_data['meta_data'].get('display_filename')
                debug_print(f"DEBUG: display_filename from metadata: '{display_filename}'")

                # If not found, try original_filename from metadata and construct .vap3 name
                if not display_filename:
                    original_filename = file_data['meta_data'].get('original_filename')
                    if original_filename:
                        # Remove extension and add .vap3
                        display_filename = os.path.splitext(original_filename)[0] + '.vap3'
                        debug_print(f"DEBUG: Constructed filename from original_filename: '{display_filename}'")

            # Final fallback to database filename field
            if not display_filename:
                display_filename = file_data['filename']
                debug_print(f"DEBUG: Using database filename as fallback: '{display_filename}'")

            # Check if we already have files loaded to determine if we should append
            append_to_existing = len(self.gui.all_filtered_sheets) > 0

            # Save the VAP3 file to a temporary location
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_file.write(file_data['file_content'])
                temp_vap3_path = temp_file.name

            try:
                # CRITICAL: Load and process VAP3 data BEFORE loading the file
                from vap_file_manager import VapFileManager
                vap_manager = VapFileManager()
                vap_data = vap_manager.load_from_vap3(temp_vap3_path)
                debug_print(f"DEBUG: VAP3 data keys: {list(vap_data.keys())}")
                debug_print(f"DEBUG: Sample images in vap_data: {list(vap_data.get('sample_images', {}).keys())}")
                debug_print(f"DEBUG: Sample image counts: {[(k, len(v)) for k, v in vap_data.get('sample_images', {}).items()]}")
                # Store VAP3 data in GUI for sample image loading
                self.gui.current_vap_data = vap_data
                debug_print(f"DEBUG: Loaded VAP3 data with keys: {list(vap_data.keys())}")

                # Use the enhanced VAP3 loading that handles sample images
                success = self.file_manager.load_vap3_file(temp_vap3_path, display_name=display_filename, append_to_existing=append_to_existing)

                if success:
                    total_files = len(self.gui.all_filtered_sheets)
                    debug_print(f"DEBUG: Successfully loaded from database. Total files: {total_files}")

                    # CRITICAL FIX: Load sample images from the VAP3 data
                    # Load sample images if they exist
                    if 'sample_images' in vap_data and vap_data['sample_images']:
                        debug_print(f"DEBUG: Loading sample images from database VAP3")

                        # Load sample images and populate main GUI
                        self.gui.load_sample_images_from_vap3(vap_data)

                        # CRITICAL: Also populate the sample_image_metadata structure
                        sample_images = vap_data.get('sample_images', {})
                        sample_crop_states = vap_data.get('sample_image_crop_states', {})
                        sample_header_data = vap_data.get('sample_images_metadata', {}).get('header_data', {})

                        debug_print(f"DEBUG: Sample images content: {sample_images}")
                        debug_print(f"DEBUG: Sample images metadata content: {sample_header_data}")
                        debug_print(f"DEBUG: Sample crop states content: {sample_crop_states}")
                        debug_print(f"DEBUG: Sample images keys: {list(sample_images.keys()) if sample_images else 'Empty'}")
                        debug_print(f"DEBUG: Total sample image files: {sum(len(imgs) for imgs in sample_images.values()) if sample_images else 0}")

                        # Also check regular sheet images for comparison
                        sheet_images = vap_data.get('sheet_images', {})
                        debug_print(f"DEBUG: Sheet images content: {sheet_images}")
                        if sheet_images:
                            for file_key, sheets in sheet_images.items():
                                debug_print(f"DEBUG: File '{file_key}' sheet images:")
                                for sheet_name, images in sheets.items():
                                    debug_print(f"DEBUG:   Sheet '{sheet_name}': {len(images)} images - {images[:2] if images else 'None'}...")

                        if sample_images:
                            if not hasattr(self.gui, 'sample_image_metadata'):
                                self.gui.sample_image_metadata = {}
                            if display_filename not in self.gui.sample_image_metadata:
                                self.gui.sample_image_metadata[display_filename] = {}

                            # Determine which sheet this belongs to
                            test_name = sample_header_data.get('test', 'Unknown Test')
                            if test_name in self.gui.filtered_sheets:
                                self.gui.sample_image_metadata[display_filename][test_name] = {
                                    'sample_images': sample_images,
                                    'sample_image_crop_states': sample_crop_states,
                                    'header_data': sample_header_data,
                                    'test_name': test_name
                                }
                                debug_print(f"DEBUG: Populated sample_image_metadata for {test_name} in file {display_filename}")
                            else:
                                # If test_name not found, try to find a matching sheet
                                for sheet_name in self.gui.filtered_sheets.keys():
                                    if sheet_name.lower() == test_name.lower() or test_name.lower() in sheet_name.lower():
                                        self.gui.sample_image_metadata[display_filename][sheet_name] = {
                                            'sample_images': sample_images,
                                            'sample_image_crop_states': sample_crop_states,
                                            'header_data': sample_header_data,
                                            'test_name': sheet_name
                                        }
                                        debug_print(f"DEBUG: Populated sample_image_metadata for matched sheet {sheet_name} in file {display_filename}")
                                        break
                    else:
                        debug_print("DEBUG: No sample images found in VAP3 data")

                    # Store database-specific metadata in the latest file entry
                    if self.gui.all_filtered_sheets:
                        latest_file = self.gui.all_filtered_sheets[-1]
                        latest_file['database_filename'] = raw_database_filename
                        latest_file['database_created_at'] = created_at

                        # Also store in the original_filename if not already set
                        if 'original_filename' not in latest_file:
                            latest_file['original_filename'] = raw_database_filename

                        debug_print(f"DEBUG: Stored database filename in metadata: {raw_database_filename}")

                    # Show the success message only if requested and not in batch mode
                    if show_success_msg and not batch_operation:
                        if total_files > 1:
                            show_success_message("Success", f"VAP3 file loaded successfully: {display_filename}\nTotal files loaded: {total_files}", self.gui.root)
                        else:
                            show_success_message("Success", f"VAP3 file loaded successfully: {display_filename}", self.gui.root)

                    return True
                else:
                    if show_success_msg:
                        messagebox.showerror("Error", f"Failed to load file: {display_filename}")
                    return False

            finally:
                # Clean up temporary file
                try:
                    os.unlink(temp_vap3_path)
                except:
                    pass

        except Exception as e:
            if show_success_msg:  # Only show error dialog if not in batch mode
                messagebox.showerror("Error", f"Error loading file from database: {e}")
            debug_print(f"ERROR: Error loading file from database: {e}")
            traceback.print_exc()
            return False
        finally:
            # Hide progress dialog only if not in batch mode
            if not batch_operation:
                self.gui.progress_dialog.hide_progress_bar()

    def load_multiple_from_database(self, file_ids):
        """Load multiple files from the database with a single progress dialog and success message."""
        if not file_ids:
            return

        loaded_files = []
        failed_files = []
        total_sample_images_loaded = 0

        try:
            # Show progress dialog for the entire batch operation
            self.gui.progress_dialog.show_progress_bar("Loading files from database...")
            self.gui.root.update_idletasks()

            total_files = len(file_ids)
            debug_print(f"DEBUG: Starting batch load of {total_files} files")

            for i, file_id in enumerate(file_ids):
                try:
                    # Update progress
                    progress = int(((i + 1) / total_files) * 100)
                    self.gui.progress_dialog.update_progress_bar(progress)
                    self.gui.root.update_idletasks()

                    debug_print(f"DEBUG: Loading file {i + 1}/{total_files} (ID: {file_id})")

                    # Load this file (suppress individual success messages and progress dialogs)
                    success = self.load_from_database(file_id, show_success_msg=False, batch_operation=True)

                    if success:
                        # Get the file info for the success message - CORRECTED METHOD NAME
                        file_data = self.db_manager.get_file_by_id(file_id)
                        if file_data:
                            display_filename = None
                            if 'meta_data' in file_data and file_data['meta_data']:
                                display_filename = file_data['meta_data'].get('display_filename')
                                if not display_filename:
                                    original_filename = file_data['meta_data'].get('original_filename')
                                    if original_filename:
                                        display_filename = os.path.splitext(original_filename)[0] + '.vap3'
                            if not display_filename:
                                display_filename = file_data['filename']

                            loaded_files.append(display_filename)
                            debug_print(f"DEBUG: Successfully loaded: {display_filename}")

                            # Count sample images for this file
                            if (hasattr(self.gui, 'sample_image_metadata') and
                                display_filename in self.gui.sample_image_metadata):
                                for sheet_metadata in self.gui.sample_image_metadata[display_filename].values():
                                    sample_images = sheet_metadata.get('sample_images', {})
                                    file_sample_count = sum(len(images) for images in sample_images.values())
                                    total_sample_images_loaded += file_sample_count
                                    debug_print(f"DEBUG: Loaded {file_sample_count} sample images for {display_filename}")
                    else:
                        failed_files.append(f"File ID {file_id}")
                        debug_print(f"DEBUG: Failed to load file ID: {file_id}")

                except Exception as e:
                    failed_files.append(f"File ID {file_id}")
                    debug_print(f"DEBUG: Exception loading file ID {file_id}: {e}")

            # Update final progress
            self.gui.progress_dialog.update_progress_bar(100)
            self.gui.root.update_idletasks()

            # Update window title with final count
            total_loaded = len(self.gui.all_filtered_sheets)
            if total_loaded > 1:
                self.gui.root.title(f"DataViewer - {total_loaded} files loaded")
            elif total_loaded == 1:
                self.gui.root.title("DataViewer - 1 file loaded")

            # Show single summary message
            if failed_files:
                if loaded_files:
                    # Partial success
                    success_count = len(loaded_files)
                    failed_count = len(failed_files)
                    message = f"Batch load completed:\n\n"
                    message += f"✓ Successfully loaded: {success_count} files\n"
                    message += f"✗ Failed to load: {failed_count} files\n\n"
                    if total_sample_images_loaded > 0:
                        message += f"📷 Sample images loaded: {total_sample_images_loaded}\n\n"
                    message += f"Total files now loaded: {len(self.gui.all_filtered_sheets)}"
                    messagebox.showwarning("Partial Success", message)
                else:
                    # Complete failure
                    messagebox.showerror("Error", f"Failed to load all {len(failed_files)} selected files.")
            else:
                # Complete success
                if len(loaded_files) == 1:
                    message = f"Successfully loaded 1 file:\n{loaded_files[0]}"
                    if total_sample_images_loaded > 0:
                        message += f"\n\n📷 Sample images loaded: {total_sample_images_loaded}"
                else:
                    message = f"Successfully loaded {len(loaded_files)} files:\n\n"
                    # Show first few filenames, then "and X more" if too many
                    if len(loaded_files) <= 5:
                        message += "\n".join([f"• {name}" for name in loaded_files])
                    else:
                        message += "\n".join([f"• {name}" for name in loaded_files[:3]])
                        message += f"\n• ... and {len(loaded_files) - 3} more files"

                    if total_sample_images_loaded > 0:
                        message += f"\n\n📷 Total sample images loaded: {total_sample_images_loaded}"
                    message += f"\n\nTotal files now loaded: {len(self.gui.all_filtered_sheets)}"

                show_success_message("Success", message, self.gui.root)

        except Exception as e:
            messagebox.showerror("Error", f"Error during batch loading: {e}")
            debug_print(f"ERROR: Batch loading error: {e}")
            traceback.print_exc()
        finally:
            # Hide progress dialog
            self.gui.progress_dialog.hide_progress_bar()

    def load_from_database_by_id(self, file_id):
        """Load a file from database by its ID."""
        debug_print(f"DEBUG: Loading file from database with ID: {file_id}")

        try:
            # Get file data from database
            file_data = self.db_manager.get_file_by_id(file_id)

            if file_data and 'filepath' in file_data:
                # Load the file
                self.file_manager.load_excel_file(file_data['filepath'], skip_database_storage=True)
                return True
            else:
                debug_print(f"DEBUG: Could not find file data for ID: {file_id}")
                return False

        except Exception as e:
            debug_print(f"DEBUG: Error loading file from database ID {file_id}: {e}")
            return False

    def show_database_browser(self, comparison_mode=False):
        """Show a dialog to browse files stored in the database with multi-select support."""
        # Create a dialog to browse the database
        dialog = Toplevel(self.gui.root)

        # Update title based on mode
        if comparison_mode:
            dialog.title("Select Files for Comparison")
        else:
            dialog.title("Database Browser")

        dialog.transient(self.gui.root)
        self.file_manager.center_window(dialog, 900, 700)

        # Create frames for UI elements
        top_frame = Frame(dialog)
        top_frame.pack(fill="x", padx=10, pady=10)

        list_frame = Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        bottom_frame = Frame(dialog)
        bottom_frame.pack(fill="x", padx=10, pady=10)

        # Store grouped files data
        file_groups = {}  # Key: base filename, Value: list of file records
        listbox_to_file_mapping = {}  # Key: listbox index, Value: file record

        def extract_base_filename(filename):
            """Extract base filename without timestamp and extension."""

            # Remove common timestamp patterns and extensions
            base = re.sub(r'\s+\d{4}-\d{2}-\d{2}.*', '', filename)  # Remove date patterns
            base = re.sub(r'\s+copy.*', '', base, flags=re.IGNORECASE)  # Remove "copy" variants
            base = re.sub(r'\.[^.]*$', '', base)  # Remove extension
            return base.strip()


        def sort_file_groups(groups, sort_method):
            """Sort file groups based on the selected method."""
            debug_print(f"DEBUG: Sorting files by {sort_method}")

            if sort_method == "newest_first":
                # Sort by most recent update (within each group, then by group)
                sorted_groups = {}
                for base_name, versions in groups.items():
                    # Each group is already sorted by creation date (newest first)
                    sorted_groups[base_name] = versions
                # Sort group keys by the most recent file in each group
                sorted_keys = sorted(sorted_groups.keys(),
                                   key=lambda name: sorted_groups[name][0]['created_at'], reverse=True)
            elif sort_method == "oldest_first":
                # Sort by oldest update
                sorted_groups = {}
                for base_name, versions in groups.items():
                    sorted_groups[base_name] = versions
                # Sort group keys by the most recent file in each group (oldest first)
                sorted_keys = sorted(sorted_groups.keys(),
                                   key=lambda name: sorted_groups[name][0]['created_at'], reverse=False)
            elif sort_method == "alphabetical_asc":
                # Sort alphabetically A to Z
                sorted_groups = groups
                sorted_keys = sorted(groups.keys())
            elif sort_method == "alphabetical_desc":
                # Sort alphabetically Z to A
                sorted_groups = groups
                sorted_keys = sorted(groups.keys(), reverse=True)
            else:
                # Default to newest first
                sorted_groups = groups
                sorted_keys = sorted(groups.keys(),
                                   key=lambda name: groups[name][0]['created_at'], reverse=True)

            return sorted_groups, sorted_keys

        def populate_listbox(filter_keyword = None):
            """Populate the listbox with files according to current sort method."""
            try:
                debug_print("DEBUG: Populating listbox with sorted files")

                # Clear current listbox
                file_listbox.delete(0, tk.END)
                listbox_to_file_mapping.clear()

                # Get current sort method
                current_sort = sort_var.get()
                if current_sort == "Newest First":
                    sort_method = "newest_first"
                elif current_sort == "Oldest First":
                    sort_method = "oldest_first"
                elif current_sort == "A to Z":
                    sort_method = "alphabetical_asc"
                elif current_sort == "Z to A":
                    sort_method = "alphabetical_desc"
                else:
                    sort_method = "newest_first"

                # Apply filter if provided
                filtered_groups = file_groups.copy()
                if filter_keyword and filter_keyword.strip():
                    keyword_lower = filter_keyword.strip().lower()
                    filtered_groups = {
                        base_name: versions
                        for base_name, versions in file_groups.items()
                        if keyword_lower in base_name.lower()
                    }
                    debug_print(f"DEBUG: Applied filter '{filter_keyword}' - {len(filtered_groups)} groups remaining")


                # Sort the file groups
                sorted_groups, sorted_keys = sort_file_groups(filtered_groups, sort_method)

                # Populate listbox with sorted groups
                listbox_index = 0
                for base_name in sorted_keys:
                    latest_file = sorted_groups[base_name][0]  # First item is newest
                    count = len(sorted_groups[base_name])

                    if count > 1:
                        display_text = f"{base_name} ({count} versions, latest: {latest_file['created_at'].strftime('%Y-%m-%d %H:%M')})"
                    else:
                        display_text = f"{base_name} ({latest_file['created_at'].strftime('%Y-%m-%d %H:%M')})"

                    file_listbox.insert(tk.END, display_text)
                    listbox_to_file_mapping[listbox_index] = latest_file
                    debug_print(f"DEBUG: Mapped listbox index {listbox_index} to file ID {latest_file['id']} for base name '{base_name}'")
                    listbox_index += 1

                update_selection_info()
                debug_print(f"DEBUG: Listbox populated with {listbox_index} items")

            except Exception as e:
                debug_print(f"DEBUG: Error populating listbox: {e}")
                messagebox.showerror("Error", f"Error displaying files: {e}")

        try:
            # Get files from database
            files = self.db_manager.list_files()
            debug_print(f"DEBUG: Found {len(files)} files in database")

            # Group files by base filename
            for file_record in files:
                base_name = extract_base_filename(file_record['filename'])
                if base_name not in file_groups:
                    file_groups[base_name] = []
                file_groups[base_name].append(file_record)

            # Sort each group by creation date (newest first)
            for group in file_groups.values():
                group.sort(key=lambda x: x['created_at'], reverse=True)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load files from database: {e}")
            dialog.destroy()
            return

        # Create header
        if comparison_mode:
            header_text = "Select files for comparison analysis (minimum 2 required)"
        else:
            header_text = "Database Files (grouped by base name, showing latest version)"

        Label(top_frame, text=header_text, font=("Arial", 12, "bold")).pack()

        # Add sorting controls
        sort_frame = Frame(top_frame)
        sort_frame.pack(fill="x", pady=5)

        Label(sort_frame, text="Sort by:", font=("Arial", 10)).pack(side="left", padx=(0, 5))

        # Sort options variable
        sort_var = tk.StringVar(value="newest_first")

        sort_options = [
            ("Newest First", "newest_first"),
            ("Oldest First", "oldest_first"),
            ("A to Z", "alphabetical_asc"),
            ("Z to A", "alphabetical_desc")
        ]

        sort_dropdown = ttk.Combobox(sort_frame, textvariable=sort_var, values=[option[1] for option in sort_options],
                                   state="readonly", width=15)
        sort_dropdown.pack(side="left", padx=5)

        # Set display values
        sort_dropdown['values'] = [option[0] for option in sort_options]
        sort_dropdown.set("Newest First")

        # Right side - Filter controls
        filter_var = tk.StringVar()
        Label(sort_frame, text="Filter:", font=("Arial", 10)).pack(side="right", padx=(5, 0))
        filter_entry = tk.Entry(sort_frame, textvariable=filter_var, width=20, font=("Arial", 10))
        filter_entry.pack(side="right", padx=(0, 5))

        # Create listbox with scrollbar
        listbox_frame = Frame(list_frame)

        # Create listbox with scrollbar
        listbox_frame = Frame(list_frame)
        listbox_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side="right", fill="y")

        # Use extended selection mode for multi-select
        file_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set,
                                 selectmode=tk.EXTENDED, font=FONT)
        file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=file_listbox.yview)

        # Selection info label (create before functions that use it)
        selection_info = Label(list_frame, text="", font=("Arial", 10))
        selection_info.pack(pady=5)

        # DEFINE update_selection_info FIRST since populate_listbox needs it
        def update_selection_info():
            """Update the selection information label."""
            debug_print("DEBUG: Updating selection info")
            try:
                selected_count = len(file_listbox.curselection())
                if comparison_mode:
                    if selected_count < 2:
                        selection_info.config(text=f"Selected: {selected_count} files (need at least 2 for comparison)",
                                            fg="red")
                    else:
                        selection_info.config(text=f"Selected: {selected_count} files (ready for comparison)",
                                            fg="green")
                else:
                    selection_info.config(text=f"Selected: {selected_count} files", fg="black")
                debug_print(f"DEBUG: Selection info updated - {selected_count} files selected")
            except Exception as e:
                debug_print(f"DEBUG: Error updating selection info: {e}")

        # Initial population with default sort
        populate_listbox()

        # Bind sort dropdown change event
        def on_sort_change(event=None):
            """Handle sort method change."""
            debug_print(f"DEBUG: Sort method changed to {sort_var.get()}")
            current_filter = filter_var.get()
            populate_listbox(current_filter if current_filter else None)

        sort_dropdown.bind('<<ComboboxSelected>>', on_sort_change)

        # Bind filter entry change event
        def on_filter_change(event=None):
            """Handle filter keyword change."""
            filter_keyword = filter_var.get()
            debug_print(f"DEBUG: Filter keyword changed to '{filter_keyword}'")
            populate_listbox(filter_keyword if filter_keyword else None)

        filter_entry.bind('<KeyRelease>', on_filter_change)
        filter_entry.bind('<FocusOut>', on_filter_change)

        def show_version_history_dialog(base_name, versions, main_listbox_index):
            """Show dialog with version history and allow version selection."""
            history_dialog = Toplevel(dialog)
            history_dialog.title(f"Version History - {base_name}")

            # Hide the window immediately to prevent stutter during positioning
            history_dialog.withdraw()

            history_dialog.transient(dialog)
            history_dialog.grab_set()

            # Create frames
            top_frame_hist = Frame(history_dialog)
            top_frame_hist.pack(fill="x", padx=10, pady=10)

            list_frame_hist = Frame(history_dialog)
            list_frame_hist.pack(fill="both", expand=True, padx=10, pady=10)

            bottom_frame_hist = Frame(history_dialog)
            bottom_frame_hist.pack(fill="x", padx=10, pady=10)

            # Header
            Label(top_frame_hist, text=f"Version History: {base_name}",
                  font=("Arial", 14, "bold")).pack()
            Label(top_frame_hist, text="Double-click a version to select it for loading. Use Ctrl+click for multiple selection.",
                  font=("Arial", 10)).pack(pady=5)

            # Version listbox with scrollbar
            hist_listbox_frame = Frame(list_frame_hist)
            hist_listbox_frame.pack(fill="both", expand=True)

            hist_scrollbar = tk.Scrollbar(hist_listbox_frame)
            hist_scrollbar.pack(side="right", fill="y")

            version_listbox = tk.Listbox(hist_listbox_frame, yscrollcommand=hist_scrollbar.set, selectmode=tk.EXTENDED, font=FONT)
            version_listbox.pack(side="left", fill="both", expand=True)
            hist_scrollbar.config(command=version_listbox.yview)

            # Populate with versions (newest first)
            for i, version in enumerate(versions):
                created_str = version['created_at'].strftime('%Y-%m-%d %H:%M:%S')
                is_current = i == 0
                status = " [CURRENT]" if is_current else f" [v{len(versions)-i}]"
                display_text = f"{version['filename']}{status} - {created_str}"
                version_listbox.insert(tk.END, display_text)

            def on_version_double_click(event):
                """Handle double-click on version to replace in main browser."""
                selection = version_listbox.curselection()
                if not selection:
                    return

                selected_version_idx = selection[0]
                selected_version = versions[selected_version_idx]

                # Update the main listbox display
                created_str = selected_version['created_at'].strftime('%Y-%m-%d %H:%M')
                if selected_version_idx == 0:
                    # Current version selected - restore normal display
                    count = len(versions)
                    if count > 1:
                        display_text = f"{base_name} ({count} versions, latest: {created_str})"
                    else:
                        display_text = f"{base_name} ({created_str})"
                    # Remove from restoration tracking if it exists
                    if hasattr(dialog, 'original_mappings') and main_listbox_index in dialog.original_mappings:
                        del dialog.original_mappings[main_listbox_index]
                else:
                    # Older version selected - show with indicator
                    version_num = len(versions) - selected_version_idx
                    display_text = f"{base_name} [v{version_num}: {created_str}] ★"

                    # Store original mapping for restoration if not already stored
                    if not hasattr(dialog, 'original_mappings'):
                        dialog.original_mappings = {}
                    if main_listbox_index not in dialog.original_mappings:
                        dialog.original_mappings[main_listbox_index] = versions[0]  # Store original (newest)

                # Update main listbox
                file_listbox.delete(main_listbox_index)
                file_listbox.insert(main_listbox_index, display_text)

                # Update mapping to point to selected version
                listbox_to_file_mapping[main_listbox_index] = selected_version

                # Keep the item selected in main listbox
                file_listbox.selection_clear(0, tk.END)
                file_listbox.selection_set(main_listbox_index)

                debug_print(f"DEBUG: Temporarily replaced {base_name} with version {selected_version_idx + 1}")

                history_dialog.destroy()

            # Bind double-click to version listbox
            version_listbox.bind("<Double-Button-1>", on_version_double_click)

            # Close button
            def on_close():
                history_dialog.destroy()

            def on_delete_versions():
                """Delete selected versions from the version history."""
                selected_items = version_listbox.curselection()
                if not selected_items:
                    messagebox.showwarning("Warning", "Please select at least one version to delete.")
                    return

                version_ids = []
                version_names = []
                for idx in selected_items:
                    if idx < len(versions):
                        version_record = versions[idx]
                        version_ids.append(version_record['id'])
                        version_names.append(version_record['filename'])

                if not version_ids:
                    messagebox.showwarning("Warning", "No valid versions selected for deletion.")
                    return

                # Confirm deletion
                if len(version_ids) == 1:
                    confirm_msg = f"Are you sure you want to delete this version:\n\n{version_names[0]}\n\nThis action cannot be undone."
                else:
                    confirm_msg = f"Are you sure you want to delete {len(version_ids)} versions?\n\nThis action cannot be undone."

                if not messagebox.askyesno("Confirm Deletion", confirm_msg):
                    return

                try:
                    debug_print(f"DEBUG: Starting deletion of {len(version_ids)} versions")
                    success_count, error_count = self.db_manager.delete_multiple_files(version_ids)

                    if success_count > 0:
                        if error_count == 0:
                            messagebox.showinfo("Success", f"Successfully deleted {success_count} version(s).")
                        else:
                            messagebox.showwarning("Partial Success", f"Deleted {success_count} version(s), but {error_count} failed.")

                        # Check if we deleted the most recent version
                        most_recent_deleted = 0 in selected_items

                        if most_recent_deleted:
                            # Find the new most recent version and update main listbox
                            remaining_versions = [v for i, v in enumerate(versions) if i not in selected_items]
                            if remaining_versions:
                                # Sort remaining versions by date (newest first)
                                remaining_versions.sort(key=lambda x: x['created_at'], reverse=True)
                                new_most_recent = remaining_versions[0]

                                # Update the main listbox display
                                base_name = extract_base_filename(new_most_recent['filename'])
                                count = len(remaining_versions)
                                created_str = new_most_recent['created_at'].strftime('%Y-%m-%d %H:%M')

                                if count > 1:
                                    display_text = f"{base_name} ({count} versions, latest: {created_str})"
                                else:
                                    display_text = f"{base_name} ({created_str})"

                                # Update the item in main listbox
                                file_listbox.delete(main_listbox_index)
                                file_listbox.insert(main_listbox_index, display_text)
                                listbox_to_file_mapping[main_listbox_index] = new_most_recent

                                debug_print(f"DEBUG: Updated main listbox after deleting most recent version")
                            else:
                                # All versions deleted, remove from main listbox
                                file_listbox.delete(main_listbox_index)
                                if main_listbox_index in listbox_to_file_mapping:
                                    del listbox_to_file_mapping[main_listbox_index]
                                debug_print(f"DEBUG: Removed entry from main listbox - all versions deleted")

                        # Close the version history dialog and refresh
                        history_dialog.destroy()
                        update_selection_info()
                    else:
                        messagebox.showerror("Error", "Failed to delete any versions.")

                except Exception as e:
                    debug_print(f"DEBUG: Error during version deletion: {e}")
                    messagebox.showerror("Error", f"Error during deletion: {e}")

            # Add delete button on the left, close on the right
            Button(bottom_frame_hist, text="Delete Selected", command=on_delete_versions,
                   bg="#f44336", fg="white", font=FONT).pack(side="left", padx=5)

            Button(bottom_frame_hist, text="Close", command=on_close,
                   bg="#f44336", fg="black", font=FONT).pack(side="right", padx=5)

            # Center the window and then show it (prevents stutter)
            self.file_manager.center_window(history_dialog, 600, 400)
            history_dialog.deiconify()

        def on_delete():
            selected_items = file_listbox.curselection()
            if not selected_items:
                messagebox.showwarning("Warning", "Please select at least one file to delete.")
                return

            file_ids = []
            filenames = []
            for idx in selected_items:
                if idx in listbox_to_file_mapping:
                    file_record = listbox_to_file_mapping[idx]
                    file_ids.append(file_record['id'])
                    filenames.append(file_record['filename'])

            if not file_ids:
                messagebox.showwarning("Warning", "No valid files selected for deletion.")
                return

            # Confirm deletion
            if len(file_ids) == 1:
                confirm_msg = f"Are you sure you want to delete the file:\n\n{filenames[0]}\n\nThis action cannot be undone."
            else:
                confirm_msg = f"Are you sure you want to delete {len(file_ids)} files?\n\nThis action cannot be undone."

            if not messagebox.askyesno("Confirm Deletion", confirm_msg):
                return

            try:
                debug_print(f"DEBUG: Starting deletion of {len(file_ids)} files")
                success_count, error_count = self.db_manager.delete_multiple_files(file_ids)

                if success_count > 0:
                    if error_count == 0:
                        messagebox.showinfo("Success", f"Successfully deleted {success_count} file(s).")
                    else:
                        messagebox.showwarning("Partial Success", f"Deleted {success_count} file(s), but {error_count} failed.")

                    # Refresh the file list after deletion
                    refresh_file_list()
                else:
                    messagebox.showerror("Error", "Failed to delete any files.")

            except Exception as e:
                debug_print(f"DEBUG: Error during file deletion: {e}")
                messagebox.showerror("Error", f"Error during deletion: {e}")

        def refresh_file_list():
            """Refresh the file list display after deletion."""
            try:
                debug_print("DEBUG: Refreshing file list after deletion")

                # Clear current listbox
                file_listbox.delete(0, tk.END)
                listbox_to_file_mapping.clear()
                file_groups.clear()

                # Reload files from database
                files = self.db_manager.list_files()
                debug_print(f"DEBUG: Reloaded {len(files)} files from database")

                # Regroup files
                for file_record in files:
                    base_name = extract_base_filename(file_record['filename'])
                    if base_name not in file_groups:
                        file_groups[base_name] = []
                    file_groups[base_name].append(file_record)

                # Sort each group by creation date (newest first)
                for base_name in file_groups:
                    file_groups[base_name].sort(key=lambda x: x['created_at'], reverse=True)

                # Repopulate listbox
                populate_listbox()

                update_selection_info()
                debug_print("DEBUG: File list refresh completed")

            except Exception as e:
                debug_print(f"DEBUG: Error refreshing file list: {e}")
                messagebox.showerror("Error", f"Error refreshing file list: {e}")

        def on_double_click(event):
            """Handle double-click to show version history for selected files."""
            selected_items = list(file_listbox.curselection())
            if not selected_items:
                return

            # Process each selected file one by one
            def process_next_file(file_indices):
                if not file_indices:
                    return  # All done

                idx = file_indices[0]
                remaining = file_indices[1:]

                if idx not in listbox_to_file_mapping:
                    # Skip this one and process next
                    if remaining:
                        dialog.after(100, lambda: process_next_file(remaining))
                    return

                selected_file = listbox_to_file_mapping[idx]
                base_name = extract_base_filename(selected_file['filename'])

                # Check if this file has multiple versions
                if base_name not in file_groups or len(file_groups[base_name]) <= 1:
                    # Single version - skip to next file
                    if remaining:
                        dialog.after(100, lambda: process_next_file(remaining))
                    return

                # Show version history dialog
                # We need to ensure the dialog is modal and wait for it to close before proceeding
                def show_dialog_and_continue():
                    show_version_history_dialog(base_name, file_groups[base_name], idx)
                    # After dialog closes, process next file
                    if remaining:
                        dialog.after(100, lambda: process_next_file(remaining))

                show_dialog_and_continue()

            # Start processing the first file
            process_next_file(selected_items)

        def on_selection_change(event):
            """Restore original versions for deselected items."""
            if hasattr(dialog, 'original_mappings'):
                current_selection = set(file_listbox.curselection())

                # Check each item that has been temporarily replaced
                items_to_restore = []
                for idx, original_version in dialog.original_mappings.items():
                    if idx not in current_selection:
                        items_to_restore.append((idx, original_version))

                # Restore items that were deselected
                for idx, original_version in items_to_restore:
                    base_name = extract_base_filename(original_version['filename'])
                    count = len(file_groups[base_name])
                    created_str = original_version['created_at'].strftime('%Y-%m-%d %H:%M')

                    if count > 1:
                        display_text = f"{base_name} ({count} versions, latest: {created_str})"
                    else:
                        display_text = f"{base_name} ({created_str})"

                    # Restore display and mapping
                    file_listbox.delete(idx)
                    file_listbox.insert(idx, display_text)
                    listbox_to_file_mapping[idx] = original_version

                    debug_print(f"DEBUG: Restored {base_name} to latest version (deselected)")

                # Remove restored items from tracking
                for idx, _ in items_to_restore:
                    del dialog.original_mappings[idx]

            update_selection_info()

        def select_all():
            file_listbox.select_set(0, tk.END)
            update_selection_info()

        def select_none():
            file_listbox.selection_clear(0, tk.END)
            update_selection_info()

        # Define load functions based on mode
        if comparison_mode:
            def on_load():
                selected_items = file_listbox.curselection()
                if len(selected_items) < 2:
                    messagebox.showwarning("Warning", "Please select at least 2 files for comparison.")
                    return

                file_ids = []
                for idx in selected_items:
                    if idx in listbox_to_file_mapping:
                        file_record = listbox_to_file_mapping[idx]
                        file_ids.append(file_record['id'])

                dialog.destroy()
                if len(file_ids) >= 2:
                    original_all_filtered_sheets = self.gui.all_filtered_sheets.copy()
                    self.load_multiple_from_database(file_ids)
                    if len(self.gui.all_filtered_sheets) >= 2:
                        from sample_comparison import SampleComparisonWindow
                        comparison_window = SampleComparisonWindow(self.gui, self.gui.all_filtered_sheets)
                        comparison_window.show()
                    else:
                        messagebox.showwarning("Warning", "Failed to load enough files for comparison.")
                        self.gui.all_filtered_sheets = original_all_filtered_sheets
        else:
            def on_load():
                selected_items = file_listbox.curselection()
                if not selected_items:
                    messagebox.showwarning("Warning", "Please select at least one file to load.")
                    return

                file_ids = []
                for idx in selected_items:
                    if idx in listbox_to_file_mapping:
                        file_record = listbox_to_file_mapping[idx]
                        file_ids.append(file_record['id'])

                dialog.destroy()
                if len(file_ids) == 1:
                    self.load_from_database(file_ids[0])
                elif len(file_ids) > 1:
                    self.load_multiple_from_database(file_ids)

        # Bind events
        file_listbox.bind("<Double-Button-1>", on_double_click)
        file_listbox.bind("<<ListboxSelect>>", on_selection_change)

        # Create buttons
        button_frame = Frame(bottom_frame)
        button_frame.pack(fill="x", pady=10)

        # Add select all/none buttons
        select_frame = Frame(button_frame)
        select_frame.pack(fill="x", pady=(0, 10))

        Button(select_frame, text="Select All", command=select_all, font=FONT).pack(side="left", padx=5)
        Button(select_frame, text="Select None", command=select_none, font=FONT).pack(side="left", padx=5)

        # Add info label
        info_label = Label(select_frame, text="Double-click files to view version history",
                          font=("Arial", 9), fg="gray")
        info_label.pack(side="right", padx=5)

        # Create main action buttons
        Button(button_frame, text="Load Selected", command=on_load,
               bg="#4CAF50", fg="black", font=FONT).pack(side="left", padx=5)

        # Add delete button in the middle
        Button(button_frame, text="Delete Selected", command=on_delete,
               bg="#f44336", fg="white", font=FONT).pack(side="left", padx=20)

        Button(button_frame, text="Close", command=dialog.destroy,
               bg="#666666", fg="black", font=FONT).pack(side="right", padx=5)

        # Initialize selection info
        update_selection_info()

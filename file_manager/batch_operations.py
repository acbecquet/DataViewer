"""
Batch Operations Module for DataViewer Application

This module handles batch loading of Excel files from folders,
including scanning, filtering, and processing multiple files.
"""

# Standard library imports
import os
import copy

# Third party imports
import tkinter as tk
from tkinter import filedialog, messagebox

# Local imports
from utils import debug_print, show_success_message


class BatchOperations:
    """Handles batch loading operations for multiple files."""
    
    def __init__(self, file_manager):
        """Initialize with reference to parent FileManager."""
        self.file_manager = file_manager
        self.gui = file_manager.gui
        self.root = file_manager.root
        
    def batch_load_folder(self):
        """
        Batch load Excel files from a selected folder and all subfolders.
        Only loads files whose names contain test names from the processing dictionary.
        Shows accurate progress during the operation.
        """
        debug_print("DEBUG: Starting batch folder loading process")

        # Get folder selection from user
        folder_path = filedialog.askdirectory(
            title="Select Folder for Batch Loading (will scan all subfolders)"
        )

        if not folder_path:
            debug_print("DEBUG: No folder selected, canceling batch load")
            return

        try:
            # Get test names from processing dictionary for filename matching
            test_names = self._get_test_names_for_matching()
            debug_print(f"DEBUG: Got {len(test_names)} test names for matching: {test_names[:5]}...")

            # Scan folder recursively for Excel files
            debug_print(f"DEBUG: Scanning folder: {folder_path}")
            excel_files = self._scan_folder_for_excel_files(folder_path)
            debug_print(f"DEBUG: Found {len(excel_files)} Excel files total")

            # Filter files based on test name matching
            matching_files = self._filter_files_by_test_names(excel_files, test_names)
            debug_print(f"DEBUG: Found {len(matching_files)} files matching test names")

            if not matching_files:
                messagebox.showinfo("No Matching Files",
                                   f"No Excel files found in '{folder_path}' with test names in their filename.\n\n"
                                   f"Looking for files containing any of these test names: {', '.join(test_names[:10])}{'...' if len(test_names) > 10 else ''}")
                return

            # Confirm with user before proceeding
            proceed = messagebox.askyesno("Batch Load Confirmation",
                                        f"Found {len(matching_files)} Excel files to load.\n\n"
                                        f"This may take several minutes. Proceed with batch loading?")
            if not proceed:
                debug_print("DEBUG: User cancelled batch loading")
                return

            # Perform batch loading with progress tracking
            self._perform_batch_loading(matching_files)

        except Exception as e:
            debug_print(f"ERROR: Batch folder loading failed: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Batch Loading Error", f"Failed to complete batch loading: {e}")

    def _get_test_names_for_matching(self):
        """Get test names from processing dictionary for filename matching."""
        debug_print("DEBUG: Extracting test names from processing functions")

        try:
            # Import processing module to get test names
            import processing

            # Get the processing functions dictionary
            processing_functions = processing.get_processing_functions()
            debug_print(f"DEBUG: Found {len(processing_functions)} processing functions")

            # Extract test names, excluding generic ones
            test_names = []
            exclude_terms = ['default', 'sheet1', 'legacy']

            for sheet_name in processing_functions.keys():
                if sheet_name.lower() not in exclude_terms:
                    test_names.append(sheet_name.lower())
                    # Also add variations for better matching
                    test_names.append(sheet_name.replace(" ", "").lower())
                    test_names.append(sheet_name.replace(" ", "_").lower())
                    test_names.append(sheet_name.replace(" ", "-").lower())

            # Remove duplicates and sort
            test_names = list(set(test_names))
            test_names.sort()

            debug_print(f"DEBUG: Generated {len(test_names)} test name variations for matching")
            return test_names

        except Exception as e:
            debug_print(f"ERROR: Failed to get test names: {e}")
            # Fallback list of common test names
            return ['test plan', 'lifetime test', 'device life test', 'quick screening',
                    'aerosol temperature', 'user test', 'horizontal puffing', 'extended test']

    def _scan_folder_for_excel_files(self, folder_path):
        """Recursively scan folder for Excel files."""
        debug_print(f"DEBUG: Starting recursive scan of folder: {folder_path}")

        excel_files = []
        excel_extensions = ('.xlsx', '.xls')

        try:
            for root, dirs, files in os.walk(folder_path):
                debug_print(f"DEBUG: Scanning directory: {root}")

                for file in files:
                    if file.lower().endswith(excel_extensions):
                        full_path = os.path.join(root, file)

                        # Skip temporary Excel files
                        if file.startswith('~$'):
                            debug_print(f"DEBUG: Skipping temporary file: {file}")
                            continue

                        # Check if file is accessible
                        try:
                            if os.access(full_path, os.R_OK):
                                excel_files.append(full_path)
                                debug_print(f"DEBUG: Added Excel file: {file}")
                            else:
                                debug_print(f"DEBUG: Skipping inaccessible file: {file}")
                        except Exception as e:
                            debug_print(f"DEBUG: Error checking file access for {file}: {e}")
                            continue

            debug_print(f"DEBUG: Completed folder scan, found {len(excel_files)} Excel files")
            return excel_files

        except Exception as e:
            debug_print(f"ERROR: Failed to scan folder: {e}")
            raise

    def _filter_files_by_test_names(self, excel_files, test_names):
        """Filter Excel files by checking if filename contains any test names."""
        debug_print(f"DEBUG: Filtering {len(excel_files)} files against {len(test_names)} test names")

        matching_files = []

        for file_path in excel_files:
            filename = os.path.basename(file_path).lower()
            debug_print(f"DEBUG: Checking file: {filename}")

            # Check if any test name is contained in the filename
            for test_name in test_names:
                if test_name in filename:
                    matching_files.append(file_path)
                    debug_print(f"DEBUG: MATCH - '{filename}' contains '{test_name}'")
                    break

        debug_print(f"DEBUG: Filtered to {len(matching_files)} matching files")
        return matching_files

    def _perform_batch_loading(self, file_paths):
        """Perform the actual batch loading with accurate progress tracking."""
        debug_print(f"DEBUG: Starting batch loading of {len(file_paths)} files")

        # Show progress dialog
        self.gui.progress_dialog.show_progress_bar("Batch loading files...")
        self.gui.root.update_idletasks()

        loaded_files = []
        failed_files = []
        skipped_files = []
        total_files = len(file_paths)

        try:
            for index, file_path in enumerate(file_paths, 1):
                try:
                    # Update progress with current file info
                    filename = os.path.basename(file_path)
                    progress = int((index / total_files) * 100)
                    progress_text = f"Loading {index}/{total_files}: {filename[:40]}{'...' if len(filename) > 40 else ''}"

                    debug_print(f"DEBUG: {progress_text}")

                    # SIMPLIFIED: Only update progress bar, skip label updates
                    try:
                        self.gui.progress_dialog.update_progress_bar(progress)
                    except Exception as e:
                        debug_print(f"DEBUG: Progress update failed: {e}")

                    self.gui.root.update_idletasks()

                    # Load the file using existing infrastructure
                    debug_print(f"DEBUG: Loading file: {file_path}")

                    # Load the file - this stores in database by default
                    self.file_manager.load_excel_file(file_path, force_reload=True)

                    # Store the loaded file data
                    if hasattr(self.gui, 'filtered_sheets') and self.gui.filtered_sheets:
                        file_data = {
                            "file_name": filename,
                            "file_path": file_path,
                            "display_filename": filename,
                            "filtered_sheets": copy.deepcopy(self.gui.filtered_sheets),
                            "source": "batch_folder_load"
                        }

                        # Add to all_filtered_sheets if not already present
                        if not any(f["file_path"] == file_path for f in self.gui.all_filtered_sheets):
                            self.gui.all_filtered_sheets.append(file_data)
                            loaded_files.append(filename)
                            debug_print(f"DEBUG: Successfully loaded and stored: {filename}")
                        else:
                            debug_print(f"DEBUG: File already loaded, skipping: {filename}")
                    else:
                        debug_print(f"ERROR: No filtered_sheets after loading {filename}")
                        failed_files.append(filename)

                except ValueError as e:
                    error_msg = str(e).lower()
                    if any(phrase in error_msg for phrase in [
                        "no valid legacy sample data found",
                        "no samples with meaningful data found after filtering",
                        "empty",
                        "no data found"
                    ]):
                        debug_print(f"DEBUG: Gracefully skipping empty file: {filename} - {e}")
                        skipped_files.append(filename)
                        continue
                    else:
                        debug_print(f"ERROR: Failed to load file {file_path}: {e}")
                        failed_files.append(os.path.basename(file_path))
                        continue

                except Exception as e:
                    error_msg = str(e).lower()
                    if any(phrase in error_msg for phrase in [
                        "permission denied",
                        "file not found",
                        "corrupted",
                        "cannot read",
                        "invalid file format"
                    ]):
                        debug_print(f"DEBUG: Gracefully skipping problematic file: {filename} - {e}")
                        skipped_files.append(filename)
                        continue
                    else:
                        debug_print(f"ERROR: Failed to load file {file_path}: {e}")
                        import traceback
                        traceback.print_exc()
                        failed_files.append(os.path.basename(file_path))
                        continue

            # Update UI after successful batch loading
            if loaded_files:
                debug_print("DEBUG: Updating UI after batch loading")
                self.file_manager.update_file_dropdown()

                # Set the last loaded file as active
                if self.gui.all_filtered_sheets:
                    last_file = self.gui.all_filtered_sheets[-1]
                    self.file_manager.set_active_file(last_file["file_name"])
                    self.file_manager.update_ui_for_current_file()

            # Show completion summary
            self._show_batch_loading_summary(loaded_files, failed_files, skipped_files, total_files)

        except Exception as e:
            debug_print(f"ERROR: Batch loading process failed: {e}")
            import traceback
            traceback.print_exc()
            raise
        finally:
            # Always hide progress dialog
            try:
                self.gui.progress_dialog.hide_progress_bar()
            except:
                pass
            self.gui.root.update_idletasks()

    def _show_batch_loading_summary(self, loaded_files, failed_files, skipped_files, total_files):
        """Show summary of batch loading results including skipped files."""
        debug_print(f"DEBUG: Showing batch loading summary - {len(loaded_files)} loaded, {len(failed_files)} failed, {len(skipped_files)} skipped")

        success_count = len(loaded_files)
        failure_count = len(failed_files)
        skipped_count = len(skipped_files)

        if success_count == 0 and skipped_count == 0:
            messagebox.showerror("Batch Loading Failed",
                               f"Failed to load any of the {total_files} files.\n\n"
                               f"Check the debug output for specific error details.")
        elif failure_count == 0 and skipped_count == 0:
            # Complete success
            message = f"Batch Loading Complete!\n\n"
            message += f"Successfully loaded {success_count} files:\n\n"

            # Show first few filenames
            if len(loaded_files) <= 8:
                message += "\n".join([f"• {name}" for name in loaded_files])
            else:
                message += "\n".join([f"• {name}" for name in loaded_files[:5]])
                message += f"\n• ... and {len(loaded_files) - 5} more files"

            message += f"\n\nTotal files now loaded: {len(self.gui.all_filtered_sheets)}"
            message += f"\n\n📁 All files have been stored in the database for future access."
            show_success_message("Batch Loading Success", message, self.gui.root)
        else:
            # Mixed results
            message = f"Batch Loading Completed\n\n"
            message += f"Results Summary:\n"
            message += f"✅ Successfully loaded: {success_count} files\n"
            if skipped_count > 0:
                message += f"⏭️ Skipped (empty): {skipped_count} files\n"
            if failure_count > 0:
                message += f"❌ Failed to load: {failure_count} files\n"
            message += f"\nTotal processed: {total_files} files\n"

            if success_count > 0:
                message += f"\n📁 Successfully loaded files have been stored in the database.\n"
                message += "\nLoaded files:\n"
                if len(loaded_files) <= 5:
                    message += "\n".join([f"• {name}" for name in loaded_files])
                else:
                    message += "\n".join([f"• {name}" for name in loaded_files[:3]])
                    message += f"\n• ... and {len(loaded_files) - 3} more"

            if skipped_count > 0:
                message += "\n\nSkipped files (empty/no data):\n"
                if len(skipped_files) <= 5:
                    message += "\n".join([f"• {name}" for name in skipped_files])
                else:
                    message += "\n".join([f"• {name}" for name in skipped_files[:3]])
                    message += f"\n• ... and {len(skipped_files) - 3} more"

            if failure_count > 0:
                message += "\n\nFailed files:\n"
                if len(failed_files) <= 5:
                    message += "\n".join([f"• {name}" for name in failed_files])
                else:
                    message += "\n".join([f"• {name}" for name in failed_files[:3]])
                    message += f"\n• ... and {len(failed_files) - 3} more"

            if failure_count > 0:
                messagebox.showwarning("Batch Loading Complete", message)
            else:
                show_success_message("Batch Loading Complete", message, self.gui.root)

"""
data_collection_data.py
Developed by Charlie Becquet
Data class for data collection window
"""

# pylint: disable=no-member
# This module is part of a multiple inheritance structure where attributes
# are defined in other parent classes (DataCollectionData, DataCollectionHandlers, etc.)

import copy
import time
import statistics
from utils import debug_print

class DataCollectionData:

    def handle_sample_count_change(self, old_count, new_count):
        """Handle changes in the number of samples."""
        debug_print(f"DEBUG: Handling sample count change: {old_count} -> {new_count}")

        if new_count > old_count:
            # Add new samples
            for i in range(old_count, new_count):
                sample_id = f"Sample {i+1}"

                # Initialize with the complete data structure (matching initialize_data)
                self.data[sample_id] = {
                    "current_row_index": 0,
                    "avg_tpm": 0.0
                }

                # Add all standard columns (must match the structure in initialize_data)
                columns = ["puffs", "before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog", "notes", "tpm"]
                for column in columns:
                    self.data[sample_id][column] = []

                # Add special columns for User Test Simulation
                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    self.data[sample_id]["chronography"] = []

                # Pre-initialize 50 rows for new sample
                for j in range(50):
                    puff = (j + 1) * self.puff_interval
                    self.data[sample_id]["puffs"].append(puff)

                    # Initialize other columns with empty values
                    for column in columns:
                        if column == "puffs":
                            continue  # Already initialized
                        elif column == "tpm":
                            self.data[sample_id][column].append(None)
                        else:
                            self.data[sample_id][column].append("")

                    # Add chronography initialization for User Test Simulation
                    if "chronography" in self.data[sample_id]:
                        self.data[sample_id]["chronography"].append("")

                debug_print(f"DEBUG: Added new sample {sample_id} with complete data structure")

            debug_print(f"DEBUG: Added {new_count - old_count} new samples")

        elif new_count < old_count:
            # Remove excess samples
            for i in range(new_count, old_count):
                sample_id = f"Sample {i+1}"
                if sample_id in self.data:
                    del self.data[sample_id]
                    debug_print(f"DEBUG: Removed sample {sample_id}")

            debug_print(f"DEBUG: Removed {old_count - new_count} samples")

        # Update the number of samples
        self.num_samples = new_count

        # Recreate the UI to reflect the new sample count
        self.recreate_sample_tabs()

    def sync_tksheet_to_data(self, sheet, sample_id):
        """Sync tksheet data back to internal data structure"""
        debug_print(f"DEBUG: Syncing tksheet data to internal structure for {sample_id}")

        try:
            sheet_data = sheet.get_sheet_data()
            # Store old weight values to detect changes
            old_before_weights = self.data[sample_id]["before_weight"].copy()
            old_after_weights = self.data[sample_id]["after_weight"].copy()

            # Find last nonzero puff index AND value (for puff auto-population)
            last_nonzero_puff_index = None
            last_puff_value = 0

            for i, row in enumerate(sheet_data):
                try:
                    puff_val = float(row[0]) if row[0] not in (None, "") else 0
                    if puff_val != 0:
                        last_nonzero_puff_index = i
                        last_puff_value = int(puff_val)  # Store the actual value, not just index
                        debug_print(f"DEBUG: Found nonzero puff at row {i} with value {last_puff_value}")
                except (ValueError, IndexError):
                    continue

            # Find last nonzero AFTER WEIGHT index (for weight auto-progression)
            last_nonzero_after_weight_index = None

            # Get correct column indices based on test type
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                before_weight_col = 2
                after_weight_col = 3
            else:
                before_weight_col = 1
                after_weight_col = 2

            for i, row in enumerate(sheet_data):
                try:
                    after_weight_val = str(row[after_weight_col]).strip() if len(row) > after_weight_col and row[after_weight_col] is not None else ""
                    if after_weight_val and after_weight_val != "":
                        last_nonzero_after_weight_index = i
                        debug_print(f"DEBUG: Found nonzero after weight at row {i} with value {after_weight_val}")
                except (ValueError, IndexError):
                    continue

            debug_print(f"DEBUG: Last puff index: {last_nonzero_puff_index}, Last after weight index: {last_nonzero_after_weight_index}")

            # Clear existing data arrays
            for key in self.data[sample_id]:
                if key not in ["current_row_index", "avg_tpm"]:
                    self.data[sample_id][key] = []

            # Track auto weight progression actions to execute later
            auto_weight_actions = []

            # Rebuild data from sheet
            for row_idx, row_data in enumerate(sheet_data):
                # Ensure row_data is a list and handle empty/None values
                if not isinstance(row_data, list):
                    row_data = list(row_data) if row_data is not None else []

                # Helper function to safely get cell value
                def safe_get_cell(data_row, col_idx):
                    if col_idx < len(data_row) and data_row[col_idx] is not None:
                        val = str(data_row[col_idx]).strip()
                        return val if val else ""
                    return ""

                # Detect before/after weight value change
                if self.old_sheet_data is not None:
                    old_row = self.old_sheet_data[row_idx] if row_idx < len(self.old_sheet_data) else []
                else:
                    old_row = sheet_data[row_idx] if row_idx < len(sheet_data) else []

                new_before = safe_get_cell(row_data, before_weight_col)
                new_after = safe_get_cell(row_data, after_weight_col)
                old_before = str(old_row[before_weight_col]).strip() if len(old_row) > before_weight_col and old_row[before_weight_col] is not None else ""
                old_after = str(old_row[after_weight_col]).strip() if len(old_row) > after_weight_col and old_row[after_weight_col] is not None else ""
                new_puffs = safe_get_cell(row_data, 0)

                # Check for auto weight progression condition - ONLY for after weight column changes
                should_trigger_auto_weight = False
                # CRITICAL: Only trigger if AFTER weight changed AND before weight did NOT change
                # This ensures we only trigger when user edited the after weight column, not before weight column
                after_weight_changed = (old_after == "" and new_after != "")
                before_weight_changed = (old_before == "" and new_before != "")

                if after_weight_changed and not before_weight_changed:  # Only after weight changed, not before weight
                    # Use AFTER WEIGHT index for weight progression logic
                    if last_nonzero_after_weight_index is None or row_idx >= last_nonzero_after_weight_index:
                        debug_print(f"DEBUG: AFTER weight column changed from empty to {new_after} at row {row_idx} - scheduling auto weight AND puff progression")
                        auto_weight_actions.append((row_idx, new_after))
                        should_trigger_auto_weight = True
                    else:
                        debug_print(f"DEBUG: After weight changed at row {row_idx} but this is before last after weight row {last_nonzero_after_weight_index} - no auto progression")
                elif before_weight_changed and not after_weight_changed:
                    debug_print(f"DEBUG: BEFORE weight column changed from empty to {new_before} at row {row_idx} - NO auto progression (only TPM update)")
                elif after_weight_changed and before_weight_changed:
                    debug_print(f"DEBUG: BOTH weight columns changed at row {row_idx} - this might be bulk data entry or auto-population, no auto progression")

                # Handle puff updates - USE PUFF INDEX for puff logic
                should_update_puffs = False
                debug_print(f"Should we update puffs? {old_row}, new before:{new_before}, new after:{new_after}")

                # Only update puffs if we're past the last known puffs row AND weights changed
                if last_nonzero_puff_index is not None and row_idx > last_nonzero_puff_index:
                    if (old_before == "" and new_before != "") or (old_after == "" and new_after != ""):
                        if new_puffs == "":
                            debug_print(f"Yes! New data detected at row {row_idx}, auto-filling puffs")
                            should_update_puffs = True
                        else:
                            debug_print(f"No - puffs already filled: {new_puffs}")
                    else:
                        debug_print(f"No - no weight changes detected")
                elif last_nonzero_puff_index is None:
                    # No existing puffs data, this is the first entry
                    if (old_before == "" and new_before != "") or (old_after == "" and new_after != ""):
                        if new_puffs == "":
                            debug_print(f"Yes! First data entry detected at row {row_idx}, auto-filling puffs")
                            should_update_puffs = True
                            # For first entry, set last_nonzero_puff_index to -1 so calculation works
                            last_nonzero_puff_index = -1
                            last_puff_value = 0
                else:
                    debug_print(f"No - editing existing puff data at row {row_idx} (last puffs row: {last_nonzero_puff_index})")

                if should_update_puffs:
                    # Calculate correct puff value: base + increments
                    puff_value = last_puff_value + (row_idx - last_nonzero_puff_index) * self.puff_interval
                    new_puffs = str(puff_value)
                    debug_print(f"DEBUG: Auto-filled puffs at row {row_idx} with value {new_puffs} (base: {last_puff_value}, distance: {row_idx - last_nonzero_puff_index}, interval: {self.puff_interval})")

                    # Update sheet visually
                    puff_col_idx = 1 if self.test_name in ["User Test Simulation", "User Simulation Test"] else 0
                    sheet.set_cell_data(row_idx, puff_col_idx, new_puffs)

                    # Populate all intermediate puff values from last_nonzero_puff_index + 1 to row_idx
                    for intermediate_row in range(last_nonzero_puff_index + 1, row_idx + 1):
                        if intermediate_row < len(sheet_data):
                            intermediate_puff_value = last_puff_value + (intermediate_row - last_nonzero_puff_index) * self.puff_interval
                            sheet.set_cell_data(intermediate_row, puff_col_idx, str(intermediate_puff_value))
                            debug_print(f"DEBUG: Auto-populated intermediate row {intermediate_row} with puffs value {intermediate_puff_value}")

                if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                    # Map columns to data structure for User Test Simulation
                    self.data[sample_id]["chronography"].append(safe_get_cell(row_data, 0))
                    self.data[sample_id]["puffs"].append(safe_get_cell(row_data, 1))
                    self.data[sample_id]["before_weight"].append(safe_get_cell(row_data, 2))
                    self.data[sample_id]["after_weight"].append(safe_get_cell(row_data, 3))
                    self.data[sample_id]["draw_pressure"].append(safe_get_cell(row_data, 4))
                    self.data[sample_id]["smell"].append(safe_get_cell(row_data, 5))  # Failure
                    self.data[sample_id]["notes"].append(safe_get_cell(row_data, 6))
                    self.data[sample_id]["tpm"].append(None)  # Will be recalculated
                else:
                    # Standard format mapping
                    self.data[sample_id]["puffs"].append(safe_get_cell(row_data, 0))
                    self.data[sample_id]["before_weight"].append(safe_get_cell(row_data, 1))
                    self.data[sample_id]["after_weight"].append(safe_get_cell(row_data, 2))
                    self.data[sample_id]["draw_pressure"].append(safe_get_cell(row_data, 3))
                    self.data[sample_id]["resistance"].append(safe_get_cell(row_data, 4))
                    self.data[sample_id]["smell"].append(safe_get_cell(row_data, 5))
                    self.data[sample_id]["clog"].append(safe_get_cell(row_data, 6))
                    self.data[sample_id]["notes"].append(safe_get_cell(row_data, 7))
                    self.data[sample_id]["tpm"].append(None)  # Will be recalculated

            # NOW execute auto weight progression actions after data arrays are fully rebuilt
            debug_print(f"DEBUG: Executing {len(auto_weight_actions)} auto weight progression actions")
            for action_row_idx, action_value in auto_weight_actions:
                debug_print(f"DEBUG: Executing auto weight progression for row {action_row_idx} with value {action_value}")
                # Update next row's before weight
                self.auto_progress_weight(sheet, sample_id, action_row_idx, action_value)

                # ALSO update next row's puffs if it doesn't have any
                next_row_idx = action_row_idx + 1
                if next_row_idx < len(self.data[sample_id]["puffs"]):
                    current_puffs = self.data[sample_id]["puffs"][next_row_idx] if next_row_idx < len(self.data[sample_id]["puffs"]) else ""
                    if not current_puffs or str(current_puffs).strip() == "":
                        # Calculate next puff value
                        current_puff_value = int(self.data[sample_id]["puffs"][action_row_idx]) if action_row_idx < len(self.data[sample_id]["puffs"]) and self.data[sample_id]["puffs"][action_row_idx] else 0
                        next_puff_value = current_puff_value + self.puff_interval

                        # Update data structure
                        self.data[sample_id]["puffs"][next_row_idx] = next_puff_value

                        # Update sheet display
                        puff_col_idx = 1 if self.test_name in ["User Test Simulation", "User Simulation Test"] else 0
                        sheet.set_cell_data(next_row_idx, puff_col_idx, str(next_puff_value))

                        debug_print(f"DEBUG: Auto weight progression also updated next row puffs: row {next_row_idx} = {next_puff_value}")

            # Check if any weights changed
            weights_changed = (
                old_before_weights != self.data[sample_id]["before_weight"] or
                old_after_weights != self.data[sample_id]["after_weight"]
            )

            if weights_changed:
                debug_print(f"DEBUG: Weight values changed, triggering TPM recalculation")
                # Force immediate TPM calculation
                self.calculate_tpm(sample_id)
                # Force immediate plot update
                self.update_stats_panel()

            debug_print(f"DEBUG: Synced {len(sheet_data)} rows to internal data structure")

        except Exception as e:
            debug_print(f"DEBUG: Error syncing tksheet data: {e}")
            import traceback
            traceback.print_exc()

    def process_sheet_changes(self, sheet, sample_id):
        """Process all changes to the sheet data."""
        debug_print(f"DEBUG: Processing sheet changes for {sample_id}")

        try:
            # Sync sheet data to internal structure
            self.sync_tksheet_to_data(sheet, sample_id)

            # Recalculate TPM
            self.calculate_tpm(sample_id)

            # Update TPM column in sheet
            self.update_tpm_in_sheet(sheet, sample_id)

            # Update stats panel
            self.window.after(100, self.update_stats_panel)

            # Mark as changed
            self.mark_unsaved_changes()

            # update old sheet data for next update
            self.old_sheet_data = sheet.get_sheet_data()

        except Exception as e:
            debug_print(f"DEBUG: Error processing sheet changes: {e}")

    def calculate_tpm(self, sample_id):
        """Calculate TPM values for all rows for a specific sample"""

        # Ensure all required lists are the same length
        max_length = max(
            len(self.data[sample_id]["puffs"]),
            len(self.data[sample_id]["before_weight"]),
            len(self.data[sample_id]["after_weight"])
        )

        # Extend lists if needed
        while len(self.data[sample_id]["tpm"]) < max_length:
            self.data[sample_id]["tpm"].append(None)

        valid_tpm_values = []

        for i in range(len(self.data[sample_id]["puffs"])):
            try:
                # Get weight values
                before_weight_str = self.data[sample_id]["before_weight"][i]
                after_weight_str = self.data[sample_id]["after_weight"][i]

                # Skip if either weight is missing
                if not before_weight_str or not after_weight_str:
                    continue

                # Convert to float
                before_weight = float(before_weight_str)
                after_weight = float(after_weight_str)

                # Validate weights
                if before_weight <= after_weight:
                    continue

                # Calculate puffs in this interval
                current_puff = int(self.data[sample_id]["puffs"][i])

                if i == 0:
                    # First row: use the current puff count (e.g., 20 puffs from 0 to 20)
                    puffs_in_interval = current_puff
                else:
                    # Subsequent rows: difference from previous puff count
                    prev_puff = int(self.data[sample_id]["puffs"][i - 1])
                    puffs_in_interval = current_puff - prev_puff

                # Skip if invalid puff interval
                if puffs_in_interval <= 0:
                    continue

                # Calculate TPM (mg/puff)
                weight_consumed = before_weight - after_weight  # in grams
                tpm = (weight_consumed * 1000) / puffs_in_interval  # Convert to mg per puff

                # Store result
                self.data[sample_id]["tpm"][i] = round(tpm, 3)
                valid_tpm_values.append(tpm)

            except Exception:
                # Ensure tpm list is long enough even for failed calculations
                while len(self.data[sample_id]["tpm"]) <= i:
                    self.data[sample_id]["tpm"].append(None)

        # Update average TPM
        self.data[sample_id]["avg_tpm"] = sum(valid_tpm_values) / len(valid_tpm_values) if valid_tpm_values else 0.0

        return len(valid_tpm_values) > 0

    def calculate_dynamic_plot_size(self, parent_frame):
        """Calculate plot size that directly scales with window size."""
        debug_print("DEBUG: Starting dynamic plot size calculation")

        # Force geometry update to get current dimensions
        self.window.update_idletasks()

        if hasattr(self, 'stats_frame') and self.stats_frame.winfo_exists():
            parent_frame = self.stats_frame
            parent_frame.update_idletasks()

        # Get the actual dimensions
        available_width = parent_frame.winfo_width()
        available_height = parent_frame.winfo_height()

        debug_print(f"DEBUG: Parent frame: {parent_frame.__class__.__name__}")
        debug_print(f"DEBUG: Available dimensions - Width: {available_width}px, Height: {available_height}px")

        # Simple fallback for initial sizing or very small windows
        if available_width < 200 or available_height < 200:
            debug_print("DEBUG: Using fallback size for small window")
            return (6, 4)

        # Reserve space for stats text and controls
        text_space = 120  # Space for text statistics
        control_space = 50  # Space for labels and padding

        # Calculate available space for the plot
        plot_height_available = available_height - text_space - control_space
        plot_width_available = available_width - 40  # 20px margin on each side

        # Use most of the available space for the plot
        plot_width_pixels = max(plot_width_available, 200)
        plot_height_pixels = max(plot_height_available, 150)

        debug_print(f"DEBUG: Plot space in pixels - Width: {plot_width_pixels}px, Height: {plot_height_pixels}px")

        # Convert to inches for matplotlib (using standard 100 DPI)
        plot_width_inches = plot_width_pixels / 100.0
        plot_height_inches = plot_height_pixels / 100.0

        # Apply minimum and maximum sizes
        plot_width_inches = max(min(plot_width_inches, 12.0), 3.0)
        plot_height_inches = max(min(plot_height_inches, 8.0), 2.0)

        debug_print(f"DEBUG: FINAL plot size - Width: {plot_width_inches:.2f} inches, Height: {plot_height_inches:.2f} inches")

        return (plot_width_inches, plot_height_inches)

    def validate_weight_entry(self, sample_id, row_idx, column_name, value):
        """
        Validate weight entries to ensure data consistency.
        Returns True if valid, False otherwise.
        """
        try:
            if not value or not value.strip():
                return True  # Empty values are allowed

            weight_value = float(value.strip())

            # Check for reasonable weight values (between 0.001g and 100g)
            if weight_value < 0.001 or weight_value > 100:
                debug_print(f"WARNING: Weight value {weight_value}g seems unreasonable for row {row_idx}")
                return False

            # If this is an after_weight, check that it's less than before_weight
            if column_name == "after_weight":
                before_weight_str = self.data[sample_id]["before_weight"][row_idx]
                if before_weight_str and before_weight_str.strip():
                    try:
                        before_weight = float(before_weight_str.strip())
                        if weight_value >= before_weight:
                            debug_print(f"WARNING: After weight ({weight_value}g) should be less than before weight ({before_weight}g)")
                            return False
                    except ValueError:
                        pass

            # If this is a before_weight, check that it's greater than after_weight
            elif column_name == "before_weight":
                after_weight_str = self.data[sample_id]["after_weight"][row_idx]
                if after_weight_str and after_weight_str.strip():
                    try:
                        after_weight = float(after_weight_str.strip())
                        if weight_value <= after_weight:
                            debug_print(f"WARNING: Before weight ({weight_value}g) should be greater than after weight ({after_weight}g)")
                            return False
                    except ValueError:
                        pass

            return True

        except ValueError:
            debug_print(f"ERROR: Invalid weight value '{value}' - must be a number")
            return False

    def start_auto_save_timer(self):
        """Start the auto-save timer."""
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)

        self.auto_save_timer = self.window.after(self.auto_save_interval, self.auto_save)

    def auto_save(self):
        """Perform automatic save without user confirmation."""
        if self.has_unsaved_changes:
            try:
                self.save_data(show_confirmation=False, auto_save=True)
                self.update_save_status(False)  # Mark as saved
                debug_print("DEBUG: Auto-save completed successfully")
            except Exception as e:
                debug_print(f"DEBUG: Auto-save failed: {e}")
                # Continue even if auto-save fails

        # Restart the timer
        self.start_auto_save_timer()

    def mark_unsaved_changes(self):
        """Mark that there are unsaved changes."""
        if not self.has_unsaved_changes:
            self.has_unsaved_changes = True
            self.update_save_status(True)

    def update_save_status(self, has_changes):
        """Update the save status indicator in the UI."""
        if has_changes:
            self.save_status_label.config(foreground="red", text="●")
            self.save_status_text.config(text="Unsaved changes")
        else:
            self.save_status_label.config(foreground="green", text="●")
            self.save_status_text.config(text="All changes saved")
            import datetime
            self.last_save_time = datetime.datetime.now()

    def initialize_data(self):
        """Initialize data structures for all samples."""
        self.data = {}
        for i in range(self.num_samples):
            sample_id = f"Sample {i+1}"

            # Basic data structure
            self.data[sample_id] = {
                "current_row_index": 0,
                "avg_tpm": 0.0,
                "sample_notes": ""
            }

            # Add standard columns (smell field will be used for both "smell" and "failure" depending on test type)
            columns = ["puffs", "before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog", "notes", "tpm"]
            for column in columns:
                self.data[sample_id][column] = []

            # Add special columns for User Test Simulation
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                self.data[sample_id]["chronography"] = []

            # Pre-initialize rows
            for j in range(50):
                puff = (j + 1) * self.puff_interval
                self.data[sample_id]["puffs"].append(puff)

                # Initialize other columns with empty values
                for column in columns:
                    if column == "puffs":
                        continue  # Already initialized
                    elif column == "tpm":
                        self.data[sample_id][column].append(None)
                    else:
                        self.data[sample_id][column].append("")

                # Initialize chronography if needed
                if "chronography" in self.data[sample_id]:
                    self.data[sample_id]["chronography"].append("")

            self.log(f"Initialized data for Sample {i+1}", "debug")

    def calculate_tpm_for_row(self, sample_id, changed_row_idx):
        """Calculate TPM for a specific row and related rows when data changes."""
        debug_print(f"DEBUG: Calculating TPM for {sample_id}, row {changed_row_idx}")

        # Calculate TPM for the changed row
        self.calculate_single_row_tpm(sample_id, changed_row_idx)

        # If puffs changed, recalculate all subsequent rows since intervals may have changed
        if changed_row_idx < len(self.data[sample_id]["puffs"]) - 1:
            for i in range(changed_row_idx + 1, len(self.data[sample_id]["puffs"])):
                self.calculate_single_row_tpm(sample_id, i)

        # Update average TPM
        valid_tpm_values = [v for v in self.data[sample_id]["tpm"] if v is not None]
        self.data[sample_id]["avg_tpm"] = sum(valid_tpm_values) / len(valid_tpm_values) if valid_tpm_values else 0.0

    def calculate_single_row_tpm(self, sample_id, row_idx):
        """Calculate TPM for a single row."""
        try:
            # Ensure TPM list is long enough
            while len(self.data[sample_id]["tpm"]) <= row_idx:
                self.data[sample_id]["tpm"].append(None)

            # Get weight values for this row
            before_weight_str = self.data[sample_id]["before_weight"][row_idx] if row_idx < len(self.data[sample_id]["before_weight"]) else ""
            after_weight_str = self.data[sample_id]["after_weight"][row_idx] if row_idx < len(self.data[sample_id]["after_weight"]) else ""

            # Skip if either weight is missing
            if not before_weight_str or not after_weight_str:
                self.data[sample_id]["tpm"][row_idx] = None
                return

            # Convert to float
            before_weight = float(before_weight_str)
            after_weight = float(after_weight_str)

            # Validate weights
            if before_weight <= after_weight:
                self.data[sample_id]["tpm"][row_idx] = None
                return

            # Calculate puffs in this interval
            current_puff = int(self.data[sample_id]["puffs"][row_idx]) if row_idx < len(self.data[sample_id]["puffs"]) else 0
            puffs_in_interval = current_puff

            if row_idx > 0:
                prev_puff = int(self.data[sample_id]["puffs"][row_idx - 1]) if (row_idx - 1) < len(self.data[sample_id]["puffs"]) else 0
                puffs_in_interval = current_puff - prev_puff

            # Skip if invalid puff interval
            if puffs_in_interval <= 0:
                self.data[sample_id]["tpm"][row_idx] = None
                return

            # Calculate TPM (mg/puff)
            weight_consumed = before_weight - after_weight  # in grams
            tpm = (weight_consumed * 1000) / puffs_in_interval  # Convert to mg per puff

            # Store result
            self.data[sample_id]["tpm"][row_idx] = round(tpm, 6)

            debug_print(f"DEBUG: Calculated TPM for {sample_id} row {row_idx}: {tpm:.6f} mg/puff")

        except (ValueError, TypeError, IndexError) as e:
            debug_print(f"DEBUG: Error calculating TPM for {sample_id} row {row_idx}: {e}")
            self.data[sample_id]["tpm"][row_idx] = None

    def convert_cell_value(self, value, column_name):
        """Convert user-entered value to appropriate type."""
        value = value.strip()
        if not value:
            return "" if column_name != "puffs" else 0

        try:
            if column_name == "puffs":
                return int(float(value))
            elif column_name in ["before_weight", "after_weight", "draw_pressure", "resistance", "smell", "clog"]:
                return float(value)
            elif column_name in ["chronography", "notes"]:
                return str(value)
        except ValueError:
            return value  # Keep as-is if conversion fails
        return value

    def get_column_name(self, col_idx):
        """Get the internal column name based on column index."""
        # Column mapping varies by test type
        if self.test_name in ["User Test Simulation", "User Simulation Test"]:
            # User Test Simulation: [Chronography, Puffs, Before Weight, After Weight, Draw Pressure, Failure, Notes, TPM]
            column_map = {
                0: "chronography",
                1: "puffs",
                2: "before_weight",
                3: "after_weight",
                4: "draw_pressure",
                5: "smell",  # "Failure" column maps to "smell" field in data structure
                6: "notes",
                7: "tpm"  # TPM column (read-only)
            }
        else:
            # Standard test: [Puffs, Before Weight, After Weight, Draw Pressure, Resistance, Smell, Clog, Notes, TPM]
            column_map = {
                0: "puffs",
                1: "before_weight",
                2: "after_weight",
                3: "draw_pressure",
                4: "resistance",
                5: "smell",
                6: "clog",
                7: "notes",
                8: "tpm"  # TPM column (read-only)
            }

        # Return appropriate column name or default to a safe value
        return column_map.get(col_idx, "notes")

    def ensure_initial_tpm_calculation(self):
        """Ensure TPM is calculated and displayed when window opens."""
        debug_print("DEBUG: Ensuring initial TPM calculation and display")

        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"

            # Calculate TPM for this sample
            self.calculate_tpm(sample_id)

        if hasattr(self, 'sample_sheets') and i < len(self.sample_sheets):
            sheet = self.sample_sheets[i]
            self.update_tksheet(sheet, sample_id)

        # Update stats panel to show current TPM data
        self.update_stats_panel()

        debug_print("DEBUG: Initial TPM calculation and display completed")
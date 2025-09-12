"""
data_collection_handlers.py
Developed by Charlie Becquet
Event handlers class for data collection window.
"""

import tkinter as tk
from tkinter import messagebox
import statistics
import time
import datetime
from utils import show_success_message, debug_print

class DataCollectionHandlers:
    def setup_event_handlers(self):
        """Set up event handlers for the window."""
        # Window close protocol
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)

        # Keyboard shortcuts
        self.setup_hotkeys()

        # Bind window resize events to update plot size
        #self.window.bind('<Configure>', self.on_window_resize_plot, add=True)
        debug_print("DEBUG: Window resize event handler bound")

        # Initialize plot sizing after window is fully set up
        #self.window.after(1000, self.initialize_plot_size)

        debug_print("DEBUG: Event handlers set up")

    def setup_hotkeys(self):
        """Set up keyboard shortcuts for navigation."""
        if not hasattr(self, 'hotkey_bindings'):
            self.hotkey_bindings = {}

        # Clear any existing bindings
        for key, binding_id in self.hotkey_bindings.items():
            try:
                self.window.unbind(key, binding_id)
            except:
                pass

        self.hotkey_bindings.clear()

        # Bind Ctrl+S for quick save
        self.window.bind("<Control-s>", lambda e: self.save_data(show_confirmation=False) if self.hotkeys_enabled else None)

        # Use Alt+Left/Right for navigation instead of Ctrl
        def handle_alt_nav(direction):
            if not self.hotkeys_enabled:
                return "break"

            if direction == "left":
                self.go_to_previous_sample()
            else:
                self.go_to_next_sample()
            return "break"

        # Bind Alt+Left/Right for sample navigation
        self.window.bind("<Alt-Left>", lambda e: handle_alt_nav("left"))
        self.window.bind("<Alt-Right>", lambda e: handle_alt_nav("right"))

        # Also bind with focus_set to ensure window has focus
        self.window.focus_set()

        debug_print("DEBUG: Hotkeys set up with Alt+Left/Right for sample navigation")

    def on_tab_changed(self, event=None):
        """Handle notebook tab changes to properly save/load notes."""
        debug_print("DEBUG: Tab changed - handling notes save/load")

        # Get the NEW tab index (after the change)
        try:
            new_tab_index = self.notebook.index(self.notebook.select())
            debug_print(f"DEBUG: New tab index: {new_tab_index}")
        except:
            new_tab_index = 0
            debug_print("DEBUG: Failed to get new tab index, defaulting to 0")

        # Save notes for the PREVIOUS tab (before the change) if we have notes widget
        if hasattr(self, 'sample_notes_text') and self.sample_notes_text.winfo_exists():
            try:
                # Get the current notes content
                current_notes = self.sample_notes_text.get('1.0', 'end-1c')

                # Save to the PREVIOUS tab's data using our tracked index
                previous_sample_id = f"Sample {self.current_tab_index + 1}"
                self.data[previous_sample_id]["sample_notes"] = current_notes
                debug_print(f"DEBUG: Saved notes from previous tab {previous_sample_id}: '{current_notes[:30]}...'")

                # Force the text widget to lose focus
                self.sample_notes_text.selection_clear()
                self.window.focus_set()  # Move focus to main window

                # Process any pending events (including FocusOut)
                self.window.update_idletasks()
                debug_print("DEBUG: Forced notes widget to lose focus and processed pending events")

            except Exception as e:
                debug_print(f"DEBUG: Error saving notes during tab change: {e}")

        # Update our tracking to the new tab
        self.current_tab_index = new_tab_index
        debug_print(f"DEBUG: Updated current_tab_index to: {new_tab_index}")

        # Small delay to ensure all events are processed, then update the panel
        self.window.after(10, self.update_stats_panel)

    def on_edit_complete(self, sheet, sample_id, row_idx):
        """Handle actions when user completes editing a cell (Enter, Tab, etc.)"""
        try:
            debug_print(f"DEBUG: Edit completed for {sample_id} at row {row_idx}")

            # Check if current row has an after weight and next row needs before weight
            if row_idx < len(self.data[sample_id]["after_weight"]) - 1:  # Not the last row
                current_after_weight = self.data[sample_id]["after_weight"][row_idx]
                next_before_weight = self.data[sample_id]["before_weight"][row_idx + 1] if (row_idx + 1) < len(self.data[sample_id]["before_weight"]) else ""

                # If current row has after weight and next row's before weight is empty
                if current_after_weight and str(current_after_weight).strip() != "" and (not next_before_weight or str(next_before_weight).strip() == ""):
                    debug_print(f"DEBUG: Edit completion - auto-progressing weight from row {row_idx} ({current_after_weight}) to row {row_idx + 1}")

                    # Update the next row's before weight in data structure
                    self.data[sample_id]["before_weight"][row_idx + 1] = current_after_weight

                    # Update the sheet display
                    before_weight_col_idx = 2 if self.test_name in ["User Test Simulation", "User Simulation Test"] else 1
                    sheet.set_cell_data(row_idx + 1, before_weight_col_idx, str(current_after_weight))

                    debug_print(f"DEBUG: Edit completion - auto-set next row before_weight to {current_after_weight}")

                    # Mark as changed and recalculate TPM
                    self.mark_unsaved_changes()
                    self.calculate_tpm(sample_id)
                    self.update_tpm_in_sheet(sheet, sample_id)

        except Exception as e:
            debug_print(f"DEBUG: Error in on_edit_complete: {e}")

    def setup_enhanced_sheet_bindings(self, sheet, sample_id, sample_index):
        """Set up simplified event bindings for reliable cell editing."""

        # Simple event handler for when cells are modified
        def on_cell_modified(event):
            """Handle any cell modification."""
            try:
                debug_print(f"DEBUG: Cell modified in {sample_id}")
                # Small delay to ensure edit is complete
                self.window.after(50, lambda: self.process_sheet_changes(sheet, sample_id))

            except Exception as e:
                debug_print(f"DEBUG: Error in cell modified handler: {e}")

        # Handle selection changes and detect edit completion
        def on_selection_changed(event):
            try:
                selections = sheet.get_selected_cells()
                if selections:
                    new_row, col = selections[0]

                    # Check if we moved to a different row (edit completed on previous row)
                    if hasattr(self, '_current_edit_row') and self._current_edit_row is not None and self._current_edit_row != new_row:
                        debug_print(f"DEBUG: Edit completed, moved from row {self._current_edit_row} to row {new_row}")
                        # User moved to different row - edit on previous row is complete
                        self.on_edit_complete(sheet, sample_id, self._current_edit_row)

                    # Update current edit row
                    self._current_edit_row = new_row
                    debug_print(f"DEBUG: Selection changed to row {new_row}, col {col}")

            except Exception as e:
                debug_print(f"DEBUG: Error in selection handler: {e}")


        # Bind to the sheet's built-in events
        sheet.bind("<<SheetModified>>", on_cell_modified)
        sheet.bind("<<SheetSelectCell>>", on_selection_changed)

        debug_print(f"DEBUG: Simplified sheet bindings set up for {sample_id}")

    def on_cancel(self):
        """Handle cancel button click or window close."""
        self.on_window_close()

    def on_notes_changed(self, event=None):
        """Handle changes to sample notes text."""
        # Add this check at the beginning
        if self.updating_notes:
            debug_print("DEBUG: Ignoring notes change during panel update")
            return

        debug_print("DEBUG: Sample notes changed")

        try:
            # Use our tracked tab index instead of querying the notebook
            current_sample_id = f"Sample {self.current_tab_index + 1}"
            debug_print(f"DEBUG: Using tracked tab index {self.current_tab_index} for sample {current_sample_id}")

            # Get the current notes content
            notes_content = self.sample_notes_text.get('1.0', 'end-1c')

            # Store in data structure
            self.data[current_sample_id]["sample_notes"] = notes_content

            debug_print(f"DEBUG: Updated notes for {current_sample_id}: '{notes_content[:30]}...'")

        except Exception as e:
            debug_print(f"DEBUG: Error updating sample notes: {e}")

    def update_tpm_plot_for_current_sample(self):
        """Update the TPM plot for the currently selected sample with smart y-axis bounds."""
        try:
            # Get currently selected sample
            current_tab_index = self.notebook.index(self.notebook.select())
            current_sample_id = f"Sample {current_tab_index + 1}"
        except:
            current_sample_id = "Sample 1"
            current_tab_index = 0

        debug_print(f"DEBUG: Updating TPM plot for {current_sample_id}")

        # Clear the plot
        self.stats_ax.clear()

        # Get data for the sample
        tpm_values = [v for v in self.data[current_sample_id]["tpm"] if v is not None]
        puff_values = []

        # Get corresponding puff values for non-None TPM values
        for i, tpm in enumerate(self.data[current_sample_id]["tpm"]):
            if tpm is not None and i < len(self.data[current_sample_id]["puffs"]):
                puff_values.append(self.data[current_sample_id]["puffs"][i])

        if tpm_values and puff_values and len(tpm_values) == len(puff_values):
            # Plot TPM over puffs
            self.stats_ax.plot(puff_values, tpm_values, marker='o', linewidth=2, markersize=4, color='blue')
            self.stats_ax.set_xlabel('Puffs', fontsize=10)
            self.stats_ax.set_ylabel('TPM (mg/puff)', fontsize=10)
            self.stats_ax.set_title(f'TPM Over Time - {current_sample_id}', fontsize=11)
            self.stats_ax.grid(True, alpha=0.3)

            # Smart y-axis bounds logic
            if len(tpm_values) >= 2:
                sorted_tpm = sorted(tpm_values, reverse=True)  # Sort descending
                max_tpm = sorted_tpm[0]
                second_max_tpm = sorted_tpm[1]

                if max_tpm <= 20:
                    # Use max + 2 when max <= 20
                    y_max = max_tpm + 2
                    debug_print(f"DEBUG: Setting y-axis bounds 0 to {y_max} (max {max_tpm} + 2)")
                elif second_max_tpm <= 20:
                    # Use second largest value when max > 20 but second max <= 20
                    y_max = second_max_tpm + 2
                    debug_print(f"DEBUG: Setting y-axis bounds 0 to {y_max} (second max {second_max_tpm} + 2, ignoring outlier {max_tpm})")
                else:
                    # Both values > 20, default to 20
                    y_max = 20
                    debug_print(f"DEBUG: Setting y-axis bounds 0 to 20 (both max {max_tpm} and second max {second_max_tpm} > 20)")
            elif len(tpm_values) == 1:
                # Single value
                max_tpm = tpm_values[0]
                if max_tpm <= 20:
                    y_max = max_tpm + 2
                else:
                    y_max = 20  # Single outlier, cap at 20
                debug_print(f"DEBUG: Setting y-axis bounds 0 to {y_max} (single value {max_tpm})")
            else:
                # No values, use default
                y_max = 9
                debug_print(f"DEBUG: Setting default y-axis bounds 0 to 9")

            self.stats_ax.set_ylim(0, y_max)

            # Adjust tick label sizes for better fit
            self.stats_ax.tick_params(axis='both', which='major', labelsize=9)

        else:
            # Show empty plot with labels
            self.stats_ax.set_xlabel('Puffs', fontsize=10)
            self.stats_ax.set_ylabel('TPM (mg/puff)', fontsize=10)
            self.stats_ax.set_title(f'TPM Over Time - {current_sample_id}', fontsize=11)
            self.stats_ax.grid(True, alpha=0.3)
            self.stats_ax.text(0.5, 0.5, 'No TPM data available',
                              transform=self.stats_ax.transAxes, ha='center', va='center', fontsize=10)
            self.stats_ax.set_ylim(0, 9)  # Default bounds for empty plot

        # Apply layout and draw
        self.stats_fig.tight_layout(pad=1.0)
        self.stats_canvas.draw()

        debug_print(f"DEBUG: TPM plot updated for {current_sample_id} with smart y-axis bounds")

    def update_stats_panel(self, event=None):
        """Update the enhanced TPM statistics panel."""
        debug_print("DEBUG: Updating enhanced TPM statistics panel")

        if not hasattr(self, 'stats_frame') or not self.stats_frame.winfo_exists():
            debug_print("DEBUG: Stats frame not available")
            return

        # Get current sample info
        current_tab = self.notebook.index(self.notebook.select())
        current_sample_id = f"Sample {current_tab + 1}"

        # Update column weights for equal distribution - FIX: Use text_frame_ref instead of text_frame
        if hasattr(self, 'text_frame_ref') and self.text_frame_ref.winfo_exists():
            for i in range(3):
                col_config = self.text_frame_ref.grid_columnconfigure(i)
                debug_print(f"DEBUG: text_frame column {i} weight: {col_config}")

        # Update stats for current sample
        debug_print(f"DEBUG: Updating stats for {current_sample_id}")

        # Get the actual sample name from header data - FIXED: Proper sample name extraction
        actual_sample_name = "Unknown Sample"
        if (hasattr(self, 'header_data') and
            'samples' in self.header_data and
            current_tab < len(self.header_data['samples'])):
            sample_info = self.header_data['samples'][current_tab]
            if isinstance(sample_info, dict):
                actual_sample_name = sample_info.get('id', f"Sample {current_tab + 1}")
            else:
                actual_sample_name = str(sample_info)

        debug_print(f"DEBUG: Using sample name: {actual_sample_name}")

        # Update the main sample label - FIXED: Show correct sample name
        if hasattr(self, 'current_sample_label') and self.current_sample_label.winfo_exists():
            self.current_sample_label.config(text=f"Sample {current_tab + 1}: {actual_sample_name}")
            debug_print(f"DEBUG: Updated main sample label to: Sample {current_tab + 1}: {actual_sample_name}")

        # Calculate TPM statistics
        sample_data = self.data[current_sample_id]
        tpm_values = [float(tpm) for tpm in sample_data.get("tpm", []) if tpm and str(tpm).strip() and tpm != ""]

        if tpm_values:
            avg_tpm = statistics.mean(tpm_values)
            latest_tpm = tpm_values[-1] if tpm_values else 0
            current_puffs = len([p for p in sample_data.get("puffs", []) if p and str(p).strip()])

            # Update statistics text - FIXED: Remove redundant headers
            sample_info_text = f"Average TPM: {avg_tpm:.5f}\nLatest TPM: {latest_tpm:.5f}\nStd Dev (last 5 sessions): {statistics.stdev(tpm_values[-5:]) if len(tpm_values) >= 2 else 0:.6f}\nCurrent Puffs: {current_puffs}"

            # Update sample information - FIXED: Remove redundant headers
            sample_info = self.header_data['samples'][current_tab] if current_tab < len(self.header_data.get('samples', [])) else {}
            sample_info_display = f"Media: {sample_info.get('media', '')}\nViscosity: {sample_info.get('viscosity', '')}\nInitial Oil Mass: {sample_info.get('oil_mass', '')}\nResistance: {sample_info.get('resistance', '')}\nVoltage: {sample_info.get('voltage', '')}\nPuffing Regime: {sample_info.get('puffing_regime', '')}\nDevice Type: {self.header_data.get('common', {}).get('device_type', '')}"

            debug_print(f"DEBUG: Formatted TPM stats text: {sample_info_text[:50]}...")
            debug_print(f"DEBUG: Formatted sample info text: {sample_info_display[:50]}...")

        else:
            sample_info_text = "No TPM data available"
            sample_info_display = "No data available"
            debug_print("DEBUG: No TPM data available for current sample")

        # Update label widgets - FIXED: Use config(text=...) for labels instead of text widget operations
        if hasattr(self, 'sample_stats_label') and self.sample_stats_label.winfo_exists():
            self.sample_stats_label.config(text=sample_info_text)
            debug_print("DEBUG: Updated TPM stats label")

        if hasattr(self, 'sample_info_label') and self.sample_info_label.winfo_exists():
            self.sample_info_label.config(text=sample_info_display)
            debug_print("DEBUG: Updated sample info label")

        # Update sample notes with proper protection against recursive events
        if hasattr(self, 'sample_notes_text') and self.sample_notes_text.winfo_exists():
            # Set flag to prevent recursive event handling
            self.updating_notes = True
            debug_print(f"DEBUG: Setting updating_notes flag to True for {current_sample_id}")

            try:
                # Completely unbind events temporarily
                self.sample_notes_text.unbind('<KeyRelease>')
                self.sample_notes_text.unbind('<FocusOut>')

                # Ensure the widget doesn't have focus before updating content
                if self.sample_notes_text == self.window.focus_get():
                    self.window.focus_set()
                    self.window.update_idletasks()
                    debug_print("DEBUG: Removed focus from notes widget before content update")

                # Clear and set notes content
                self.sample_notes_text.delete('1.0', 'end')
                notes_content = self.data[current_sample_id].get("sample_notes", "")
                if notes_content:
                    self.sample_notes_text.insert('1.0', notes_content)

                debug_print(f"DEBUG: Updated notes widget content for {current_sample_id}: '{notes_content[:50]}...'")

            finally:
                # Always re-enable event binding and clear flag
                self.sample_notes_text.bind('<KeyRelease>', self.on_notes_changed)
                self.sample_notes_text.bind('<FocusOut>', self.on_notes_changed)
                self.updating_notes = False
                debug_print(f"DEBUG: Clearing updating_notes flag and re-binding events for {current_sample_id}")

        # Update TPM plot for current sample
        self.update_tpm_plot_for_current_sample()

        debug_print(f"DEBUG: Enhanced stats updated for {current_sample_id} with correct sample name: {actual_sample_name}")

    def go_to_previous_sample(self):
        """Navigate to the previous sample tab with proper notes saving."""
        if not self.hotkeys_enabled:
            return

        # SAVE CURRENT NOTES BEFORE SWITCHING
        self._save_current_notes_before_tab_switch()

        current_tab = self.notebook.index(self.notebook.select())
        target_tab = (current_tab - 1) % len(self.sample_frames)

        debug_print(f"DEBUG: go_to_previous_sample switching from tab {current_tab} to {target_tab}")

        # Update our tracking to current tab (this will be the "previous" tab after the switch)
        self.current_tab_index = current_tab

        # Switch to the target tab
        self.notebook.select(target_tab)

    def go_to_next_sample(self):
        """Navigate to the next sample tab with proper notes saving."""
        if not self.hotkeys_enabled:
            return

        # SAVE CURRENT NOTES BEFORE SWITCHING
        self._save_current_notes_before_tab_switch()

        current_tab = self.notebook.index(self.notebook.select())
        target_tab = (current_tab + 1) % len(self.sample_frames)

        debug_print(f"DEBUG: go_to_next_sample switching from tab {current_tab} to {target_tab}")

        # Update our tracking to current tab (this will be the "previous" tab after the switch)
        self.current_tab_index = current_tab

        # Switch to the target tab
        self.notebook.select(target_tab)

    def auto_progress_weight(self, sheet, sample_id, row_idx, value):
        """Auto-fill the next row's before_weight with current after_weight."""
        if value in ["", 0, None]:
            return
        try:
            val = float(value)
            next_row = row_idx + 1
            if next_row < len(self.data[sample_id]["before_weight"]):
                self.data[sample_id]["before_weight"][next_row] = val
                # Update sheet display
                col_idx = 2 if self.test_name in ["User Test Simulation", "User Simulation Test"] else 1
                sheet.set_cell_data(next_row, col_idx, str(val))
                debug_print(f"DEBUG: Auto-set Sample {sample_id} row {next_row} before_weight to {val}")
        except (ValueError, TypeError):
            debug_print("DEBUG: Failed auto weight progression due to invalid value")

    def on_window_resize_plot(self, event):
        """Handle window resize events with debouncing for plot updates."""
        debug_print(f"DEBUG: RESIZE EVENT DETECTED - Widget: {event.widget}, Window: {self.window}")
        debug_print(f"DEBUG: Event widget type: {type(event.widget)}")
        debug_print(f"DEBUG: Window type: {type(self.window)}")
        debug_print(f"DEBUG: Event widget == window? {event.widget == self.window}")

        # Only handle main window resize events, not child widgets
        if event.widget != self.window:
            debug_print(f"DEBUG: Ignoring resize event from child widget: {event.widget}")
            return

        debug_print("DEBUG: MAIN WINDOW RESIZE CONFIRMED - Processing...")

        # Get current window dimensions for verification
        current_width = self.window.winfo_width()
        current_height = self.window.winfo_height()
        debug_print(f"DEBUG: Current window dimensions: {current_width}x{current_height}")

        # Debounce rapid resize events
        if hasattr(self, '_resize_timer'):
            self.window.after_cancel(self._resize_timer)
            debug_print("DEBUG: Cancelled previous resize timer")

        # Schedule plot size update with a small delay to avoid excessive updates
        self._resize_timer = self.window.after(1000, self.update_plot_size_for_resize)
        debug_print("DEBUG: Scheduled plot resize update in 1000ms")

    def update_plot_size_for_resize(self):
        """Update plot size with artifact prevention and frame validation."""
        try:
            # Check if we have the necessary components
            if not hasattr(self, 'stats_canvas') or not self.stats_canvas.get_tk_widget().winfo_exists():
                debug_print("DEBUG: Stats canvas not available for resize")
                return

            if not hasattr(self, 'stats_fig') or not self.stats_fig:
                debug_print("DEBUG: Stats figure not available for resize")
                return

            # Wait for frame geometry to stabilize
            self.window.update_idletasks()

            # Use the actual stats container for sizing
            if hasattr(self, 'stats_frame_container') and self.stats_frame_container.winfo_exists():
                parent_for_sizing = self.stats_frame_container
            else:
                debug_print("DEBUG: Stats frame container not available, skipping resize")
                return

            parent_for_sizing.update_idletasks()

            # Validate that frames have reasonable dimensions before proceeding
            parent_width = parent_for_sizing.winfo_width()
            parent_height = parent_for_sizing.winfo_height()

            debug_print(f"DEBUG: Parent frame size for stats: {parent_width}x{parent_height}")

            if parent_width < 200 or parent_height < 200:
                debug_print("DEBUG: Parent frame size too small, skipping this resize update")
                return

            # Calculate new size based on validated frame dimensions
            new_width, new_height = self.calculate_dynamic_plot_size(parent_for_sizing)

            # Get current figure size for comparison
            current_width, current_height = self.stats_fig.get_size_inches()

            # Only update if change is significant
            width_diff = abs(new_width - current_width)
            height_diff = abs(new_height - current_height)
            threshold = 0.5  # Threshold to reduce excessive updates

            if width_diff > threshold or height_diff > threshold:
                debug_print(f"DEBUG: Significant size change detected - updating plot from {current_width:.2f}x{current_height:.2f} to {new_width:.2f}x{new_height:.2f}")

                # Apply the new size
                self.stats_fig.set_size_inches(new_width, new_height)

                # Redraw the canvas
                self.stats_canvas.draw_idle()
                debug_print("DEBUG: Plot resize completed")
            else:
                debug_print("DEBUG: Size change below threshold, skipping update to prevent artifacts")

        except Exception as e:
            debug_print(f"DEBUG: Error during plot resize: {str(e)}")
            import traceback
            debug_print(f"DEBUG: Full traceback: {traceback.format_exc()}")

    def on_tab_changed(self, event=None):
        """Handle notebook tab changes to properly save/load notes."""
        debug_print("DEBUG: Tab changed - handling notes save/load")

        # Save current notes before switching if notes widget exists and has focus
        if hasattr(self, 'sample_notes_text') and self.sample_notes_text.winfo_exists():
            try:
                # Get the current notes content before tab switch
                current_notes = self.sample_notes_text.get('1.0', 'end-1c')

                # Get the previous tab that was selected (before this change)
                if hasattr(self, 'previous_tab_index'):
                    prev_sample_id = f"Sample {self.previous_tab_index + 1}"
                    self.data[prev_sample_id]["sample_notes"] = current_notes
                    debug_print(f"DEBUG: Saved notes for previous tab {prev_sample_id}: '{current_notes[:30]}...'")

                # Force the text widget to lose focus
                self.sample_notes_text.selection_clear()
                self.window.focus_set()  # Move focus to main window

                # Process any pending events (including FocusOut)
                self.window.update_idletasks()
                debug_print("DEBUG: Forced notes widget to lose focus and processed pending events")

            except Exception as e:
                debug_print(f"DEBUG: Error handling focus during tab change: {e}")

        # Store current tab for next change
        try:
            current_tab_index = self.notebook.index(self.notebook.select())
            self.previous_tab_index = current_tab_index
            debug_print(f"DEBUG: Stored current tab index: {current_tab_index}")
        except:
            pass

        # Small delay to ensure all events are processed, then update the panel
        self.window.after(10, self.update_stats_panel)

    def update_all_sample_info(self):
        """Update sample information display for all samples."""
        debug_print("DEBUG: Updating sample information for all samples")

        for i, sample_data in enumerate(self.header_data.get("samples", [])):
            try:
                sample_id = sample_data.get("id", f"Sample {i+1}") if isinstance(sample_data, dict) else str(sample_data)

                # Update the sample info display if it exists
                if hasattr(self, 'sample_info_displays') and i < len(self.sample_info_displays):
                    info_display = self.sample_info_displays[i]

                    # Update the text content
                    info_text = f"Media: {sample_data.get('media', 'Unknown')}\n"
                    info_text += f"Viscosity: {sample_data.get('viscosity', 'Unknown')}\n"
                    info_text += f"Initial Oil Mass: {sample_data.get('oil_mass', 'Unknown')}\n"
                    info_text += f"Resistance: {sample_data.get('resistance', 'Unknown')}\n"
                    info_text += f"Voltage: {sample_data.get('voltage', 'Unknown')}\n"
                    info_text += f"Puffing Regime: {sample_data.get('puffing_regime', 'Unknown')}\n"
                    info_text += f"Device Type: {self.header_data.get('common', {}).get('device_type', 'Unknown')}"

                    info_display.config(state='normal')
                    info_display.delete('1.0', tk.END)
                    info_display.insert('1.0', info_text)
                    info_display.config(state='disabled')

            except Exception as e:
                debug_print(f"ERROR: Failed to update sample info for sample {i}: {e}")

    def update_header_labels_recursive(self, widget):
        """Recursively update header labels in the widget tree."""
        try:
            if isinstance(widget, ttk.Label):
                text = widget.cget('text')
                if 'Tester:' in text:
                    new_text = f"Tester: {self.header_data['common']['tester']}"
                    widget.config(text=new_text)

            # Recurse through children
            for child in widget.winfo_children():
                self.update_header_labels_recursive(child)

        except Exception as e:
            debug_print(f"DEBUG: Error updating label: {e}")

    def load_existing_sample_images_from_vap3(self):
        """Load existing sample images from VAP3 file or main GUI when opening data collection window."""
        try:
            debug_print("DEBUG: Loading existing sample images from VAP3")

            images_loaded = False

            # Method 1: Check if parent has loaded VAP3 data with sample images
            if hasattr(self.parent, 'filtered_sheets'):
                # Check for VAP3 sample images data
                vap_data = getattr(self.parent, 'current_vap_data', {})
                sample_images = vap_data.get('sample_images', {})
                sample_image_crop_states = vap_data.get('sample_image_crop_states', {})

                if sample_images:
                    debug_print(f"DEBUG: Found existing sample images in VAP3: {len(sample_images)} samples")

                    # Load the images into our data structure
                    self.sample_images = sample_images.copy()
                    self.sample_image_crop_states = sample_image_crop_states.copy()
                    images_loaded = True

                    debug_print(f"DEBUG: Loaded existing sample images from VAP3: {list(self.sample_images.keys())}")

            # Method 2: Check if parent has stored sample metadata from previous transfers
            if not images_loaded and hasattr(self.parent, 'sample_image_metadata'):
                current_file = getattr(self.parent, 'current_file', None)
                if (current_file and
                    current_file in self.parent.sample_image_metadata and
                    self.test_name in self.parent.sample_image_metadata[current_file]):

                    metadata = self.parent.sample_image_metadata[current_file][self.test_name]
                    stored_sample_images = metadata.get('sample_images', {})
                    stored_crop_states = metadata.get('sample_image_crop_states', {})

                    if stored_sample_images:
                        debug_print(f"DEBUG: Found existing sample images in main GUI metadata: {len(stored_sample_images)} samples")

                        # Load the images into our data structure
                        self.sample_images = stored_sample_images.copy()
                        self.sample_image_crop_states = stored_crop_states.copy()
                        images_loaded = True

                        debug_print(f"DEBUG: Loaded existing sample images from main GUI: {list(self.sample_images.keys())}")

            # Method 3: Try to reconstruct from main GUI sheet images (fallback method)
            if not images_loaded and hasattr(self.parent, 'sheet_images'):
                current_file = getattr(self.parent, 'current_file', None)
                if (current_file and
                    current_file in self.parent.sheet_images and
                    self.test_name in self.parent.sheet_images[current_file]):

                    sheet_image_paths = self.parent.sheet_images[current_file][self.test_name]
                    if sheet_image_paths:
                        debug_print(f"DEBUG: Found {len(sheet_image_paths)} sheet images, attempting to organize by sample")

                        # Try to organize images by sample based on filename patterns or stored metadata
                        reconstructed_sample_images = {}
                        reconstructed_crop_states = {}

                        # Check if we have image crop states that might contain sample info
                        image_crop_states = getattr(self.parent, 'image_crop_states', {})

                        for img_path in sheet_image_paths:
                            # For now, put all images in "Sample 1" as fallback
                            # This could be enhanced to parse filenames or use other metadata
                            sample_key = "Sample 1"
                            if sample_key not in reconstructed_sample_images:
                                reconstructed_sample_images[sample_key] = []
                            reconstructed_sample_images[sample_key].append(img_path)
                            reconstructed_crop_states[img_path] = image_crop_states.get(img_path, False)

                        if reconstructed_sample_images:
                            self.sample_images = reconstructed_sample_images
                            self.sample_image_crop_states = reconstructed_crop_states
                            images_loaded = True
                            debug_print(f"DEBUG: Reconstructed sample images from sheet images")

            if images_loaded:
                # Log summary of loaded images
                total_images = sum(len(imgs) for imgs in self.sample_images.values())
                debug_print(f"DEBUG: Successfully loaded {total_images} total images across {len(self.sample_images)} samples")
                for sample_id, images in self.sample_images.items():
                    debug_print(f"DEBUG: {sample_id}: {len(images)} images")
            else:
                debug_print("DEBUG: No existing sample images found")

        except Exception as e:
            debug_print(f"ERROR: Failed to load existing sample images: {e}")
            import traceback
            traceback.print_exc()

    def highlight_tpm_cells(self, sheet, sample_id):
        """Highlight TPM cells that have calculated values (like the green highlighting in treeview)"""
        # Get TPM column index
        if hasattr(sheet, 'headers'):
            tpm_col_idx = len(sheet.headers) - 1
        else:
            if self.test_name in ["User Test Simulation", "User Simulation Test"]:
                tpm_col_idx = 7
            else:
                tpm_col_idx = 8

        for row_idx in range(len(self.data[sample_id]["tpm"])):
            tpm_value = self.data[sample_id]["tpm"][row_idx]
            if tpm_value is not None:
                # Highlight the cell with green background (like your original TPM highlighting)
                sheet.highlight_cells(row=row_idx, column=tpm_col_idx, bg="#C6EFCE", fg="black")

    def on_window_close(self):
        """Handle window close event with auto-save and sample image transfer."""
        self.log("Window close event triggered", "debug")

        # Cancel auto-save timer
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)

        # Only handle unsaved changes if result is not already set
        if self.result is None:
            # Auto-save if there are unsaved changes
            if self.has_unsaved_changes:
                if messagebox.askyesno("Save Changes",
                                     "You have unsaved changes. Save before closing?"):
                    try:
                        self.save_data(show_confirmation=False)
                        self.result = "load_file"
                    except Exception as e:
                        messagebox.showerror("Save Error", f"Failed to save: {e}")
                        return  # Don't close if save failed
                else:
                    self.result = "cancel"
            else:
                self.result = "load_file" if self.last_save_time else "cancel"

        # NEW: Transfer sample images to main GUI for display
        self._transfer_sample_images_to_main_gui()

        self._transfer_sample_notes_to_main_gui()

        # Show the main GUI window again before destroying
        if hasattr(self.parent, 'root') and self.main_window_was_visible:
            debug_print("DEBUG: Restoring main GUI window from data collection window")
            self.parent.root.deiconify()  # Show main window
            self.parent.root.state('zoomed')
            self.parent.root.lift()  # Bring to front
            self.parent.root.focus_set()  # Give focus to main window

            debug_print("DEBUG: Main GUI window restored")

        self.window.destroy()

    def _save_current_notes_before_tab_switch(self):
        """Save the current notes before switching tabs."""
        if hasattr(self, 'sample_notes_text') and self.sample_notes_text.winfo_exists():
            try:
                # Get current notes content
                current_notes = self.sample_notes_text.get('1.0', 'end-1c')

                # Get current tab index
                current_tab = self.notebook.index(self.notebook.select())
                current_sample_id = f"Sample {current_tab + 1}"

                # Save notes to data structure
                self.data[current_sample_id]["sample_notes"] = current_notes
                debug_print(f"DEBUG: Saved notes for {current_sample_id} before tab switch: '{current_notes[:30]}...'")

                # Force the text widget to lose focus to ensure any pending changes are captured
                self.sample_notes_text.selection_clear()
                self.window.focus_set()
                self.window.update_idletasks()

            except Exception as e:
                debug_print(f"DEBUG: Error saving notes before tab switch: {e}")

    def recalculate_all_tpm(self):
        """Recalculate TPM for all samples."""
        debug_print("DEBUG: Recalculating TPM for all samples")
        for i in range(self.num_samples):
            sample_id = f"Sample {i + 1}"
            self.calculate_tpm(sample_id)

        self.update_stats_panel()
        self.mark_unsaved_changes()
        show_success_message("Recalculation Complete", "TPM values have been recalculated for all samples.", self.window)

    def get_sample_images_summary(self):
        """Get a summary of all loaded sample images."""
        summary = {}
        total_images = 0

        for sample_id, images in self.sample_images.items():
            summary[sample_id] = len(images)
            total_images += len(images)

        debug_print(f"DEBUG: Sample images summary: {summary} (Total: {total_images})")
        return summary, total_images

    def _transfer_sample_notes_to_main_gui(self):
        """Transfer sample notes to main GUI with proper structure for display."""
        try:
            debug_print("DEBUG: Transferring sample notes to main GUI")

            # Ensure we save current notes before transfer
            self._save_current_notes_before_tab_switch()

            # Check if we have sample notes to transfer
            if not hasattr(self, 'data') or not self.data:
                debug_print("DEBUG: No sample data to transfer notes from")
                return

            debug_print(f"DEBUG: Processing notes for {len(self.data)} samples")

            # Collect all sample notes
            sample_notes_data = {}
            notes_found = False

            for sample_id, sample_data in self.data.items():
                sample_notes = sample_data.get("sample_notes", "")
                if sample_notes.strip():  # Only include non-empty notes
                    sample_notes_data[sample_id] = sample_notes
                    notes_found = True
                    debug_print(f"DEBUG: Found notes for {sample_id}: '{sample_notes[:50]}...'")

            if not notes_found:
                debug_print("DEBUG: No sample notes found to transfer")
                return

            # Ensure header_data structure exists and is properly formatted
            if not hasattr(self, 'header_data') or not self.header_data:
                self.header_data = {'samples': []}

            # Ensure header_data has enough sample entries
            while len(self.header_data['samples']) < self.num_samples:
                self.header_data['samples'].append({})

            # Update header_data with current notes
            for i in range(self.num_samples):
                sample_id = f"Sample {i + 1}"
                if sample_id in sample_notes_data:
                    self.header_data['samples'][i]['sample_notes'] = sample_notes_data[sample_id]
                    debug_print(f"DEBUG: Updated header_data with notes for {sample_id}")

            # Store sample notes in parent for main GUI processing
            if sample_notes_data:
                self.parent.pending_sample_notes = {
                    'test_name': self.test_name,
                    'header_data': self.header_data.copy(),
                    'notes_data': sample_notes_data.copy()
                }
                debug_print(f"DEBUG: Stored sample notes for main GUI - test: {self.test_name}")

                # Store notes metadata for reverse lookup (similar to images)
                if not hasattr(self.parent, 'sample_notes_metadata'):
                    self.parent.sample_notes_metadata = {}
                if self.parent.current_file not in self.parent.sample_notes_metadata:
                    self.parent.sample_notes_metadata[self.parent.current_file] = {}

                self.parent.sample_notes_metadata[self.parent.current_file][self.test_name] = {
                    'header_data': self.header_data.copy(),
                    'notes_data': sample_notes_data.copy(),
                    'test_name': self.test_name
                }
                debug_print(f"DEBUG: Stored sample notes metadata for reverse lookup")

                # If parent has a method to immediately process notes, call it
                if hasattr(self.parent, 'process_pending_sample_notes'):
                    self.parent.process_pending_sample_notes()

        except Exception as e:
            debug_print(f"ERROR: Failed to transfer sample notes to main GUI: {e}")
            import traceback
            traceback.print_exc()

    def _transfer_sample_images_to_main_gui(self):
        """Transfer sample images to main GUI with proper labeling for display."""
        try:
            debug_print("DEBUG: Transferring sample images to main GUI")

            # Check if we have sample images to transfer
            if not hasattr(self, 'sample_images') or not self.sample_images:
                debug_print("DEBUG: No sample images to transfer")
                return

            debug_print(f"DEBUG: Processing {len(self.sample_images)} sample groups for main GUI")

            # Create formatted images for main GUI display
            formatted_images = []

            for sample_id, image_paths in self.sample_images.items():
                try:
                    # Extract sample index from sample_id (e.g., "Sample 1" -> 0)
                    sample_index = int(sample_id.split()[-1]) - 1

                    if sample_index < len(self.header_data['samples']):
                        sample_info = self.header_data['samples'][sample_index]

                        # Create labels for each image with comprehensive information
                        for img_path in image_paths:
                            # Create descriptive label: "Sample 1 - Test Name - Media - Viscosity - Date"
                            label_parts = [
                                sample_info.get('id', sample_id),
                                self.test_name,
                                sample_info.get('media', 'Unknown Media'),
                                f"{sample_info.get('viscosity', 'Unknown')} cP",
                                datetime.datetime.now().strftime("%Y-%m-%d")
                            ]

                            # Filter out empty parts and join
                            formatted_label = " - ".join(filter(lambda x: x and str(x).strip(), label_parts))

                            formatted_images.append({
                                'path': img_path,
                                'label': formatted_label,
                                'sample_id': sample_id,
                                'sample_info': sample_info,
                                'crop_state': getattr(self, 'sample_image_crop_states', {}).get(img_path, False)
                            })

                            debug_print(f"DEBUG: Created formatted image: {formatted_label}")

                except (ValueError, IndexError) as e:
                    debug_print(f"DEBUG: Error processing sample {sample_id}: {e}")
                    continue

            # Store formatted images in parent for main GUI processing
            if formatted_images:
                self.parent.pending_formatted_images = formatted_images
                # Store the target sheet information
                self.parent.pending_images_target_sheet = self.test_name
                debug_print(f"DEBUG: Stored {len(formatted_images)} formatted images for main GUI")
                debug_print(f"DEBUG: Target sheet stored as: {self.test_name}")

                # Store sample-specific image metadata for reverse lookup
                if not hasattr(self.parent, 'sample_image_metadata'):
                    self.parent.sample_image_metadata = {}
                if self.parent.current_file not in self.parent.sample_image_metadata:
                    self.parent.sample_image_metadata[self.parent.current_file] = {}
                if self.test_name not in self.parent.sample_image_metadata[self.parent.current_file]:
                    self.parent.sample_image_metadata[self.parent.current_file][self.test_name] = {}

                # Store the sample-to-image mapping for later retrieval
                self.parent.sample_image_metadata[self.parent.current_file][self.test_name] = {
                    'sample_images': self.sample_images.copy(),
                    'sample_image_crop_states': getattr(self, 'sample_image_crop_states', {}).copy(),
                    'header_data': self.header_data.copy(),
                    'test_name': self.test_name
                }
                debug_print(f"DEBUG: Stored sample metadata for reverse lookup")

                # If parent has a method to immediately process these, call it
                if hasattr(self.parent, 'process_pending_sample_images'):
                    self.parent.process_pending_sample_images()

        except Exception as e:
            debug_print(f"ERROR: Failed to transfer sample images to main GUI: {e}")
            import traceback
            traceback.print_exc()


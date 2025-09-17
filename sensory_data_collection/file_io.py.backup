"""
File I/O Manager
Handles saving and loading of sensory sessions and Excel export
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import json
import pandas as pd
import os
from datetime import datetime
from utils import debug_print, show_success_message


class FileIOManager:
    """Manages file operations for sensory data collection."""
    def __init__(self, sensory_window, session_manager, sample_manager):
        """Initialize the file I/O manager with reference to main window."""
        self.sensory_window = sensory_window
        self.session_manager = session_manager
        self.sample_manager = sample_manager
        
    def save_session(self):
        """Save the current session to a JSON file."""
        if not self.session_manager.current_session_id or not self.session_manager.sessions:
            messagebox.showwarning("Warning", "No session to save!")
            return

        # Make sure current samples are saved to the session
        if self.session_manager.current_session_id in self.session_manager.sessions:
            self.session_manager.sessions[self.session_manager.current_session_id]['samples'] = self.sensory_window.samples
            self.session_manager.sessions[self.session_manager.current_session_id]['header'] = {field: var.get() for field, var in self.sensory_window.header_vars.items()}

        current_session = self.session_manager.sessions[self.session_manager.current_session_id]

        if not current_session['samples']:
            messagebox.showwarning("Warning", "No sample data to save!")
            return

        # Default filename based on session name and assessor
        assessor_name = current_session['header'].get('Assessor Name', 'Unknown')
        safe_assessor = "".join(c for c in assessor_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_session = "".join(c for c in self.session_manager.current_session_id if c.isalnum() or c in (' ', '-', '_')).strip()

        default_filename = f"{safe_assessor}_{safe_session}_sensory_session.json"

        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Sensory Session",
            initialfile=default_filename
        )

        if filename:
            try:
                # Create session data with updated timestamp
                session_data = {
                    'header': current_session['header'],
                    'samples': current_session['samples'],
                    'timestamp': datetime.now().isoformat(),
                    'session_name': self.session_manager.current_session_id,
                    'source_file': current_session.get('source_file', ''),
                    'source_image': current_session.get('source_image', '')
                }

                with open(filename, 'w') as f:
                    json.dump(session_data, f, indent=2)

                debug_print(f"DEBUG: Saved session {self.session_manager.current_session_id} to {filename}")
                show_success_message("Success",
                                  f"Session '{self.session_manager.current_session_id}' saved to {os.path.basename(filename)}\n"
                                  f"Saved {len(current_session['samples'])} samples", self.sensory_window.window)
                debug_print(f"Saved sensory session to: {filename}")

            except Exception as e:
                debug_print(f"DEBUG: Error saving session: {e}")
                messagebox.showerror("Error", f"Failed to save session: {e}")

        self.sensory_window.mode_manager.bring_to_front()

    def load_session(self):
        """Load one or more sessions from JSON files as new sessions."""
        debug_print("DEBUG: Starting load session with multiple file selection")

        # Use askopenfilenames to allow multiple file selection
        filenames = filedialog.askopenfilenames(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Load Sensory Sessions (Hold Ctrl to select multiple)"
        )

        if not filenames:
            debug_print("DEBUG: No files selected for loading")
            return

        debug_print(f"DEBUG: Selected {len(filenames)} files for loading")

        successful_loads = 0
        failed_loads = []
        loaded_session_names = []

        for filename in filenames:
            try:
                debug_print(f"DEBUG: Processing file: {filename}")

                with open(filename, 'r') as f:
                    session_data = json.load(f)

                debug_print(f"DEBUG: Loading session from {filename}")
                debug_print(f"DEBUG: Session data keys: {list(session_data.keys())}")

                # Validate session data
                if not self.session_manager.validate_session_data(session_data):
                    debug_print(f"DEBUG: Invalid session data in {filename}")
                    failed_loads.append(f"{os.path.basename(filename)} - Invalid format")
                    continue

                # Create session name from filename
                base_filename = os.path.splitext(os.path.basename(filename))[0]
                session_name = base_filename

                # Ensure unique session name
                counter = 1
                original_name = session_name
                while session_name in self.session_manager.sessions:
                    session_name = f"{original_name}_{counter}"
                    counter += 1

                debug_print(f"DEBUG: Creating new session: {session_name}")

                # Create new session with loaded data
                self.session_manager.sessions[session_name] = {
                    'header': session_data.get('header', {}),
                    'samples': session_data.get('samples', {}),
                    'timestamp': session_data.get('timestamp', datetime.now().isoformat()),
                    'source_file': filename
                }

                debug_print(f"DEBUG: Session created with {len(self.session_manager.sessions[session_name]['samples'])} samples")
                successful_loads += 1
                loaded_session_names.append(session_name)

            except Exception as e:
                debug_print(f"DEBUG: Error loading session from {filename}: {e}")
                import traceback
                traceback.print_exc()
                failed_loads.append(f"{os.path.basename(filename)} - {str(e)}")

        # Report results to user
        if successful_loads > 0:
            # Switch to the last loaded session
            last_session = loaded_session_names[-1]
            self.session_manager.switch_to_session(last_session)

            # Update session selector UI
            self.session_manager.update_session_combo()
            if hasattr(self, 'session_var'):
                self.session_var.set(last_session)

            # Update other UI components
            self.sample_manager.update_sample_combo()
            self.sample_manager.update_sample_checkboxes()

            # Select first sample if available
            if self.sensory_window.samples:
                first_sample = list(self.sensory_window.samples.keys())[0]
                self.sample_var.set(first_sample)
                self.sample_manager.load_sample_data(first_sample)
            else:
                self.sample_var.set('')
                self.sample_manager.clear_form()

            self.plot_manager.update_plot()

            # Create success message
            success_msg = f"Successfully loaded {successful_loads} session(s):\n"
            success_msg += "\n".join([f"• {name}" for name in loaded_session_names])
            success_msg += f"\n\nCurrently viewing: {last_session}"
            success_msg += "\nUse session selector to switch between sessions."

            debug_print(f"DEBUG: Successfully loaded {successful_loads} sessions")
            show_success_message("Sessions Loaded", success_msg, self.sensory_window.window)

        # Report any failures
        if failed_loads:
            failure_msg = f"Failed to load {len(failed_loads)} file(s):\n"
            failure_msg += "\n".join([f"• {fail}" for fail in failed_loads])
            debug_print(f"DEBUG: Failed to load {len(failed_loads)} files")
            messagebox.showwarning("Load Errors", failure_msg)

        # Overall result
        if successful_loads == 0:
            messagebox.showerror("Load Failed", "No valid session files could be loaded.")
        else:
            debug_print(f"DEBUG: Load session completed - {successful_loads} successful, {len(failed_loads)} failed")

        self.sensory_window.mode_manager.bring_to_front()

    def export_to_excel(self):
        """Export the sensory data to an Excel file."""
        if not self.sensory_window.samples:
            messagebox.showwarning("Warning", "No data to export!")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Sensory Data"
        )

        if filename:
            try:
                # Create DataFrame for sensory data
                data_rows = []
                for sample_name, sample_data in self.sensory_window.samples.items():
                    row = {'Sample': sample_name}

                    # Add header information
                    for field, var in self.sensory_window.header_vars.items():
                        row[field] = var.get()

                    # Add sensory ratings
                    for metric in self.sensory_window.metrics:
                        row[metric] = sample_data.get(metric, 0)

                    row['Comments'] = sample_data.get('comments', '')
                    data_rows.append(row)

                df = pd.DataFrame(data_rows)

                # Save to Excel
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sensory Data', index=False)

                    # Save spider plot as image
                    plot_filename = filename.replace('.xlsx', '_spider_plot.png')
                    self.fig.savefig(plot_filename, dpi=300, bbox_inches='tight')

                show_success_message("Success", f"Data exported to {filename}\nSpider plot saved as {plot_filename}", self.sensory_window.window)
                debug_print(f"Exported sensory data to: {filename}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {e}")

        self.sensory_window.mode_manager.bring_to_front()

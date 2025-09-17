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
    
    def __init__(self, sensory_window):
        """Initialize the file I/O manager with reference to main window."""
        self.sensory_window = sensory_window
        
    def save_session(self):
        """Save the current session to a JSON file."""
        if not self.current_session_id or not self.sessions:
            messagebox.showwarning("Warning", "No session to save!")
            return

        # Make sure current samples are saved to the session
        if self.current_session_id in self.sessions:
            self.sessions[self.current_session_id]['samples'] = self.samples
            self.sessions[self.current_session_id]['header'] = {field: var.get() for field, var in self.header_vars.items()}

        current_session = self.sessions[self.current_session_id]

        if not current_session['samples']:
            messagebox.showwarning("Warning", "No sample data to save!")
            return

        # Default filename based on session name and assessor
        assessor_name = current_session['header'].get('Assessor Name', 'Unknown')
        safe_assessor = "".join(c for c in assessor_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_session = "".join(c for c in self.current_session_id if c.isalnum() or c in (' ', '-', '_')).strip()

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
                    'session_name': self.current_session_id,
                    'source_file': current_session.get('source_file', ''),
                    'source_image': current_session.get('source_image', '')
                }

                with open(filename, 'w') as f:
                    json.dump(session_data, f, indent=2)

                debug_print(f"DEBUG: Saved session {self.current_session_id} to {filename}")
                show_success_message("Success",
                                  f"Session '{self.current_session_id}' saved to {os.path.basename(filename)}\n"
                                  f"Saved {len(current_session['samples'])} samples", self.window)
                debug_print(f"Saved sensory session to: {filename}")

            except Exception as e:
                debug_print(f"DEBUG: Error saving session: {e}")
                messagebox.showerror("Error", f"Failed to save session: {e}")

        self.bring_to_front()

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
                if not self.validate_session_data(session_data):
                    debug_print(f"DEBUG: Invalid session data in {filename}")
                    failed_loads.append(f"{os.path.basename(filename)} - Invalid format")
                    continue

                # Create session name from filename
                base_filename = os.path.splitext(os.path.basename(filename))[0]
                session_name = base_filename

                # Ensure unique session name
                counter = 1
                original_name = session_name
                while session_name in self.sessions:
                    session_name = f"{original_name}_{counter}"
                    counter += 1

                debug_print(f"DEBUG: Creating new session: {session_name}")

                # Create new session with loaded data
                self.sessions[session_name] = {
                    'header': session_data.get('header', {}),
                    'samples': session_data.get('samples', {}),
                    'timestamp': session_data.get('timestamp', datetime.now().isoformat()),
                    'source_file': filename
                }

                debug_print(f"DEBUG: Session created with {len(self.sessions[session_name]['samples'])} samples")
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
            self.switch_to_session(last_session)

            # Update session selector UI
            self.update_session_combo()
            if hasattr(self, 'session_var'):
                self.session_var.set(last_session)

            # Update other UI components
            self.update_sample_combo()
            self.update_sample_checkboxes()

            # Select first sample if available
            if self.samples:
                first_sample = list(self.samples.keys())[0]
                self.sample_var.set(first_sample)
                self.load_sample_data(first_sample)
            else:
                self.sample_var.set('')
                self.clear_form()

            self.update_plot()

            # Create success message
            success_msg = f"Successfully loaded {successful_loads} session(s):\n"
            success_msg += "\n".join([f"• {name}" for name in loaded_session_names])
            success_msg += f"\n\nCurrently viewing: {last_session}"
            success_msg += "\nUse session selector to switch between sessions."

            debug_print(f"DEBUG: Successfully loaded {successful_loads} sessions")
            show_success_message("Sessions Loaded", success_msg, self.window)

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

        self.bring_to_front()

    def export_to_excel(self):
        """Export the sensory data to an Excel file."""
        if not self.samples:
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
                for sample_name, sample_data in self.samples.items():
                    row = {'Sample': sample_name}

                    # Add header information
                    for field, var in self.header_vars.items():
                        row[field] = var.get()

                    # Add sensory ratings
                    for metric in self.metrics:
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

                show_success_message("Success", f"Data exported to {filename}\nSpider plot saved as {plot_filename}", self.window)
                debug_print(f"Exported sensory data to: {filename}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {e}")

        self.bring_to_front()

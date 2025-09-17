"""
Session Manager
Handles all session-related operations including creation, switching, merging, and management
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import os
from datetime import datetime
from utils import debug_print, show_success_message, FONT


class SessionManager:
    """Manages sensory data collection sessions."""
    
    def __init__(self, sensory_window, sample_manager=None, plot_manager=None):
        """Initialize the session manager with reference to main window."""
        self.sensory_window = sensory_window
        self.sample_manager = sample_manager
        self.plot_manager = plot_manager
        self.sessions = {}
        self.current_session_id = None
        self.session_counter = 1
        
    def create_new_session(self, session_name=None, source_image=None):
        """Create a new session for data collection."""
        if session_name is None:
            session_name = f"Session_{self.session_counter}"
            self.session_counter += 1

        debug_print(f"DEBUG: Creating new session: {session_name}")
        debug_print(f"DEBUG: Source image: {source_image}")

        # Create new session structure
        self.sessions[session_name] = {
            'header': {field: var.get() for field, var in self.sensory_window.header_vars.items()},
            'samples': {},
            'timestamp': datetime.now().isoformat(),
            'source_image': source_image or ''
        }

        # Switch to new session
        self.current_session_id = session_name
        self.sensory_window.samples = self.sessions[session_name]['samples']

        debug_print(f"DEBUG: Session created successfully")
        debug_print(f"DEBUG: Current session ID: {self.current_session_id}")
        debug_print(f"DEBUG: Session structure: {self.sessions[session_name]}")

        self.update_session_combo()
        self.sample_manager.update_sample_combo()
        self.sample_manager.update_sample_checkboxes()

        return session_name

    def switch_to_session(self, session_id):
        """Switch to a specific session."""
        if session_id not in self.sessions:
            debug_print(f"DEBUG: Session {session_id} not found")
            return False

        debug_print(f"DEBUG: Switching from session {self.current_session_id} to {session_id}")

        # Save current session data before switching
        if self.current_session_id and self.current_session_id in self.sessions:
            self.sessions[self.current_session_id]['samples'] = self.sensory_window.samples
            self.sessions[self.current_session_id]['header'] = {field: var.get() for field, var in self.sensory_window.header_vars.items()}
            debug_print(f"DEBUG: Saved {len(self.sensory_window.samples)} samples to previous session")

        # Switch to new session
        self.current_session_id = session_id
        self.sensory_window.samples = self.sessions[session_id]['samples']

        # Update header fields with session data
        session_header = self.sessions[session_id]['header']
        for field, var in self.sensory_window.header_vars.items():
            if field in session_header:
                var.set(session_header[field])
            else:
                if field == "Date":
                    var.set(datetime.now().strftime("%Y-%m-%d"))
                else:
                    var.set('')

        debug_print(f"DEBUG: Switched to session {session_id} with {len(self.sensory_window.samples)} samples")

        # Update UI components
        self.sample_manager.update_sample_combo()
        self.sample_manager.update_sample_checkboxes()

        # Select first sample if available
        if self.sensory_window.samples:
            first_sample = list(self.sensory_window.samples.keys())[0]
            self.sensory_window.sample_var.set(first_sample)
            self.sample_manager.load_sample_data(first_sample)
            self.refresh_value_displays()
        else:
            self.sensory_window.sample_var.set('')
            self.sample_manager.clear_form()

        self.plot_manager.update_plot()
        debug_print("DEBUG: Session switch completed with display refresh")
        return True

    def update_session_combo(self):
        """Update the session selection combo box."""
        session_names = list(self.sessions.keys())
        if hasattr(self, 'session_combo'):
            self.sensory_window.session_combo['values'] = session_names

    def setup_session_selector(self, parent_frame):
        """Add session selector to the interface."""
        # Add session selector frame with reduced width
        session_frame = ttk.LabelFrame(parent_frame, text="Session Management", padding=10)
        session_frame.pack(fill='x', padx=5, pady=5)

        # Configure session_frame for centered grid layout
        session_frame.grid_columnconfigure(0, weight=1)
        session_frame.grid_columnconfigure(1, weight=1)
        debug_print("DEBUG: Configured session_frame for centered layout")

        # Top row - Session selection centered
        top_frame = ttk.Frame(session_frame)
        top_frame.grid(row=0, column=0, columnspan=2, pady=(0, 5))

        session_label = ttk.Label(top_frame, text="Current Session:", font=FONT)
        session_label.pack(side='left', padx=(0, 5))

        self.sensory_window.session_var = tk.StringVar()
        self.sensory_window.session_combo = ttk.Combobox(top_frame, textvariable=self.sensory_window.session_var,
                                         font=FONT, state='readonly', width=15)
        self.sensory_window.session_combo.pack(side='left', padx=(0, 10))
        self.sensory_window.session_combo.bind('<<ComboboxSelected>>', self.on_session_selected)
        debug_print("DEBUG: Session dropdown centered on top row")

        # Second row - Session management buttons centered
        button_frame = ttk.Frame(session_frame)
        button_frame.grid(row=1, column=0, columnspan=2)

        ttk.Button(button_frame, text="New Session",
                   command=self.add_new_session).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Combine Sessions",
                   command=self.show_combine_sessions_dialog).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Delete Session",
                   command=self.delete_current_session).pack(side='left', padx=2)
        debug_print("DEBUG: Session management buttons centered on second row")

        debug_print("DEBUG: Session selector UI setup complete with centered layout")

    def on_session_selected(self, event=None):
        """Handle session selection change."""
        selected_session = self.sensory_window.session_var.get()
        if selected_session and selected_session != self.current_session_id:
            debug_print(f"DEBUG: Session selection changed to: {selected_session}")
            self.switch_to_session(selected_session)

    def add_new_session(self):
        """Add a new empty session."""
        session_name = tk.simpledialog.askstring("New Session", "Enter session name:")
        if session_name and session_name.strip():
            session_name = session_name.strip()
            if session_name in self.sessions:
                messagebox.showerror("Session Exists", f"Session '{session_name}' already exists.")
                return

            debug_print(f"DEBUG: Creating new session: {session_name}")
            self.create_new_session(session_name)
            self.sensory_window.session_var.set(session_name)
            show_success_message("Success", f"Created new session: {session_name}", self.sensory_window.window)

    def delete_current_session(self):
        """Delete the current session."""
        if not self.current_session_id:
            messagebox.showwarning("No Session", "No session selected to delete.")
            return

        if len(self.sessions) <= 1:
            messagebox.showwarning("Cannot Delete", "Cannot delete the last session.")
            return

        if messagebox.askyesno("Confirm Delete",
                              f"Delete session '{self.current_session_id}'?\n"
                              f"This will permanently remove all data in this session."):

            session_to_delete = self.current_session_id
            debug_print(f"DEBUG: Deleting session: {session_to_delete}")

            # Switch to another session first
            remaining_sessions = [s for s in self.sessions.keys() if s != session_to_delete]
            if remaining_sessions:
                self.switch_to_session(remaining_sessions[0])

            # Delete the session
            del self.sessions[session_to_delete]
            self.update_session_combo()

            debug_print(f"DEBUG: Session {session_to_delete} deleted successfully")
            show_success_message("Success", f"Session '{session_to_delete}' deleted.", self.sensory_window.window)

    def show_combine_sessions_dialog(self):
        """Show dialog to select and combine multiple sessions."""
        if len(self.sessions) < 2:
            show_success_message("Insufficient Sessions",
                              "Need at least 2 sessions to combine.", self.sensory_window.window)
            return

        # Create dialog window
        combine_window = tk.Toplevel(self.sensory_window.window)
        combine_window.title("Combine Sessions")
        combine_window.geometry("400x300")
        combine_window.transient(self.sensory_window.window)
        combine_window.grab_set()

        # Instructions
        ttk.Label(combine_window,
                 text="Select sessions to combine into a new session:",
                 font=FONT).pack(pady=10)

        # Session selection with checkboxes
        selection_frame = ttk.Frame(combine_window)
        selection_frame.pack(fill='both', expand=True, padx=20, pady=10)

        session_vars = {}
        for session_id in self.sessions.keys():
            var = tk.BooleanVar()
            session_vars[session_id] = var

            # Create checkbox with session info
            sample_count = len(self.sessions[session_id]['samples'])
            source_image = self.sessions[session_id].get('source_image', 'Manual')
            source_name = os.path.basename(source_image) if source_image else 'Manual'

            checkbox_text = f"{session_id} ({sample_count} samples) - {source_name}"
            ttk.Checkbutton(selection_frame, text=checkbox_text,
                           variable=var).pack(anchor='w', pady=2)

        # New session name
        name_frame = ttk.Frame(combine_window)
        name_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(name_frame, text="New session name:").pack(side='left')
        name_var = tk.StringVar(value=f"Combined_Session_{self.session_counter}")
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=20)
        name_entry.pack(side='right')

        # Buttons
        button_frame = ttk.Frame(combine_window)
        button_frame.pack(fill='x', padx=20, pady=20)

        def combine_selected_sessions():
            selected_sessions = [sid for sid, var in session_vars.items() if var.get()]
            new_session_name = name_var.get().strip()

            debug_print(f"DEBUG: Combining sessions: {selected_sessions}")
            debug_print(f"DEBUG: New session name: {new_session_name}")

            if len(selected_sessions) < 2:
                messagebox.showwarning("Insufficient Selection",
                                     "Select at least 2 sessions to combine.")
                return

            if not new_session_name:
                messagebox.showwarning("Invalid Name", "Enter a name for the new session.")
                return

            if new_session_name in self.sessions:
                messagebox.showerror("Name Exists",
                                   f"Session '{new_session_name}' already exists.")
                return

            # Combine sessions
            combined_samples = {}
            total_sample_count = 0

            for session_id in selected_sessions:
                session_samples = self.sessions[session_id]['samples']
                for sample_name, sample_data in session_samples.items():
                    # Create unique sample name if conflict
                    unique_name = sample_name
                    counter = 1
                    while unique_name in combined_samples:
                        unique_name = f"{sample_name}_{counter}"
                        counter += 1

                    combined_samples[unique_name] = sample_data
                    total_sample_count += 1
                    debug_print(f"DEBUG: Added sample {unique_name} from session {session_id}")

            # Create new combined session
            self.create_new_session(new_session_name)
            self.sessions[new_session_name]['samples'] = combined_samples
            self.sensory_window.samples = combined_samples

            # Update UI
            self.sensory_window.session_var.set(new_session_name)
            self.sample_manager.update_sample_combo()
            self.sample_manager.update_sample_checkboxes()
            self.plot_manager.update_plot()

            combine_window.destroy()

            debug_print(f"DEBUG: Successfully combined {len(selected_sessions)} sessions")
            debug_print(f"DEBUG: New session has {total_sample_count} samples")

            show_success_message("Success",
                              f"Combined {len(selected_sessions)} sessions into '{new_session_name}'!\n"
                              f"Total samples: {total_sample_count}", self.sensory_window.window)

        ttk.Button(button_frame, text="Combine Sessions",
                   command=combine_selected_sessions).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel",
                   command=combine_window.destroy).pack(side='right', padx=5)

        debug_print("DEBUG: Combine sessions dialog created")

    def merge_sessions_from_files(self):
        """Merge multiple session JSON files into a new session."""

        debug_print("DEBUG: Starting merge sessions from files")

        # Select multiple JSON files
        filenames = filedialog.askopenfilenames(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Select Session Files to Merge (Hold Ctrl to select multiple)"
        )

        if not filenames or len(filenames) < 2:
            show_success_message("Insufficient Files",
                              "Please select at least 2 session files to merge.", self.sensory_window.window)
            return

        debug_print(f"DEBUG: Selected {len(filenames)} files for merging")

        # Load and validate all session files
        loaded_sessions = {}
        failed_files = []

        for filename in filenames:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)

                # Validate session data structure
                if self.validate_session_data(session_data):
                    session_name = os.path.splitext(os.path.basename(filename))[0]
                    loaded_sessions[session_name] = {
                        'file_path': filename,
                        'data': session_data
                    }
                    debug_print(f"DEBUG: Successfully loaded session from {filename}")
                else:
                    failed_files.append(filename)
                    debug_print(f"DEBUG: Invalid session format in {filename}")

            except Exception as e:
                failed_files.append(filename)
                debug_print(f"DEBUG: Failed to load {filename}: {e}")

        if not loaded_sessions:
            messagebox.showerror("Load Error",
                               "No valid session files could be loaded.\n"
                               "Ensure files are in the correct JSON format.")
            return

        if failed_files:
            failed_list = '\n'.join([os.path.basename(f) for f in failed_files])
            messagebox.showwarning("Some Files Failed",
                                 f"Failed to load {len(failed_files)} files:\n{failed_list}\n\n"
                                 f"Continuing with {len(loaded_sessions)} valid files.")

        # Show merge configuration dialog
        self.show_merge_sessions_dialog(loaded_sessions)

    def validate_session_data(self, session_data):
        """Validate that the JSON file has the correct session format."""

        debug_print("DEBUG: Validating session data structure")

        if not isinstance(session_data, dict):
            debug_print("DEBUG: Session data is not a dictionary")
            return False

        # Check for required top-level keys
        required_keys = ['samples']
        if not all(key in session_data for key in required_keys):
            debug_print(f"DEBUG: Missing required keys. Found: {list(session_data.keys())}")
            return False

        # Check samples structure
        samples = session_data.get('samples', {})
        if not isinstance(samples, dict):
            debug_print("DEBUG: Samples is not a dictionary")
            return False

        # Validate sample data structure
        for sample_name, sample_data in samples.items():
            if not isinstance(sample_data, dict):
                debug_print(f"DEBUG: Sample {sample_name} data is not a dictionary")
                return False

            # Check for expected metrics (at least some should be present)
            metrics_found = sum(1 for metric in self.sensory_window.metrics if metric in sample_data)
            if metrics_found == 0:
                debug_print(f"DEBUG: No valid metrics found in sample {sample_name}")
                return False


        return True

    def show_merge_sessions_dialog(self, loaded_sessions):
        """Show dialog to configure session merging."""

        # Create dialog window
        merge_window = tk.Toplevel(self.sensory_window.window)
        merge_window.title("Merge Session Files")
        merge_window.geometry("600x500")
        merge_window.transient(self.sensory_window.window)
        merge_window.grab_set()

        # Main frame
        main_frame = ttk.Frame(merge_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Title
        title_label = ttk.Label(main_frame,
                               text="Configure Session Merge",
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 10))

        # Instructions
        instructions = ("Select which sessions to include in the merge.\n"
                       "Conflicting sample names will be automatically renamed.\n"
                       "Header information will be merged where possible.")

        ttk.Label(main_frame, text=instructions,
                 font=('Arial', 9), justify='left').pack(pady=(0, 15))

        # Session selection frame with scrollbar
        selection_frame = ttk.LabelFrame(main_frame, text="Sessions to Merge", padding=10)
        selection_frame.pack(fill='both', expand=True, pady=(0, 10))

        # Create scrollable frame
        canvas = tk.Canvas(selection_frame, height=200)
        scrollbar = ttk.Scrollbar(selection_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Session checkboxes with details
        session_vars = {}

        for session_name, session_info in loaded_sessions.items():
            session_data = session_info['data']
            sample_count = len(session_data.get('samples', {}))

            # Extract some header info for display
            header = session_data.get('header', {})
            assessor = header.get('Assessor Name', 'Unknown')
            date = header.get('Date', 'Unknown')

            var = tk.BooleanVar(value=True)  # Default to checked
            session_vars[session_name] = {
                'var': var,
                'data': session_data,
                'file_path': session_info['file_path']
            }

            # Create session info frame
            session_frame = ttk.Frame(scrollable_frame)
            session_frame.pack(fill='x', pady=2)

            # Checkbox and main info
            info_text = f"{session_name} ({sample_count} samples)"
            ttk.Checkbutton(session_frame, text=info_text,
                       variable=var).pack(anchor='w')

            # Additional details
            details = f"   Assessor: {assessor} | Date: {date} | File: {os.path.basename(session_info['file_path'])}"
            ttk.Label(session_frame, text=details,
                     font=('Arial', 8), foreground='gray').pack(anchor='w', padx=(20, 0))

        # Merge options frame
        options_frame = ttk.LabelFrame(main_frame, text="Merge Options", padding=10)
        options_frame.pack(fill='x', pady=(0, 10))

        # New session name
        name_frame = ttk.Frame(options_frame)
        name_frame.pack(fill='x', pady=(0, 5))

        ttk.Label(name_frame, text="Merged session name:").pack(side='left')
        default_name = f"Merged_Session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        name_var = tk.StringVar(value=default_name)
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=30)
        name_entry.pack(side='right')

        # Header merge strategy
        header_frame = ttk.Frame(options_frame)
        header_frame.pack(fill='x', pady=5)

        ttk.Label(header_frame, text="Header merge strategy:").pack(side='left')
        header_strategy = tk.StringVar(value="first")
        header_combo = ttk.Combobox(header_frame, textvariable=header_strategy,
                                   values=["first", "most_recent", "manual"],
                                   state='readonly', width=15)
        header_combo.pack(side='right')

        # Sample naming strategy
        naming_frame = ttk.Frame(options_frame)
        naming_frame.pack(fill='x', pady=5)

        ttk.Label(naming_frame, text="Duplicate sample naming:").pack(side='left')
        naming_strategy = tk.StringVar(value="auto_rename")
        naming_combo = ttk.Combobox(naming_frame, textvariable=naming_strategy,
                                   values=["auto_rename", "prefix_session", "skip_duplicates"],
                                   state='readonly', width=15)
        naming_combo.pack(side='right')

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))

        def perform_merge():
            debug_print("DEBUG: Starting merge process")

            # Get selected sessions
            selected_sessions = {}
            for session_name, session_info in session_vars.items():
                if session_info['var'].get():
                    selected_sessions[session_name] = session_info

            if len(selected_sessions) < 2:
                messagebox.showwarning("Insufficient Selection",
                                     "Please select at least 2 sessions to merge.")
                return

            new_session_name = name_var.get().strip()
            if not new_session_name:
                messagebox.showwarning("Invalid Name", "Please enter a name for the merged session.")
                return

            debug_print(f"DEBUG: Merging {len(selected_sessions)} sessions into '{new_session_name}'")

            # Perform the merge
            success = self.execute_session_merge(
                selected_sessions,
                new_session_name,
                header_strategy.get(),
                naming_strategy.get()
            )

            if success:
                merge_window.destroy()
                show_success_message("Merge Complete",
                                  f"Successfully merged {len(selected_sessions)} sessions!\n"
                                  f"New session: {new_session_name}", self.sensory_window.window)

        def select_all():
            for session_info in session_vars.values():
                session_info['var'].set(True)

        def select_none():
            for session_info in session_vars.values():
                session_info['var'].set(False)

        # Selection buttons
        select_frame = ttk.Frame(button_frame)
        select_frame.pack(side='left')

        ttk.Button(select_frame, text="Select All", command=select_all).pack(side='left', padx=2)
        ttk.Button(select_frame, text="Select None", command=select_none).pack(side='left', padx=2)

        # Action buttons
        action_frame = ttk.Frame(button_frame)
        action_frame.pack(side='right')

        ttk.Button(action_frame, text="Cancel",
                   command=merge_window.destroy).pack(side='left', padx=5)
        ttk.Button(action_frame, text="Merge Sessions",
                   command=perform_merge).pack(side='left', padx=5)

        debug_print("DEBUG: Merge dialog created successfully")

    def execute_session_merge(self, selected_sessions, new_session_name, header_strategy, naming_strategy):
        """Execute the actual merging of sessions."""

        debug_print(f"DEBUG: Executing merge with strategy - header: {header_strategy}, naming: {naming_strategy}")

        try:
            # Initialize merged session structure
            merged_session = {
                'header': {},
                'samples': {},
                'timestamp': datetime.now().isoformat(),
                'merge_info': {
                    'source_sessions': list(selected_sessions.keys()),
                    'merge_date': datetime.now().isoformat(),
                    'header_strategy': header_strategy,
                    'naming_strategy': naming_strategy
                }
            }

            # Merge headers based on strategy
            merged_session['header'] = self.merge_headers(selected_sessions, header_strategy)

            # Merge samples with conflict resolution
            merged_samples, conflicts_resolved = self.merge_samples(selected_sessions, naming_strategy)
            merged_session['samples'] = merged_samples

            # Create the new session in memory
            if hasattr(self, 'sessions'):
                # Using session-based structure
                self.sessions[new_session_name] = merged_session
                self.switch_to_session(new_session_name)
                if hasattr(self, 'session_var'):
                    self.sensory_window.session_var.set(new_session_name)
                self.update_session_combo()
            else:
                # Fallback to old structure
                self.sensory_window.samples = merged_samples

                # Update header fields
                for field, value in merged_session['header'].items():
                    if field in self.sensory_window.header_vars:
                        self.sensory_window.header_vars[field].set(value)

            # Update UI
            self.sample_manager.update_sample_combo()
            self.sample_manager.update_sample_checkboxes()

            # Select first sample if available
            if merged_samples:
                first_sample = list(merged_samples.keys())[0]
                self.sensory_window.sample_var.set(first_sample)
                self.sample_manager.load_sample_data(first_sample)

            self.plot_manager.update_plot()

            # Log merge details
            total_samples = len(merged_samples)
            total_sessions = len(selected_sessions)

            debug_print(f"DEBUG: Merge completed successfully")
            debug_print(f"DEBUG: Total samples: {total_samples}")
            debug_print(f"DEBUG: Conflicts resolved: {conflicts_resolved}")
            debug_print(f"DEBUG: Sessions merged: {total_sessions}")

            return True

        except Exception as e:
            debug_print(f"DEBUG: Merge execution failed: {e}")
            messagebox.showerror("Merge Error", f"Failed to merge sessions: {e}")
            return False

    def merge_headers(self, selected_sessions, strategy):
        """Merge header information based on the selected strategy."""

        debug_print(f"DEBUG: Merging headers with strategy: {strategy}")

        all_headers = []
        for session_name, session_info in selected_sessions.items():
            header = session_info['data'].get('header', {})
            if header:
                all_headers.append((session_name, header))

        if not all_headers:
            return {}

        if strategy == "first":
            return all_headers[0][1].copy()

        elif strategy == "most_recent":
            # Find header with most recent timestamp
            most_recent = None
            most_recent_time = None

            for session_name, header in all_headers:
                session_data = selected_sessions[session_name]['data']
                timestamp_str = session_data.get('timestamp', '')

                try:
                    if timestamp_str:
                        timestamp = datetime.fromisoformat(timestamp_str.replace('Z', '+00:00'))
                        if most_recent_time is None or timestamp > most_recent_time:
                            most_recent_time = timestamp
                            most_recent = header
                except:
                    pass

            return most_recent.copy() if most_recent else all_headers[0][1].copy()

        elif strategy == "manual":
            # For now, return first header - could be extended to show manual selection dialog
            return all_headers[0][1].copy()

        return {}

    def merge_samples(self, selected_sessions, naming_strategy):
        """Merge samples with conflict resolution."""

        debug_print(f"DEBUG: Merging samples with naming strategy: {naming_strategy}")

        merged_samples = {}
        conflicts_resolved = 0

        for session_name, session_info in selected_sessions.items():
            session_samples = session_info['data'].get('samples', {})

            for original_sample_name, sample_data in session_samples.items():
                final_sample_name = original_sample_name

                # Handle naming conflicts
                if final_sample_name in merged_samples:
                    conflicts_resolved += 1

                    if naming_strategy == "auto_rename":
                        counter = 1
                        while f"{original_sample_name}_{counter}" in merged_samples:
                            counter += 1
                        final_sample_name = f"{original_sample_name}_{counter}"

                    elif naming_strategy == "prefix_session":
                        final_sample_name = f"{session_name}_{original_sample_name}"
                        counter = 1
                        while final_sample_name in merged_samples:
                            final_sample_name = f"{session_name}_{original_sample_name}_{counter}"
                            counter += 1

                    elif naming_strategy == "skip_duplicates":
                        debug_print(f"DEBUG: Skipping duplicate sample: {original_sample_name}")
                        continue

                # Copy sample data
                merged_samples[final_sample_name] = sample_data.copy()

                debug_print(f"DEBUG: Added sample {original_sample_name} as {final_sample_name}")

        debug_print(f"DEBUG: Sample merge complete - {len(merged_samples)} total samples, {conflicts_resolved} conflicts resolved")

        return merged_samples, conflicts_resolved
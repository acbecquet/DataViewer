"""
Mode Manager
Handles switching between collection mode and comparison mode
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime
from utils import debug_print, show_success_message


class ModeManager:
    """Manages collection vs comparison mode switching and operations."""
    
    def __init__(self, sensory_window, session_manager, plot_manager):
        """Initialize the mode manager with reference to main window."""
        self.sensory_window = sensory_window
        self.session_manager = session_manager
        self.plot_manager = plot_manager
        
    def toggle_mode(self):
        """Toggle between collection mode and comparison mode."""
        if self.sensory_window.current_mode == "collection":
            # Switch to comparison mode
            self.switch_to_comparison_mode()
        else:
            # Switch to collection mode
            self.switch_to_collection_mode()

    def switch_to_comparison_mode(self):
        """Switch to comparison mode - show averages across users."""
        debug_print("DEBUG: Switching to comparison mode")

        self.sensory_window.current_mode = "comparison"
        self.sensory_window.mode_button.config(text="Switch to Collection Mode")

        # Add comparison title
        self.setup_comparison_title()

        # Gray out sensory evaluation panel
        self.disable_sensory_evaluation()

        # Load multiple sessions if needed
        if not self.sensory_window.all_sessions_data:
            self.load_multiple_sessions()

        # Calculate averages
        self.calculate_sample_averages()

        # Update plot with averages
        self.update_comparison_plot()

        # Bring to front after mode switch
        self.bring_to_front()

        debug_print("Switched to comparison mode - showing averaged data across users")
        show_success_message("Comparison Mode", "Now showing averaged data across multiple users.\nSensory evaluation is disabled in this mode.", self.sensory_window.window)

    def switch_to_collection_mode(self):
        """Switch to collection mode - normal single user operation."""
        debug_print("DEBUG: Switching to collection mode")

        self.sensory_window.current_mode = "collection"
        self.sensory_window.mode_button.config(text="Switch to Comparison Mode")

        # Remove comparison title if it exists
        if hasattr(self, 'comparison_title_frame'):
            self.comparison_title_frame.destroy()

        # Re-enable sensory evaluation panel
        self.enable_sensory_evaluation()

        # Update plot with current user's data
        self.plot_manager.update_plot()

        # Bring to front after mode switch
        self.bring_to_front()

        debug_print("Switched to collection mode - showing single user data")
        show_success_message("Collection Mode", "Now showing single user data collection mode.\nSensory evaluation is enabled.", self.sensory_window.window)

    def disable_sensory_evaluation(self):
        """Gray out and disable all sensory evaluation controls."""
        # Find the sensory evaluation frame and disable all children
        for widget in self.sensory_window.left_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and widget.cget('text') == 'Sensory Evaluation':
                self.set_widget_state(widget, 'disabled')
                widget.configure(style='Disabled.TLabelframe')
        debug_print("Disabled sensory evaluation panel for comparison mode")

    def enable_sensory_evaluation(self):
        """Re-enable all sensory evaluation controls."""
        # Find the sensory evaluation frame and enable all children
        for widget in self.sensory_window.left_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and widget.cget('text') == 'Sensory Evaluation':
                self.set_widget_state(widget, 'normal')
                widget.configure(style='TLabelframe')
        debug_print("Enabled sensory evaluation panel for collection mode")

    def bring_to_front(self):
        """Bring the sensory window to front after user actions."""
        if self.sensory_window.window and self.sensory_window.window.winfo_exists():
            self.sensory_window.window.lift()
            self.sensory_window.window.focus_set()
            debug_print("DEBUG: Brought sensory window to front")

    def set_widget_state(self, parent, state):
        """Recursively set state of all child widgets."""
        try:
            parent.configure(state=state)
        except:
            pass  # Some widgets don't support state

        for child in parent.winfo_children():
            self.set_widget_state(child, state)

    def load_multiple_sessions(self):
        """Enhanced method to load multiple session files for comparison."""
        debug_print("DEBUG: Loading multiple sessions for comparison mode")

        filenames = filedialog.askopenfilenames(
            title="Select Session Files for Comparison",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not filenames:
            debug_print("DEBUG: No files selected for comparison")
            return False

        if len(filenames) < 2:
            messagebox.showwarning("Warning", "Please select at least 2 session files for comparison.")
            return False

        successful_loads = 0

        for filename in filenames:
            try:
                debug_print(f"DEBUG: Loading session file: {filename}")

                with open(filename, 'r') as f:
                    session_data = json.load(f)

                # Create session name from filename
                base_filename = os.path.splitext(os.path.basename(filename))[0]
                session_name = base_filename

                # Ensure unique session name
                counter = 1
                original_name = session_name
                while session_name in self.session_manager.sessions:
                    session_name = f"{original_name}_{counter}"
                    counter += 1

                # Create new session with loaded data
                self.session_manager.sessions[session_name] = {
                    'header': session_data.get('header', {}),
                    'samples': session_data.get('samples', {}),
                    'timestamp': session_data.get('timestamp', datetime.now().isoformat()),
                    'source_file': filename
                }

                successful_loads += 1
                debug_print(f"DEBUG: Successfully loaded session {session_name} with {len(self.session_manager.sessions[session_name]['samples'])} samples")

            except Exception as e:
                debug_print(f"DEBUG: Error loading session from {filename}: {e}")
                messagebox.showerror("Error", f"Failed to load session from {os.path.basename(filename)}: {e}")

        if successful_loads >= 2:
            debug_print(f"DEBUG: Successfully loaded {successful_loads} sessions for comparison")
            show_success_message("Success", f"Loaded {successful_loads} sessions for comparison.", self.sensory_window.window)
            return True
        else:
            messagebox.showerror("Error", "Failed to load enough sessions for comparison (minimum 2 required).")
            return False

        self.bring_to_front()

    def calculate_sample_averages(self):
        """Calculate averages for each sample across all loaded sessions."""
        debug_print("DEBUG: Calculating sample averages across all sessions")

        if len(self.session_manager.sessions) < 2:
            debug_print("DEBUG: Not enough sessions for comparison")
            return

        sample_data = {}

        # Collect all values for each sample/metric combination
        for session_name, session_info in self.session_manager.sessions.items():
            samples = session_info.get('samples', {})
            header = session_info.get('header', {})
            assessor_name = header.get('Assessor Name', session_name)

            debug_print(f"DEBUG: Processing session {session_name} with assessor {assessor_name}")

            for sample_name, sample_values in samples.items():
                if sample_name not in sample_data:
                    sample_data[sample_name] = {metric: [] for metric in self.sensory_window.metrics}
                    sample_data[sample_name]['comments'] = []
                    sample_data[sample_name]['assessors'] = []

                # Collect metric values
                for metric in self.sensory_window.metrics:
                    if metric in sample_values and sample_values[metric] is not None:
                        try:
                            value = float(sample_values[metric])
                            sample_data[sample_name][metric].append(value)
                        except (ValueError, TypeError):
                            debug_print(f"DEBUG: Invalid value for {metric} in {sample_name}: {sample_values[metric]}")

                # Collect comments
                if 'comments' in sample_values and sample_values['comments'].strip():
                    sample_data[sample_name]['comments'].append(f"{assessor_name}: {sample_values['comments']}")

                sample_data[sample_name]['assessors'].append(assessor_name)

        # Calculate averages
        self.average_samples = {}
        for sample_name, data in sample_data.items():
            self.average_samples[sample_name] = {}

            for metric in self.sensory_window.metrics:
                if data[metric]:  # If we have values
                    avg_value = sum(data[metric]) / len(data[metric])
                    self.average_samples[sample_name][metric] = round(avg_value, 1)
                    debug_print(f"DEBUG: {sample_name} {metric} average: {avg_value:.1f} (from {len(data[metric])} values)")
                else:
                    self.average_samples[sample_name][metric] = 5  # Default middle value

            # Combine comments
            self.average_samples[sample_name]['comments'] = '\n'.join(data['comments'])
            self.average_samples[sample_name]['assessor_count'] = len(set(data['assessors']))

        debug_print(f"DEBUG: Calculated averages for {len(self.average_samples)} samples across {len(self.session_manager.sessions)} sessions")

    def update_comparison_plot(self):
        """Update plot to show averaged data across users."""
        if not self.average_samples:
            return

        # Temporarily replace samples with averages for plotting
        original_samples = self.sensory_window.samples.copy()
        original_checkboxes = self.sample_manager.sample_checkboxes.copy()

        # Set up average samples for plotting
        self.sensory_window.samples = self.average_samples.copy()

        # Update sample checkboxes to show average samples
        self.sample_manager.sample_checkboxes = {}
        for sample_name in self.average_samples.keys():
            var = tk.BooleanVar(value=True)  # Show all by default
            self.sample_manager.sample_checkboxes[sample_name] = var

        # Update the checkbox display
        self.sample_manager.update_sample_checkboxes()

        # Update the plot
        self.plot_manager.create_spider_plot()

        debug_print("Updated plot with averaged comparison data")

        self.bring_to_front()

    def setup_comparison_title(self):
        """Add or update the comparison mode title."""
        # Remove existing title if it exists
        if hasattr(self, 'comparison_title_frame'):
            self.comparison_title_frame.destroy()

        if self.sensory_window.current_mode == "comparison":
            # Create title frame at the top of the window
            self.comparison_title_frame = ttk.Frame(self.sensory_window.window)
            self.comparison_title_frame.pack(side='top', fill='x', pady=10)

            # Add the title label with white background
            title_label = ttk.Label(
                self.comparison_title_frame,
                text="Comparing Average Sensory Results",
                font=("Arial", 16, "bold"),
                anchor='center'
            )
            title_label.pack(expand=True)

            # Ensure window stays on top after adding title
            self.sensory_window.window.update_idletasks()
            self.bring_to_front()

            debug_print("DEBUG: Added comparison mode title with white background and brought to front")
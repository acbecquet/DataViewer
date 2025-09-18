"""
Sample Manager
Handles all sample-related operations including CRUD operations, data loading, and UI updates
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from utils import debug_print, show_success_message


class SampleManager:
    """Manages sensory evaluation samples."""
    
    def __init__(self, sensory_window, plot_manager=None):
        """Initialize the sample manager with reference to main window."""
        self.sensory_window = sensory_window
        self.plot_manager = plot_manager
        self.sample_checkboxes = {}
        
    def add_sample(self):
        """Add a new sample for evaluation."""
        import tkinter.simpledialog
        sample_name = tk.simpledialog.askstring("Add Sample", "Enter sample name:")
        if sample_name and sample_name not in self.sensory_window.samples:
            # Initialize sample data
            self.sensory_window.samples[sample_name] = {metric: 0 for metric in self.sensory_window.metrics}
            self.sensory_window.samples[sample_name]['comments'] = ''

            # Update UI
            self.update_sample_combo()
            self.update_sample_checkboxes()
            self.sensory_window.sample_var.set(sample_name)
            self.load_sample_data(sample_name)

            # force plot update
            self.plot_manager.update_plot()

            debug_print(f"Added sample: {sample_name}")
        elif sample_name in self.sensory_window.samples:
            messagebox.showwarning("Warning", "Sample name already exists!")

    def remove_sample(self):
        """Remove the currently selected sample."""
        current_sample = self.sensory_window.sample_var.get()
        if current_sample and current_sample in self.sensory_window.samples:
            if messagebox.askyesno("Confirm", f"Remove sample '{current_sample}'?"):
                del self.sensory_window.samples[current_sample]
                self.update_sample_combo()
                self.update_sample_checkboxes()

                # Select first available sample or clear
                if self.sensory_window.samples:
                    first_sample = list(self.sensory_window.samples.keys())[0]
                    self.sensory_window.sample_var.set(first_sample)
                    self.load_sample_data(first_sample)
                else:
                    self.sensory_window.sample_var.set('')
                    self.clear_form()

                self.plot_manager.update_plot()
                debug_print(f"Removed sample: {current_sample}")

    def clear_current_sample(self):
        """Clear data for the current sample."""
        current_sample = self.sensory_window.sample_var.get()
        if current_sample and current_sample in self.sensory_window.samples:
            if messagebox.askyesno("Confirm", f"Clear data for '{current_sample}'?"):
                # Reset all ratings to 5 (neutral)
                for metric in self.sensory_window.metrics:
                    self.sensory_window.samples[current_sample][metric] = 5
                    self.sensory_window.rating_vars[metric].set(5)

                self.sensory_window.samples[current_sample]['comments'] = ''
                self.sensory_window.comments_text.delete('1.0', tk.END)

                self.plot_manager.update_plot()
                debug_print(f"Cleared data for sample: {current_sample}")

    def rename_current_sample(self):
        """Rename the currently selected sample."""
        current_sample = self.sensory_window.sample_var.get()
        if not current_sample:
            messagebox.showwarning("No Sample Selected", "Please select a sample to rename.")
            return

        if current_sample not in self.sensory_window.samples:
            messagebox.showwarning("Sample Not Found", f"Sample '{current_sample}' not found.")
            return

        # Get new name from user
        new_name = tk.simpledialog.askstring(
            "Rename Sample",
            f"Current name: {current_sample}\n\nEnter new name:",
            initialvalue=current_sample
        )

        if not new_name or new_name.strip() == "":
            return  # User cancelled or entered empty name

        new_name = new_name.strip()

        # Check if new name already exists
        if new_name in self.sensory_window.samples and new_name != current_sample:
            messagebox.showerror("Name Conflict",
                               f"A sample named '{new_name}' already exists.\n"
                               f"Please choose a different name.")
            return

        # Perform the rename
        if new_name != current_sample:
            # Copy data to new key
            self.sensory_window.samples[new_name] = self.sensory_window.samples[current_sample]

            # Remove old key
            del self.sensory_window.samples[current_sample]

            # Update UI components
            self.update_sample_combo()
            self.update_sample_checkboxes()

            # Select the renamed sample
            self.sensory_window.sample_var.set(new_name)
            self.load_sample_data(new_name)

            # Update plot
            self.plot_manager.update_plot()

            debug_print(f"Renamed sample '{current_sample}' to '{new_name}'")
            show_success_message("Success", f"Sample renamed to '{new_name}'", self.sensory_window.window)

    def batch_rename_samples(self):
        """Rename multiple samples at once."""
        if not self.sensory_window.samples:
            messagebox.showwarning("No Samples", "No samples available to rename.")
            return

        # Create batch rename dialog
        rename_window = tk.Toplevel(self.sensory_window.window)
        rename_window.title("Batch Rename Samples")
        rename_window.geometry("600x400")
        rename_window.transient(self.sensory_window.window)
        rename_window.grab_set()

        # Center the window
        rename_window.update_idletasks()
        x = (rename_window.winfo_screenwidth() // 2) - (300)
        y = (rename_window.winfo_screenheight() // 2) - (200)
        rename_window.geometry(f"600x400+{x}+{y}")

        main_frame = ttk.Frame(rename_window, padding=10)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="Batch Rename Samples", font=('Arial', 14, 'bold')).pack(pady=5)
        ttk.Label(main_frame, text="Edit the names below, then click Apply Changes", font=('Arial', 10)).pack(pady=2)

        # Create scrollable frame for sample entries
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Store the name variables
        name_vars = {}
        original_names = list(self.sensory_window.samples.keys())

        # Create entry fields for each sample
        for i, sample_name in enumerate(original_names):
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill='x', pady=2, padx=5)

            ttk.Label(row_frame, text=f"Sample {i+1}:", width=10).pack(side='left')

            name_var = tk.StringVar(value=sample_name)
            name_vars[sample_name] = name_var

            entry = ttk.Entry(row_frame, textvariable=name_var, width=40)
            entry.pack(side='left', padx=5, fill='x', expand=True)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Quick action buttons
        quick_frame = ttk.Frame(main_frame)
        quick_frame.pack(fill='x', pady=10)

        def apply_prefix():
            prefix = tk.simpledialog.askstring("Add Prefix", "Enter prefix to add:")
            if prefix:
                for sample_name in original_names:
                    current = name_vars[sample_name].get()
                    name_vars[sample_name].set(f"{prefix}{current}")

        def apply_suffix():
            suffix = tk.simpledialog.askstring("Add Suffix", "Enter suffix to add:")
            if suffix:
                for sample_name in original_names:
                    current = name_vars[sample_name].get()
                    name_vars[sample_name].set(f"{current}{suffix}")

        def number_samples():
            base = tk.simpledialog.askstring("Number Samples", "Enter base name (e.g., 'Test'):")
            if base:
                for i, sample_name in enumerate(original_names):
                    name_vars[sample_name].set(f"{base} {i+1}")

        ttk.Button(quick_frame, text="Add Prefix", command=apply_prefix).pack(side='left', padx=5)
        ttk.Button(quick_frame, text="Add Suffix", command=apply_suffix).pack(side='left', padx=5)
        ttk.Button(quick_frame, text="Number Samples", command=number_samples).pack(side='left', padx=5)

        # Apply/Cancel buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)

        def apply_changes():
            # Collect all new names
            new_names = {}
            for original_name in original_names:
                new_name = name_vars[original_name].get().strip()
                if not new_name:
                    messagebox.showerror("Empty Name", "All samples must have names.")
                    return
                new_names[original_name] = new_name

            # Check for duplicates
            name_counts = {}
            for new_name in new_names.values():
                name_counts[new_name] = name_counts.get(new_name, 0) + 1

            duplicates = [name for name, count in name_counts.items() if count > 1]
            if duplicates:
                messagebox.showerror("Duplicate Names",
                                   f"The following names are used more than once:\n{', '.join(duplicates)}\n\n"
                                   f"Please ensure all names are unique.")
                return

            # Apply the changes
            new_samples = {}
            for original_name in original_names:
                new_name = new_names[original_name]
                new_samples[new_name] = self.sensory_window.samples[original_name]

            self.sensory_window.samples = new_samples

            # Update UI
            current_selection = self.sensory_window.sample_var.get()
            if current_selection in new_names:
                new_selection = new_names[current_selection]
            else:
                new_selection = list(self.sensory_window.samples.keys())[0] if self.sensory_window.samples else ""

            self.update_sample_combo()
            self.update_sample_checkboxes()

            if new_selection:
                self.sensory_window.sample_var.set(new_selection)
                self.load_sample_data(new_selection)

            self.plot_manager.update_plot()

            rename_window.destroy()
            show_success_message("Success", f"Successfully renamed {len(original_names)} samples.", self.sensory_window.window)

        ttk.Button(button_frame, text="Apply Changes", command=apply_changes).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=rename_window.destroy).pack(side='right', padx=5)

    def update_sample_combo(self):
        """Update the sample selection combo box."""
        sample_names = list(self.sensory_window.samples.keys())
        self.sensory_window.sample_combo['values'] = sample_names

    def update_sample_checkboxes(self):
        """Update the checkboxes for sample selection in plotting."""
        if not hasattr(self.sensory_window, 'checkbox_frame') or not self.sensory_window.checkbox_frame:
            debug_print("DEBUG: No checkbox frame available for sample checkboxes")
            return

        # Clear existing checkboxes
        for widget in self.sensory_window.checkbox_frame.winfo_children():
            widget.destroy()

        # Show different header based on mode
        if self.sensory_window.current_mode == "comparison":
            header_text = "Select Averaged Samples to Display:"
        else:
            header_text = "Select Samples to Display:"

        # Update the header label
        for widget in self.sensory_window.checkbox_frame.master.winfo_children():
            if isinstance(widget, ttk.Label):
                widget.config(text=header_text)
                break

        self.sample_checkboxes = {}

        # Create checkboxes for each sample
        for i, sample_name in enumerate(self.sensory_window.samples.keys()):
            var = tk.BooleanVar(value=True)  # Default to checked
            checkbox = ttk.Checkbutton(self.sensory_window.checkbox_frame, text=sample_name,
                                     variable=var, command=self.plot_manager.update_plot)
            checkbox.grid(row=i//3, column=i%3, sticky='w', padx=5, pady=2)
            self.sample_checkboxes[sample_name] = var

    def on_sample_changed(self, event=None):
        """Handle sample selection change."""
        selected_sample = self.sensory_window.sample_var.get()
        if selected_sample in self.sensory_window.samples:
            self.load_sample_data(selected_sample)

    def load_sample_data(self, sample_name):
        """Load data for the specified sample into the form."""
        debug_print(f"DEBUG: Loading sample data for: {sample_name}")

        if sample_name in self.sensory_window.samples:
            sample_data = self.sensory_window.samples[sample_name]
            debug_print(f"DEBUG: Found sample data: {sample_data}")

            # Load ratings and update both sliders AND display labels
            for metric in self.sensory_window.metrics:
                value = sample_data.get(metric, 5)
                debug_print(f"DEBUG: Setting {metric} to {value}")

                # Update the slider value
                self.sensory_window.rating_vars[metric].set(value)

                # Manually update the display label
                if hasattr(self.sensory_window, 'value_labels') and metric in self.sensory_window.value_labels:
                    self.sensory_window.value_labels[metric].config(text=str(value))
                    debug_print(f"DEBUG: Updated display label for {metric} to {value}")
                else:
                    debug_print(f"DEBUG: No value label found for {metric}")

            # Load comments
            comments = sample_data.get('comments', '')
            self.sensory_window.comments_text.delete('1.0', tk.END)
            self.sensory_window.comments_text.insert('1.0', comments)
            debug_print(f"DEBUG: Loaded comments: '{comments[:50]}...'")

            debug_print(f"DEBUG: Successfully loaded all data for {sample_name}")
        else:
            debug_print(f"DEBUG: Sample {sample_name} not found in samples")

    def refresh_value_displays(self):
        """Refresh all value display labels to match current slider values."""
        debug_print("DEBUG: Refreshing all value displays")

        if not hasattr(self.sensory_window, 'value_labels'):
            debug_print("DEBUG: No value labels found, skipping refresh")
            return

        for metric in self.sensory_window.metrics:
            if metric in self.sensory_window.value_labels and metric in self.sensory_window.rating_vars:
                current_value = self.sensory_window.rating_vars[metric].get()
                self.sensory_window.value_labels[metric].config(text=str(current_value))
                debug_print(f"DEBUG: Refreshed {metric} display to {current_value}")

    def save_current_sample(self):
        """Save the current form data to the selected sample."""
        current_sample = self.sensory_window.sample_var.get()
        if not current_sample:
            messagebox.showwarning("Warning", "No sample selected!")
            return

        if current_sample not in self.sensory_window.samples:
            # Create new sample if it doesn't exist
            self.sensory_window.samples[current_sample] = {}

        # Save ratings
        for metric in self.sensory_window.metrics:
            self.sensory_window.samples[current_sample][metric] = self.sensory_window.rating_vars[metric].get()

        # Save comments
        comments = self.sensory_window.comments_text.get('1.0', tk.END).strip()
        self.sensory_window.samples[current_sample]['comments'] = comments

        self.plot_manager.update_plot()
        debug_print(f"Saved data for sample: {current_sample}")
        show_success_message("Success", f"Data saved for {current_sample}", self.sensory_window.window)

    def clear_form(self):
        """Clear all form fields."""
        debug_print("DEBUG: Clearing form and refreshing displays")

        for metric in self.sensory_window.metrics:
            self.sensory_window.rating_vars[metric].set(5)

            # Also update the display labels
            if hasattr(self.sensory_window, 'value_labels') and metric in self.sensory_window.value_labels:
                self.sensory_window.value_labels[metric].config(text="5")
                debug_print(f"DEBUG: Reset {metric} display to 5")

        self.sensory_window.comments_text.delete('1.0', tk.END)
        debug_print("DEBUG: Form cleared and displays refreshed")

    def auto_save_and_update(self):
        """Automatically save current sample data and update plot."""
        current_sample = self.sensory_window.sample_var.get()
        if current_sample and current_sample in self.sensory_window.samples:
            # Auto-save ratings
            for metric in self.sensory_window.metrics:
                self.sensory_window.samples[current_sample][metric] = self.sensory_window.rating_vars[metric].get()

            # Also auto-save comments
            comments = self.sensory_window.comments_text.get('1.0', tk.END).strip()
            self.sensory_window.samples[current_sample]['comments'] = comments

            # Update plot immediately
            self.plot_manager.update_plot()

            debug_print(f"DEBUG: Auto-saved all data for {current_sample}")
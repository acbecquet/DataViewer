"""
Enhanced Notes Manager Module for DataViewer Application

This module handles interactive sample notes display and editing with auto-save functionality.
"""

import tkinter as tk
from tkinter import ttk
from utils import debug_print


class EnhancedNotesManager:
    """Manages interactive sample notes with auto-save and table integration."""
    
    def __init__(self, gui):
        """Initialize the enhanced notes manager.
        
        Args:
            gui: Reference to the main DataViewer GUI instance
        """
        self.gui = gui
        self.notes_text_widgets = {}
        self.current_sheet_name = None
        self.auto_save_after_id = None
        self.last_saved_content = {}
        
    def create_interactive_notes_display(self):
        """Create an interactive notes display area with editable fields."""
        if not hasattr(self.gui, 'notes_frame') or not self.gui.notes_frame.winfo_exists():
            debug_print("DEBUG: Notes frame does not exist, cannot create notes display")
            return
        
        debug_print("DEBUG: Creating interactive notes display")
        
        # Clear existing widgets
        for widget in self.gui.notes_frame.winfo_children():
            widget.destroy()
        
        # Create notebook for tabs
        self.gui.notes_notebook = ttk.Notebook(self.gui.notes_frame)
        self.gui.notes_notebook.pack(fill="both", expand=True)
        
        self.notes_text_widgets.clear()
        
    def update_interactive_notes_display(self, sheet_name):
        """Update notes display with editable fields for each sample.
        
        Args:
            sheet_name: Name of the sheet to display notes for
        """
        if not hasattr(self.gui, 'notes_frame') or not self.gui.notes_frame.winfo_exists():
            debug_print("DEBUG: Notes frame does not exist")
            return
        
        debug_print(f"DEBUG: Updating interactive notes display for sheet: {sheet_name}")
        self.current_sheet_name = sheet_name
        
        # Get sheet info
        sheet_info = self.gui.filtered_sheets.get(sheet_name)
        if not sheet_info:
            debug_print(f"DEBUG: No sheet info found for {sheet_name}")
            self.create_empty_editable_notes()
            return
        
        # Get header data
        header_data = sheet_info.get('header_data')
        if not header_data:
            debug_print(f"DEBUG: No header data for {sheet_name}, creating empty editable notes")
            self.create_empty_editable_notes()
            return
        
        samples_data = header_data.get('samples', [])
        if not samples_data:
            debug_print(f"DEBUG: No samples data, creating empty editable notes")
            self.create_empty_editable_notes()
            return
        
        debug_print(f"DEBUG: Found {len(samples_data)} samples for notes")
        
        # Clear existing tabs
        for tab in self.gui.notes_notebook.tabs():
            self.gui.notes_notebook.forget(tab)
        
        self.notes_text_widgets.clear()
        self.last_saved_content.clear()
        
        # Get processed data to extract sample names
        try:
            data = sheet_info.get('data')
            sample_names = self._extract_sample_names_from_data(data, len(samples_data))
        except Exception as e:
            debug_print(f"DEBUG: Could not extract sample names: {e}")
            sample_names = [f"Sample {i+1}" for i in range(len(samples_data))]
        
        # Create editable tab for each sample
        for i, sample_data in enumerate(samples_data):
            sample_id = sample_names[i] if i < len(sample_names) else f"Sample {i+1}"
            sample_notes = sample_data.get('sample_notes', '')
            
            # Create tab
            tab_frame = ttk.Frame(self.gui.notes_notebook)
            self.gui.notes_notebook.add(tab_frame, text=f"Sample {i+1}")
            
            # Create content container
            notes_container = ttk.Frame(tab_frame)
            notes_container.pack(fill="both", expand=True, padx=5, pady=5)
            
            # Sample info header
            info_frame = ttk.Frame(notes_container)
            info_frame.pack(fill="x", pady=(0, 5))
            
            sample_info_label = ttk.Label(
                info_frame,
                text=f"Sample: {sample_id}",
                font=('Arial', 10, 'bold')
            )
            sample_info_label.pack(anchor='w')
            
            # Additional info
            if sample_data.get('media'):
                media_label = ttk.Label(
                    info_frame,
                    text=f"Media: {sample_data.get('media', 'N/A')}"
                )
                media_label.pack(anchor='w')
            
            if sample_data.get('viscosity'):
                viscosity_label = ttk.Label(
                    info_frame,
                    text=f"Viscosity: {sample_data.get('viscosity', 'N/A')}"
                )
                viscosity_label.pack(anchor='w')
            
            # Editable notes area
            text_frame = ttk.Frame(notes_container)
            text_frame.pack(fill="both", expand=True)
            
            # CRITICAL: Make text widget editable (state='normal')
            notes_text = tk.Text(
                text_frame,
                wrap='word',
                font=('Arial', 9),
                bg='white',
                state='normal'
            )
            notes_scrollbar = ttk.Scrollbar(
                text_frame,
                orient='vertical',
                command=notes_text.yview
            )
            notes_text.configure(yscrollcommand=notes_scrollbar.set)
            
            notes_text.pack(side='left', fill='both', expand=True)
            notes_scrollbar.pack(side='right', fill='y')
            
            # Insert existing notes
            if sample_notes:
                notes_text.insert('1.0', sample_notes)
                self.last_saved_content[f"Sample_{i}"] = sample_notes
            else:
                placeholder = "Enter test notes for this sample here..."
                notes_text.insert('1.0', placeholder)
                notes_text.tag_add("placeholder", "1.0", "end")
                notes_text.tag_config("placeholder", foreground="gray")
                self.last_saved_content[f"Sample_{i}"] = ""
            
            # Bind auto-save events
            notes_text.bind('<KeyRelease>', lambda e, idx=i: self._schedule_auto_save(idx))
            notes_text.bind('<FocusOut>', lambda e, idx=i: self._save_sample_notes(idx))
            notes_text.bind('<FocusIn>', lambda e, widget=notes_text: self._remove_placeholder(widget))
            
            # Store reference
            self.notes_text_widgets[f"Sample_{i}"] = notes_text
            
            debug_print(f"DEBUG: Created editable notes tab for Sample {i+1}")
        
        debug_print(f"DEBUG: Interactive notes display created with {len(self.notes_text_widgets)} tabs")
    
    def create_empty_editable_notes(self):
        """Create empty editable notes for samples when no data exists."""
        debug_print("DEBUG: Creating empty editable notes display")
        
        # Clear existing tabs
        for tab in self.gui.notes_notebook.tabs():
            self.gui.notes_notebook.forget(tab)
        
        self.notes_text_widgets.clear()
        
        # Determine number of samples from data if available
        current_sheet = self.gui.selected_sheet.get()
        num_samples = 4  # Default
        
        if current_sheet and current_sheet in self.gui.filtered_sheets:
            sheet_info = self.gui.filtered_sheets[current_sheet]
            data = sheet_info.get('data')
            if data is not None and not data.empty:
                # Try to determine number of samples from data structure
                num_samples = self._count_samples_from_data(data)
        
        # Create empty editable tabs
        for i in range(num_samples):
            tab_frame = ttk.Frame(self.gui.notes_notebook)
            self.gui.notes_notebook.add(tab_frame, text=f"Sample {i+1}")
            
            notes_container = ttk.Frame(tab_frame)
            notes_container.pack(fill="both", expand=True, padx=5, pady=5)
            
            # Sample header
            info_frame = ttk.Frame(notes_container)
            info_frame.pack(fill="x", pady=(0, 5))
            
            sample_label = ttk.Label(
                info_frame,
                text=f"Sample {i+1}",
                font=('Arial', 10, 'bold')
            )
            sample_label.pack(anchor='w')
            
            # Editable text area
            text_frame = ttk.Frame(notes_container)
            text_frame.pack(fill="both", expand=True)
            
            notes_text = tk.Text(
                text_frame,
                wrap='word',
                font=('Arial', 9),
                bg='white',
                state='normal'
            )
            notes_scrollbar = ttk.Scrollbar(
                text_frame,
                orient='vertical',
                command=notes_text.yview
            )
            notes_text.configure(yscrollcommand=notes_scrollbar.set)
            
            notes_text.pack(side='left', fill='both', expand=True)
            notes_scrollbar.pack(side='right', fill='y')
            
            # Add placeholder
            placeholder = "Enter test notes for this sample here..."
            notes_text.insert('1.0', placeholder)
            notes_text.tag_add("placeholder", "1.0", "end")
            notes_text.tag_config("placeholder", foreground="gray")
            
            # Bind events
            notes_text.bind('<KeyRelease>', lambda e, idx=i: self._schedule_auto_save(idx))
            notes_text.bind('<FocusOut>', lambda e, idx=i: self._save_sample_notes(idx))
            notes_text.bind('<FocusIn>', lambda e, widget=notes_text: self._remove_placeholder(widget))
            
            self.notes_text_widgets[f"Sample_{i}"] = notes_text
            self.last_saved_content[f"Sample_{i}"] = ""
        
        debug_print(f"DEBUG: Created {num_samples} empty editable note tabs")
    
    def _remove_placeholder(self, text_widget):
        """Remove placeholder text on focus."""
        try:
            content = text_widget.get('1.0', 'end-1c')
            if "Enter test notes for this sample here..." in content:
                text_widget.delete('1.0', 'end')
                text_widget.tag_remove("placeholder", "1.0", "end")
        except Exception as e:
            debug_print(f"DEBUG: Error removing placeholder: {e}")
    
    def _schedule_auto_save(self, sample_index):
        """Schedule auto-save after typing stops.
        
        Args:
            sample_index: Index of the sample being edited
        """
        # Cancel previous scheduled save
        if self.auto_save_after_id:
            self.gui.root.after_cancel(self.auto_save_after_id)
        
        # Schedule new save after 2 seconds of inactivity
        self.auto_save_after_id = self.gui.root.after(2000, lambda: self._save_sample_notes(sample_index))
    
    def _save_sample_notes(self, sample_index):
        """Save sample notes to the data structure.
        
        Args:
            sample_index: Index of the sample to save notes for
        """
        try:
            widget_key = f"Sample_{sample_index}"
            
            if widget_key not in self.notes_text_widgets:
                debug_print(f"DEBUG: No text widget found for {widget_key}")
                return
            
            text_widget = self.notes_text_widgets[widget_key]
            content = text_widget.get('1.0', 'end-1c').strip()
            
            # Skip if content unchanged
            if widget_key in self.last_saved_content:
                if content == self.last_saved_content[widget_key]:
                    return
            
            # Remove placeholder text from save
            if content == "Enter test notes for this sample here...":
                content = ""
            
            debug_print(f"DEBUG: Auto-saving notes for Sample {sample_index+1}: {len(content)} characters")
            
            # Update filtered_sheets with notes
            if not self.current_sheet_name:
                debug_print("DEBUG: No current sheet name set")
                return
            
            sheet_info = self.gui.filtered_sheets.get(self.current_sheet_name)
            if not sheet_info:
                debug_print(f"DEBUG: No sheet info for {self.current_sheet_name}")
                return
            
            # Ensure header_data structure exists
            if 'header_data' not in sheet_info:
                sheet_info['header_data'] = {'samples': []}
            
            header_data = sheet_info['header_data']
            
            # Ensure samples list is long enough
            while len(header_data['samples']) <= sample_index:
                header_data['samples'].append({})
            
            # Save the notes
            header_data['samples'][sample_index]['sample_notes'] = content
            self.last_saved_content[widget_key] = content
            
            # Mark file as modified
            if hasattr(self.gui, 'all_filtered_sheets'):
                for file_data in self.gui.all_filtered_sheets:
                    if file_data.get('file_path') == self.gui.file_path:
                        file_data['is_modified'] = True
                        debug_print(f"DEBUG: Marked file as modified")
                        break
            
            debug_print(f"DEBUG: Successfully saved notes for Sample {sample_index+1}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to save sample notes: {e}")
            import traceback
            traceback.print_exc()
    
    def switch_to_sample_tab(self, sample_index):
        """Switch to the notes tab for a specific sample.
        
        Args:
            sample_index: Index of sample to switch to (0-based)
        """
        try:
            if not hasattr(self.gui, 'notes_notebook'):
                debug_print("DEBUG: No notes notebook exists")
                return
            
            # Get all tabs
            tabs = self.gui.notes_notebook.tabs()
            
            if sample_index >= len(tabs):
                debug_print(f"DEBUG: Sample index {sample_index} out of range (max {len(tabs)-1})")
                return
            
            # Switch to the tab
            self.gui.notes_notebook.select(tabs[sample_index])
            debug_print(f"DEBUG: Switched to notes tab for Sample {sample_index+1}")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to switch to sample tab: {e}")
    
    def _extract_sample_names_from_data(self, data, num_samples):
        """Extract sample names from the processed data.
        
        Args:
            data: Processed DataFrame
            num_samples: Expected number of samples
            
        Returns:
            List of sample names
        """
        sample_names = []
        
        try:
            # Look for 'Sample Name' column
            if 'Sample Name' in data.columns:
                sample_names = data['Sample Name'].dropna().unique().tolist()
                debug_print(f"DEBUG: Extracted {len(sample_names)} sample names from data")
            
            # Ensure we have enough names
            while len(sample_names) < num_samples:
                sample_names.append(f"Sample {len(sample_names)+1}")
            
        except Exception as e:
            debug_print(f"DEBUG: Error extracting sample names: {e}")
            sample_names = [f"Sample {i+1}" for i in range(num_samples)]
        
        return sample_names
    
    def _count_samples_from_data(self, data):
        """Count number of samples from data structure.
        
        Args:
            data: DataFrame to analyze
            
        Returns:
            Number of samples detected
        """
        try:
            # Try counting unique sample names
            if 'Sample Name' in data.columns:
                return len(data['Sample Name'].dropna().unique())
            
            # Default to 4 samples
            return 4
            
        except Exception as e:
            debug_print(f"DEBUG: Error counting samples: {e}")
            return 4


def bind_table_double_click(gui):
    """Bind double-click event on table to switch notes tabs.
    
    Args:
        gui: Main DataViewer GUI instance
    """
    try:
        # Get current sheet widget
        current_sheet = gui.selected_sheet.get()
        
        if not hasattr(gui, 'current_sheet_widget'):
            debug_print("DEBUG: No current_sheet_widget attribute")
            return
        
        if current_sheet not in gui.current_sheet_widget:
            debug_print(f"DEBUG: No sheet widget for {current_sheet}")
            return
        
        sheet_widget = gui.current_sheet_widget[current_sheet]
        
        # Bind double-click to switch notes tab
        def on_double_click(event):
            try:
                # Get selected cell
                selected = sheet_widget.get_currently_selected()
                if not selected:
                    return
                
                row = selected.row
                debug_print(f"DEBUG: Double-clicked row {row}")
                
                # Switch to corresponding sample notes tab
                # Row index maps to sample index
                if hasattr(gui, 'notes_manager'):
                    gui.notes_manager.switch_to_sample_tab(row)
                
            except Exception as e:
                debug_print(f"DEBUG: Error handling table double-click: {e}")
        
        # Bind the event
        sheet_widget.bind("<Double-Button-1>", on_double_click)
        debug_print("DEBUG: Bound double-click event to table")
        
    except Exception as e:
        debug_print(f"DEBUG: Error binding table double-click: {e}")
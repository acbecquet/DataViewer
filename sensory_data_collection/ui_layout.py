"""
UI Layout Manager
Handles UI layout, window management, and resize operations
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime
from utils import APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT, debug_print


class UILayoutManager:
    """Manages UI layout and window operations."""
    
    def __init__(self, sensory_window):
        """Initialize the UI layout manager with reference to main window."""
        self.sensory_window = sensory_window
        
    def setup_layout(self):
        """Create the main layout with proper canvas sizing."""
        # Create main paned window
        main_paned = tk.PanedWindow(self.window, orient='horizontal', sashrelief='raised', sashwidth=4)
        main_paned.pack(fill='both', expand=True, padx=5, pady=5)

        # Store reference to main_paned for resize handling
        self.main_paned = main_paned

        left_canvas = tk.Canvas(main_paned, bg=APP_BACKGROUND_COLOR, highlightthickness=0)
        self.left_canvas = left_canvas
        left_scrollbar = ttk.Scrollbar(main_paned, orient="vertical", command=left_canvas.yview)
        self.left_frame = ttk.Frame(left_canvas)

        # Better scroll configuration
        def configure_left_scroll(event=None):
            """Configure scroll region and handle resizing."""
            self.left_frame.update_idletasks()

            # Update scroll region
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

            # Get the current canvas size
            canvas_width = left_canvas.winfo_width()
            if canvas_width > 50:  # Valid width
                # Configure the interior frame to fill the canvas width
                left_canvas.itemconfig(left_canvas.find_all()[0], width=canvas_width-4)

        self.left_frame.bind("<Configure>", configure_left_scroll)

        # Create the canvas window
        left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        # Better canvas resize handling
        def on_left_canvas_configure(event):
            """Handle left canvas resize events."""
            if event.widget == left_canvas:
                # Update scroll region
                left_canvas.configure(scrollregion=left_canvas.bbox("all"))

                # Ensure interior frame fills width
                canvas_width = event.width
                if canvas_width > 50 and left_canvas.find_all():
                    left_canvas.itemconfig(left_canvas.find_all()[0], width=canvas_width-4)

        left_canvas.bind('<Configure>', on_left_canvas_configure)

        # Mouse wheel scrolling
        def _on_mousewheel_left(event):
            left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        left_canvas.bind("<MouseWheel>", _on_mousewheel_left)

        # Add to paned window
        main_paned.add(left_canvas, stretch="always")

        right_canvas = tk.Canvas(main_paned, bg=APP_BACKGROUND_COLOR, highlightthickness=0)
        right_scrollbar = ttk.Scrollbar(main_paned, orient="vertical", command=right_canvas.yview)
        self.right_frame = ttk.Frame(right_canvas)

        self.right_canvas = right_canvas

        # Right panel configuration
        def configure_right_scroll(event=None):
            """Configure right scroll region."""
            self.right_frame.update_idletasks()
            bbox = right_canvas.bbox("all")
            if bbox:
                right_canvas.configure(scrollregion=bbox)

                # Get the height of the paned window
                if hasattr(self, 'main_paned'):
                    paned_height = self.main_paned.winfo_height()
                    if paned_height > 100:  # Valid height
                        # Set canvas to use full paned window height
                        right_canvas.configure(height=paned_height)
                        debug_print(f"DEBUG: Set right canvas height to match paned window: {paned_height}px")

        self.right_frame.bind("<Configure>", configure_right_scroll)

        self.right_canvas_window = right_canvas.create_window((0, 0), window=self.right_frame, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)

        def on_right_canvas_configure(event):
            """Handle right canvas resize events and update interior frame."""
            if event.widget == right_canvas:
                canvas_width = event.width
                canvas_height = event.height
                debug_print(f"DEBUG: Right canvas resized to: {canvas_width}x{canvas_height}")

                if canvas_width > 50 and canvas_height > 50:
                    right_canvas.itemconfig(self.right_canvas_window, width=canvas_width-4, height=canvas_height-4)
                    self.right_frame.configure(width=canvas_width-4, height=canvas_height-4)

                    # Force update of all children
                    self.right_frame.update_idletasks()

                    # Trigger plot resize after canvas updates (only if we have valid dimensions)
                    if hasattr(self, 'update_plot_size_for_resize') and canvas_width > 200 and canvas_height > 200:
                        self.window.after(1000, self.update_plot_size_for_resize)

        right_canvas.bind('<Configure>', on_right_canvas_configure)

        # Mouse wheel scrolling
        def _on_mousewheel_right(event):
            right_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        right_canvas.bind("<MouseWheel>", _on_mousewheel_right)

        # Add right canvas with stretch="always" like the left canvas
        main_paned.add(right_canvas, stretch="always")

        # Store the sash position function
        def set_initial_sash_position():
            try:
                window_width = self.window.winfo_width()
                if window_width > 100:
                    sash_position = int(window_width * 0.40)
                    main_paned.sash_place(0, sash_position, 0)
                    debug_print(f"DEBUG: Set sash position to {sash_position}")

                    # Force canvas height update after sash positioning
                    self.window.after(50, self.equalize_canvas_heights)
            except Exception as e:
                debug_print(f"DEBUG: Sash positioning failed: {e}")

        self.set_initial_sash_position = set_initial_sash_position

        debug_print("DEBUG: Enhanced layout setup complete")

        # Add session management and panels
        self.setup_session_selector(self.left_frame)
        self.setup_data_entry_panel()
        self.setup_plot_panel()

        # Apply sizing optimization after content is added
        self.window.after(100, self.optimize_window_size)

        # Initialize default session
        if not self.sessions:
            self.create_new_session("Default_Session")

    def setup_data_entry_panel(self):
        """Setup the left panel for data entry."""
        # Header section
        header_frame = ttk.LabelFrame(self.left_frame, text="Session Information", padding=10)
        header_frame.pack(fill='x', padx=5, pady=5)

        # Configure header_frame for 2x2 + button layout
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(2, weight=1)
        header_frame.grid_columnconfigure(3, weight=1)
        debug_print("DEBUG: Configured header_frame for optimized 2x2 layout")

        self.header_vars = {}

        # Row 0: Assessor Name and Media
        ttk.Label(header_frame, text="Assessor Name:", font=FONT).grid(
            row=0, column=0, sticky='e', padx=5, pady=2)
        assessor_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=assessor_var, font=FONT, width=15).grid(
            row=0, column=1, sticky='w', padx=5, pady=2)
        self.header_vars["Assessor Name"] = assessor_var
        debug_print("DEBUG: Added Assessor Name to row 0, column 0-1")

        ttk.Label(header_frame, text="Media:", font=FONT).grid(
            row=0, column=2, sticky='e', padx=5, pady=2)
        media_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=media_var, font=FONT, width=15).grid(
            row=0, column=3, sticky='w', padx=5, pady=2)
        self.header_vars["Media"] = media_var
        debug_print("DEBUG: Added Media to row 0, column 2-3")

        # Row 1: Puff Length and Date
        ttk.Label(header_frame, text="Puff Length:", font=FONT).grid(
            row=1, column=0, sticky='e', padx=5, pady=2)
        puff_var = tk.StringVar()
        ttk.Entry(header_frame, textvariable=puff_var, font=FONT, width=15).grid(
            row=1, column=1, sticky='w', padx=5, pady=2)
        self.header_vars["Puff Length"] = puff_var
        debug_print("DEBUG: Added Puff Length to row 1, column 0-1")

        ttk.Label(header_frame, text="Date:", font=FONT).grid(
            row=1, column=2, sticky='e', padx=5, pady=2)
        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(header_frame, textvariable=date_var, font=FONT, width=15).grid(
            row=1, column=3, sticky='w', padx=5, pady=2)
        self.header_vars["Date"] = date_var
        debug_print("DEBUG: Added Date to row 1, column 2-3")

        # Row 2: Mode switch button (centered across all columns)
        mode_button_frame = ttk.Frame(header_frame)
        mode_button_frame.grid(row=2, column=0, columnspan=4, pady=10)

        self.mode_button = ttk.Button(mode_button_frame, text="Switch to Comparison Mode",
                                     command=self.toggle_mode, width=25)
        self.mode_button.pack()
        debug_print("Added mode switch button to header section")

        # Sample management sectionG
        sample_frame = ttk.LabelFrame(self.left_frame, text="Sample Management", padding=10)
        sample_frame.pack(fill='x', padx=5, pady=5)

        debug_print("Setting up sample management with simple centering")

        # ROW 1: Sample selection
        sample_select_outer = ttk.Frame(sample_frame)
        sample_select_outer.pack(fill='x', pady=5)

        sample_select_frame = ttk.Frame(sample_select_outer)
        sample_select_frame.pack(expand=True)

        ttk.Label(sample_select_frame, text="Current Sample:", font=FONT).pack(side='left')
        self.sample_var = tk.StringVar()
        self.sample_combo = ttk.Combobox(sample_select_frame, textvariable=self.sample_var,
                                        state="readonly", width=15)
        self.sample_combo.pack(side='left', padx=5)
        self.sample_combo.bind('<<ComboboxSelected>>', self.on_sample_changed)

        debug_print("Sample selection centered with expand=True")

        # ROW 2: Buttons - simple center using expand
        button_outer = ttk.Frame(sample_frame)
        button_outer.pack(fill='x', pady=5)

        button_frame = ttk.Frame(button_outer)
        button_frame.pack(expand=True)

        ttk.Button(button_frame, text="Add Sample",
                  command=self.add_sample).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Remove Sample",
                  command=self.remove_sample).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Clear Data",
                  command=self.clear_current_sample).pack(side='left', padx=2)

        debug_print("Buttons centered with expand=True")

        # Sensory evaluation section
        eval_frame = ttk.LabelFrame(self.left_frame, text="Sensory Evaluation", padding=10)
        eval_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Create rating scales for each metric
        self.rating_vars = {}
        self.value_labels = {}

        for i, metric in enumerate(self.metrics):
            metric_container = ttk.Frame(eval_frame)
            metric_container.pack(fill='x', pady=4)

            metric_frame = ttk.Frame(metric_container)
            metric_frame.pack(anchor='center')

            # Metric label
            ttk.Label(metric_frame, text=f"{metric}:", font=FONT, width=12).pack(side='left')

            scale_container = ttk.Frame(metric_frame)
            scale_container.pack(side='left', padx=5)

            # Rating scale (1-9)
            self.rating_vars[metric] = tk.IntVar(value=5)
            scale = tk.Scale(scale_container, from_=1, to=9, orient='horizontal',
                           variable=self.rating_vars[metric], font=FONT,
                           length=300, showvalue=0, tickinterval=1,
                           sliderlength=20, sliderrelief='raised', width=15)
            scale.pack(side='left')
            debug_print(f"DEBUG: Created centered scale for {metric} with length=200, smaller pointer (sliderlength=15), and tickmarks every 1 point")

            # Current value display
            value_label = ttk.Label(metric_frame, text="5", width=2)
            value_label.pack(side='left', padx=(10, 0))

            self.value_labels[metric] = value_label
            debug_print(f"DEBUG: Stored reference to value label for {metric}")

            # Update value display AND plot when scale changes (LIVE UPDATES)
            def update_live(val, label=value_label, var=self.rating_vars[metric], metric_name=metric):
                label.config(text=str(var.get()))
                self.auto_save_and_update()
            scale.config(command=update_live)
            debug_print(f"DEBUG: Centered scale for {metric} configured with smaller pointer and tickmarks from 1-9")

        # Comments section
        comments_frame = ttk.Frame(eval_frame)
        comments_frame.pack(fill='x', pady=10)

        ttk.Label(comments_frame, text="Additional Comments:", font=FONT).pack(anchor='w')
        self.comments_text = tk.Text(comments_frame, height=4, font=FONT)
        self.comments_text.pack(fill='x', pady=2)

        # Auto-save comments when user types
        def on_comment_change(event=None):
            """Auto-save comments when user types."""
            current_sample = self.sample_var.get()
            if current_sample and current_sample in self.samples:
                comments = self.comments_text.get('1.0', tk.END).strip()
                self.samples[current_sample]['comments'] = comments
                debug_print(f"DEBUG: Auto-saved comments for {current_sample}: '{comments[:50]}...'")

        # Bind to text change events
        self.comments_text.bind('<KeyRelease>', on_comment_change)
        self.comments_text.bind('<FocusOut>', on_comment_change)
        self.comments_text.bind('<Button-1>', lambda e: self.window.after(100, on_comment_change))

    def center_window(self):
        """Center the window on screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def on_window_resize(self, event):
        """Handle general window resize events with dynamic sash positioning."""
        # Only process if this is the main window resize, not child widgets
        if event.widget != self.window:
            return

        # Get current window dimensions
        window_width = self.window.winfo_width()
        window_height = self.window.winfo_height()

        if hasattr(self, 'main_paned'):
            # Force geometry update first
            self.main_paned.update_idletasks()

            left_panel_proportion = 0.3
            new_sash_position = int(window_width * left_panel_proportion)


            min_left_width = 350
            max_left_width = window_width - 400

            new_sash_position = max(min_left_width, min(new_sash_position, max_left_width))

            debug_print(f"DEBUG: Setting proportional sash position to: {new_sash_position}px ({left_panel_proportion*100:.0f}% of window width)")

            # Update the sash position
            try:
                self.main_paned.sash_place(0, new_sash_position, 0)
                debug_print("DEBUG: Sash position updated successfully")
            except Exception as e:
                debug_print(f"DEBUG: Error updating sash position: {e}")

        # Force all frames to update their geometry
        debug_print("DEBUG: Forcing frame geometry updates...")
        self.window.update_idletasks()

        # Equalize canvas heights
        self.equalize_canvas_heights()

        # Force right canvas to update its size
        if hasattr(self, 'right_canvas'):
            self.right_canvas.update_idletasks()
            # Trigger the canvas configure event to update interior frame
            self.right_canvas.event_generate('<Configure>')

        # Trigger plot-specific resize with a slight delay to allow frame updates
        if hasattr(self, 'on_window_resize_plot'):
            # Add a small delay to let the sash repositioning complete
            self.window.after(50, lambda: self.on_window_resize_plot(event))

    def equalize_canvas_heights(self):
        """Ensure both canvases have the same height."""
        if hasattr(self, 'left_canvas') and hasattr(self, 'right_canvas'):
            # Get the paned window height
            if hasattr(self, 'main_paned'):
                paned_height = self.main_paned.winfo_height()

                # Set both canvases to the same height
                if paned_height > 100:
                    self.left_canvas.configure(height=paned_height - 10)
                    self.right_canvas.configure(height=paned_height - 10)

                    # Force update
                    self.left_canvas.update_idletasks()
                    self.right_canvas.update_idletasks()

    def optimize_window_size(self):
        """Calculate window size based on actual frame dimensions after layout."""

        # Force complete layout update
        self.window.update_idletasks()
        self.window.update()

        # Measure what the left_frame actually uses after layout
        self.left_frame.update_idletasks()
        actual_left_frame_height = self.left_frame.winfo_reqheight()

        # Also measure right frame actual height for comparison
        self.right_frame.update_idletasks()
        actual_right_frame_height = self.right_frame.winfo_reqheight()

        debug_print(f"DEBUG: Actual frame heights - Left: {actual_left_frame_height}px, Right: {actual_right_frame_height}px")

        # Width calculations (existing logic)
        left_frame_width = self.left_frame.winfo_reqwidth()
        right_frame_width = self.right_frame.winfo_reqwidth()

        min_plot_width = 500
        optimal_left_width = max(left_frame_width + 40, 450)
        optimal_right_width = max(min_plot_width, right_frame_width + 20)
        total_optimal_width = optimal_left_width + optimal_right_width + 50

        # Reduce window chrome overhead and use required height instead of actual height
        governing_content_height = max(actual_left_frame_height, actual_right_frame_height)
        window_chrome = 10  # REDUCED from 120 to 30 - just enough for title bar and borders
        total_optimal_height = governing_content_height + window_chrome

        debug_print(f"DEBUG: Window sized for actual frame height: {governing_content_height}px")

        # Screen constraints
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        max_usable_width = int(screen_width * 0.9)
        max_usable_height = int(screen_height * 0.85)

        final_width = min(total_optimal_width, max_usable_width)
        final_height = min(total_optimal_height, max_usable_height)

        final_width = max(final_width, 800)
        final_height = max(final_height, 500)

        if final_height > screen_height*0.91:
            final_height = screen_height*0.91

        debug_print(f"DEBUG: Final window size matching actual content: {final_width}x{final_height}")

        # Apply the sizing
        self.window.geometry(f"{final_width}x{final_height}")

        # Pass the actual required height, not the full window height
        available_height = governing_content_height
        self.window.after(50, lambda: self.configure_canvas_sizing(available_height))

        self.center_window()

        if hasattr(self, 'set_initial_sash_position'):
            self.window.after(200, self.set_initial_sash_position)

    def configure_canvas_sizing(self, available_content_height):
        """Configure canvas sizing using the frame's actual rendered size."""
        debug_print(f"DEBUG: Configuring canvas sizing")

        self.window.update_idletasks()

        # Get the required height of the left frame content
        self.left_frame.update_idletasks()
        required_frame_height = self.left_frame.winfo_reqheight()

        debug_print(f"DEBUG: left_frame required height: {required_frame_height}px")

        # Set canvas to exactly match what the frame requires
        if hasattr(self, 'left_canvas'):
            # Add small padding but not excessive
            canvas_height = required_frame_height
            self.left_canvas.configure(height=canvas_height)
            debug_print(f"DEBUG: Canvas set to frame's required height + padding: {canvas_height}px")

            # Update scroll region
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))

        debug_print("DEBUG: Canvas sized to exact frame height - minimal gray space")

    def coordinate_panel_heights(self, final_window_height, window_chrome_height):
        """Coordinate both panels to work optimally at the chosen window height."""
        debug_print("DEBUG: Coordinating panel heights for optimal layout")

        # Calculate available height for panel content
        available_panel_height = final_window_height - window_chrome_height

        # Left Panel Strategy: Optimize scrolling behavior
        # If content fits, disable scrolling; if not, optimize scroll region
        left_content_height = self.left_frame.winfo_reqheight()

        if left_content_height <= available_panel_height:
            # Content fits! Set canvas to exact content height to eliminate gray space
            optimal_left_height = left_content_height + 5  # Small buffer
            debug_print(f"DEBUG: Left panel content fits - setting canvas to {optimal_left_height}px")
        else:
            # Content is taller - use available height and enable smooth scrolling
            optimal_left_height = available_panel_height - 10  # Account for scrollbar
            debug_print(f"DEBUG: Left panel content scrollable - setting canvas to {optimal_left_height}px")

        # Apply the left panel height optimization
        if hasattr(self, 'left_canvas'):
            self.left_canvas.configure(height=optimal_left_height)

        # Right Panel Strategy: Ensure plot area uses available space efficiently
        right_content_height = self.right_frame.winfo_reqheight()

        if right_content_height < available_panel_height:
            # Right panel has extra space - we could expand plot or center it
            extra_space = available_panel_height - right_content_height
            debug_print(f"DEBUG: Right panel has {extra_space}px extra space - content will be naturally centered")
        else:
            debug_print(f"DEBUG: Right panel content fits exactly in available space")

    def setup_interface(self):
        """Set up the main interface with session management."""
        # Create main frames
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Add session management at the top
        self.setup_session_selector(main_frame)

        # Initialize with default session
        if not self.sessions:
            self.create_new_session("Default_Session")

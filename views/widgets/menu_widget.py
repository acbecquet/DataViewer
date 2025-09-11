# views/widgets/menu_widget.py
"""
views/widgets/menu_widget.py
Menu widget for application menu components.
This contains menu functionality and control widgets.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List, Callable, Union


class MenuWidget:
    """Widget for application menus and controls."""
    
    def __init__(self, parent: tk.Widget, controller: Optional[Any] = None):
        """Initialize the menu widget."""
        self.parent = parent
        self.controller = controller
        
        # Menu components
        self.menubar: Optional[tk.Menu] = None
        self.menus: Dict[str, tk.Menu] = {}
        
        # Control components
        self.control_frame: Optional[ttk.Frame] = None
        self.dropdowns: Dict[str, ttk.Combobox] = {}
        self.buttons: Dict[str, ttk.Button] = {}
        self.checkboxes: Dict[str, ttk.Checkbutton] = {}
        self.labels: Dict[str, ttk.Label] = {}
        
        # Variables for controls
        self.variables: Dict[str, tk.Variable] = {}
        
        # Callbacks
        self.menu_callbacks: Dict[str, Callable] = {}
        self.control_callbacks: Dict[str, Callable] = {}
        
        print("DEBUG: MenuWidget initialized")
    
    def create_main_menubar(self, root: tk.Tk) -> tk.Menu:
        """Create main application menubar."""
        try:
            self.menubar = tk.Menu(root)
            root.config(menu=self.menubar)
            
            print("DEBUG: MenuWidget - main menubar created")
            return self.menubar
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating menubar: {e}")
            return None
    
    def add_menu(self, menu_name: str, label: str, tearoff: bool = False) -> tk.Menu:
        """Add a menu to the menubar."""
        try:
            if not self.menubar:
                print("ERROR: MenuWidget - no menubar available")
                return None
            
            menu = tk.Menu(self.menubar, tearoff=tearoff)
            self.menubar.add_cascade(label=label, menu=menu)
            self.menus[menu_name] = menu
            
            print(f"DEBUG: MenuWidget - menu added: {menu_name}")
            return menu
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error adding menu {menu_name}: {e}")
            return None
    
    def add_menu_item(self, menu_name: str, item_type: str, label: str, 
                     command: Callable = None, **kwargs) -> bool:
        """Add an item to a menu."""
        try:
            if menu_name not in self.menus:
                print(f"ERROR: MenuWidget - menu not found: {menu_name}")
                return False
            
            menu = self.menus[menu_name]
            callback_key = f"{menu_name}_{label.replace(' ', '_').replace('&', '').lower()}"
            
            if command:
                self.menu_callbacks[callback_key] = command
            
            if item_type == "command":
                menu.add_command(label=label, command=command, **kwargs)
            elif item_type == "separator":
                menu.add_separator()
            elif item_type == "checkbutton":
                menu.add_checkbutton(label=label, command=command, **kwargs)
            elif item_type == "radiobutton":
                menu.add_radiobutton(label=label, command=command, **kwargs)
            elif item_type == "cascade":
                submenu = tk.Menu(menu, tearoff=kwargs.get('tearoff', False))
                menu.add_cascade(label=label, menu=submenu)
                submenu_name = f"{menu_name}_{label.replace(' ', '_').lower()}"
                self.menus[submenu_name] = submenu
            else:
                print(f"ERROR: MenuWidget - unknown menu item type: {item_type}")
                return False
            
            print(f"DEBUG: MenuWidget - menu item added: {menu_name}.{label}")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error adding menu item: {e}")
            return False
    
    def create_file_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create standard file menu."""
        try:
            menu = self.add_menu("file", "File")
            if not menu:
                return False
            
            # File menu items
            items = [
                ("command", "New", callbacks.get("new")),
                ("command", "Load Excel", callbacks.get("load_excel")),
                ("command", "Load VAP3", callbacks.get("load_vap3")),
                ("separator", None, None),
                ("command", "Save As VAP3", callbacks.get("save_vap3")),
                ("separator", None, None),
                ("command", "Exit", callbacks.get("exit"))
            ]
            
            for item_type, label, command in items:
                if item_type == "separator":
                    self.add_menu_item("file", "separator", "")
                else:
                    self.add_menu_item("file", item_type, label, command)
            
            print("DEBUG: MenuWidget - file menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating file menu: {e}")
            return False
    
    def create_view_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create view menu."""
        try:
            menu = self.add_menu("view", "View")
            if not menu:
                return False
            
            items = [
                ("command", "View Raw Data", callbacks.get("view_raw_data")),
                ("command", "Trend Analysis", callbacks.get("trend_analysis"))
            ]
            
            for item_type, label, command in items:
                self.add_menu_item("view", item_type, label, command)
            
            print("DEBUG: MenuWidget - view menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating view menu: {e}")
            return False
    
    def create_database_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create database menu."""
        try:
            menu = self.add_menu("database", "Database")
            if not menu:
                return False
            
            items = [
                ("command", "Browse Database", callbacks.get("browse_database")),
                ("command", "Load from Database", callbacks.get("load_from_database")),
                ("command", "Update Database", callbacks.get("update_database"))
            ]
            
            for item_type, label, command in items:
                self.add_menu_item("database", item_type, label, command)
            
            print("DEBUG: MenuWidget - database menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating database menu: {e}")
            return False
    
    def create_data_collection_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create data collection menu."""
        try:
            menu = self.add_menu("data_collection", "Data Collection")
            if not menu:
                return False
            
            items = [
                ("command", "Collect Data", callbacks.get("collect_data")),
                ("command", "Sensory Data Collection", callbacks.get("sensory_data_collection")),
                ("command", "Sample Comparison", callbacks.get("sample_comparison"))
            ]
            
            for item_type, label, command in items:
                self.add_menu_item("data_collection", item_type, label, command)
            
            print("DEBUG: MenuWidget - data collection menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating data collection menu: {e}")
            return False
    
    def create_calculate_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create calculate menu."""
        try:
            menu = self.add_menu("calculate", "Calculate")
            if not menu:
                return False
            
            items = [
                ("command", "Viscosity", callbacks.get("viscosity"))
            ]
            
            for item_type, label, command in items:
                self.add_menu_item("calculate", item_type, label, command)
            
            print("DEBUG: MenuWidget - calculate menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating calculate menu: {e}")
            return False
    
    def create_reports_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create reports menu."""
        try:
            menu = self.add_menu("reports", "Reports")
            if not menu:
                return False
            
            items = [
                ("command", "Generate Test Report", callbacks.get("test_report")),
                ("command", "Generate Full Report", callbacks.get("full_report"))
            ]
            
            for item_type, label, command in items:
                self.add_menu_item("reports", item_type, label, command)
            
            print("DEBUG: MenuWidget - reports menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating reports menu: {e}")
            return False
    
    def create_help_menu(self, callbacks: Dict[str, Callable]) -> bool:
        """Create help menu."""
        try:
            menu = self.add_menu("help", "Help")
            if not menu:
                return False
            
            items = [
                ("command", "Help", callbacks.get("help")),
                ("separator", None, None),
                ("command", "About", callbacks.get("about"))
            ]
            
            for item_type, label, command in items:
                if item_type == "separator":
                    self.add_menu_item("help", "separator", "")
                else:
                    self.add_menu_item("help", item_type, label, command)
            
            print("DEBUG: MenuWidget - help menu created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating help menu: {e}")
            return False
    
    def create_standard_menubar(self, root: tk.Tk, callbacks: Dict[str, Dict[str, Callable]]) -> bool:
        """Create standard application menubar with all menus."""
        try:
            if not self.create_main_menubar(root):
                return False
            
            # Create all standard menus
            menu_creators = [
                ("file", self.create_file_menu),
                ("view", self.create_view_menu),
                ("database", self.create_database_menu),
                ("data_collection", self.create_data_collection_menu),
                ("calculate", self.create_calculate_menu),
                ("reports", self.create_reports_menu),
                ("help", self.create_help_menu)
            ]
            
            for menu_name, creator_func in menu_creators:
                menu_callbacks = callbacks.get(menu_name, {})
                if not creator_func(menu_callbacks):
                    print(f"WARNING: MenuWidget - failed to create {menu_name} menu")
            
            print("DEBUG: MenuWidget - standard menubar created")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating standard menubar: {e}")
            return False
    
    def create_control_frame(self, parent: tk.Widget) -> ttk.Frame:
        """Create frame for control widgets."""
        try:
            self.control_frame = ttk.Frame(parent)
            print("DEBUG: MenuWidget - control frame created")
            return self.control_frame
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating control frame: {e}")
            return None
    
    def add_dropdown(self, name: str, label: str, values: List[str], 
                    variable: tk.StringVar = None, callback: Callable = None,
                    pack_options: Dict = None) -> ttk.Combobox:
        """Add a dropdown control."""
        try:
            if not self.control_frame:
                print("ERROR: MenuWidget - no control frame available")
                return None
            
            # Create label
            if label:
                label_widget = ttk.Label(self.control_frame, text=label)
                pack_opts = {"side": "left", "padx": (5, 2)}
                if pack_options and "label_pack" in pack_options:
                    pack_opts.update(pack_options["label_pack"])
                label_widget.pack(**pack_opts)
                self.labels[f"{name}_label"] = label_widget
            
            # Create variable if not provided
            if not variable:
                variable = tk.StringVar()
            self.variables[name] = variable
            
            # Create dropdown
            dropdown = ttk.Combobox(self.control_frame, textvariable=variable, 
                                  values=values, state="readonly")
            
            # Pack dropdown
            pack_opts = {"side": "left", "padx": (0, 10)}
            if pack_options and "dropdown_pack" in pack_options:
                pack_opts.update(pack_options["dropdown_pack"])
            dropdown.pack(**pack_opts)
            
            # Bind callback
            if callback:
                self.control_callbacks[name] = callback
                dropdown.bind('<<ComboboxSelected>>', lambda e: callback())
            
            self.dropdowns[name] = dropdown
            
            print(f"DEBUG: MenuWidget - dropdown added: {name}")
            return dropdown
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error adding dropdown {name}: {e}")
            return None
    
    def add_button(self, name: str, text: str, command: Callable = None,
                  pack_options: Dict = None, **kwargs) -> ttk.Button:
        """Add a button control."""
        try:
            if not self.control_frame:
                print("ERROR: MenuWidget - no control frame available")
                return None
            
            # Create button
            button = ttk.Button(self.control_frame, text=text, command=command, **kwargs)
            
            # Pack button
            pack_opts = {"side": "left", "padx": 5}
            if pack_options:
                pack_opts.update(pack_options)
            button.pack(**pack_opts)
            
            # Store callback
            if command:
                self.control_callbacks[name] = command
            
            self.buttons[name] = button
            
            print(f"DEBUG: MenuWidget - button added: {name}")
            return button
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error adding button {name}: {e}")
            return None
    
    def add_checkbox(self, name: str, text: str, variable: tk.BooleanVar = None,
                    callback: Callable = None, pack_options: Dict = None) -> ttk.Checkbutton:
        """Add a checkbox control."""
        try:
            if not self.control_frame:
                print("ERROR: MenuWidget - no control frame available")
                return None
            
            # Create variable if not provided
            if not variable:
                variable = tk.BooleanVar()
            self.variables[name] = variable
            
            # Create checkbox
            checkbox = ttk.Checkbutton(self.control_frame, text=text, variable=variable)
            
            # Pack checkbox
            pack_opts = {"side": "left", "padx": 5}
            if pack_options:
                pack_opts.update(pack_options)
            checkbox.pack(**pack_opts)
            
            # Bind callback
            if callback:
                self.control_callbacks[name] = callback
                checkbox.configure(command=callback)
            
            self.checkboxes[name] = checkbox
            
            print(f"DEBUG: MenuWidget - checkbox added: {name}")
            return checkbox
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error adding checkbox {name}: {e}")
            return None
    
    def create_standard_controls(self, parent: tk.Widget, callbacks: Dict[str, Callable],
                               file_list: List[str] = None, sheet_list: List[str] = None,
                               plot_options: List[str] = None) -> ttk.Frame:
        """Create standard control layout."""
        try:
            # Create control frame
            control_frame = self.create_control_frame(parent)
            if not control_frame:
                return None
            
            # Left side controls
            left_frame = ttk.Frame(control_frame)
            left_frame.pack(side="left", fill="x", expand=True)
            
            # Temporarily set control_frame to left_frame for adding controls
            original_frame = self.control_frame
            self.control_frame = left_frame
            
            # File dropdown
            if file_list is not None:
                self.add_dropdown("file", "File:", file_list, 
                                callback=callbacks.get("file_changed"),
                                pack_options={"dropdown_pack": {"width": 30}})
            
            # Sheet dropdown
            if sheet_list is not None:
                self.add_dropdown("sheet", "Sheet:", sheet_list,
                                callback=callbacks.get("sheet_changed"),
                                pack_options={"dropdown_pack": {"width": 25}})
            
            # Plot dropdown
            if plot_options is not None:
                self.add_dropdown("plot", "Plot:", plot_options,
                                callback=callbacks.get("plot_changed"),
                                pack_options={"dropdown_pack": {"width": 15}})
            
            # Right side controls
            right_frame = ttk.Frame(control_frame)
            right_frame.pack(side="right")
            
            # Set control_frame to right_frame for right-side controls
            self.control_frame = right_frame
            
            # Crop checkbox
            self.add_checkbox("crop", "Enable Image Cropping",
                            callback=callbacks.get("crop_changed"))
            
            # Restore original control frame
            self.control_frame = original_frame
            
            print("DEBUG: MenuWidget - standard controls created")
            return control_frame
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error creating standard controls: {e}")
            return None
    
    # Update methods for dynamic control content
    def update_dropdown_values(self, dropdown_name: str, values: List[str], 
                             selected_value: str = None) -> bool:
        """Update dropdown values."""
        try:
            if dropdown_name not in self.dropdowns:
                print(f"ERROR: MenuWidget - dropdown not found: {dropdown_name}")
                return False
            
            dropdown = self.dropdowns[dropdown_name]
            current_value = dropdown.get()
            
            # Update values
            dropdown['values'] = values
            
            # Set selection
            if selected_value and selected_value in values:
                dropdown.set(selected_value)
            elif current_value in values:
                dropdown.set(current_value)
            elif values:
                dropdown.set(values[0])
            else:
                dropdown.set("")
            
            print(f"DEBUG: MenuWidget - dropdown updated: {dropdown_name} with {len(values)} values")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error updating dropdown {dropdown_name}: {e}")
            return False
    
    def set_dropdown_value(self, dropdown_name: str, value: str) -> bool:
        """Set dropdown value."""
        try:
            if dropdown_name not in self.dropdowns:
                return False
            
            if dropdown_name in self.variables:
                self.variables[dropdown_name].set(value)
            else:
                self.dropdowns[dropdown_name].set(value)
            
            print(f"DEBUG: MenuWidget - dropdown value set: {dropdown_name} = {value}")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error setting dropdown value: {e}")
            return False
    
    def get_dropdown_value(self, dropdown_name: str) -> Optional[str]:
        """Get dropdown value."""
        try:
            if dropdown_name in self.variables:
                return self.variables[dropdown_name].get()
            elif dropdown_name in self.dropdowns:
                return self.dropdowns[dropdown_name].get()
            return None
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error getting dropdown value: {e}")
            return None
    
    def set_button_state(self, button_name: str, state: str) -> bool:
        """Set button state (normal, disabled)."""
        try:
            if button_name not in self.buttons:
                return False
            
            self.buttons[button_name].configure(state=state)
            print(f"DEBUG: MenuWidget - button state set: {button_name} = {state}")
            return True
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error setting button state: {e}")
            return False
    
    def set_checkbox_value(self, checkbox_name: str, value: bool) -> bool:
        """Set checkbox value."""
        try:
            if checkbox_name in self.variables:
                self.variables[checkbox_name].set(value)
                print(f"DEBUG: MenuWidget - checkbox value set: {checkbox_name} = {value}")
                return True
            return False
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error setting checkbox value: {e}")
            return False
    
    def get_checkbox_value(self, checkbox_name: str) -> Optional[bool]:
        """Get checkbox value."""
        try:
            if checkbox_name in self.variables:
                return self.variables[checkbox_name].get()
            return None
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error getting checkbox value: {e}")
            return None
    
    # Show/hide controls
    def show_control(self, control_name: str) -> bool:
        """Show a control widget."""
        try:
            widget = None
            if control_name in self.dropdowns:
                widget = self.dropdowns[control_name]
            elif control_name in self.buttons:
                widget = self.buttons[control_name]
            elif control_name in self.checkboxes:
                widget = self.checkboxes[control_name]
            elif f"{control_name}_label" in self.labels:
                widget = self.labels[f"{control_name}_label"]
            
            if widget:
                widget.pack()
                print(f"DEBUG: MenuWidget - control shown: {control_name}")
                return True
            return False
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error showing control: {e}")
            return False
    
    def hide_control(self, control_name: str) -> bool:
        """Hide a control widget."""
        try:
            widget = None
            if control_name in self.dropdowns:
                widget = self.dropdowns[control_name]
            elif control_name in self.buttons:
                widget = self.buttons[control_name]
            elif control_name in self.checkboxes:
                widget = self.checkboxes[control_name]
            elif f"{control_name}_label" in self.labels:
                widget = self.labels[f"{control_name}_label"]
            
            if widget:
                widget.pack_forget()
                print(f"DEBUG: MenuWidget - control hidden: {control_name}")
                return True
            return False
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error hiding control: {e}")
            return False
    
    # Cleanup
    def clear_all_controls(self):
        """Clear all control widgets."""
        try:
            if self.control_frame:
                for widget in self.control_frame.winfo_children():
                    widget.destroy()
            
            self.dropdowns.clear()
            self.buttons.clear()
            self.checkboxes.clear()
            self.labels.clear()
            self.variables.clear()
            self.control_callbacks.clear()
            
            print("DEBUG: MenuWidget - all controls cleared")
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error clearing controls: {e}")
    
    def destroy_menubar(self):
        """Destroy the menubar."""
        try:
            if self.menubar:
                self.menubar.destroy()
                self.menubar = None
            
            self.menus.clear()
            self.menu_callbacks.clear()
            
            print("DEBUG: MenuWidget - menubar destroyed")
            
        except Exception as e:
            print(f"ERROR: MenuWidget - error destroying menubar: {e}")
    
    def get_menu(self, menu_name: str) -> Optional[tk.Menu]:
        """Get a menu by name."""
        return self.menus.get(menu_name)
    
    def get_control_widget(self, control_name: str) -> Optional[tk.Widget]:
        """Get a control widget by name."""
        if control_name in self.dropdowns:
            return self.dropdowns[control_name]
        elif control_name in self.buttons:
            return self.buttons[control_name]
        elif control_name in self.checkboxes:
            return self.checkboxes[control_name]
        elif f"{control_name}_label" in self.labels:
            return self.labels[f"{control_name}_label"]
        return None
    
    def get_all_dropdown_values(self) -> Dict[str, str]:
        """Get all dropdown values."""
        values = {}
        for name in self.dropdowns:
            values[name] = self.get_dropdown_value(name)
        return values
    
    def get_all_checkbox_values(self) -> Dict[str, bool]:
        """Get all checkbox values."""
        values = {}
        for name in self.checkboxes:
            values[name] = self.get_checkbox_value(name)
        return values
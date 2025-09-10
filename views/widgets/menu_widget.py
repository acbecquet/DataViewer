# views/widgets/menu_widget.py
"""
views/widgets/menu_widget.py
Menu bar widget.
This will contain menu creation logic from main_gui.py.
"""

import tkinter as tk
from typing import Dict, Callable, Optional


class MenuWidget:
    """Widget for creating and managing application menus."""
    
    def __init__(self, parent: tk.Tk):
        """Initialize the menu widget."""
        self.parent = parent
        self.menubar: Optional[tk.Menu] = None
        self.menus: Dict[str, tk.Menu] = {}
        self.callbacks: Dict[str, Callable] = {}
        
        print("DEBUG: MenuWidget initialized")
    
    def create_menubar(self):
        """Create the main menu bar."""
        self.menubar = tk.Menu(self.parent)
        self.parent.config(menu=self.menubar)
        
        print("DEBUG: MenuWidget created menubar")
    
    def add_menu(self, menu_name: str, label: str):
        """Add a menu to the menu bar."""
        if not self.menubar:
            self.create_menubar()
        
        menu = tk.Menu(self.menubar, tearoff=0)
        self.menus[menu_name] = menu
        self.menubar.add_cascade(label=label, menu=menu)
        
        print(f"DEBUG: MenuWidget added menu: {menu_name}")
    
    def add_menu_item(self, menu_name: str, label: str, command_name: str, 
                     accelerator: str = "", separator_before: bool = False):
        """Add an item to a menu."""
        if menu_name not in self.menus:
            print(f"WARNING: MenuWidget - menu {menu_name} not found")
            return
        
        menu = self.menus[menu_name]
        
        if separator_before:
            menu.add_separator()
        
        command = self.callbacks.get(command_name, lambda: print(f"Command {command_name} not implemented"))
        
        menu.add_command(
            label=label,
            command=command,
            accelerator=accelerator
        )
        
        print(f"DEBUG: MenuWidget added item '{label}' to {menu_name}")
    
    def set_callback(self, command_name: str, callback: Callable):
        """Set a callback for a menu command."""
        self.callbacks[command_name] = callback
        print(f"DEBUG: MenuWidget set callback for {command_name}")
    
    def set_callbacks(self, callbacks: Dict[str, Callable]):
        """Set multiple callbacks at once."""
        self.callbacks.update(callbacks)
        print(f"DEBUG: MenuWidget set {len(callbacks)} callbacks")
    
    def create_standard_menus(self):
        """Create the standard application menus."""
        print("DEBUG: MenuWidget creating standard menus")
        
        # File menu
        self.add_menu("file", "File")
        self.add_menu_item("file", "New", "file_new", "Ctrl+N")
        self.add_menu_item("file", "Load Excel", "file_load_excel", "Ctrl+O")
        self.add_menu_item("file", "Load VAP3", "file_load_vap3")
        self.add_menu_item("file", "Save As VAP3", "file_save_vap3", "Ctrl+S", separator_before=True)
        self.add_menu_item("file", "Exit", "file_exit", "Ctrl+Q", separator_before=True)
        
        # View menu
        self.add_menu("view", "View")
        self.add_menu_item("view", "View Raw Data", "view_raw_data")
        self.add_menu_item("view", "Trend Analysis", "view_trend_analysis")
        
        # Database menu
        self.add_menu("database", "Database")
        self.add_menu_item("database", "Browse Database", "db_browse")
        self.add_menu_item("database", "Load from Database", "db_load")
        
        # Calculate menu
        self.add_menu("calculate", "Calculate")
        self.add_menu_item("calculate", "Viscosity", "calc_viscosity")
        
        # Reports menu
        self.add_menu("reports", "Reports")
        self.add_menu_item("reports", "Generate Test Report", "report_test")
        self.add_menu_item("reports", "Generate Full Report", "report_full")
        
        # Help menu
        self.add_menu("help", "Help")
        self.add_menu_item("help", "Help", "help_help", "F1")
        self.add_menu_item("help", "About", "help_about", separator_before=True)
        
        print("DEBUG: MenuWidget standard menus created")
    
    def enable_menu_item(self, menu_name: str, label: str, enabled: bool = True):
        """Enable or disable a menu item."""
        if menu_name not in self.menus:
            return
        
        menu = self.menus[menu_name]
        
        # Find menu item by label
        for i in range(menu.index("end") + 1):
            try:
                if menu.entrycget(i, "label") == label:
                    state = "normal" if enabled else "disabled"
                    menu.entryconfig(i, state=state)
                    print(f"DEBUG: MenuWidget {label} {'enabled' if enabled else 'disabled'}")
                    break
            except tk.TclError:
                continue
    
    def get_menubar(self) -> Optional[tk.Menu]:
        """Get the menu bar."""
        return self.menubar

"""
viscosity_calculator.py
Module for calculating terpene percentages based on viscosity.
"""
import tkinter as tk
from tkinter import ttk, Toplevel
from utils import FONT, APP_BACKGROUND_COLOR, BUTTON_COLOR

class ViscosityCalculator:
    def __init__(self, gui):
        """
        Initialize the ViscosityCalculator with a reference to the main GUI.
        
        Args:
            gui (TestingGUI): The main application GUI instance
        """
        self.gui = gui
        self.root = gui.root
        
    def show_calculator(self):
        """
        Show the viscosity calculator window.
        
        Returns:
            tk.Toplevel: The calculator window
        """
        # Create a new top-level window
        calculator_window = Toplevel(self.root)
        calculator_window.title("Calculate Terpene % for Viscosity")
        calculator_window.geometry("400x400")
        calculator_window.resizable(False, False)
        calculator_window.configure(bg=APP_BACKGROUND_COLOR)
        
        # Make the window modal
        calculator_window.transient(self.root)
        calculator_window.grab_set()
        
        # Center the window on the screen
        self.gui.center_window(calculator_window, 400, 400)
        
        # For now, just create an empty window
        # Content will be added in future implementations
        
        return calculator_window
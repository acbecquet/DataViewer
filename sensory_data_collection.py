"""
Sensory Data Collection Window for DataViewer Application
Developed by Charlie Becquet
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
import math
from datetime import datetime
import json
import os
from utils import APP_BACKGROUND_COLOR, BUTTON_COLOR, FONT, debug_print, show_success_message

from sensory_ml_training import SensoryMLTrainer, SensoryAIProcessor





class SensoryDataCollectionWindow:
    """Main window for sensory data collection and visualization."""

    def __init__(self, parent, close_callback=None):
        self.parent = parent
        self.close_callback = close_callback
        self.window = None
        self.data = {}
        self.sessions = {}
        self.current_session_id = None
        self.session_counter = 1
        self.samples = {}
        self.current_sample = None
        debug_print("DEBUG: Initialized session-based data structure")

        # Sensory metrics
        self.metrics = [
            "Burnt Taste",
            "Vapor Volume",
            "Overall Flavor",
            "Smoothness",
            "Overall Liking"
        ]

        self.current_mode = "collection"
        self.all_sessions_data = {}
        self.average_samples = {}


        self.ml_trainer = SensoryMLTrainer(self)
        self.ai_processor = SensoryAIProcessor(self)


        # Header data fields
        self.header_fields = [
            "Assessor Name",
            "Media",
            "Puff Length",
            "Date"
        ]

        # SOP text
        self.sop_text = """
SENSORY EVALUATION STANDARD OPERATING PROCEDURE

1. PREPARATION:
   - Ensure all cartridges are at room temperature
   - Use clean, odor-free environment, ideally in a fume hood

2. EVALUATION:
   - Take 2-3 moderate puffs per sample
   - Be sure to compare flavor with original oil if available
   - Rate each attribute on 1-9 scale (1=poor, 9=excellent)

3. SCALE INTERPRETATION:
   - Burnt Taste: 1=Very Burnt, 9=No Burnt Taste
   - Vapor Volume: 1=Very Low, 9=Very High
   - Overall Flavor: 1=Poor, 9=Excellent
   - Smoothness: 1=Very Harsh, 9=Very Smooth
   - Overall Liking: 1=Dislike Extremely, 9=Like Extremely

4. NOTES:
   - Record any unusual observations
   - Note any technical issues with samples
        """


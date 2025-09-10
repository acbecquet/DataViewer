# models/report_model.py
"""
models/report_model.py
Report generation models and configuration.
These models will replace report-related structures from report_generator.py.
"""

from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class ReportConfig:
    """Configuration for report generation."""
    report_type: str = "test"  # test, full, comparison
    include_plots: bool = True
    include_images: bool = True
    include_raw_data: bool = False
    output_format: str = "excel"  # excel, powerpoint, pdf
    template_path: Optional[str] = None
    output_directory: str = "reports"
    filename_prefix: str = "report"
    created_at: datetime = field(default_factory=datetime.now)
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created ReportConfig for {self.report_type} report")
        print(f"DEBUG: Output format: {self.output_format}, Include plots: {self.include_plots}")


@dataclass
class ReportData:
    """Container for report data and metadata."""
    title: str
    sheets_data: Dict[str, Any] = field(default_factory=dict)
    plot_data: Dict[str, Any] = field(default_factory=dict)
    image_data: Dict[str, Any] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)
    generation_time: Optional[datetime] = None
    file_path: Optional[str] = None
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created ReportData '{self.title}' with {len(self.sheets_data)} sheets")


class ReportModel:
    """Main report model for managing report generation."""
    
    def __init__(self):
        """Initialize the report model."""
        self.config = ReportConfig()
        self.report_data: Optional[ReportData] = None
        self.generation_progress = 0
        self.is_generating = False
        
        print("DEBUG: ReportModel initialized")
    
    def set_config(self, config: ReportConfig):
        """Set the report configuration."""
        self.config = config
        print(f"DEBUG: Set report config for {config.report_type} report")
    
    def start_generation(self, report_data: ReportData):
        """Start report generation process."""
        self.report_data = report_data
        self.is_generating = True
        self.generation_progress = 0
        print(f"DEBUG: Started report generation for '{report_data.title}'")
    
    def update_progress(self, progress: int):
        """Update generation progress."""
        self.generation_progress = progress
        print(f"DEBUG: Report generation progress: {progress}%")
    
    def complete_generation(self, file_path: str):
        """Complete report generation."""
        if self.report_data:
            self.report_data.file_path = file_path
            self.report_data.generation_time = datetime.now()
        self.is_generating = False
        self.generation_progress = 100
        print(f"DEBUG: Report generation completed: {file_path}")
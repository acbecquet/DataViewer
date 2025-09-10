# controllers/report_controller.py
"""
controllers/report_controller.py
Report generation controller that coordinates report creation.
This replaces the coordination logic currently in report_generator.py.
"""

from typing import Optional, Dict, Any, List
from models.data_model import DataModel
from models.report_model import ReportModel, ReportConfig, ReportData


class ReportController:
    """Controller for report generation."""
    
    def __init__(self, data_model: DataModel, report_service: Any):
        """Initialize the report controller."""
        self.data_model = data_model
        self.report_service = report_service
        self.report_model = ReportModel()
        
        # Cross-controller references
        self.file_controller: Optional['FileController'] = None
        self.plot_controller: Optional['PlotController'] = None
        
        print("DEBUG: ReportController initialized")
    
    def set_file_controller(self, file_controller: 'FileController'):
        """Set reference to file controller."""
        self.file_controller = file_controller
        print("DEBUG: ReportController connected to FileController")
    
    def set_plot_controller(self, plot_controller: 'PlotController'):
        """Set reference to plot controller."""
        self.plot_controller = plot_controller
        print("DEBUG: ReportController connected to PlotController")
    
    def generate_test_report(self, sheet_name: str) -> bool:
        """Generate a test report for a specific sheet."""
        print(f"DEBUG: ReportController generating test report for {sheet_name}")
        
        try:
            # Create report configuration
            config = ReportConfig(
                report_type="test",
                output_format="excel"
            )
            self.report_model.set_config(config)
            
            # Get sheet data
            sheet_data = self.data_model.get_current_filtered_sheets().get(sheet_name)
            if not sheet_data:
                print(f"ERROR: Sheet {sheet_name} not found")
                return False
            
            # Create report data
            report_data = ReportData(
                title=f"Test Report - {sheet_name}",
                sheets_data={sheet_name: sheet_data}
            )
            
            # Start generation
            self.report_model.start_generation(report_data)
            
            # Generate through service (placeholder)
            # result = self.report_service.generate_report(report_data, config)
            
            # Complete generation
            output_path = f"reports/test_report_{sheet_name}.xlsx"
            self.report_model.complete_generation(output_path)
            
            print(f"DEBUG: ReportController completed test report: {output_path}")
            return True
            
        except Exception as e:
            print(f"ERROR: ReportController failed to generate test report: {e}")
            return False
    
    def generate_full_report(self) -> bool:
        """Generate a full report with all data."""
        print("DEBUG: ReportController generating full report")
        
        try:
            # Create report configuration
            config = ReportConfig(
                report_type="full",
                output_format="powerpoint",
                include_plots=True,
                include_images=True
            )
            self.report_model.set_config(config)
            
            # Get all data
            all_sheets = self.data_model.get_current_filtered_sheets()
            
            # Create report data
            report_data = ReportData(
                title="Full Data Report",
                sheets_data=all_sheets
            )
            
            # Generate report
            self.report_model.start_generation(report_data)
            
            # Generate through service (placeholder)
            # result = self.report_service.generate_full_report(report_data, config)
            
            # Complete generation
            output_path = "reports/full_report.pptx"
            self.report_model.complete_generation(output_path)
            
            print(f"DEBUG: ReportController completed full report: {output_path}")
            return True
            
        except Exception as e:
            print(f"ERROR: ReportController failed to generate full report: {e}")
            return False
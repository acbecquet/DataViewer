# utils/config.py
"""
utils/config.py
Configuration management for the DataViewer application.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class DatabaseConfig:
    """Database configuration settings."""
    type: str = "sqlite"
    host: str = "localhost"
    port: int = 5432
    database: str = "dataviewer.db"
    username: str = ""
    password: str = ""
    timeout: int = 30


@dataclass
class UIConfig:
    """UI configuration settings."""
    theme: str = "default"
    window_width: int = 1200
    window_height: int = 800
    font_family: str = "Arial"
    font_size: int = 10
    auto_save: bool = True
    show_tooltips: bool = True


@dataclass
class ProcessingConfig:
    """Data processing configuration settings."""
    default_plot_type: str = "TPM"
    auto_detect_sheets: bool = True
    cache_processed_data: bool = True
    max_cache_size_mb: int = 100
    processing_timeout: int = 60


@dataclass
class ReportConfig:
    """Report generation configuration settings."""
    default_format: str = "excel"
    include_plots: bool = True
    include_images: bool = True
    include_raw_data: bool = False
    output_directory: str = "reports"
    dpi: int = 300


class Config:
    """Main configuration manager."""
    
    def __init__(self, config_file: str = "config.json"):
        """Initialize configuration manager."""
        self.config_file = config_file
        self.config_dir = Path("config")
        self.full_config_path = self.config_dir / config_file
        
        # Configuration sections
        self.database = DatabaseConfig()
        self.ui = UIConfig()
        self.processing = ProcessingConfig()
        self.reports = ReportConfig()
        
        # Application metadata
        self.app_version = "3.0.0"
        self.last_updated = datetime.now()
        
        print("DEBUG: Config initialized")
        print(f"DEBUG: Config file path: {self.full_config_path}")
        
        # Load existing config if available
        self.load_config()
    
    def load_config(self) -> bool:
        """Load configuration from file."""
        try:
            if not self.full_config_path.exists():
                print("DEBUG: Config file not found, using defaults")
                return self.save_config()  # Create default config
            
            with open(self.full_config_path, 'r') as f:
                config_data = json.load(f)
            
            # Load database config
            if 'database' in config_data:
                db_data = config_data['database']
                self.database = DatabaseConfig(**db_data)
            
            # Load UI config
            if 'ui' in config_data:
                ui_data = config_data['ui']
                self.ui = UIConfig(**ui_data)
            
            # Load processing config
            if 'processing' in config_data:
                proc_data = config_data['processing']
                self.processing = ProcessingConfig(**proc_data)
            
            # Load reports config
            if 'reports' in config_data:
                report_data = config_data['reports']
                self.reports = ReportConfig(**report_data)
            
            # Load metadata
            self.app_version = config_data.get('app_version', self.app_version)
            last_updated_str = config_data.get('last_updated')
            if last_updated_str:
                self.last_updated = datetime.fromisoformat(last_updated_str)
            
            print("DEBUG: Config loaded successfully")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to load config: {e}")
            return False
    
    def save_config(self) -> bool:
        """Save current configuration to file."""
        try:
            # Ensure config directory exists
            self.config_dir.mkdir(parents=True, exist_ok=True)
            
            # Prepare config data
            config_data = {
                'database': asdict(self.database),
                'ui': asdict(self.ui),
                'processing': asdict(self.processing),
                'reports': asdict(self.reports),
                'app_version': self.app_version,
                'last_updated': self.last_updated.isoformat()
            }
            
            # Write to file
            with open(self.full_config_path, 'w') as f:
                json.dump(config_data, f, indent=2)
            
            print(f"DEBUG: Config saved to {self.full_config_path}")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to save config: {e}")
            return False
    
    def get_database_connection_string(self) -> str:
        """Get database connection string."""
        if self.database.type.lower() == "sqlite":
            return f"sqlite:///{self.database.database}"
        elif self.database.type.lower() == "postgresql":
            return f"postgresql://{self.database.username}:{self.database.password}@{self.database.host}:{self.database.port}/{self.database.database}"
        else:
            return ""
    
    def update_database_config(self, **kwargs):
        """Update database configuration."""
        for key, value in kwargs.items():
            if hasattr(self.database, key):
                setattr(self.database, key, value)
                print(f"DEBUG: Updated database.{key} = {value}")
        
        self.last_updated = datetime.now()
    
    def update_ui_config(self, **kwargs):
        """Update UI configuration."""
        for key, value in kwargs.items():
            if hasattr(self.ui, key):
                setattr(self.ui, key, value)
                print(f"DEBUG: Updated ui.{key} = {value}")
        
        self.last_updated = datetime.now()
    
    def reset_to_defaults(self):
        """Reset configuration to defaults."""
        self.database = DatabaseConfig()
        self.ui = UIConfig()
        self.processing = ProcessingConfig()
        self.reports = ReportConfig()
        self.last_updated = datetime.now()
        
        print("DEBUG: Config reset to defaults")
    
    def validate_config(self) -> tuple[bool, list[str]]:
        """Validate current configuration."""
        issues = []
        
        # Validate database config
        if not self.database.database:
            issues.append("Database name is empty")
        
        # Validate UI config
        if self.ui.window_width < 800:
            issues.append("Window width too small (minimum 800)")
        
        if self.ui.window_height < 600:
            issues.append("Window height too small (minimum 600)")
        
        # Validate processing config
        if self.processing.max_cache_size_mb < 10:
            issues.append("Cache size too small (minimum 10MB)")
        
        # Validate reports config
        if not self.reports.output_directory:
            issues.append("Report output directory not specified")
        
        is_valid = len(issues) == 0
        print(f"DEBUG: Config validation - {'VALID' if is_valid else 'INVALID'} ({len(issues)} issues)")
        
        return is_valid, issues
    
    def get_config_summary(self) -> Dict[str, Any]:
        """Get configuration summary."""
        return {
            'app_version': self.app_version,
            'last_updated': self.last_updated.isoformat(),
            'config_file': str(self.full_config_path),
            'database_type': self.database.type,
            'ui_theme': self.ui.theme,
            'default_plot_type': self.processing.default_plot_type,
            'report_format': self.reports.default_format
        }


# Global config instance
_config_instance: Optional[Config] = None


def get_config() -> Config:
    """Get the global configuration instance."""
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    return _config_instance


def reload_config():
    """Reload configuration from file."""
    global _config_instance
    if _config_instance:
        _config_instance.load_config()
        print("DEBUG: Config reloaded")
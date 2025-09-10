# utils/config.py
"""
utils/config.py
Configuration management for the DataViewer application.
Enhanced with Synology database support and offline mode.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class DatabaseConfig:
    """Database configuration settings with enhanced Synology support."""
    type: str = "sqlite"
    host: str = "localhost"
    port: int = 5432
    database: str = "dataviewer.db"
    username: str = ""
    password: str = ""
    timeout: int = 30
    
    # Enhanced database settings from refactor_files
    max_retries: int = 3
    retry_delay: float = 2.0
    wal_mode: bool = True  # Better for network/concurrent access
    cache_size: int = 10000  # Larger cache for network latency
    synchronous_mode: str = "NORMAL"  # Balance safety vs performance
    
    # Synology-specific settings
    synology_host: str = ""
    synology_port: int = 5432
    synology_database: str = "dataviewer"
    synology_username: str = ""
    synology_password: str = ""
    connection_timeout: int = 30


@dataclass
class OfflineModeConfig:
    """Offline mode configuration settings."""
    enabled: bool = True
    cache_duration_hours: int = 24
    essential_functions_only: bool = False
    fallback_behavior: str = "offline_mode"  # "offline_mode" or "error"
    auto_sync_on_reconnect: bool = True


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
    
    # Analysis configuration from sample_comparison.py
    model_keywords_raw: list = None
    grouped_tests: list = None
    
    def __post_init__(self):
        """Initialize default values for analysis configuration."""
        if self.model_keywords_raw is None:
            self.model_keywords_raw = []
        if self.grouped_tests is None:
            self.grouped_tests = []


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
    """Main configuration manager with enhanced database and offline support."""
    
    def __init__(self, config_file: str = "config.json"):
        """Initialize configuration manager."""
        self.config_file = config_file
        self.config_dir = Path("config")
        self.full_config_path = self.config_dir / config_file
        
        # Configuration sections
        self.database = DatabaseConfig()
        self.offline_mode = OfflineModeConfig()  # New offline mode config
        self.ui = UIConfig()
        self.processing = ProcessingConfig()
        self.reports = ReportConfig()
        
        # Application metadata
        self.app_version = "3.0.0"
        self.last_updated = datetime.now()
        
        print("DEBUG: Config initialized with enhanced database support")
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
            
            # Load offline mode config (new)
            if 'offline_mode' in config_data:
                offline_data = config_data['offline_mode']
                self.offline_mode = OfflineModeConfig(**offline_data)
            
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
                'offline_mode': asdict(self.offline_mode),  # New offline mode
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
        """Get database connection string with enhanced support."""
        if self.database.type.lower() == "sqlite":
            return f"sqlite:///{self.database.database}"
        elif self.database.type.lower() == "postgresql":
            return f"postgresql://{self.database.username}:{self.database.password}@{self.database.host}:{self.database.port}/{self.database.database}"
        elif self.database.type.lower() == "synology":
            # Enhanced Synology support
            return f"postgresql://{self.database.synology_username}:{self.database.synology_password}@{self.database.synology_host}:{self.database.synology_port}/{self.database.synology_database}"
        else:
            return ""
    
    def get_synology_connection_string(self) -> str:
        """Get Synology-specific database connection string."""
        return f"postgresql://{self.database.synology_username}:{self.database.synology_password}@{self.database.synology_host}:{self.database.synology_port}/{self.database.synology_database}"
    
    def update_database_config(self, **kwargs):
        """Update database configuration."""
        for key, value in kwargs.items():
            if hasattr(self.database, key):
                setattr(self.database, key, value)
                print(f"DEBUG: Updated database.{key} = {value}")
        
        self.last_updated = datetime.now()
    
    def update_synology_config(self, host: str, port: int, database: str, username: str, password: str):
        """Update Synology database configuration."""
        self.database.synology_host = host
        self.database.synology_port = port
        self.database.synology_database = database
        self.database.synology_username = username
        self.database.synology_password = password
        self.last_updated = datetime.now()
        
        print(f"DEBUG: Updated Synology config: {host}:{port}/{database}")
    
    def update_offline_mode_config(self, **kwargs):
        """Update offline mode configuration."""
        for key, value in kwargs.items():
            if hasattr(self.offline_mode, key):
                setattr(self.offline_mode, key, value)
                print(f"DEBUG: Updated offline_mode.{key} = {value}")
        
        self.last_updated = datetime.now()
    
    def update_ui_config(self, **kwargs):
        """Update UI configuration."""
        for key, value in kwargs.items():
            if hasattr(self.ui, key):
                setattr(self.ui, key, value)
                print(f"DEBUG: Updated ui.{key} = {value}")
        
        self.last_updated = datetime.now()
    
    def update_processing_config(self, **kwargs):
        """Update processing configuration."""
        for key, value in kwargs.items():
            if hasattr(self.processing, key):
                setattr(self.processing, key, value)
                print(f"DEBUG: Updated processing.{key} = {value}")
        
        self.last_updated = datetime.now()
    
    def update_analysis_config(self, model_keywords: list = None, grouped_tests: list = None):
        """Update analysis configuration (from sample_comparison.py)."""
        if model_keywords is not None:
            self.processing.model_keywords_raw = model_keywords
            print(f"DEBUG: Updated model keywords: {len(model_keywords)} items")
        
        if grouped_tests is not None:
            self.processing.grouped_tests = grouped_tests
            print(f"DEBUG: Updated grouped tests: {len(grouped_tests)} items")
        
        self.last_updated = datetime.now()
    
    def reset_to_defaults(self):
        """Reset configuration to defaults."""
        self.database = DatabaseConfig()
        self.offline_mode = OfflineModeConfig()
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
        
        if self.database.timeout < 5:
            issues.append("Database timeout too low (minimum 5 seconds)")
        
        # Validate Synology config if using Synology
        if self.database.type.lower() == "synology":
            if not self.database.synology_host:
                issues.append("Synology host not specified")
            if not self.database.synology_username:
                issues.append("Synology username not specified")
        
        # Validate offline mode config
        if self.offline_mode.cache_duration_hours < 1:
            issues.append("Cache duration too low (minimum 1 hour)")
        
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
            'synology_configured': bool(self.database.synology_host),
            'offline_mode_enabled': self.offline_mode.enabled,
            'ui_theme': self.ui.theme,
            'default_plot_type': self.processing.default_plot_type,
            'report_format': self.reports.default_format,
            'analysis_keywords_count': len(self.processing.model_keywords_raw),
            'grouped_tests_count': len(self.processing.grouped_tests)
        }
    
    def create_synology_sample_config(self) -> bool:
        """Create a sample Synology configuration file."""
        try:
            sample_config_path = self.config_dir / "synology_sample.json"
            
            sample_config = {
                "database_type": "synology",
                "synology_settings": {
                    "host": "your-synology-ip-or-hostname",
                    "port": 5432,
                    "database": "dataviewer",
                    "username": "your_username",
                    "password": "your_password",
                    "connection_timeout": 30
                },
                "offline_mode": {
                    "enabled": True,
                    "cache_duration_hours": 24,
                    "essential_functions_only": False,
                    "fallback_behavior": "offline_mode"
                }
            }
            
            self.config_dir.mkdir(parents=True, exist_ok=True)
            
            with open(sample_config_path, 'w') as f:
                json.dump(sample_config, f, indent=2)
            
            print(f"DEBUG: Sample Synology config created: {sample_config_path}")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to create sample Synology config: {e}")
            return False


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


def create_database_config_file():
    """Create database configuration file (from release_workflow.py)."""
    config = get_config()
    return config.create_synology_sample_config()


# Initialize config on import
print("DEBUG: config.py - Configuration system loaded")
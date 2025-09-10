# models/database_model.py
"""
models/database_model.py
Database models and connection management.
These models will replace database-related structures from file_manager.py.
"""

from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class DatabaseConfig:
    """Configuration for database connections."""
    database_type: str = "sqlite"  # sqlite, postgresql, mysql
    host: Optional[str] = None
    port: Optional[int] = None
    database: str = "dataviewer.db"
    username: Optional[str] = None
    password: Optional[str] = None
    connection_timeout: int = 30
    max_connections: int = 10
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created DatabaseConfig for {self.database_type} database")
        if self.host:
            print(f"DEBUG: Remote database: {self.host}:{self.port}/{self.database}")
        else:
            print(f"DEBUG: Local database: {self.database}")


@dataclass
class DatabaseRecord:
    """Model for database record metadata."""
    record_id: str
    filename: str
    file_path: str
    created_at: datetime
    updated_at: datetime
    file_size: int
    checksum: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created DatabaseRecord for '{self.filename}' (ID: {self.record_id})")


class DatabaseModel:
    """Main database model for managing connections and records."""
    
    def __init__(self, config: Optional[DatabaseConfig] = None):
        """Initialize the database model."""
        self.config = config or DatabaseConfig()
        self.connection = None
        self.is_connected = False
        self.stored_files: Dict[str, DatabaseRecord] = {}
        self.connection_error: Optional[str] = None
        
        print("DEBUG: DatabaseModel initialized")
        print(f"DEBUG: Database type: {self.config.database_type}")
    
    def connect(self) -> bool:
        """Connect to the database."""
        try:
            # Placeholder for actual database connection
            self.is_connected = True
            self.connection_error = None
            print(f"DEBUG: Connected to {self.config.database_type} database")
            return True
        except Exception as e:
            self.connection_error = str(e)
            print(f"ERROR: Failed to connect to database: {e}")
            return False
    
    def disconnect(self):
        """Disconnect from the database."""
        self.is_connected = False
        self.connection = None
        print("DEBUG: Disconnected from database")
    
    def add_stored_file(self, record: DatabaseRecord):
        """Add a file record to stored files."""
        self.stored_files[record.record_id] = record
        print(f"DEBUG: Added database record for '{record.filename}'")
    
    def get_stored_files(self) -> List[DatabaseRecord]:
        """Get all stored file records."""
        return list(self.stored_files.values())
    
    def find_record_by_filename(self, filename: str) -> Optional[DatabaseRecord]:
        """Find a database record by filename."""
        for record in self.stored_files.values():
            if record.filename == filename:
                return record
        return None
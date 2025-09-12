# models/database_model.py
"""
models/database_model.py
Database model for managing connections, operations, and data persistence.
Consolidated from database_manager.py, database_explorer.py, config database settings, and database operations.
"""

import os
import sys
import time
import json
import sqlite3
import shutil
import hashlib
import tempfile
import threading
import platform
from typing import Dict, List, Optional, Set, Any, Tuple, Union
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, Toplevel, Label, Button, ttk, Frame, Listbox, Scrollbar


def debug_print(message: str):
    """Debug print function for database operations."""
    print(f"DEBUG: DatabaseModel - {message}")


def get_synology_paths() -> List[str]:
    """Get potential Synology database paths based on the platform."""
    paths = []
    
    if platform.system() == "Windows":
        # Standard Synology Drive paths on Windows
        drive_letters = ['Z:', 'Y:', 'X:', 'W:', 'V:', 'U:', 'T:', 'S:']
        common_synology_paths = ['SDR-DataViewer', 'dataviewer', 'shared']
        
        for drive in drive_letters:
            for folder in common_synology_paths:
                paths.append(os.path.join(drive, folder))
    
    elif platform.system() == "Darwin":  # macOS
        # macOS Synology paths
        synology_bases = ['/Volumes/SDR-DataViewer', '/Volumes/dataviewer', '/Volumes/shared']
        paths.extend(synology_bases)
    
    elif platform.system() == "Linux":
        # Linux Synology mount points
        synology_bases = ['/mnt/synology', '/mount/synology', '/media/synology']
        paths.extend(synology_bases)
    
    debug_print(f"Generated {len(paths)} potential Synology paths")
    return paths


def get_database_path() -> str:
    """Determine the appropriate database path with Synology support."""
    DATABASE_FILENAME = "dataviewer.db"
    
    debug_print("Starting database path discovery...")
    
    # Method 1: Check for explicit Synology environment variable
    synology_path = os.environ.get('SYNOLOGY_DATABASE_PATH')
    if synology_path and os.path.exists(synology_path):
        debug_print(f"Using Synology path from environment: {synology_path}")
        return os.path.join(synology_path, DATABASE_FILENAME)
    
    # Method 2: Try Synology Drive paths
    debug_print("Checking for accessible Synology paths...")
    synology_paths = get_synology_paths()
    
    for path in synology_paths:
        try:
            if os.path.exists(path) and os.access(path, os.W_OK):
                db_path = os.path.join(path, DATABASE_FILENAME)
                debug_print(f"Found accessible Synology path: {db_path}")
                return db_path
        except Exception as e:
            debug_print(f"Cannot access {path}: {e}")
            continue
    
    # Method 3: Check user drive locations (Windows backup locations)
    if platform.system() == "Windows":
        user_drives = ['D:', 'E:', 'F:', 'G:']
        for drive in user_drives:
            for folder in ['Synology Drive', 'SynologyDrive', 'SDR-DataViewer']:
                try:
                    path = os.path.join(drive, folder)
                    if os.path.exists(path) and os.access(path, os.W_OK):
                        db_path = os.path.join(path, DATABASE_FILENAME)
                        debug_print(f"Found backup Synology Drive path: {db_path}")
                        return db_path
                except:
                    continue
    
    # Method 4: Fallback to local database
    debug_print("WARNING: No accessible shared database found")
    debug_print("INFO: Using local database - changes won't be shared with team")
    
    # Determine appropriate local database location
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller environment - use user's app data directory
        if platform.system() == "Windows":
            app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
            local_dir = os.path.join(app_data, 'DataViewer')
        else:
            local_dir = os.path.expanduser('~/.dataviewer')
        
        os.makedirs(local_dir, exist_ok=True)
        local_path = os.path.join(local_dir, DATABASE_FILENAME)
        debug_print(f"PyInstaller environment, using app data: {local_path}")
    else:
        # Development environment - use script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        local_path = os.path.join(script_dir, DATABASE_FILENAME)
        debug_print(f"Development environment, using script directory: {local_path}")
    
    # Ensure directory exists
    local_dir = os.path.dirname(os.path.abspath(local_path))
    try:
        os.makedirs(local_dir, exist_ok=True)
        debug_print(f"Ensured local database directory exists: {local_dir}")
    except Exception as e:
        debug_print(f"WARNING: Could not create local database directory {local_dir}: {e}")
    
    debug_print(f"FALLBACK: Using local database: {local_path}")
    return local_path


@dataclass
class DatabaseConfig:
    """Enhanced configuration for database connections with Synology support."""
    database_type: str = "sqlite"
    host: Optional[str] = None
    port: Optional[int] = None
    database: str = "dataviewer.db"
    username: Optional[str] = None
    password: Optional[str] = None
    connection_timeout: int = 30
    max_connections: int = 10
    
    # Enhanced settings for network databases
    max_retries: int = 3
    retry_delay: float = 2.0
    wal_mode: bool = True  # Better for network/concurrent access
    cache_size: int = 10000  # Larger cache for network latency
    synchronous_mode: str = "NORMAL"  # Balance safety vs performance
    temp_store: str = "MEMORY"  # Use memory for temp storage
    
    # Synology-specific settings
    synology_host: str = ""
    synology_port: int = 5432
    synology_database: str = "dataviewer"
    synology_username: str = ""
    synology_password: str = ""
    auto_discover_synology: bool = True
    
    def __post_init__(self):
        """Post-initialization processing."""
        debug_print(f"Created DatabaseConfig for {self.database_type} database")
        if self.host:
            debug_print(f"Remote database: {self.host}:{self.port}/{self.database}")
        else:
            debug_print(f"Local database: {self.database}")
    
    def get_connection_string(self) -> str:
        """Get appropriate connection string."""
        if self.database_type.lower() == "sqlite":
            return f"sqlite:///{self.database}"
        elif self.database_type.lower() == "postgresql":
            return f"postgresql://{self.username}:{self.password}@{self.host}:{self.port}/{self.database}"
        elif self.database_type.lower() == "synology":
            return f"postgresql://{self.synology_username}:{self.synology_password}@{self.synology_host}:{self.synology_port}/{self.synology_database}"
        else:
            return ""


@dataclass
class DatabaseRecord:
    """Enhanced model for database record metadata."""
    record_id: int
    filename: str
    file_path: str
    created_at: datetime
    updated_at: Optional[datetime] = None
    file_size: int = 0
    checksum: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    sheet_names: List[str] = field(default_factory=list)
    is_plotting: bool = False
    is_empty: bool = True
    
    def __post_init__(self):
        """Post-initialization processing."""
        debug_print(f"Created DatabaseRecord for '{self.filename}' (ID: {self.record_id})")
        debug_print(f"File size: {self.file_size} bytes, Sheets: {len(self.sheet_names)}")
    
    def format_size(self) -> str:
        """Format file size for display."""
        if self.file_size == 0:
            return "0 B"
        elif self.file_size < 1024:
            return f"{self.file_size} B"
        elif self.file_size < 1024 * 1024:
            return f"{self.file_size / 1024:.1f} KB"
        elif self.file_size < 1024 * 1024 * 1024:
            return f"{self.file_size / (1024 * 1024):.1f} MB"
        else:
            return f"{self.file_size / (1024 * 1024 * 1024):.1f} GB"
    
    def get_display_name(self) -> str:
        """Get display name from metadata or filename."""
        if self.metadata and 'display_filename' in self.metadata:
            return self.metadata['display_filename']
        return self.filename
    
    def update_metadata(self, **kwargs):
        """Update metadata fields."""
        self.metadata.update(kwargs)
        self.updated_at = datetime.now()


@dataclass
class DatabaseStats:
    """Statistics about database contents."""
    total_files: int = 0
    total_size_bytes: int = 0
    files_by_type: Dict[str, int] = field(default_factory=dict)
    recent_files_count: int = 0
    oldest_file_date: Optional[datetime] = None
    newest_file_date: Optional[datetime] = None
    average_file_size: float = 0.0
    
    def format_total_size(self) -> str:
        """Format total size for display."""
        if self.total_size_bytes == 0:
            return "0 B"
        elif self.total_size_bytes < 1024:
            return f"{self.total_size_bytes} B"
        elif self.total_size_bytes < 1024 * 1024:
            return f"{self.total_size_bytes / 1024:.1f} KB"
        elif self.total_size_bytes < 1024 * 1024 * 1024:
            return f"{self.total_size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{self.total_size_bytes / (1024 * 1024 * 1024):.1f} GB"


class DatabaseModel:
    """
    Main database model for managing connections, operations, and data persistence.
    Consolidated from database_manager.py, database_explorer.py, and configuration management.
    """
    
    def __init__(self, config: Optional[DatabaseConfig] = None):
        """Initialize the database model."""
        # Configuration
        self.config = config or DatabaseConfig()
        
        # Connection management
        self.connection = None
        self.is_connected = False
        self.connection_error: Optional[str] = None
        self.database_path: Optional[str] = None
        
        # Data management
        self.stored_files: Dict[int, DatabaseRecord] = {}
        self.stored_files_cache: Set[str] = set()  # Track files already stored
        self.query_cache: Dict[str, Any] = {}
        
        # Statistics tracking
        self.query_count = 0
        self.connection_count = 0
        
        # Threading support
        self.db_lock = threading.Lock()
        
        debug_print("DatabaseModel initialized")
        debug_print(f"Database type: {self.config.database_type}")
        debug_print(f"Auto-discover Synology: {self.config.auto_discover_synology}")
    
    # ===================== CONNECTION MANAGEMENT =====================
    
    def connect(self, db_path: str = None) -> bool:
        """Connect to the database with enhanced Synology support."""
        debug_print("Attempting database connection...")
        
        try:
            # Determine database path
            if not db_path:
                if self.config.auto_discover_synology:
                    db_path = get_database_path()
                else:
                    db_path = self.config.database
            
            self.database_path = db_path
            
            # Create directory if it doesn't exist
            db_dir = os.path.dirname(os.path.abspath(db_path))
            try:
                os.makedirs(db_dir, exist_ok=True)
                debug_print(f"Ensured database directory exists: {db_dir}")
            except Exception as dir_error:
                debug_print(f"WARNING: Could not create database directory {db_dir}: {dir_error}")
            
            # Enhanced connection with retry logic for network databases
            for attempt in range(self.config.max_retries):
                try:
                    debug_print(f"Database connection attempt {attempt + 1}/{self.config.max_retries}")
                    
                    # Connect with enhanced settings for network use
                    self.connection = sqlite3.connect(
                        db_path,
                        detect_types=sqlite3.PARSE_DECLTYPES,
                        timeout=float(self.config.connection_timeout),
                        check_same_thread=False  # Allow multi-threading
                    )
                    
                    # Configure for network use
                    self.connection.execute("PRAGMA foreign_keys = ON")
                    if self.config.wal_mode:
                        self.connection.execute("PRAGMA journal_mode = WAL")
                    self.connection.execute(f"PRAGMA synchronous = {self.config.synchronous_mode}")
                    self.connection.execute(f"PRAGMA cache_size = {self.config.cache_size}")
                    self.connection.execute(f"PRAGMA temp_store = {self.config.temp_store}")
                    
                    # Test the connection
                    self.connection.execute("SELECT 1").fetchone()
                    
                    debug_print(f"SUCCESS: Database connected on attempt {attempt + 1}")
                    break
                    
                except sqlite3.OperationalError as db_error:
                    if "database is locked" in str(db_error).lower() and attempt < self.config.max_retries - 1:
                        debug_print(f"Database locked, retrying in {self.config.retry_delay} seconds...")
                        time.sleep(self.config.retry_delay)
                        self.config.retry_delay *= 1.5  # Exponential backoff
                        continue
                    else:
                        raise
                except Exception as conn_error:
                    if attempt < self.config.max_retries - 1:
                        debug_print(f"Connection failed, retrying in {self.config.retry_delay} seconds: {conn_error}")
                        time.sleep(self.config.retry_delay)
                        self.config.retry_delay *= 1.5
                        continue
                    else:
                        raise
            
            # Create tables if they don't exist
            self._create_tables()
            
            # Update status
            self.is_connected = True
            self.connection_error = None
            self.connection_count += 1
            
            debug_print(f"Database initialized successfully at: {db_path}")
            return True
            
        except Exception as e:
            error_msg = f"Failed to connect to database: {e}"
            debug_print(f"ERROR: {error_msg}")
            self.connection_error = error_msg
            self.is_connected = False
            if not hasattr(self, 'connection'):
                self.connection = None
            return False
    
    def disconnect(self):
        """Disconnect from the database."""
        if self.connection:
            try:
                self.connection.close()
                self.connection = None
                self.is_connected = False
                debug_print("Database connection closed")
            except sqlite3.Error as e:
                debug_print(f"Error closing database connection: {e}")
        else:
            debug_print("Database was not connected")
    
    def _check_connection(self):
        """Ensure the database connection is open."""
        if self.connection is None:
            raise ConnectionError("Database connection is not initialized")
    
    def _create_tables(self):
        """Create necessary database tables if they don't exist."""
        debug_print("Creating database tables...")
        
        cursor = self.connection.cursor()
        
        # Files table - enhanced with more metadata
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            file_content BLOB NOT NULL,
            meta_data TEXT,
            created_at TIMESTAMP NOT NULL,
            updated_at TIMESTAMP,
            file_size INTEGER DEFAULT 0,
            checksum TEXT,
            file_type TEXT DEFAULT 'vap3'
        )
        ''')
        
        # Sheets table - enhanced with more sheet metadata
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS sheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            is_plotting BOOLEAN NOT NULL,
            is_empty BOOLEAN NOT NULL,
            row_count INTEGER DEFAULT 0,
            column_count INTEGER DEFAULT 0,
            sheet_type TEXT DEFAULT 'standard',
            FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
        )
        ''')
        
        # Images table - enhanced with more image metadata
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            image_path TEXT NOT NULL,
            image_data BLOB NOT NULL,
            crop_enabled BOOLEAN NOT NULL,
            image_type TEXT DEFAULT 'standard',
            image_size INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
        )
        ''')
        
        # Sample images table - for advanced sample image tracking
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS sample_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER NOT NULL,
            sample_id TEXT NOT NULL,
            image_path TEXT NOT NULL,
            image_data BLOB NOT NULL,
            crop_enabled BOOLEAN DEFAULT FALSE,
            header_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
        )
        ''')
        
        # Create indexes for better performance
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_files_filename ON files (filename)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_files_created_at ON files (created_at)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sheets_file_id ON sheets (file_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_images_file_id ON images (file_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sample_images_file_id ON sample_images (file_id)')
        
        self.connection.commit()
        debug_print("Database tables created successfully")
    
    # ===================== FILE OPERATIONS =====================
    
    def store_vap3_file(self, file_path: str, meta_data: Dict[str, Any]) -> Optional[int]:
        """Store a VAP3 file in the database."""
        debug_print(f"Storing VAP3 file: {file_path}")
        
        try:
            self._check_connection()
            
            # Verify file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            # Read the file content
            with open(file_path, 'rb') as f:
                file_content = f.read()
            
            # Calculate file size and checksum
            file_size = len(file_content)
            checksum = hashlib.md5(file_content).hexdigest()
            
            # Use the display filename from meta_data
            filename = meta_data.get('display_filename', os.path.basename(file_path))
            
            debug_print(f"Storing file: {filename} (from {file_path})")
            debug_print(f"File size: {file_size} bytes, Checksum: {checksum[:8]}...")
            debug_print(f"Meta data keys: {list(meta_data.keys())}")
            
            # Convert meta_data to JSON for storage
            meta_data_json = json.dumps(meta_data)
            
            cursor = self.connection.cursor()
            cursor.execute(
                """INSERT INTO files (filename, file_content, meta_data, created_at, 
                   file_size, checksum, file_type) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (filename, file_content, meta_data_json, datetime.now(),
                 file_size, checksum, 'vap3')
            )
            self.connection.commit()
            
            # Get the ID of the newly inserted record
            file_id = cursor.lastrowid
            
            # Add to cache
            self.stored_files_cache.add(file_path)
            
            debug_print(f"File stored in database with ID {file_id} and filename '{filename}'")
            return file_id
            
        except Exception as e:
            if self.connection:
                try:
                    self.connection.rollback()
                except sqlite3.Error:
                    pass
            error_msg = f"Error storing file in database: {e}"
            debug_print(f"ERROR: {error_msg}")
            raise Exception(error_msg)
    
    def store_sheet_info(self, file_id: int, sheet_name: str, is_plotting: bool, 
                        is_empty: bool, row_count: int = 0, column_count: int = 0) -> int:
        """Store sheet information in the database."""
        debug_print(f"Storing sheet info: {sheet_name} (file_id={file_id})")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                """INSERT INTO sheets (file_id, sheet_name, is_plotting, is_empty, 
                   row_count, column_count) VALUES (?, ?, ?, ?, ?, ?)""",
                (file_id, sheet_name, is_plotting, is_empty, row_count, column_count)
            )
            self.connection.commit()
            
            sheet_id = cursor.lastrowid
            debug_print(f"Sheet info stored with ID {sheet_id}")
            return sheet_id
            
        except Exception as e:
            if self.connection:
                try:
                    self.connection.rollback()
                except sqlite3.Error:
                    pass
            error_msg = f"Error storing sheet info: {e}"
            debug_print(f"ERROR: {error_msg}")
            raise Exception(error_msg)
    
    def store_image(self, file_id: int, sheet_name: str, image_path: str, 
                   crop_enabled: bool = False, image_type: str = "standard") -> int:
        """Store image information and data in the database."""
        debug_print(f"Storing image: {image_path} (file_id={file_id})")
        
        try:
            self._check_connection()
            
            # Read image data
            if not os.path.exists(image_path):
                raise FileNotFoundError(f"Image not found: {image_path}")
            
            with open(image_path, 'rb') as f:
                image_data = f.read()
            
            image_size = len(image_data)
            
            cursor = self.connection.cursor()
            cursor.execute(
                """INSERT INTO images (file_id, sheet_name, image_path, image_data, 
                   crop_enabled, image_type, image_size) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (file_id, sheet_name, os.path.basename(image_path), image_data, 
                 crop_enabled, image_type, image_size)
            )
            self.connection.commit()
            
            image_id = cursor.lastrowid
            debug_print(f"Image stored with ID {image_id}")
            return image_id
            
        except Exception as e:
            if self.connection:
                try:
                    self.connection.rollback()
                except sqlite3.Error:
                    pass
            error_msg = f"Error storing image: {e}"
            debug_print(f"ERROR: {error_msg}")
            raise Exception(error_msg)
    
    # ===================== FILE RETRIEVAL OPERATIONS =====================
    
    def list_files(self, limit: int = None, order_by: str = "created_at DESC") -> List[DatabaseRecord]:
        """List all files stored in the database."""
        debug_print(f"Listing files (limit={limit}, order_by={order_by})")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            query = f"SELECT id, filename, created_at, file_size, checksum, meta_data FROM files ORDER BY {order_by}"
            if limit:
                query += f" LIMIT {limit}"
            
            cursor.execute(query)
            rows = cursor.fetchall()
            
            records = []
            for row in rows:
                try:
                    # Parse metadata
                    try:
                        metadata = json.loads(row[5]) if row[5] else {}
                    except json.JSONDecodeError:
                        metadata = {}
                    
                    # Parse created_at
                    try:
                        if isinstance(row[2], datetime):
                            created_at = row[2]
                        else:
                            created_at = datetime.fromisoformat(row[2])
                    except (ValueError, TypeError):
                        created_at = datetime.now()
                    
                    record = DatabaseRecord(
                        record_id=row[0],
                        filename=row[1],
                        file_path="",  # Not stored for database files
                        created_at=created_at,
                        file_size=row[3] or 0,
                        checksum=row[4],
                        metadata=metadata
                    )
                    records.append(record)
                    
                except Exception as record_error:
                    debug_print(f"Error processing record {row[0]}: {record_error}")
                    continue
            
            self.query_count += 1
            debug_print(f"Retrieved {len(records)} file records")
            return records
            
        except Exception as e:
            error_msg = f"Error listing files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def get_file_by_id(self, file_id: int) -> Optional[Dict[str, Any]]:
        """Get a file by its ID."""
        debug_print(f"Getting file by ID: {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                "SELECT id, filename, file_content, meta_data, created_at FROM files WHERE id = ?",
                (file_id,)
            )
            row = cursor.fetchone()
            
            if row:
                try:
                    metadata = json.loads(row[3]) if row[3] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                try:
                    if isinstance(row[4], datetime):
                        created_at = row[4]
                    else:
                        created_at = datetime.fromisoformat(row[4])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                self.query_count += 1
                return {
                    "id": row[0],
                    "filename": row[1],
                    "file_content": row[2],
                    "meta_data": metadata,
                    "created_at": created_at
                }
            else:
                debug_print(f"File with ID {file_id} not found")
                return None
                
        except Exception as e:
            error_msg = f"Error getting file by ID {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return None
    
    def get_files_with_sheet_info(self) -> List[Dict[str, Any]]:
        """Get all files with their associated sheet information."""
        debug_print("Getting files with sheet information")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            query = """
            SELECT DISTINCT f.id, f.filename, f.meta_data, f.created_at, f.file_size,
                   GROUP_CONCAT(s.sheet_name) as sheet_names
            FROM files f
            LEFT JOIN sheets s ON f.id = s.file_id
            GROUP BY f.id, f.filename, f.meta_data, f.created_at, f.file_size
            ORDER BY f.created_at DESC
            """
            
            cursor.execute(query)
            rows = cursor.fetchall()
            
            files = []
            for row in rows:
                try:
                    metadata = json.loads(row[2]) if row[2] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                try:
                    if isinstance(row[3], datetime):
                        created_at = row[3]
                    else:
                        created_at = datetime.fromisoformat(row[3])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                sheet_names = row[5].split(',') if row[5] else []
                
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "meta_data": metadata,
                    "created_at": created_at,
                    "file_size": row[4] or 0,
                    "sheet_names": sheet_names
                })
            
            self.query_count += 1
            debug_print(f"Retrieved {len(files)} files with sheet info")
            return files
            
        except Exception as e:
            error_msg = f"Error getting files with sheet info: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def search_files(self, search_term: str, limit: int = 50) -> List[DatabaseRecord]:
        """Search for files by filename or metadata."""
        debug_print(f"Searching files for: '{search_term}'")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            query = """
            SELECT id, filename, created_at, file_size, checksum, meta_data 
            FROM files 
            WHERE filename LIKE ? OR meta_data LIKE ?
            ORDER BY created_at DESC
            LIMIT ?
            """
            
            search_pattern = f"%{search_term}%"
            cursor.execute(query, (search_pattern, search_pattern, limit))
            rows = cursor.fetchall()
            
            records = []
            for row in rows:
                try:
                    metadata = json.loads(row[5]) if row[5] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                try:
                    created_at = datetime.fromisoformat(row[2]) if isinstance(row[2], str) else row[2]
                except:
                    created_at = datetime.now()
                
                record = DatabaseRecord(
                    record_id=row[0],
                    filename=row[1],
                    file_path="",
                    created_at=created_at,
                    file_size=row[3] or 0,
                    checksum=row[4],
                    metadata=metadata
                )
                records.append(record)
            
            self.query_count += 1
            debug_print(f"Found {len(records)} files matching '{search_term}'")
            return records
            
        except Exception as e:
            error_msg = f"Error searching files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    # ===================== FILE MANAGEMENT OPERATIONS =====================
    
    def delete_file(self, file_id: int) -> bool:
        """Delete a file and all associated data from the database."""
        debug_print(f"Deleting file with ID: {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            
            # Get filename for logging
            cursor.execute("SELECT filename FROM files WHERE id = ?", (file_id,))
            result = cursor.fetchone()
            filename = result[0] if result else f"ID {file_id}"
            
            # Delete the file (cascading deletes will handle sheets and images)
            cursor.execute("DELETE FROM files WHERE id = ?", (file_id,))
            deleted_count = cursor.rowcount
            
            self.connection.commit()
            
            if deleted_count > 0:
                debug_print(f"Successfully deleted file: {filename}")
                # Remove from stored files cache if present
                self.stored_files.pop(file_id, None)
                return True
            else:
                debug_print(f"File with ID {file_id} not found")
                return False
                
        except Exception as e:
            if self.connection:
                try:
                    self.connection.rollback()
                except sqlite3.Error:
                    pass
            error_msg = f"Error deleting file {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    def export_file_to_vap3(self, file_id: int, output_path: str) -> bool:
        """Export a file from database to VAP3 file."""
        debug_print(f"Exporting file {file_id} to: {output_path}")
        
        try:
            file_data = self.get_file_by_id(file_id)
            if not file_data:
                debug_print(f"File with ID {file_id} not found")
                return False
            
            # Write file content to output path
            with open(output_path, 'wb') as f:
                f.write(file_data['file_content'])
            
            debug_print(f"Successfully exported file to: {output_path}")
            return True
            
        except Exception as e:
            error_msg = f"Error exporting file {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    def get_most_recent_version_by_base_name(self, base_filename: str) -> Optional[Dict[str, Any]]:
        """Get the most recent version of a file by its base filename."""
        debug_print(f"Getting most recent version of: {base_filename}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                """SELECT id, filename, file_content, meta_data, created_at 
                   FROM files 
                   WHERE filename LIKE ? 
                   ORDER BY created_at DESC 
                   LIMIT 1""",
                (f"%{base_filename}%",)
            )
            
            row = cursor.fetchone()
            if row:
                try:
                    metadata = json.loads(row[3]) if row[3] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                try:
                    created_at = datetime.fromisoformat(row[4]) if isinstance(row[4], str) else row[4]
                except:
                    created_at = datetime.now()
                
                self.query_count += 1
                return {
                    "id": row[0],
                    "filename": row[1],
                    "file_content": row[2],
                    "meta_data": metadata,
                    "created_at": created_at
                }
            else:
                debug_print(f"No files found matching: {base_filename}")
                return None
                
        except Exception as e:
            error_msg = f"Error getting most recent version: {e}"
            debug_print(f"ERROR: {error_msg}")
            return None
    
    # ===================== DATABASE STATISTICS AND ANALYSIS =====================
    
    def get_database_stats(self) -> DatabaseStats:
        """Get comprehensive database statistics."""
        debug_print("Calculating database statistics")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            
            # Get total files and size
            cursor.execute("SELECT COUNT(*), SUM(file_size) FROM files")
            total_files, total_size = cursor.fetchone()
            total_size = total_size or 0
            
            # Get files by type
            cursor.execute("SELECT file_type, COUNT(*) FROM files GROUP BY file_type")
            files_by_type = dict(cursor.fetchall())
            
            # Get recent files (last 7 days)
            week_ago = datetime.now() - timedelta(days=7)
            cursor.execute("SELECT COUNT(*) FROM files WHERE created_at > ?", (week_ago,))
            recent_files_count = cursor.fetchone()[0]
            
            # Get oldest and newest file dates
            cursor.execute("SELECT MIN(created_at), MAX(created_at) FROM files")
            oldest_date, newest_date = cursor.fetchone()
            
            # Parse dates
            try:
                oldest_file_date = datetime.fromisoformat(oldest_date) if oldest_date else None
            except:
                oldest_file_date = None
            
            try:
                newest_file_date = datetime.fromisoformat(newest_date) if newest_date else None
            except:
                newest_file_date = None
            
            # Calculate average file size
            average_file_size = total_size / total_files if total_files > 0 else 0.0
            
            stats = DatabaseStats(
                total_files=total_files,
                total_size_bytes=total_size,
                files_by_type=files_by_type,
                recent_files_count=recent_files_count,
                oldest_file_date=oldest_file_date,
                newest_file_date=newest_file_date,
                average_file_size=average_file_size
            )
            
            self.query_count += 1
            debug_print(f"Database stats: {total_files} files, {stats.format_total_size()}")
            return stats
            
        except Exception as e:
            error_msg = f"Error getting database stats: {e}"
            debug_print(f"ERROR: {error_msg}")
            return DatabaseStats()
    
    def get_largest_files(self, limit: int = 10) -> List[DatabaseRecord]:
        """Get the largest files in the database."""
        debug_print(f"Getting {limit} largest files")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                """SELECT id, filename, created_at, file_size, checksum, meta_data 
                   FROM files 
                   ORDER BY file_size DESC 
                   LIMIT ?""",
                (limit,)
            )
            rows = cursor.fetchall()
            
            records = []
            for row in rows:
                try:
                    metadata = json.loads(row[5]) if row[5] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                try:
                    created_at = datetime.fromisoformat(row[2]) if isinstance(row[2], str) else row[2]
                except:
                    created_at = datetime.now()
                
                record = DatabaseRecord(
                    record_id=row[0],
                    filename=row[1],
                    file_path="",
                    created_at=created_at,
                    file_size=row[3] or 0,
                    checksum=row[4],
                    metadata=metadata
                )
                records.append(record)
            
            self.query_count += 1
            debug_print(f"Retrieved {len(records)} largest files")
            return records
            
        except Exception as e:
            error_msg = f"Error getting largest files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def get_recent_files(self, days: int = 7, limit: int = 50) -> List[DatabaseRecord]:
        """Get recent files from the database."""
        debug_print(f"Getting files from last {days} days (limit={limit})")
        
        try:
            self._check_connection()
            
            cutoff_date = datetime.now() - timedelta(days=days)
            
            cursor = self.connection.cursor()
            cursor.execute(
                """SELECT id, filename, created_at, file_size, checksum, meta_data 
                   FROM files 
                   WHERE created_at > ? 
                   ORDER BY created_at DESC 
                   LIMIT ?""",
                (cutoff_date, limit)
            )
            rows = cursor.fetchall()
            
            records = []
            for row in rows:
                try:
                    metadata = json.loads(row[5]) if row[5] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                try:
                    created_at = datetime.fromisoformat(row[2]) if isinstance(row[2], str) else row[2]
                except:
                    created_at = datetime.now()
                
                record = DatabaseRecord(
                    record_id=row[0],
                    filename=row[1],
                    file_path="",
                    created_at=created_at,
                    file_size=row[3] or 0,
                    checksum=row[4],
                    metadata=metadata
                )
                records.append(record)
            
            self.query_count += 1
            debug_print(f"Retrieved {len(records)} recent files")
            return records
            
        except Exception as e:
            error_msg = f"Error getting recent files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    # ===================== DATABASE MAINTENANCE OPERATIONS =====================
    
    def backup_database(self, backup_path: str) -> bool:
        """Create a backup of the database."""
        debug_print(f"Creating database backup: {backup_path}")
        
        try:
            if not self.database_path or not os.path.exists(self.database_path):
                debug_print("Database file does not exist")
                return False
            
            # Ensure backup directory exists
            backup_dir = os.path.dirname(backup_path)
            if backup_dir and not os.path.exists(backup_dir):
                os.makedirs(backup_dir, exist_ok=True)
            
            # Create backup
            shutil.copy2(self.database_path, backup_path)
            
            debug_print(f"Database backup created successfully: {backup_path}")
            return True
            
        except Exception as e:
            error_msg = f"Failed to create database backup: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    def vacuum_database(self) -> bool:
        """Vacuum the database to reclaim space and optimize performance."""
        debug_print("Vacuuming database")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute("VACUUM")
            
            debug_print("Database vacuum completed successfully")
            return True
            
        except Exception as e:
            error_msg = f"Error vacuuming database: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    def analyze_database(self) -> bool:
        """Analyze the database to update query planner statistics."""
        debug_print("Analyzing database")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute("ANALYZE")
            
            debug_print("Database analysis completed successfully")
            return True
            
        except Exception as e:
            error_msg = f"Error analyzing database: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    def check_database_integrity(self) -> bool:
        """Check database integrity."""
        debug_print("Checking database integrity")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute("PRAGMA integrity_check")
            result = cursor.fetchall()
            
            if result and result[0][0] == "ok":
                debug_print("Database integrity check passed")
                return True
            else:
                debug_print(f"Database integrity check failed: {result}")
                return False
                
        except Exception as e:
            error_msg = f"Error checking database integrity: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    # ===================== CACHE AND UTILITY METHODS =====================
    
    def clear_cache(self):
        """Clear all internal caches."""
        debug_print("Clearing database caches")
        
        self.stored_files.clear()
        self.stored_files_cache.clear()
        self.query_cache.clear()
        
        debug_print("Database caches cleared")
    
    def is_file_stored(self, file_path: str) -> bool:
        """Check if a file path is already stored in cache."""
        return file_path in self.stored_files_cache
    
    def get_service_status(self) -> Dict[str, Any]:
        """Get comprehensive service status information."""
        return {
            'initialized': True,
            'connected': self.is_connected,
            'database_path': self.database_path,
            'database_type': self.config.database_type,
            'query_count': self.query_count,
            'connection_count': self.connection_count,
            'cached_files': len(self.stored_files_cache),
            'database_exists': os.path.exists(self.database_path) if self.database_path else False,
            'database_size': os.path.getsize(self.database_path) if self.database_path and os.path.exists(self.database_path) else 0,
            'connection_error': self.connection_error,
            'synology_auto_discover': self.config.auto_discover_synology,
            'config': {
                'max_retries': self.config.max_retries,
                'retry_delay': self.config.retry_delay,
                'wal_mode': self.config.wal_mode,
                'cache_size': self.config.cache_size,
                'connection_timeout': self.config.connection_timeout
            }
        }
    
    def __del__(self):
        """Cleanup when model is destroyed."""
        try:
            self.disconnect()
        except:
            pass  # Ignore cleanup errors


# Export the main classes
__all__ = ['DatabaseModel', 'DatabaseConfig', 'DatabaseRecord', 'DatabaseStats']

# Debug output for model initialization
debug_print("DatabaseModel module loaded successfully")
debug_print("Available classes: DatabaseModel, DatabaseConfig, DatabaseRecord, DatabaseStats")
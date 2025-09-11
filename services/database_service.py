# services/database_service.py
"""
services/database_service.py
Consolidated database service for data persistence operations.
This consolidates database logic from database_manager.py and related database operations.
"""

import os
import json
import sqlite3
import time
import tempfile
import shutil
from typing import Optional, Dict, Any, List, Tuple, Union
from pathlib import Path
from datetime import datetime, timedelta
import traceback


def debug_print(message: str):
    """Debug print function for database operations."""
    print(f"DEBUG: DatabaseService - {message}")


class DatabaseService:
    """Service for all database operations and data persistence."""
    
    def __init__(self, database_path: Optional[str] = None):
        """Initialize the database service."""
        debug_print("Initializing DatabaseService")
        
        # Database configuration
        self.database_path = database_path or os.path.join("data", "database.db")
        self.connection = None
        self.is_connected = False
        self.db_path = None
        
        # Connection management
        self.max_retries = 3
        self.base_retry_delay = 1.0  # seconds
        self.connection_timeout = 30
        
        # Cache management
        self.stored_files_cache = set()  # Track files already stored
        self.query_cache = {}  # Cache for frequent queries
        self.cache_timeout = 300  # 5 minutes
        
        # Performance tracking
        self.query_count = 0
        self.total_query_time = 0.0
        self.connection_count = 0
        
        # Initialize database connection
        self._initialize_database()
        
        debug_print("DatabaseService initialized successfully")
        debug_print(f"Database path: {self.database_path}")
        debug_print(f"Connection status: {'Connected' if self.is_connected else 'Disconnected'}")
    
    # ===================== CONNECTION MANAGEMENT =====================
    
    def _initialize_database(self):
        """Initialize database connection and create tables."""
        debug_print("Initializing database connection")
        
        try:
            # Ensure database directory exists
            db_dir = os.path.dirname(self.database_path)
            if db_dir and not os.path.exists(db_dir):
                os.makedirs(db_dir, exist_ok=True)
                debug_print(f"Created database directory: {db_dir}")
            
            # Establish connection with retry logic
            self._establish_connection()
            
            # Create tables if they don't exist
            self._create_tables()
            
            self.is_connected = True
            self.db_path = os.path.abspath(self.database_path)
            self.connection_count += 1
            
            debug_print(f"Database initialized successfully at: {self.db_path}")
            
        except Exception as e:
            error_msg = f"Failed to initialize database: {e}"
            debug_print(f"ERROR: {error_msg}")
            self.is_connected = False
            self.connection = None
            raise
    
    def _establish_connection(self):
        """Establish database connection with retry logic and exponential backoff."""
        debug_print("Establishing database connection")
        
        retry_delay = self.base_retry_delay
        
        for attempt in range(self.max_retries):
            try:
                debug_print(f"Connection attempt {attempt + 1}/{self.max_retries}")
                
                # Create connection with timeout
                self.connection = sqlite3.connect(
                    self.database_path,
                    timeout=self.connection_timeout,
                    check_same_thread=False
                )
                
                # Enable foreign key constraints
                self.connection.execute("PRAGMA foreign_keys = ON")
                
                # Test connection
                cursor = self.connection.cursor()
                cursor.execute("SELECT 1")
                cursor.fetchone()
                
                debug_print("Database connection established successfully")
                return
                
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e).lower() and attempt < self.max_retries - 1:
                    debug_print(f"Database locked, retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    retry_delay *= 1.5  # Exponential backoff
                    continue
                else:
                    raise
            except Exception as conn_error:
                if attempt < self.max_retries - 1:
                    debug_print(f"Connection failed, retrying in {retry_delay} seconds: {conn_error}")
                    time.sleep(retry_delay)
                    retry_delay *= 1.5
                    continue
                else:
                    raise
    
    def _check_connection(self):
        """Ensure database connection is valid."""
        if self.connection is None:
            raise ConnectionError("Database connection is not initialized")
        
        try:
            # Test connection with a simple query
            cursor = self.connection.cursor()
            cursor.execute("SELECT 1")
            cursor.fetchone()
        except sqlite3.Error:
            debug_print("Connection test failed, attempting to reconnect")
            self._establish_connection()
    
    def reconnect(self) -> bool:
        """Reconnect to the database."""
        debug_print("Attempting to reconnect to database")
        
        try:
            if self.connection:
                self.connection.close()
            
            self._establish_connection()
            self.is_connected = True
            self.connection_count += 1
            
            debug_print("Database reconnection successful")
            return True
            
        except Exception as e:
            debug_print(f"Failed to reconnect to database: {e}")
            self.is_connected = False
            return False
    
    def close(self):
        """Close database connection and cleanup resources."""
        debug_print("Closing database connection")
        
        try:
            if self.connection:
                self.connection.close()
                self.connection = None
            
            self.is_connected = False
            self.query_cache.clear()
            
            debug_print("Database connection closed successfully")
            debug_print(f"Session stats - Queries: {self.query_count}, Connections: {self.connection_count}")
            
        except Exception as e:
            debug_print(f"Error closing database connection: {e}")
    
    # ===================== TABLE MANAGEMENT =====================
    
    def _create_tables(self):
        """Create necessary database tables if they don't exist."""
        debug_print("Creating database tables")
        
        try:
            cursor = self.connection.cursor()
            
            # Files table - stores file content and metadata
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                file_content BLOB NOT NULL,
                meta_data TEXT,
                created_at TIMESTAMP NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                file_size INTEGER,
                checksum TEXT
            )
            ''')
            
            # Sheets table - stores sheet information for each file
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS sheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id INTEGER NOT NULL,
                sheet_name TEXT NOT NULL,
                is_plotting BOOLEAN NOT NULL,
                is_empty BOOLEAN NOT NULL,
                sheet_data TEXT,
                processing_metadata TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
            )
            ''')
            
            # Images table - stores image data and crop information
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id INTEGER NOT NULL,
                sheet_name TEXT NOT NULL,
                image_path TEXT NOT NULL,
                image_data BLOB NOT NULL,
                crop_enabled BOOLEAN NOT NULL,
                crop_metadata TEXT,
                image_type TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
            )
            ''')
            
            # Create indexes for performance
            cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_files_filename ON files(filename)
            ''')
            
            cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_files_created_at ON files(created_at)
            ''')
            
            cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_sheets_file_id ON sheets(file_id)
            ''')
            
            cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_images_file_id ON images(file_id)
            ''')
            
            self.connection.commit()
            debug_print("Database tables created successfully")
            
        except Exception as e:
            if self.connection:
                self.connection.rollback()
            error_msg = f"Failed to create database tables: {e}"
            debug_print(f"ERROR: {error_msg}")
            raise
    
    # ===================== FILE OPERATIONS =====================
    
    def store_vap3_file(self, file_path: str, meta_data: Dict[str, Any]) -> Optional[int]:
        """
        Store a VAP3 file in the database.
        
        Args:
            file_path: Path to the VAP3 file
            meta_data: Dictionary containing file metadata
            
        Returns:
            ID of the newly inserted file record or None if failed
        """
        debug_print(f"Storing VAP3 file: {file_path}")
        
        try:
            self._check_connection()
            
            # Verify file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            # Read file content
            with open(file_path, 'rb') as f:
                file_content = f.read()
            
            file_size = len(file_content)
            
            # Get display filename from metadata
            filename = meta_data.get('display_filename', os.path.basename(file_path))
            
            debug_print(f"File size: {file_size} bytes")
            debug_print(f"Display filename: {filename}")
            debug_print(f"Metadata keys: {list(meta_data.keys())}")
            
            # Convert metadata to JSON
            meta_data_json = json.dumps(meta_data, default=str)
            
            # Insert file record
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO files (filename, file_content, meta_data, created_at, file_size) VALUES (?, ?, ?, ?, ?)",
                (filename, file_content, meta_data_json, datetime.now(), file_size)
            )
            
            file_id = cursor.lastrowid
            self.connection.commit()
            
            # Add to cache
            self.stored_files_cache.add(file_path)
            self.query_count += 1
            
            debug_print(f"File stored successfully with ID: {file_id}")
            return file_id
            
        except Exception as e:
            if self.connection:
                self.connection.rollback()
            error_msg = f"Failed to store VAP3 file: {e}"
            debug_print(f"ERROR: {error_msg}")
            raise
    
    def get_file_by_id(self, file_id: int) -> Optional[Dict[str, Any]]:
        """
        Retrieve a file by its ID.
        
        Args:
            file_id: ID of the file to retrieve
            
        Returns:
            Dictionary containing file data or None if not found
        """
        debug_print(f"Retrieving file by ID: {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                "SELECT id, filename, file_content, meta_data, created_at, file_size FROM files WHERE id = ?",
                (file_id,)
            )
            
            row = cursor.fetchone()
            self.query_count += 1
            
            if row:
                # Parse metadata
                try:
                    meta_data = json.loads(row[3]) if row[3] else {}
                except json.JSONDecodeError:
                    debug_print(f"WARNING: Invalid JSON in metadata for file {file_id}")
                    meta_data = {}
                
                # Parse created_at
                try:
                    if isinstance(row[4], datetime):
                        created_at = row[4]
                    else:
                        created_at = datetime.fromisoformat(row[4])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                file_data = {
                    "id": row[0],
                    "filename": row[1],
                    "file_content": row[2],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "file_size": row[5] or 0
                }
                
                debug_print(f"Retrieved file: {file_data['filename']} ({file_data['file_size']} bytes)")
                return file_data
            else:
                debug_print(f"File with ID {file_id} not found")
                return None
                
        except Exception as e:
            error_msg = f"Failed to retrieve file by ID {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return None
    
    def list_files(self, limit: Optional[int] = None, offset: int = 0, 
                   order_by: str = "created_at", order_desc: bool = True) -> List[Dict[str, Any]]:
        """
        List files stored in the database.
        
        Args:
            limit: Maximum number of files to return
            offset: Number of files to skip
            order_by: Column to order by
            order_desc: Whether to order in descending order
            
        Returns:
            List of file records
        """
        debug_print(f"Listing files (limit={limit}, offset={offset}, order_by={order_by})")
        
        try:
            self._check_connection()
            
            # Build query
            order_clause = f"ORDER BY {order_by} {'DESC' if order_desc else 'ASC'}"
            limit_clause = f"LIMIT {limit} OFFSET {offset}" if limit else ""
            
            query = f"""
            SELECT id, filename, created_at, file_size, meta_data 
            FROM files 
            {order_clause} 
            {limit_clause}
            """
            
            cursor = self.connection.cursor()
            cursor.execute(query)
            rows = cursor.fetchall()
            self.query_count += 1
            
            files = []
            for row in rows:
                # Parse created_at
                try:
                    if isinstance(row[2], datetime):
                        created_at = row[2]
                    else:
                        created_at = datetime.fromisoformat(row[2])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                # Parse metadata for display filename
                try:
                    meta_data = json.loads(row[4]) if row[4] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                display_filename = meta_data.get('display_filename', row[1])
                
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "display_filename": display_filename,
                    "created_at": created_at,
                    "file_size": row[3] or 0,
                    "meta_data": meta_data
                })
            
            debug_print(f"Retrieved {len(files)} files")
            return files
            
        except Exception as e:
            error_msg = f"Failed to list files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def delete_file(self, file_id: int) -> bool:
        """
        Delete a file from the database.
        
        Args:
            file_id: ID of the file to delete
            
        Returns:
            True if successful, False otherwise
        """
        debug_print(f"Deleting file with ID: {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            
            # Get filename for logging
            cursor.execute("SELECT filename FROM files WHERE id = ?", (file_id,))
            result = cursor.fetchone()
            
            if not result:
                debug_print(f"File with ID {file_id} not found")
                return False
            
            filename = result[0]
            debug_print(f"Deleting file: {filename}")
            
            # Delete file (cascade will handle sheets and images)
            cursor.execute("DELETE FROM files WHERE id = ?", (file_id,))
            deleted_count = cursor.rowcount
            
            self.connection.commit()
            self.query_count += 1
            
            debug_print(f"Successfully deleted {deleted_count} record(s)")
            return deleted_count > 0
            
        except Exception as e:
            if self.connection:
                self.connection.rollback()
            error_msg = f"Failed to delete file {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    def delete_multiple_files(self, file_ids: List[int]) -> Tuple[int, int]:
        """
        Delete multiple files from the database.
        
        Args:
            file_ids: List of file IDs to delete
            
        Returns:
            Tuple of (success_count, error_count)
        """
        debug_print(f"Batch deleting {len(file_ids)} files")
        
        success_count = 0
        error_count = 0
        
        for file_id in file_ids:
            if self.delete_file(file_id):
                success_count += 1
            else:
                error_count += 1
        
        debug_print(f"Batch deletion complete: {success_count} successful, {error_count} errors")
        return success_count, error_count
    
    def get_most_recent_version_by_base_name(self, base_filename: str) -> Optional[Dict[str, Any]]:
        """
        Get the most recent version of a file by its base filename.
        
        Args:
            base_filename: Base filename to search for
            
        Returns:
            Most recent file record or None if not found
        """
        debug_print(f"Finding most recent version of: {base_filename}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT id, filename, file_content, meta_data, created_at, file_size
                FROM files 
                WHERE filename LIKE ? 
                ORDER BY created_at DESC 
                LIMIT 1
            """, (f"%{base_filename}%",))
            
            row = cursor.fetchone()
            self.query_count += 1
            
            if row:
                # Parse metadata
                try:
                    meta_data = json.loads(row[3]) if row[3] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                # Parse created_at
                try:
                    if isinstance(row[4], datetime):
                        created_at = row[4]
                    else:
                        created_at = datetime.fromisoformat(row[4])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                file_data = {
                    "id": row[0],
                    "filename": row[1],
                    "file_content": row[2],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "file_size": row[5] or 0
                }
                
                debug_print(f"Found most recent version: {file_data['filename']}")
                return file_data
            else:
                debug_print(f"No versions found for base filename: {base_filename}")
                return None
                
        except Exception as e:
            error_msg = f"Failed to get most recent version of {base_filename}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return None
    
    # ===================== SHEET OPERATIONS =====================
    
    def store_sheet_info(self, file_id: int, sheet_name: str, is_plotting: bool, 
                         is_empty: bool, sheet_data: Optional[str] = None) -> Optional[int]:
        """
        Store sheet information in the database.
        
        Args:
            file_id: ID of the parent file
            sheet_name: Name of the sheet
            is_plotting: Whether this is a plotting sheet
            is_empty: Whether this sheet is empty
            sheet_data: Optional serialized sheet data
            
        Returns:
            ID of the newly inserted sheet record or None if failed
        """
        debug_print(f"Storing sheet info: {sheet_name} for file {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO sheets (file_id, sheet_name, is_plotting, is_empty, sheet_data) VALUES (?, ?, ?, ?, ?)",
                (file_id, sheet_name, is_plotting, is_empty, sheet_data)
            )
            
            sheet_id = cursor.lastrowid
            self.connection.commit()
            self.query_count += 1
            
            debug_print(f"Sheet info stored with ID: {sheet_id}")
            return sheet_id
            
        except Exception as e:
            if self.connection:
                self.connection.rollback()
            error_msg = f"Failed to store sheet info: {e}"
            debug_print(f"ERROR: {error_msg}")
            return None
    
    def get_sheets_for_file(self, file_id: int) -> List[Dict[str, Any]]:
        """
        Get all sheets for a specific file.
        
        Args:
            file_id: ID of the file
            
        Returns:
            List of sheet records
        """
        debug_print(f"Getting sheets for file: {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                "SELECT id, sheet_name, is_plotting, is_empty, sheet_data, created_at FROM sheets WHERE file_id = ?",
                (file_id,)
            )
            
            rows = cursor.fetchall()
            self.query_count += 1
            
            sheets = []
            for row in rows:
                sheets.append({
                    "id": row[0],
                    "sheet_name": row[1],
                    "is_plotting": bool(row[2]),
                    "is_empty": bool(row[3]),
                    "sheet_data": row[4],
                    "created_at": row[5]
                })
            
            debug_print(f"Retrieved {len(sheets)} sheets for file {file_id}")
            return sheets
            
        except Exception as e:
            error_msg = f"Failed to get sheets for file {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def get_files_with_sheet_info(self) -> List[Dict[str, Any]]:
        """
        Get all files with their associated sheet information.
        
        Returns:
            List of file records with sheet data
        """
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
            self.query_count += 1
            
            files = []
            for row in rows:
                # Parse metadata
                try:
                    meta_data = json.loads(row[2]) if row[2] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                # Parse created_at
                try:
                    if isinstance(row[3], datetime):
                        created_at = row[3]
                    else:
                        created_at = datetime.fromisoformat(row[3])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                # Parse sheet names
                sheet_names = row[5].split(',') if row[5] else []
                
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "file_size": row[4] or 0,
                    "sheet_names": sheet_names
                })
            
            debug_print(f"Retrieved {len(files)} files with sheet info")
            return files
            
        except Exception as e:
            error_msg = f"Failed to get files with sheet info: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    # ===================== IMAGE OPERATIONS =====================
    
    def store_image(self, file_id: int, sheet_name: str, image_path: str, 
                    image_data: bytes, crop_enabled: bool = False, 
                    crop_metadata: Optional[Dict[str, Any]] = None) -> Optional[int]:
        """
        Store image data in the database.
        
        Args:
            file_id: ID of the parent file
            sheet_name: Name of the associated sheet
            image_path: Path to the image file
            image_data: Binary image data
            crop_enabled: Whether cropping is enabled
            crop_metadata: Optional crop metadata
            
        Returns:
            ID of the newly inserted image record or None if failed
        """
        debug_print(f"Storing image: {image_path} for file {file_id}")
        
        try:
            self._check_connection()
            
            # Convert crop metadata to JSON
            crop_metadata_json = json.dumps(crop_metadata) if crop_metadata else None
            
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO images (file_id, sheet_name, image_path, image_data, crop_enabled, crop_metadata) VALUES (?, ?, ?, ?, ?, ?)",
                (file_id, sheet_name, os.path.basename(image_path), image_data, crop_enabled, crop_metadata_json)
            )
            
            image_id = cursor.lastrowid
            self.connection.commit()
            self.query_count += 1
            
            debug_print(f"Image stored with ID: {image_id}")
            return image_id
            
        except Exception as e:
            if self.connection:
                self.connection.rollback()
            error_msg = f"Failed to store image: {e}"
            debug_print(f"ERROR: {error_msg}")
            return None
    
    def get_images_for_file(self, file_id: int, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Get all images for a specific file and optionally sheet.
        
        Args:
            file_id: ID of the file
            sheet_name: Optional sheet name filter
            
        Returns:
            List of image records
        """
        debug_print(f"Getting images for file: {file_id}, sheet: {sheet_name}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            
            if sheet_name:
                cursor.execute(
                    "SELECT id, sheet_name, image_path, image_data, crop_enabled, crop_metadata FROM images WHERE file_id = ? AND sheet_name = ?",
                    (file_id, sheet_name)
                )
            else:
                cursor.execute(
                    "SELECT id, sheet_name, image_path, image_data, crop_enabled, crop_metadata FROM images WHERE file_id = ?",
                    (file_id,)
                )
            
            rows = cursor.fetchall()
            self.query_count += 1
            
            images = []
            for row in rows:
                # Parse crop metadata
                try:
                    crop_metadata = json.loads(row[5]) if row[5] else {}
                except json.JSONDecodeError:
                    crop_metadata = {}
                
                images.append({
                    "id": row[0],
                    "sheet_name": row[1],
                    "image_path": row[2],
                    "image_data": row[3],
                    "crop_enabled": bool(row[4]),
                    "crop_metadata": crop_metadata
                })
            
            debug_print(f"Retrieved {len(images)} images")
            return images
            
        except Exception as e:
            error_msg = f"Failed to get images for file {file_id}: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    # ===================== SEARCH AND QUERY OPERATIONS =====================
    
    def search_files(self, search_term: str, search_fields: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """
        Search for files based on filename or metadata.
        
        Args:
            search_term: Term to search for
            search_fields: Fields to search in (filename, meta_data)
            
        Returns:
            List of matching file records
        """
        debug_print(f"Searching files for term: '{search_term}'")
        
        if not search_fields:
            search_fields = ['filename', 'meta_data']
        
        try:
            self._check_connection()
            
            # Build search query
            where_conditions = []
            params = []
            
            if 'filename' in search_fields:
                where_conditions.append("filename LIKE ?")
                params.append(f"%{search_term}%")
            
            if 'meta_data' in search_fields:
                where_conditions.append("meta_data LIKE ?")
                params.append(f"%{search_term}%")
            
            where_clause = " OR ".join(where_conditions)
            query = f"""
            SELECT id, filename, meta_data, created_at, file_size
            FROM files 
            WHERE {where_clause}
            ORDER BY created_at DESC
            """
            
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()
            self.query_count += 1
            
            files = []
            for row in rows:
                # Parse metadata
                try:
                    meta_data = json.loads(row[2]) if row[2] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                # Parse created_at
                try:
                    if isinstance(row[3], datetime):
                        created_at = row[3]
                    else:
                        created_at = datetime.fromisoformat(row[3])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "file_size": row[4] or 0
                })
            
            debug_print(f"Found {len(files)} files matching search term")
            return files
            
        except Exception as e:
            error_msg = f"Failed to search files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def get_recent_files(self, days: int = 7, limit: int = 50) -> List[Dict[str, Any]]:
        """
        Get files created within the specified number of days.
        
        Args:
            days: Number of days to look back
            limit: Maximum number of files to return
            
        Returns:
            List of recent file records
        """
        debug_print(f"Getting files from last {days} days (limit: {limit})")
        
        try:
            self._check_connection()
            
            cutoff_date = datetime.now() - timedelta(days=days)
            
            cursor = self.connection.cursor()
            cursor.execute(
                "SELECT id, filename, meta_data, created_at, file_size FROM files WHERE created_at > ? ORDER BY created_at DESC LIMIT ?",
                (cutoff_date, limit)
            )
            
            rows = cursor.fetchall()
            self.query_count += 1
            
            files = []
            for row in rows:
                # Parse metadata
                try:
                    meta_data = json.loads(row[2]) if row[2] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                # Parse created_at
                try:
                    if isinstance(row[3], datetime):
                        created_at = row[3]
                    else:
                        created_at = datetime.fromisoformat(row[3])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "file_size": row[4] or 0
                })
            
            debug_print(f"Retrieved {len(files)} recent files")
            return files
            
        except Exception as e:
            error_msg = f"Failed to get recent files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    def get_largest_files(self, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Get the largest files by size.
        
        Args:
            limit: Maximum number of files to return
            
        Returns:
            List of largest file records
        """
        debug_print(f"Getting {limit} largest files")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute(
                "SELECT id, filename, meta_data, created_at, file_size FROM files ORDER BY file_size DESC LIMIT ?",
                (limit,)
            )
            
            rows = cursor.fetchall()
            self.query_count += 1
            
            files = []
            for row in rows:
                # Parse metadata
                try:
                    meta_data = json.loads(row[2]) if row[2] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                # Parse created_at
                try:
                    if isinstance(row[3], datetime):
                        created_at = row[3]
                    else:
                        created_at = datetime.fromisoformat(row[3])
                except (ValueError, TypeError):
                    created_at = datetime.now()
                
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "file_size": row[4] or 0
                })
            
            debug_print(f"Retrieved {len(files)} largest files")
            return files
            
        except Exception as e:
            error_msg = f"Failed to get largest files: {e}"
            debug_print(f"ERROR: {error_msg}")
            return []
    
    # ===================== DATABASE ANALYSIS AND STATISTICS =====================
    
    def get_database_statistics(self) -> Dict[str, Any]:
        """
        Get comprehensive database statistics.
        
        Returns:
            Dictionary containing database statistics
        """
        debug_print("Getting database statistics")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            
            # File statistics
            cursor.execute("""
                SELECT 
                    COUNT(*) as file_count,
                    SUM(file_size) as total_size,
                    AVG(file_size) as avg_size,
                    MAX(file_size) as max_size,
                    MIN(file_size) as min_size
                FROM files
            """)
            file_stats = cursor.fetchone()
            
            # Sheet statistics
            cursor.execute("SELECT COUNT(*) FROM sheets")
            sheet_count = cursor.fetchone()[0]
            
            # Image statistics
            cursor.execute("SELECT COUNT(*) FROM images")
            image_count = cursor.fetchone()[0]
            
            # Recent activity (last 7 days)
            cutoff_date = datetime.now() - timedelta(days=7)
            cursor.execute("SELECT COUNT(*) FROM files WHERE created_at > ?", (cutoff_date,))
            recent_files = cursor.fetchone()[0]
            
            # Database file size
            db_file_size = 0
            if os.path.exists(self.database_path):
                db_file_size = os.path.getsize(self.database_path)
            
            self.query_count += 4
            
            stats = {
                'file_count': file_stats[0] or 0,
                'total_content_size': file_stats[1] or 0,
                'avg_file_size': file_stats[2] or 0,
                'max_file_size': file_stats[3] or 0,
                'min_file_size': file_stats[4] or 0,
                'sheet_count': sheet_count,
                'image_count': image_count,
                'recent_files_7_days': recent_files,
                'database_file_size': db_file_size,
                'database_path': self.database_path,
                'is_connected': self.is_connected,
                'query_count': self.query_count,
                'connection_count': self.connection_count
            }
            
            debug_print(f"Database statistics: {stats['file_count']} files, {stats['sheet_count']} sheets, {stats['image_count']} images")
            return stats
            
        except Exception as e:
            error_msg = f"Failed to get database statistics: {e}"
            debug_print(f"ERROR: {error_msg}")
            return {}
    
    def get_file_size_info(self, file_id: int) -> int:
        """
        Get file size information for a specific file.
        
        Args:
            file_id: ID of the file
            
        Returns:
            File size in bytes
        """
        debug_print(f"Getting file size for ID: {file_id}")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            cursor.execute("SELECT file_size FROM files WHERE id = ?", (file_id,))
            result = cursor.fetchone()
            self.query_count += 1
            
            size = result[0] if result else 0
            debug_print(f"File size: {size} bytes")
            return size
            
        except Exception as e:
            debug_print(f"Error getting file size: {e}")
            return 0
    
    def optimize_database(self) -> bool:
        """
        Optimize database performance by running VACUUM and ANALYZE.
        
        Returns:
            True if successful, False otherwise
        """
        debug_print("Optimizing database")
        
        try:
            self._check_connection()
            
            cursor = self.connection.cursor()
            
            # Analyze database statistics
            cursor.execute("ANALYZE")
            
            # Vacuum database to reclaim space
            cursor.execute("VACUUM")
            
            self.connection.commit()
            self.query_count += 2
            
            debug_print("Database optimization completed")
            return True
            
        except Exception as e:
            error_msg = f"Failed to optimize database: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False
    
    # ===================== UTILITY METHODS =====================
    
    def format_size(self, size_bytes: int) -> str:
        """Format bytes into human readable size."""
        if size_bytes == 0:
            return "0 B"
        elif size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
    
    def backup_database(self, backup_path: str) -> bool:
        """
        Create a backup of the database.
        
        Args:
            backup_path: Path for the backup file
            
        Returns:
            True if successful, False otherwise
        """
        debug_print(f"Creating database backup: {backup_path}")
        
        try:
            if not os.path.exists(self.database_path):
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
    
    def is_file_stored(self, file_path: str) -> bool:
        """Check if a file path is already stored in cache."""
        return file_path in self.stored_files_cache
    
    def clear_cache(self):
        """Clear all caches."""
        debug_print("Clearing database service caches")
        self.stored_files_cache.clear()
        self.query_cache.clear()
    
    def get_service_status(self) -> Dict[str, Any]:
        """
        Get comprehensive service status information.
        
        Returns:
            Dictionary containing service status
        """
        return {
            'initialized': True,
            'connected': self.is_connected,
            'database_path': self.database_path,
            'query_count': self.query_count,
            'connection_count': self.connection_count,
            'cache_size': len(self.stored_files_cache),
            'database_exists': os.path.exists(self.database_path) if self.database_path else False,
            'database_size': os.path.getsize(self.database_path) if self.database_path and os.path.exists(self.database_path) else 0
        }
    
    def __del__(self):
        """Cleanup when service is destroyed."""
        try:
            self.close()
        except:
            pass  # Ignore cleanup errors
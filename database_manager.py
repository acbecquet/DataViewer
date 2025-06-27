import os
import sqlite3
import json
import datetime
from typing import Dict, List, Any, Optional
from utils import debug_print

class DatabaseManager:
    """Manages interactions with the SQLite database for storing VAP3 files and metadata."""
    
    def __init__(self, db_path=None):
        """
        Initialize the DatabaseManager with a connection to the SQLite database.
        
        Args:
            db_path (str, optional): Path to the SQLite database file. 
                                    If None, creates a 'dataviewer.db' in the current directory.
        """
        try:
            if db_path is None:
                # Use a default path in the application directory
                db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dataviewer.db')
            
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
            
            # Initialize database connection with extended error codes to help with debugging
            self.conn = sqlite3.connect(db_path, detect_types=sqlite3.PARSE_DECLTYPES)
            self.conn.execute("PRAGMA foreign_keys = ON")
            
            # Create tables if they don't exist
            self._create_tables()
            
            debug_print(f"Database initialized at: {db_path}")
        except Exception as e:
            print(f"Error initializing database: {e}")
            # Make sure conn exists even if initialization fails
            if not hasattr(self, 'conn'):
                self.conn = None
            raise
    
    def _create_tables(self):
        """Create necessary database tables if they don't exist."""
        cursor = self.conn.cursor()
        
        # Files table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            file_content BLOB NOT NULL,
            meta_data TEXT,
            created_at TIMESTAMP NOT NULL
        )
        ''')
        
        # Sheets table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS sheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            is_plotting BOOLEAN NOT NULL,
            is_empty BOOLEAN NOT NULL,
            FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
        )
        ''')
        
        # Images table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            image_path TEXT NOT NULL,
            image_data BLOB NOT NULL,
            crop_enabled BOOLEAN NOT NULL,
            FOREIGN KEY (file_id) REFERENCES files (id) ON DELETE CASCADE
        )
        ''')
        
        self.conn.commit()
    
    def _check_connection(self):
        """Ensure the database connection is open."""
        if self.conn is None:
            raise ConnectionError("Database connection is not initialized")
    
    def store_vap3_file(self, file_path, meta_data):
        """
        Store a VAP3 file in the database.
        
        Args:
            file_path (str): Path to the temporary VAP3 file
            meta_data (dict): Dictionary containing file metadata, including display_filename
            
        Returns:
            int: ID of the newly inserted file record
        """
        try:
            self._check_connection()
            
            # Verify file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            # Read the file content
            with open(file_path, 'rb') as f:
                file_content = f.read()
            
            # Use the display filename from meta_data instead of the temp filename
            filename = meta_data.get('display_filename')
            if not filename:
                # Fallback only if display_filename isn't provided
                filename = os.path.basename(file_path)
            
            # Print debug info before database operation
            debug_print(f"Storing file: {filename} (from {file_path})")
            debug_print(f"Meta data keys: {list(meta_data.keys())}")
            
            # Convert meta_data to JSON for storage
            meta_data_json = json.dumps(meta_data)
            
            cursor = self.conn.cursor()
            cursor.execute(
                "INSERT INTO files (filename, file_content, meta_data, created_at) VALUES (?, ?, ?, ?)",
                (filename, file_content, meta_data_json, datetime.datetime.now())
            )
            self.conn.commit()
            
            # Get the ID of the newly inserted record
            file_id = cursor.lastrowid
            
            debug_print(f"File stored in database with ID {file_id} and filename '{filename}'")
            return file_id
            
        except Exception as e:
            if hasattr(self, 'conn') and self.conn is not None:
                try:
                    self.conn.rollback()
                except sqlite3.Error:
                    pass  # Ignore rollback errors
            print(f"Error storing file in database: {e}")
            raise
    
    def store_sheet_info(self, file_id, sheet_name, is_plotting, is_empty):
        """
        Store sheet information in the database.
        
        Args:
            file_id (int): ID of the file this sheet belongs to
            sheet_name (str): Name of the sheet
            is_plotting (bool): Whether this is a plotting sheet
            is_empty (bool): Whether this sheet is empty
        """
        try:
            self._check_connection()
            
            cursor = self.conn.cursor()
            cursor.execute(
                "INSERT INTO sheets (file_id, sheet_name, is_plotting, is_empty) VALUES (?, ?, ?, ?)",
                (file_id, sheet_name, 1 if is_plotting else 0, 1 if is_empty else 0)
            )
            self.conn.commit()
            return cursor.lastrowid
        except Exception as e:
            if self.conn is not None:
                self.conn.rollback()
            print(f"Error storing sheet info: {e}")
            raise
    
    def store_image(self, file_id, image_path, sheet_name, crop_enabled):
        """
        Store an image in the database.
        
        Args:
            file_id (int): ID of the file this image belongs to
            image_path (str): Path to the image file
            sheet_name (str): Name of the sheet this image belongs to
            crop_enabled (bool): Whether auto-crop is enabled for this image
        """
        try:
            self._check_connection()
            
            # Verify file exists
            if not os.path.exists(image_path):
                raise FileNotFoundError(f"Image not found: {image_path}")
            
            # Read the image data
            with open(image_path, 'rb') as f:
                image_data = f.read()
            
            cursor = self.conn.cursor()
            cursor.execute(
                "INSERT INTO images (file_id, sheet_name, image_path, image_data, crop_enabled) VALUES (?, ?, ?, ?, ?)",
                (file_id, sheet_name, os.path.basename(image_path), image_data, 1 if crop_enabled else 0)
            )
            self.conn.commit()
            return cursor.lastrowid
        except Exception as e:
            if self.conn is not None:
                self.conn.rollback()
            print(f"Error storing image: {e}")
            raise
    
    def list_files(self):
        """
        List all files stored in the database.
        
        Returns:
            list: List of file records with id, filename, and created_at
        """
        try:
            self._check_connection()
            
            cursor = self.conn.cursor()
            cursor.execute("SELECT id, filename, created_at FROM files ORDER BY created_at DESC")
            rows = cursor.fetchall()
            
            result = []
            for row in rows:
                try:
                    # Handle different date formats safely
                    if isinstance(row[2], datetime.datetime):
                        created_at = row[2]
                    else:
                        created_at = datetime.datetime.fromisoformat(row[2])
                except (ValueError, TypeError):
                    # If date parsing fails, use current time
                    created_at = datetime.datetime.now()
                
                result.append({
                    "id": row[0],
                    "filename": row[1],
                    "created_at": created_at
                })
            
            return result
        except Exception as e:
            print(f"Error listing files: {e}")
            return []
    
    def get_file_by_id(self, file_id):
        """
        Get a file by its ID.
        
        Args:
            file_id (int): ID of the file to retrieve
            
        Returns:
            dict: File record with id, filename, file_content, meta_data, and created_at
        """
        try:
            self._check_connection()
            
            cursor = self.conn.cursor()
            cursor.execute("SELECT id, filename, file_content, meta_data, created_at FROM files WHERE id = ?", (file_id,))
            row = cursor.fetchone()
            
            if row:
                try:
                    meta_data = json.loads(row[3]) if row[3] else {}
                except json.JSONDecodeError:
                    meta_data = {}
                
                try:
                    if isinstance(row[4], datetime.datetime):
                        created_at = row[4]
                    else:
                        created_at = datetime.datetime.fromisoformat(row[4])
                except (ValueError, TypeError):
                    created_at = datetime.datetime.now()
                
                return {
                    "id": row[0],
                    "filename": row[1],
                    "file_content": row[2],
                    "meta_data": meta_data,
                    "created_at": created_at
                }
            else:
                return None
        except Exception as e:
            print(f"Error getting file: {e}")
            return None
    
    def delete_file(self, file_id):
        """
        Delete a file from the database.
        
        Args:
            file_id (int): ID of the file to delete
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            self._check_connection()
            
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM files WHERE id = ?", (file_id,))
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            if self.conn is not None:
                self.conn.rollback()
            print(f"Error deleting file: {e}")
            return False
    
    def close(self):
        """Close the database connection."""
        if hasattr(self, 'conn') and self.conn is not None:
            try:
                self.conn.close()
                self.conn = None
                debug_print("Database connection closed")
            except sqlite3.Error as e:
                print(f"Error closing database connection: {e}")
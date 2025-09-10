# services/database_service.py
"""
services/database_service.py
Database operations service.
This will contain the database logic from file_manager.py.
"""

from typing import Optional, Dict, Any, List, Tuple
import sqlite3
import json
from datetime import datetime
from pathlib import Path


class DatabaseService:
    """Service for database operations."""
    
    def __init__(self, db_path: str = "dataviewer.db"):
        """Initialize the database service."""
        self.db_path = db_path
        self.connection: Optional[sqlite3.Connection] = None
        self.is_connected = False
        
        print("DEBUG: DatabaseService initialized")
        print(f"DEBUG: Database path: {db_path}")
    
    def connect(self) -> Tuple[bool, str]:
        """Connect to the database."""
        try:
            self.connection = sqlite3.connect(self.db_path)
            self.connection.row_factory = sqlite3.Row  # Enable column access by name
            self.is_connected = True
            
            # Initialize tables if they don't exist
            self._initialize_tables()
            
            print("DEBUG: DatabaseService connected to database")
            return True, "Connected successfully"
            
        except Exception as e:
            error_msg = f"Failed to connect to database: {e}"
            print(f"ERROR: DatabaseService - {error_msg}")
            return False, error_msg
    
    def disconnect(self):
        """Disconnect from the database."""
        if self.connection:
            self.connection.close()
            self.connection = None
        self.is_connected = False
        print("DEBUG: DatabaseService disconnected from database")
    
    def _initialize_tables(self):
        """Initialize database tables."""
        if not self.connection:
            return
        
        cursor = self.connection.cursor()
        
        # Create files table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                file_path TEXT NOT NULL,
                file_size INTEGER,
                checksum TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                metadata TEXT
            )
        ''')
        
        # Create sheets table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id INTEGER,
                sheet_name TEXT NOT NULL,
                sheet_data TEXT,
                processed_data TEXT,
                metadata TEXT,
                FOREIGN KEY (file_id) REFERENCES files (id)
            )
        ''')
        
        # Create images table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id INTEGER,
                sheet_name TEXT,
                image_path TEXT,
                image_type TEXT,
                crop_data TEXT,
                metadata TEXT,
                FOREIGN KEY (file_id) REFERENCES files (id)
            )
        ''')
        
        self.connection.commit()
        print("DEBUG: DatabaseService initialized tables")
    
    def store_file(self, filename: str, file_path: str, file_data: Dict[str, Any]) -> Tuple[bool, Optional[int], str]:
        """Store file data in the database."""
        print(f"DEBUG: DatabaseService storing file: {filename}")
        
        try:
            if not self.is_connected:
                return False, None, "Not connected to database"
            
            cursor = self.connection.cursor()
            
            # Insert file record
            cursor.execute('''
                INSERT INTO files (filename, file_path, file_size, metadata)
                VALUES (?, ?, ?, ?)
            ''', (
                filename,
                file_path,
                file_data.get('file_size', 0),
                json.dumps(file_data.get('metadata', {}))
            ))
            
            file_id = cursor.lastrowid
            
            # Store sheet data
            sheets_data = file_data.get('sheets', {})
            for sheet_name, sheet_info in sheets_data.items():
                cursor.execute('''
                    INSERT INTO sheets (file_id, sheet_name, sheet_data, metadata)
                    VALUES (?, ?, ?, ?)
                ''', (
                    file_id,
                    sheet_name,
                    json.dumps(sheet_info),  # Serialize sheet data
                    json.dumps(sheet_info.get('metadata', {}))
                ))
            
            self.connection.commit()
            
            print(f"DEBUG: DatabaseService stored file {filename} with ID {file_id}")
            return True, file_id, "File stored successfully"
            
        except Exception as e:
            error_msg = f"Failed to store file: {e}"
            print(f"ERROR: DatabaseService - {error_msg}")
            return False, None, error_msg
    
    def get_file_list(self) -> Tuple[bool, List[Dict[str, Any]], str]:
        """Get list of files from database."""
        print("DEBUG: DatabaseService retrieving file list")
        
        try:
            if not self.is_connected:
                return False, [], "Not connected to database"
            
            cursor = self.connection.cursor()
            cursor.execute('''
                SELECT id, filename, file_path, file_size, created_at, updated_at
                FROM files
                ORDER BY updated_at DESC
            ''')
            
            files = []
            for row in cursor.fetchall():
                files.append({
                    'id': row['id'],
                    'filename': row['filename'],
                    'file_path': row['file_path'],
                    'file_size': row['file_size'],
                    'created_at': row['created_at'],
                    'updated_at': row['updated_at']
                })
            
            print(f"DEBUG: DatabaseService retrieved {len(files)} files")
            return True, files, "Success"
            
        except Exception as e:
            error_msg = f"Failed to get file list: {e}"
            print(f"ERROR: DatabaseService - {error_msg}")
            return False, [], error_msg
    
    def load_file_data(self, file_id: int) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """Load complete file data from database."""
        print(f"DEBUG: DatabaseService loading file data for ID: {file_id}")
        
        try:
            if not self.is_connected:
                return False, None, "Not connected to database"
            
            cursor = self.connection.cursor()
            
            # Get file info
            cursor.execute('SELECT * FROM files WHERE id = ?', (file_id,))
            file_row = cursor.fetchone()
            
            if not file_row:
                return False, None, f"File with ID {file_id} not found"
            
            # Get sheet data
            cursor.execute('SELECT * FROM sheets WHERE file_id = ?', (file_id,))
            sheet_rows = cursor.fetchall()
            
            # Reconstruct file data
            file_data = {
                'filename': file_row['filename'],
                'file_path': file_row['file_path'],
                'metadata': json.loads(file_row['metadata'] or '{}'),
                'sheets': {}
            }
            
            for sheet_row in sheet_rows:
                sheet_data = json.loads(sheet_row['sheet_data'] or '{}')
                file_data['sheets'][sheet_row['sheet_name']] = sheet_data
            
            print(f"DEBUG: DatabaseService loaded file data for {file_row['filename']}")
            return True, file_data, "Success"
            
        except Exception as e:
            error_msg = f"Failed to load file data: {e}"
            print(f"ERROR: DatabaseService - {error_msg}")
            return False, None, error_msg
    
    def delete_file(self, file_id: int) -> Tuple[bool, str]:
        """Delete a file and all associated data from database."""
        print(f"DEBUG: DatabaseService deleting file ID: {file_id}")
        
        try:
            if not self.is_connected:
                return False, "Not connected to database"
            
            cursor = self.connection.cursor()
            
            # Delete associated data first
            cursor.execute('DELETE FROM images WHERE file_id = ?', (file_id,))
            cursor.execute('DELETE FROM sheets WHERE file_id = ?', (file_id,))
            cursor.execute('DELETE FROM files WHERE id = ?', (file_id,))
            
            self.connection.commit()
            
            print(f"DEBUG: DatabaseService deleted file ID {file_id}")
            return True, "File deleted successfully"
            
        except Exception as e:
            error_msg = f"Failed to delete file: {e}"
            print(f"ERROR: DatabaseService - {error_msg}")
            return False, error_msg
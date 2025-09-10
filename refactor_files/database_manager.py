import os
import sqlite3
import json
import datetime
import sys
import time

from typing import Dict, List, Any, Optional
from utils import debug_print

def get_database_path():
    """
    Get database path prioritizing working Synology Drive client setup.
    Network discovery added as secondary option.
    """
    import platform
    import string
    from utils import debug_print
    import os
    
    # CONFIGURATION
    SYNOLOGY_IP = "192.168.222.10"  # Your confirmed working IP
    DATABASE_RELATIVE_PATH = r"SDR\Device Group\Database"
    DATABASE_FILENAME = "dataviewer.db"
    
    debug_print("DEBUG: Starting database path detection...")
    
    # Method 1: PRIORITIZE your working local Synology Drive setup
    if platform.system() == "Windows":
        # Check your known working locations first
        priority_locations = [
            r"C:\Synology2024",  # Your confirmed working setup
            r"C:\SynologyDrive"  # Alternative found location
        ]
        
        for base_path in priority_locations:
            if os.path.exists(base_path):
                db_path = os.path.join(base_path, DATABASE_RELATIVE_PATH, DATABASE_FILENAME)
                db_dir = os.path.dirname(db_path)
                
                if os.path.exists(db_dir):
                    debug_print(f"DEBUG: Found working Synology Drive path: {db_dir}")
                    
                    # Test write access
                    try:
                        test_file = os.path.join(db_dir, "test_write.tmp")
                        with open(test_file, 'w') as f:
                            f.write("test")
                        os.remove(test_file)
                        debug_print(f"SUCCESS: Using Synology Drive with write access: {db_path}")
                        return db_path
                    except Exception as e:
                        debug_print(f"WARNING: Synology Drive path found but no write access: {e}")
    
    # Method 2: Try to discover the correct network share name
    if platform.system() == "Windows":
        debug_print(f"DEBUG: Attempting to discover correct share name for {SYNOLOGY_IP}...")
        
        # Common Synology share names to try
        common_shares = ["volume1", "shared", "drive", "homes", "public", "data"]
        
        for share_name in common_shares:
            for sub_path in ["", "SDR", r"SDR\Device Group\Database"]:
                if sub_path:
                    test_unc = f"\\\\{SYNOLOGY_IP}\\{share_name}\\{sub_path}"
                else:
                    test_unc = f"\\\\{SYNOLOGY_IP}\\{share_name}"
                
                try:
                    if os.path.exists(test_unc):
                        debug_print(f"SUCCESS: Found accessible share: {test_unc}")
                        
                        # If this is the database directory, use it
                        if sub_path == r"SDR\Device Group\Database":
                            network_db_path = os.path.join(test_unc, DATABASE_FILENAME)
                            try:
                                # Test write access
                                test_file = os.path.join(test_unc, "test_write.tmp")
                                with open(test_file, 'w') as f:
                                    f.write("test")
                                os.remove(test_file)
                                debug_print(f"SUCCESS: Network path with write access: {network_db_path}")
                                return network_db_path
                            except:
                                debug_print(f"INFO: Network path found but no write access: {test_unc}")
                        
                        # If this has the SDR folder, build the full path
                        elif sub_path == "SDR":
                            full_db_path = os.path.join(test_unc, "Device Group", "Database", DATABASE_FILENAME)
                            if os.path.exists(os.path.dirname(full_db_path)):
                                try:
                                    test_file = os.path.join(os.path.dirname(full_db_path), "test_write.tmp")
                                    with open(test_file, 'w') as f:
                                        f.write("test")
                                    os.remove(test_file)
                                    debug_print(f"SUCCESS: Found full network path: {full_db_path}")
                                    return full_db_path
                                except:
                                    debug_print(f"INFO: Found network SDR folder but no write access")
                        
                        break  # Found this share, don't need to test sub-paths
                except:
                    continue  # Try next path
    
    # Method 3: Scan all drives for any Synology folders (backup)
    if platform.system() == "Windows":
        debug_print("DEBUG: Scanning all drives for Synology folders...")
        
        for drive_letter in string.ascii_uppercase:
            drive_root = f"{drive_letter}:\\"
            if not os.path.exists(drive_root):
                continue
            
            synology_patterns = ["Synology2024", "Synology2025", "SynologyDrive", "Synology"]
            
            for pattern in synology_patterns:
                synology_base = os.path.join(drive_root, pattern)
                if os.path.exists(synology_base):
                    db_path = os.path.join(synology_base, DATABASE_RELATIVE_PATH, DATABASE_FILENAME)
                    db_dir = os.path.dirname(db_path)
                    
                    if os.path.exists(db_dir):
                        try:
                            test_file = os.path.join(db_dir, "test_write.tmp")
                            with open(test_file, 'w') as f:
                                f.write("test")
                            os.remove(test_file)
                            debug_print(f"SUCCESS: Backup Synology Drive path: {db_path}")
                            return db_path
                        except:
                            continue
    
    # Method 4: Fallback to local database - FIXED VERSION
    debug_print("WARNING: No accessible shared database found")
    debug_print("INFO: Using local database - changes won't be shared with team")

    # Determine appropriate local database location
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller environment - use user's app data directory, not the temp extraction dir

        if platform.system() == "Windows":
            app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
            local_dir = os.path.join(app_data, 'DataViewer')
        else:
            local_dir = os.path.expanduser('~/.dataviewer')
        
        os.makedirs(local_dir, exist_ok=True)
        local_path = os.path.join(local_dir, DATABASE_FILENAME)
        debug_print(f"DEBUG: PyInstaller environment, using app data: {local_path}")
    else:
        # Development environment - use script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        local_path = os.path.join(script_dir, DATABASE_FILENAME)
        debug_print(f"DEBUG: Development environment, using script directory: {local_path}")

    # Ensure directory exists (though DatabaseManager will also handle this)
    local_dir = os.path.dirname(os.path.abspath(local_path))
    try:
        os.makedirs(local_dir, exist_ok=True)
        debug_print(f"DEBUG: Ensured local database directory exists: {local_dir}")
    except Exception as e:
        debug_print(f"WARNING: Could not create local database directory {local_dir}: {e}")

    debug_print(f"FALLBACK: Using local database: {local_path}")
    return local_path

class DatabaseManager:
    """Manages interactions with the SQLite database for storing VAP3 files and metadata."""
    def __init__(self, db_path=None):
        """
        Initialize the DatabaseManager with enhanced Synology support and connection retry.
        """
        try:
            db_path = get_database_path()
            if db_path is None:
                db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dataviewer.db')
            
            # Create directory if it doesn't exist (important for network paths)
            db_dir = os.path.dirname(os.path.abspath(db_path))
            try:
                os.makedirs(db_dir, exist_ok=True)
                debug_print(f"DEBUG: Ensured database directory exists: {db_dir}")
            except Exception as dir_error:
                debug_print(f"WARNING: Could not create database directory {db_dir}: {dir_error}")
                # This might be normal for network paths
            
            # Enhanced connection with retry logic for network databases
            max_retries = 3
            retry_delay = 2  # seconds
            
            for attempt in range(max_retries):
                try:
                    debug_print(f"DEBUG: Database connection attempt {attempt + 1}/{max_retries}")
                    
                    # Use WAL mode for better network performance and concurrent access
                    self.conn = sqlite3.connect(
                        db_path, 
                        detect_types=sqlite3.PARSE_DECLTYPES,
                        timeout=30.0,  # 30 second timeout for network operations
                        check_same_thread=False  # Allow multi-threading
                    )
                    
                    # Configure for network use
                    self.conn.execute("PRAGMA foreign_keys = ON")
                    self.conn.execute("PRAGMA journal_mode = WAL")  # Better for network/concurrent access
                    self.conn.execute("PRAGMA synchronous = NORMAL")  # Balance safety vs performance
                    self.conn.execute("PRAGMA cache_size = 10000")  # Larger cache for network latency
                    self.conn.execute("PRAGMA temp_store = MEMORY")  # Use memory for temp storage
                    
                    # Test the connection
                    self.conn.execute("SELECT 1").fetchone()
                    
                    debug_print(f"SUCCESS: Database connected on attempt {attempt + 1}")
                    break
                    
                except sqlite3.OperationalError as db_error:
                    if "database is locked" in str(db_error).lower() and attempt < max_retries - 1:
                        debug_print(f"DEBUG: Database locked, retrying in {retry_delay} seconds...")
                        time.sleep(retry_delay)
                        retry_delay *= 1.5  # Exponential backoff
                        continue
                    else:
                        raise
                except Exception as conn_error:
                    if attempt < max_retries - 1:
                        debug_print(f"DEBUG: Connection failed, retrying in {retry_delay} seconds: {conn_error}")
                        time.sleep(retry_delay)
                        retry_delay *= 1.5
                        continue
                    else:
                        raise
            
            # Create tables if they don't exist
            self._create_tables()
            
            debug_print(f"Database initialized successfully at: {db_path}")
            
            # Store the path for reference
            self.db_path = db_path
            
        except Exception as e:
            print(f"Error initializing database: {e}")
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
    
    def get_files_with_sheet_info(self):
        """
        Get all files with their associated sheet information for filtering.
    
        Returns:
            list: List of file records with sheet data
        """
        try:
            self._check_connection()
        
            cursor = self.conn.cursor()
            query = """
            SELECT DISTINCT f.id, f.filename, f.meta_data, f.created_at,
                   GROUP_CONCAT(s.sheet_name) as sheet_names
            FROM files f
            LEFT JOIN sheets s ON f.id = s.file_id
            GROUP BY f.id, f.filename, f.meta_data, f.created_at
            ORDER BY f.created_at DESC
            """
        
            cursor.execute(query)
            rows = cursor.fetchall()
        
            files = []
            for row in rows:
                try:
                    meta_data = json.loads(row[2]) if row[2] else {}
                except json.JSONDecodeError:
                    meta_data = {}
            
                try:
                    if isinstance(row[3], datetime.datetime):
                        created_at = row[3]
                    else:
                        created_at = datetime.datetime.fromisoformat(row[3])
                except (ValueError, TypeError):
                    created_at = datetime.datetime.now()
            
                sheet_names = row[4].split(',') if row[4] else []
            
                files.append({
                    "id": row[0],
                    "filename": row[1],
                    "meta_data": meta_data,
                    "created_at": created_at,
                    "sheet_names": sheet_names
                })
        
            debug_print(f"Retrieved {len(files)} files with sheet info")
            return files
        
        except Exception as e:
            print(f"Error getting files with sheet info: {e}")
            return []

    def delete_file_and_versions(self, file_id):
        """
        Delete a file and all its versions from the database.
    
        Args:
            file_id (int): ID of the file to delete
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            self._check_connection()
        
            cursor = self.conn.cursor()
        
            # Get filename for logging
            cursor.execute("SELECT filename FROM files WHERE id = ?", (file_id,))
            result = cursor.fetchone()
            if not result:
                debug_print(f"File ID {file_id} not found")
                return False
            
            filename = result[0]
            debug_print(f"Deleting file and versions: {filename} (ID: {file_id})")
        
            # Delete from files table (cascade will handle sheets and images)
            cursor.execute("DELETE FROM files WHERE id = ?", (file_id,))
            deleted_count = cursor.rowcount
        
            self.conn.commit()
            debug_print(f"Successfully deleted {deleted_count} record(s)")
            return deleted_count > 0
        
        except Exception as e:
            if self.conn is not None:
                self.conn.rollback()
            print(f"Error deleting file and versions: {e}")
            return False

    def delete_multiple_files(self, file_ids):
        """
        Delete multiple files from the database.
    
        Args:
            file_ids (list): List of file IDs to delete
        
        Returns:
            tuple: (success_count, error_count)
        """
        success_count = 0
        error_count = 0
    
        for file_id in file_ids:
            if self.delete_file_and_versions(file_id):
                success_count += 1
            else:
                error_count += 1
    
        debug_print(f"Batch deletion complete: {success_count} successful, {error_count} errors")
        return success_count, error_count

    def get_file_size_info(self, file_id):
        """
        Get file size information for a specific file.
    
        Args:
            file_id (int): ID of the file
        
        Returns:
            int: File size in bytes
        """
        try:
            self._check_connection()
        
            cursor = self.conn.cursor()
            cursor.execute("SELECT LENGTH(file_content) FROM files WHERE id = ?", (file_id,))
            result = cursor.fetchone()
        
            return result[0] if result else 0
        
        except Exception as e:
            debug_print(f"Error getting file size: {e}")
            return 0

    def get_most_recent_version_by_base_name(self, base_filename):
        """
        Get the most recent version of a file by its base filename.
        
        Args:
            base_filename (str): Base filename to search for
            
        Returns:
            dict: File record or None if not found
        """
        try:
            self._check_connection()
            
            cursor = self.conn.cursor()
            # Search for files with similar base names and get the most recent
            cursor.execute("""
                SELECT id, filename, file_content, meta_data, created_at 
                FROM files 
                WHERE filename LIKE ? 
                ORDER BY created_at DESC 
                LIMIT 1
            """, (f"%{base_filename}%",))
            
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
            print(f"Error getting most recent version: {e}")
            return None

    def close(self):
        """Close the database connection."""
        if hasattr(self, 'conn') and self.conn is not None:
            try:
                self.conn.close()
                self.conn = None
                debug_print("Database connection closed")
            except sqlite3.Error as e:
                print(f"Error closing database connection: {e}")

    def show_database_browser_for_comparison(self, parent_window):
        """Show database browser specifically for file selection for comparison."""
        debug_print("DEBUG: Opening database browser for comparison file selection")
        import tkinter as tk
        from tkinter import ttk
        try:
            # Create the database browser window
            browser_window = tk.Toplevel(parent_window)
            browser_window.title("Select Files for Comparison")
            browser_window.geometry("1000x700")
            browser_window.transient(parent_window)
            browser_window.grab_set()
        
            # Store selected files
            selected_files = []
            selection_confirmed = False
        
            # Create main frame
            main_frame = ttk.Frame(browser_window)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
            # Add instruction label
            instruction_label = ttk.Label(main_frame, 
                                        text="Select files for comparison analysis (minimum 2 required):", 
                                        font=("Arial", 12, "bold"))
            instruction_label.pack(pady=(0, 10))
        
            # Create treeview for file listing
            columns = ("Filename", "Created", "Test Count", "File Size")
            file_tree = ttk.Treeview(main_frame, columns=columns, show="headings", selectmode="extended")
        
            # Configure columns
            file_tree.heading("Filename", text="Filename")
            file_tree.heading("Created", text="Created")
            file_tree.heading("Test Count", text="Test Count")
            file_tree.heading("File Size", text="File Size")
        
            file_tree.column("Filename", width=400)
            file_tree.column("Created", width=150)
            file_tree.column("Test Count", width=100)
            file_tree.column("File Size", width=100)
        
            # Add scrollbars
            tree_scrollbar_v = ttk.Scrollbar(main_frame, orient="vertical", command=file_tree.yview)
            tree_scrollbar_h = ttk.Scrollbar(main_frame, orient="horizontal", command=file_tree.xview)
            file_tree.configure(yscrollcommand=tree_scrollbar_v.set, xscrollcommand=tree_scrollbar_h.set)
        
            # Pack treeview and scrollbars
            tree_frame = ttk.Frame(main_frame)
            tree_frame.pack(fill="both", expand=True, pady=(0, 10))
        
            file_tree.pack(side="left", fill="both", expand=True)
            tree_scrollbar_v.pack(side="right", fill="y")
            tree_scrollbar_h.pack(side="bottom", fill="x")
        
            # Load files from database
            files = self.list_files()
            file_data_map = {}  # Map tree item IDs to file data
        
            for file_info in files:
                # Format the display information
                filename = file_info.get('filename', 'Unknown')
                created = file_info.get('created_at', 'Unknown')
                if isinstance(created, str) and len(created) > 16:
                    created = created[:16]  # Truncate long timestamps
            
                test_count = file_info.get('test_count', 'Unknown')
                file_size = file_info.get('file_size', 'Unknown')
                if isinstance(file_size, (int, float)):
                    file_size = f"{file_size / 1024:.1f} KB"
            
                # Insert into treeview
                item_id = file_tree.insert("", "end", values=(filename, created, test_count, file_size))
                file_data_map[item_id] = file_info
        
            # Selection tracking
            selection_label = ttk.Label(main_frame, text="Selected: 0 files", font=("Arial", 10))
            selection_label.pack(pady=(0, 5))
        
            def update_selection_label():
                selected_items = file_tree.selection()
                count = len(selected_items)
                selection_label.config(text=f"Selected: {count} files")
            
            file_tree.bind("<<TreeviewSelect>>", lambda e: update_selection_label())
        
            # Button frame
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x", pady=(10, 0))
        
            def on_select_all():
                """Select all files."""
                for item in file_tree.get_children():
                    file_tree.selection_add(item)
                update_selection_label()
        
            def on_deselect_all():
                """Deselect all files."""
                file_tree.selection_remove(file_tree.selection())
                update_selection_label()
        
            def on_confirm():
                """Confirm selection and close dialog."""
                nonlocal selected_files, selection_confirmed
                selected_items = file_tree.selection()
            
                if len(selected_items) < 2:
                    messagebox.showwarning("Warning", "Please select at least 2 files for comparison.")
                    return
            
                # Collect selected file data
                selected_files = [file_data_map[item_id] for item_id in selected_items]
                selection_confirmed = True
                browser_window.destroy()
        
            def on_cancel():
                """Cancel selection and close dialog."""
                nonlocal selection_confirmed
                selection_confirmed = False
                browser_window.destroy()
        
            # Add buttons
            ttk.Button(button_frame, text="Select All", command=on_select_all).pack(side="left", padx=5)
            ttk.Button(button_frame, text="Deselect All", command=on_deselect_all).pack(side="left", padx=5)
        
            ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="right", padx=5)
            ttk.Button(button_frame, text="Compare Selected Files", command=on_confirm).pack(side="right", padx=5)
        
            # Handle window close
            browser_window.protocol("WM_DELETE_WINDOW", on_cancel)
        
            # Wait for user selection
            browser_window.wait_window()
        
            if selection_confirmed:
                debug_print(f"DEBUG: User selected {len(selected_files)} files for comparison")
                return selected_files
            else:
                debug_print("DEBUG: User cancelled database file selection")
                return []
            
        except Exception as e:
            debug_print(f"DEBUG: Error in show_database_browser_for_comparison: {e}")
            messagebox.showerror("Error", f"Failed to open database browser: {e}")
            return []
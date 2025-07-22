import os
import sqlite3
import json
import datetime
import sys

from typing import Dict, List, Any, Optional
from utils import debug_print

def get_database_path():
    """Get the correct database path whether running as script or executable"""
    if hasattr(sys, '_MEIPASS'):
        # Running as PyInstaller executable
        return os.path.join(sys._MEIPASS, 'dataviewer.db')
    else:
        # Running as script
        return 'dataviewer.db'

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
            db_path = get_database_path()
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
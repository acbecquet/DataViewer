# models/file_model.py
"""
models/file_model.py
File model for managing file operations, state tracking, caching, and file format handling.
Consolidated from file_manager.py, vap_file_manager.py, utils file functions, and file processing operations.
"""

import os
import re
import copy
import shutil
import time
import uuid
import tempfile
import threading
import subprocess
import traceback
import json
import zipfile
import hashlib
import math
import statistics
from typing import Dict, List, Optional, Set, Any, Tuple, Union
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox


def debug_print(message: str):
    """Debug print function for file operations."""
    print(f"DEBUG: FileModel - {message}")


def round_values(value: float, decimal_places: int = 3) -> float:
    """Round values for display consistency."""
    try:
        return round(float(value), decimal_places)
    except (ValueError, TypeError):
        return 0.0


@dataclass
class FileState:
    """Enhanced model for tracking file loading and processing state."""
    filepath: str
    filename: str
    file_size: int
    file_hash: Optional[str] = None
    last_modified: datetime = field(default_factory=datetime.now)
    load_status: str = "pending"  # pending, loading, loaded, error
    processing_status: str = "pending"  # pending, processing, processed, error
    error_message: Optional[str] = None
    load_time_seconds: Optional[float] = None
    is_cached: bool = False
    cache_expiry: Optional[datetime] = None
    file_type: str = "unknown"  # excel, vap3, csv, etc.
    legacy_mode: Optional[str] = None  # None, standard, legacy
    sheet_count: int = 0
    database_stored: bool = False
    
    def __post_init__(self):
        """Post-initialization processing."""
        self._calculate_file_hash()
        self._detect_file_type()
        debug_print(f"Created FileState for '{self.filename}'")
        debug_print(f"File size: {self.file_size} bytes, Hash: {self.file_hash[:8] if self.file_hash else 'None'}...")
        debug_print(f"File type: {self.file_type}, Legacy mode: {self.legacy_mode}")
    
    def _calculate_file_hash(self):
        """Calculate MD5 hash of the file for change detection."""
        try:
            if Path(self.filepath).exists():
                hasher = hashlib.md5()
                with open(self.filepath, 'rb') as f:
                    for chunk in iter(lambda: f.read(4096), b""):
                        hasher.update(chunk)
                self.file_hash = hasher.hexdigest()
                debug_print(f"Calculated hash for '{self.filename}': {self.file_hash[:8]}...")
            else:
                debug_print(f"File not found for hash calculation: {self.filepath}")
        except Exception as e:
            debug_print(f"Failed to calculate hash for '{self.filename}': {e}")
            self.file_hash = None
    
    def _detect_file_type(self):
        """Detect file type based on extension."""
        try:
            ext = Path(self.filepath).suffix.lower()
            if ext in ['.xlsx', '.xls']:
                self.file_type = "excel"
            elif ext == '.vap3':
                self.file_type = "vap3"
            elif ext == '.csv':
                self.file_type = "csv"
            else:
                self.file_type = "unknown"
        except Exception as e:
            debug_print(f"Error detecting file type: {e}")
            self.file_type = "unknown"
    
    def set_loading(self):
        """Mark file as currently loading."""
        self.load_status = "loading"
        debug_print(f"File '{self.filename}' status: loading")
    
    def set_loaded(self, load_time: float, sheet_count: int = 0):
        """Mark file as successfully loaded."""
        self.load_status = "loaded"
        self.load_time_seconds = load_time
        self.sheet_count = sheet_count
        debug_print(f"File '{self.filename}' loaded successfully in {load_time:.2f}s ({sheet_count} sheets)")
    
    def set_error(self, error_message: str):
        """Mark file as having an error."""
        self.load_status = "error"
        self.error_message = error_message
        debug_print(f"File '{self.filename}' failed to load: {error_message}")
    
    def set_processing(self):
        """Mark file as currently processing."""
        self.processing_status = "processing"
        debug_print(f"File '{self.filename}' status: processing")
    
    def set_processed(self):
        """Mark file as successfully processed."""
        self.processing_status = "processed"
        debug_print(f"File '{self.filename}' processed successfully")
    
    def has_file_changed(self) -> bool:
        """Check if the file has changed since last hash calculation."""
        if not self.file_hash:
            return True
        
        try:
            current_hash = self._calculate_current_hash()
            changed = current_hash != self.file_hash
            if changed:
                debug_print(f"File '{self.filename}' has changed (hash mismatch)")
            return changed
        except Exception as e:
            debug_print(f"Error checking file change for '{self.filename}': {e}")
            return True
    
    def _calculate_current_hash(self) -> Optional[str]:
        """Calculate current file hash without updating stored hash."""
        try:
            if Path(self.filepath).exists():
                hasher = hashlib.md5()
                with open(self.filepath, 'rb') as f:
                    for chunk in iter(lambda: f.read(4096), b""):
                        hasher.update(chunk)
                return hasher.hexdigest()
        except Exception as e:
            debug_print(f"Error calculating current hash: {e}")
        return None
    
    def is_cache_valid(self) -> bool:
        """Check if cache is still valid."""
        if not self.is_cached or not self.cache_expiry:
            return False
        return datetime.now() < self.cache_expiry
    
    def set_cached(self, duration_hours: int = 24):
        """Mark as cached with expiry time."""
        self.is_cached = True
        self.cache_expiry = datetime.now() + timedelta(hours=duration_hours)
        debug_print(f"File '{self.filename}' cached until {self.cache_expiry}")
    
    def detect_legacy_mode(self) -> Optional[str]:
        """Detect if file is legacy based on sheet structure."""
        if self.file_type != "excel":
            return None
        
        try:
            # Load sheet names to check structure
            sheets_dict = pd.read_excel(self.filepath, sheet_name=None, header=None, nrows=1)
            sheet_names = list(sheets_dict.keys())
            num_sheets = len(sheet_names)
            
            debug_print(f"File contains {num_sheets} sheet(s): {sheet_names}")
            
            # Legacy file criteria: exactly 1 sheet named 'Sheet1'
            if num_sheets == 1 and sheet_names[0] == 'Sheet1':
                self.legacy_mode = "legacy"
                debug_print("File detected as legacy format")
            else:
                self.legacy_mode = "standard"
                debug_print("File detected as standard format")
            
            return self.legacy_mode
            
        except Exception as e:
            debug_print(f"Error detecting legacy mode: {e}")
            return None


@dataclass
class FileCache:
    """Enhanced cache for file data with access tracking."""
    cache_data: Dict[str, Any] = field(default_factory=dict)
    created_at: datetime = field(default_factory=datetime.now)
    access_count: int = 0
    last_accessed: Optional[datetime] = None
    cache_size_bytes: int = 0
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Post-initialization processing."""
        debug_print(f"Created FileCache with {len(self.cache_data)} items")
    
    def store_data(self, key: str, data: Any, metadata: Dict[str, Any] = None):
        """Store data in the cache with optional metadata."""
        self.cache_data[key] = data
        self.last_accessed = datetime.now()
        
        if metadata:
            self.metadata[key] = metadata
        
        # Estimate cache size
        try:
            import sys
            self.cache_size_bytes = sys.getsizeof(self.cache_data)
        except:
            pass
        
        debug_print(f"Stored cache data for key '{key}' - {len(self.cache_data)} total items")
    
    def get_data(self, key: str) -> Optional[Any]:
        """Retrieve data from the cache."""
        if key in self.cache_data:
            self.access_count += 1
            self.last_accessed = datetime.now()
            debug_print(f"Retrieved cache data for key '{key}' (access #{self.access_count})")
            return self.cache_data[key]
        
        debug_print(f"Cache miss for key '{key}'")
        return None
    
    def has_data(self, key: str) -> bool:
        """Check if cache has data for a specific key."""
        return key in self.cache_data
    
    def remove_data(self, key: str) -> bool:
        """Remove data from cache."""
        if key in self.cache_data:
            del self.cache_data[key]
            if key in self.metadata:
                del self.metadata[key]
            debug_print(f"Removed cache data for key '{key}'")
            return True
        return False
    
    def clear_cache(self):
        """Clear all cached data."""
        item_count = len(self.cache_data)
        self.cache_data.clear()
        self.metadata.clear()
        self.access_count = 0
        self.cache_size_bytes = 0
        debug_print(f"Cleared cache ({item_count} items removed)")
    
    def get_cache_info(self) -> Dict[str, Any]:
        """Get comprehensive information about the cache."""
        return {
            'item_count': len(self.cache_data),
            'created_at': self.created_at,
            'last_accessed': self.last_accessed,
            'access_count': self.access_count,
            'cache_size_bytes': self.cache_size_bytes,
            'keys': list(self.cache_data.keys()),
            'metadata_keys': list(self.metadata.keys())
        }


@dataclass
class VAP3FileData:
    """Data structure for VAP3 file contents."""
    version: str = "1.0"
    filtered_sheets: Dict[str, Any] = field(default_factory=dict)
    sheet_images: Dict[str, Any] = field(default_factory=dict)
    plot_options: List[str] = field(default_factory=list)
    image_crop_states: Dict[str, bool] = field(default_factory=dict)
    plot_settings: Dict[str, Any] = field(default_factory=dict)
    sample_images: Dict[str, List[str]] = field(default_factory=dict)
    sample_image_crop_states: Dict[str, bool] = field(default_factory=dict)
    sample_header_data: Dict[str, Any] = field(default_factory=dict)
    sample_notes_data: Dict[str, Any] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    
    def __post_init__(self):
        """Post-initialization processing."""
        debug_print(f"Created VAP3FileData with {len(self.filtered_sheets)} sheets")
        debug_print(f"Sample images: {len(self.sample_images)} groups")
        debug_print(f"Sheet images: {len(self.sheet_images)} files")


class FileModel:
    """
    Main file model for managing file operations, state tracking, caching, and file format handling.
    Consolidated from file_manager.py, vap_file_manager.py, utils file functions, and processing operations.
    """
    
    def __init__(self, cache_duration_hours: int = 24):
        """Initialize the file model."""
        # Core file state management
        self.file_states: Dict[str, FileState] = {}  # filepath -> FileState
        self.file_caches: Dict[str, FileCache] = {}  # filepath -> FileCache
        self.loaded_files_cache: Set[str] = set()  # filepaths of loaded files
        self.stored_files_cache: Set[str] = set()  # filepaths of database-stored files
        
        # Configuration
        self.cache_duration_hours = cache_duration_hours
        self.supported_formats = ['.xlsx', '.xls', '.csv', '.vap3']
        self.excel_extensions = ['.xlsx', '.xls']
        
        # File processing parameters
        self.default_puff_time = 3.0
        self.default_puffs = 10
        self.precision_digits = 3
        self.statistical_threshold = 2
        
        # Threading support
        self.file_lock = threading.Lock()
        
        # Temporary files tracking for cleanup
        self.temp_files: List[str] = []
        
        # Database integration
        self.database_service = None
        
        debug_print("FileModel initialized")
        debug_print(f"Default cache duration: {cache_duration_hours} hours")
        debug_print(f"Supported formats: {', '.join(self.supported_formats)}")
    
    def set_database_service(self, database_service):
        """Set the database service for file operations."""
        self.database_service = database_service
        debug_print("Database service connected to FileModel")
    
    # ===================== FILE STATE MANAGEMENT =====================
    
    def add_file_state(self, filepath: str, filename: str = None) -> FileState:
        """Add or update file state tracking."""
        try:
            if filename is None:
                filename = Path(filepath).name
            
            file_path = Path(filepath)
            file_size = file_path.stat().st_size if file_path.exists() else 0
            
            file_state = FileState(
                filepath=filepath,
                filename=filename,
                file_size=file_size
            )
            
            # Detect legacy mode for Excel files
            if file_state.file_type == "excel":
                file_state.detect_legacy_mode()
            
            self.file_states[filepath] = file_state
            debug_print(f"Added file state for '{filename}' at '{filepath}'")
            return file_state
            
        except Exception as e:
            debug_print(f"Failed to add file state for '{filename}': {e}")
            # Create a minimal file state even if we can't read the file
            file_state = FileState(
                filepath=filepath,
                filename=filename or Path(filepath).name,
                file_size=0
            )
            file_state.set_error(str(e))
            self.file_states[filepath] = file_state
            return file_state
    
    def get_file_state(self, filepath: str) -> Optional[FileState]:
        """Get file state for a specific filepath."""
        return self.file_states.get(filepath)
    
    def has_file_state(self, filepath: str) -> bool:
        """Check if we have state tracking for a file."""
        return filepath in self.file_states
    
    def should_reload_file(self, filepath: str) -> bool:
        """Determine if a file should be reloaded."""
        file_state = self.get_file_state(filepath)
        if not file_state:
            debug_print(f"No file state found for '{filepath}' - should reload")
            return True
        
        if file_state.load_status == "error":
            debug_print(f"Previous error for '{filepath}' - should reload")
            return True
        
        if file_state.has_file_changed():
            debug_print(f"File changed '{filepath}' - should reload")
            return True
        
        if file_state.is_cached and not file_state.is_cache_valid():
            debug_print(f"Cache expired for '{filepath}' - should reload")
            return True
        
        debug_print(f"File '{filepath}' is up to date - no reload needed")
        return False
    
    def clear_file_state(self, filepath: str):
        """Clear state and cache for a specific file."""
        if filepath in self.file_states:
            del self.file_states[filepath]
            debug_print(f"Cleared file state for '{filepath}'")
        
        if filepath in self.file_caches:
            del self.file_caches[filepath]
            debug_print(f"Cleared file cache for '{filepath}'")
        
        self.loaded_files_cache.discard(filepath)
        self.stored_files_cache.discard(filepath)
    
    def clear_all_caches(self):
        """Clear all file states and caches."""
        state_count = len(self.file_states)
        cache_count = len(self.file_caches)
        
        self.file_states.clear()
        self.file_caches.clear()
        self.loaded_files_cache.clear()
        self.stored_files_cache.clear()
        
        debug_print(f"Cleared all caches - {state_count} states, {cache_count} caches")
    
    def cleanup_expired_caches(self):
        """Remove expired caches to free memory."""
        expired_files = []
        
        for filepath, file_state in self.file_states.items():
            if file_state.is_cached and not file_state.is_cache_valid():
                expired_files.append(filepath)
        
        for filepath in expired_files:
            self.clear_file_state(filepath)
        
        if expired_files:
            debug_print(f"Cleaned up {len(expired_files)} expired file caches")
    
    # ===================== FILE CACHE MANAGEMENT =====================
    
    def create_cache(self, filepath: str) -> FileCache:
        """Create a cache for a specific file."""
        cache = FileCache()
        self.file_caches[filepath] = cache
        debug_print(f"Created cache for file '{filepath}'")
        return cache
    
    def get_cache(self, filepath: str) -> Optional[FileCache]:
        """Get cache for a specific file."""
        return self.file_caches.get(filepath)
    
    def has_cache(self, filepath: str) -> bool:
        """Check if we have cache for a file."""
        return filepath in self.file_caches
    
    def add_to_loaded_cache(self, filepath: str):
        """Mark a file as loaded in cache."""
        self.loaded_files_cache.add(filepath)
        debug_print(f"Added '{filepath}' to loaded files cache")
    
    def is_in_loaded_cache(self, filepath: str) -> bool:
        """Check if file is in loaded cache."""
        return filepath in self.loaded_files_cache
    
    def add_to_stored_cache(self, filepath: str):
        """Mark a file as stored in database cache."""
        self.stored_files_cache.add(filepath)
        debug_print(f"Added '{filepath}' to stored files cache")
    
    def is_in_stored_cache(self, filepath: str) -> bool:
        """Check if file is in database stored cache."""
        return filepath in self.stored_files_cache
    
    # ===================== FILE VALIDATION METHODS =====================
    
    def validate_file_path(self, file_path: str) -> Tuple[bool, str]:
        """Validate if a file path exists and is accessible."""
        try:
            if not file_path:
                return False, "File path is empty"
            
            path = Path(file_path)
            
            if not path.exists():
                return False, f"File does not exist: {file_path}"
            
            if not path.is_file():
                return False, f"Path is not a file: {file_path}"
            
            if not os.access(file_path, os.R_OK):
                return False, f"File is not readable: {file_path}"
            
            debug_print(f"File path validation passed: {file_path}")
            return True, "Valid file path"
            
        except Exception as e:
            error_msg = f"Error validating file path {file_path}: {str(e)}"
            debug_print(error_msg)
            return False, error_msg
    
    def validate_excel_file(self, file_path: str) -> Tuple[bool, str]:
        """Validate if a file is a valid Excel file."""
        try:
            # First validate the file path
            path_valid, path_error = self.validate_file_path(file_path)
            if not path_valid:
                return False, path_error
            
            # Check file extension
            ext = Path(file_path).suffix.lower()
            if ext not in self.excel_extensions:
                return False, f"Not an Excel file: {file_path} (extension: {ext})"
            
            # Check for temporary Excel files
            filename = Path(file_path).name
            if filename.startswith('~$'):
                return False, f"Temporary Excel file: {file_path}"
            
            # Try to load the file to verify it's a valid Excel file
            try:
                sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
                if not sheets:
                    return False, f"Excel file contains no sheets: {file_path}"
                
                debug_print(f"Excel file validation passed: {file_path} ({len(sheets)} sheets)")
                return True, "Valid Excel file"
                
            except Exception as excel_error:
                return False, f"Cannot read Excel file {file_path}: {str(excel_error)}"
            
        except Exception as e:
            error_msg = f"Error validating Excel file {file_path}: {str(e)}"
            debug_print(error_msg)
            return False, error_msg
    
    def is_valid_excel_file(self, filename: str) -> bool:
        """Quick check if filename indicates a valid Excel file."""
        return filename.endswith(('.xlsx', '.xls')) and not filename.startswith('~$')
    
    def is_standard_file(self, file_path: str) -> bool:
        """Determine if the file is standard format (not legacy)."""
        try:
            file_state = self.get_file_state(file_path)
            if not file_state:
                file_state = self.add_file_state(file_path)
            
            if file_state.legacy_mode is None:
                file_state.detect_legacy_mode()
            
            # Return True for standard, False for legacy
            return file_state.legacy_mode == "standard"
            
        except Exception as e:
            debug_print(f"Error determining file format for {file_path}: {e}")
            return True  # Default to standard
    
    # ===================== EXCEL FILE OPERATIONS =====================
    
    def load_excel_file(self, file_path: str, legacy_mode: str = None, 
                       force_reload: bool = False) -> Tuple[bool, Dict[str, Any], str]:
        """Load an Excel file and return processed sheet data."""
        debug_print(f"Loading Excel file: {file_path}")
        
        start_time = time.time()
        
        try:
            # Validate file first
            is_valid, validation_error = self.validate_excel_file(file_path)
            if not is_valid:
                return False, {}, validation_error
            
            # Check if we should reload
            if not force_reload and not self.should_reload_file(file_path):
                debug_print("File is already loaded and up to date")
                cache = self.get_cache(file_path)
                if cache and cache.has_data('processed_sheets'):
                    return True, cache.get_data('processed_sheets'), "Loaded from cache"
            
            # Get or create file state
            file_state = self.get_file_state(file_path) or self.add_file_state(file_path)
            file_state.set_loading()
            
            # Determine legacy mode if not specified
            if legacy_mode is None:
                legacy_mode = file_state.legacy_mode or file_state.detect_legacy_mode()
            
            # Load Excel sheets
            debug_print(f"Loading Excel sheets with engine=openpyxl")
            sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            
            if not sheets:
                error_msg = "Excel file contains no readable sheets"
                file_state.set_error(error_msg)
                return False, {}, error_msg
            
            # Process sheets based on legacy mode
            processed_sheets = {}
            
            for sheet_name, sheet_data in sheets.items():
                try:
                    debug_print(f"Processing sheet: {sheet_name}")
                    
                    # Clean and validate sheet data
                    cleaned_data = self._clean_sheet_data(sheet_data, sheet_name)
                    
                    # Process based on format
                    if legacy_mode == "legacy":
                        processed_data = self._process_legacy_sheet(cleaned_data, sheet_name)
                    else:
                        processed_data = self._process_standard_sheet(cleaned_data, sheet_name)
                    
                    processed_sheets[sheet_name] = processed_data
                    
                except Exception as sheet_error:
                    debug_print(f"Error processing sheet {sheet_name}: {sheet_error}")
                    # Continue with other sheets
                    continue
            
            if not processed_sheets:
                error_msg = "No sheets could be processed successfully"
                file_state.set_error(error_msg)
                return False, {}, error_msg
            
            # Cache the results
            cache = self.get_cache(file_path) or self.create_cache(file_path)
            cache.store_data('processed_sheets', processed_sheets)
            cache.store_data('legacy_mode', legacy_mode)
            
            # Update file state
            load_time = time.time() - start_time
            file_state.set_loaded(load_time, len(processed_sheets))
            file_state.legacy_mode = legacy_mode
            
            # Mark as cached
            file_state.set_cached(self.cache_duration_hours)
            
            # Add to loaded cache
            self.add_to_loaded_cache(file_path)
            
            debug_print(f"Successfully loaded Excel file with {len(processed_sheets)} sheets in {load_time:.2f}s")
            return True, processed_sheets, f"Loaded {len(processed_sheets)} sheets successfully"
            
        except Exception as e:
            error_msg = f"Failed to load Excel file {file_path}: {str(e)}"
            debug_print(error_msg)
            
            # Update file state with error
            file_state = self.get_file_state(file_path)
            if file_state:
                file_state.set_error(str(e))
            
            return False, {}, error_msg
    
    def _clean_sheet_data(self, sheet_data: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
        """Clean and prepare sheet data for processing."""
        debug_print(f"Cleaning sheet data for {sheet_name}: {sheet_data.shape}")
        
        try:
            # Remove completely empty columns
            sheet_data = sheet_data.dropna(axis=1, how='all')
            
            # Remove completely empty rows
            sheet_data = sheet_data.dropna(axis=0, how='all')
            
            # Reset index
            sheet_data = sheet_data.reset_index(drop=True)
            
            debug_print(f"Cleaned sheet data: {sheet_data.shape}")
            return sheet_data
            
        except Exception as e:
            debug_print(f"Error cleaning sheet data: {e}")
            return sheet_data
    
    def _process_standard_sheet(self, sheet_data: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """Process a standard format sheet."""
        debug_print(f"Processing standard sheet: {sheet_name}")
        
        try:
            # Determine if this is a plotting sheet
            is_plotting_sheet = self._is_plotting_sheet(sheet_data, sheet_name)
            is_empty = sheet_data.empty or sheet_data.shape[0] < 3
            
            processed_sheet = {
                "data": sheet_data,
                "is_plotting": is_plotting_sheet,
                "is_empty": is_empty,
                "sheet_name": sheet_name,
                "format": "standard",
                "row_count": len(sheet_data),
                "column_count": len(sheet_data.columns)
            }
            
            debug_print(f"Standard sheet processed: plotting={is_plotting_sheet}, empty={is_empty}")
            return processed_sheet
            
        except Exception as e:
            debug_print(f"Error processing standard sheet: {e}")
            return {
                "data": sheet_data,
                "is_plotting": False,
                "is_empty": True,
                "sheet_name": sheet_name,
                "format": "standard",
                "error": str(e)
            }
    
    def _process_legacy_sheet(self, sheet_data: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """Process a legacy format sheet."""
        debug_print(f"Processing legacy sheet: {sheet_name}")
        
        try:
            # Legacy sheets are typically plotting sheets with specific structure
            is_plotting_sheet = True  # Legacy sheets are usually plotting sheets
            is_empty = sheet_data.empty or sheet_data.shape[0] < 3
            
            processed_sheet = {
                "data": sheet_data,
                "is_plotting": is_plotting_sheet,
                "is_empty": is_empty,
                "sheet_name": sheet_name,
                "format": "legacy",
                "row_count": len(sheet_data),
                "column_count": len(sheet_data.columns)
            }
            
            debug_print(f"Legacy sheet processed: plotting={is_plotting_sheet}, empty={is_empty}")
            return processed_sheet
            
        except Exception as e:
            debug_print(f"Error processing legacy sheet: {e}")
            return {
                "data": sheet_data,
                "is_plotting": False,
                "is_empty": True,
                "sheet_name": sheet_name,
                "format": "legacy",
                "error": str(e)
            }
    
    def _is_plotting_sheet(self, data: pd.DataFrame, sheet_name: str) -> bool:
        """Determine if a sheet is a plotting sheet based on data structure."""
        try:
            # Check sheet name patterns first
            plotting_indicators = [
                'test', 'plot', 'data', 'sample', 'measurement', 'result',
                'tpm', 'resistance', 'pressure', 'power'
            ]
            
            sheet_name_lower = sheet_name.lower()
            if any(indicator in sheet_name_lower for indicator in plotting_indicators):
                debug_print(f"Sheet '{sheet_name}' identified as plotting sheet by name")
                return True
            
            # Check data structure
            if data.empty or data.shape[0] < 3:
                return False
            
            # Look for typical plotting sheet structure
            # Check if there are numeric columns that could be TPM data
            try:
                numeric_columns = data.select_dtypes(include=[np.number]).columns
                if len(numeric_columns) >= 3:  # At least a few numeric columns
                    debug_print(f"Sheet '{sheet_name}' identified as plotting sheet by structure")
                    return True
            except:
                pass
            
            debug_print(f"Sheet '{sheet_name}' not identified as plotting sheet")
            return False
            
        except Exception as e:
            debug_print(f"Error determining if sheet is plotting sheet: {e}")
            return False
    
    # ===================== VAP3 FILE OPERATIONS =====================
    
    def save_to_vap3(self, filepath: str, vap3_data: VAP3FileData) -> bool:
        """Save data to a VAP3 file format."""
        debug_print(f"Saving to VAP3 file: {filepath}")
        
        try:
            # Ensure directory exists
            Path(filepath).parent.mkdir(parents=True, exist_ok=True)
            
            with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as archive:
                # Save metadata
                metadata = {
                    'version': vap3_data.version,
                    'timestamp': vap3_data.timestamp,
                    'sheet_count': len(vap3_data.filtered_sheets),
                    'sample_count': len(vap3_data.sample_images),
                    'has_sample_images': bool(vap3_data.sample_images),
                    'plot_options': vap3_data.plot_options
                }
                metadata.update(vap3_data.metadata)
                
                archive.writestr('metadata.json', json.dumps(metadata, indent=2))
                
                # Save sheet data as CSV files
                for sheet_name, sheet_info in vap3_data.filtered_sheets.items():
                    try:
                        if isinstance(sheet_info, dict) and 'data' in sheet_info:
                            sheet_data = sheet_info['data']
                            if isinstance(sheet_data, pd.DataFrame):
                                csv_buffer = sheet_data.to_csv(index=False)
                                archive.writestr(f'sheets/{sheet_name}.csv', csv_buffer)
                                
                                # Save sheet metadata
                                sheet_metadata = {
                                    'is_plotting': sheet_info.get('is_plotting', False),
                                    'is_empty': sheet_info.get('is_empty', True),
                                    'format': sheet_info.get('format', 'standard'),
                                    'row_count': len(sheet_data),
                                    'column_count': len(sheet_data.columns)
                                }
                                archive.writestr(f'sheets/{sheet_name}_metadata.json', 
                                               json.dumps(sheet_metadata))
                                
                                debug_print(f"Saved sheet: {sheet_name}")
                    except Exception as sheet_error:
                        debug_print(f"Error saving sheet {sheet_name}: {sheet_error}")
                        continue
                
                # Save sheet images
                if vap3_data.sheet_images:
                    for file_name, sheets in vap3_data.sheet_images.items():
                        for sheet_name, image_paths in sheets.items():
                            for i, img_path in enumerate(image_paths):
                                if os.path.exists(img_path):
                                    img_ext = os.path.splitext(img_path)[1]
                                    archive_path = f'images/{file_name}/{sheet_name}/image_{i}{img_ext}'
                                    archive.write(img_path, archive_path)
                
                # Save sample images and metadata
                if vap3_data.sample_images:
                    sample_images_metadata = {
                        'version': vap3_data.version,
                        'timestamp': vap3_data.timestamp,
                        'sample_count': len(vap3_data.sample_images),
                        'header_data': vap3_data.sample_header_data
                    }
                    archive.writestr('sample_images/metadata.json', 
                                   json.dumps(sample_images_metadata))
                    
                    for sample_id, image_paths in vap3_data.sample_images.items():
                        debug_print(f"Storing {len(image_paths)} images for {sample_id}")
                        
                        for i, img_path in enumerate(image_paths):
                            if os.path.exists(img_path):
                                img_ext = os.path.splitext(img_path)[1]
                                sample_img_path = f'sample_images/{sample_id}/image_{i}{img_ext}'
                                archive.write(img_path, sample_img_path)
                                debug_print(f"Stored sample image: {sample_img_path}")
                
                # Save various settings and states
                if vap3_data.image_crop_states:
                    archive.writestr('image_crop_states.json', 
                                   json.dumps(vap3_data.image_crop_states))
                
                if vap3_data.sample_image_crop_states:
                    archive.writestr('sample_images/crop_states.json', 
                                   json.dumps(vap3_data.sample_image_crop_states))
                
                if vap3_data.plot_settings:
                    archive.writestr('plot_settings.json', 
                                   json.dumps(vap3_data.plot_settings))
                
                if vap3_data.sample_notes_data:
                    archive.writestr('sample_notes.json', 
                                   json.dumps(vap3_data.sample_notes_data))
            
            debug_print(f"Successfully saved VAP3 file: {filepath}")
            return True
            
        except Exception as e:
            debug_print(f"Error saving VAP3 file: {e}")
            # Clean up potentially corrupted file
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass
            return False
    
    def load_from_vap3(self, filepath: str) -> Optional[VAP3FileData]:
        """Load data from a VAP3 file format."""
        debug_print(f"Loading from VAP3 file: {filepath}")
        
        try:
            # Validate file
            if not os.path.exists(filepath):
                debug_print(f"VAP3 file does not exist: {filepath}")
                return None
            
            vap3_data = VAP3FileData()
            temp_image_dir = None
            
            with zipfile.ZipFile(filepath, 'r') as archive:
                # Load metadata
                try:
                    metadata_content = archive.read('metadata.json')
                    metadata = json.loads(metadata_content)
                    vap3_data.version = metadata.get('version', '1.0')
                    vap3_data.timestamp = metadata.get('timestamp', '')
                    vap3_data.plot_options = metadata.get('plot_options', [])
                    vap3_data.metadata = metadata
                    debug_print(f"Loaded metadata: version {vap3_data.version}")
                except Exception as meta_error:
                    debug_print(f"Error loading metadata: {meta_error}")
                
                # Load sheet data
                sheet_files = [name for name in archive.namelist() if name.startswith('sheets/') and name.endswith('.csv')]
                
                for sheet_file in sheet_files:
                    try:
                        sheet_name = os.path.splitext(os.path.basename(sheet_file))[0]
                        
                        # Load CSV data
                        csv_content = archive.read(sheet_file)
                        sheet_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')))
                        
                        # Load sheet metadata if available
                        metadata_file = f'sheets/{sheet_name}_metadata.json'
                        sheet_metadata = {}
                        if metadata_file in archive.namelist():
                            try:
                                metadata_content = archive.read(metadata_file)
                                sheet_metadata = json.loads(metadata_content)
                            except:
                                pass
                        
                        # Create sheet info
                        sheet_info = {
                            'data': sheet_data,
                            'is_plotting': sheet_metadata.get('is_plotting', False),
                            'is_empty': sheet_metadata.get('is_empty', sheet_data.empty),
                            'format': sheet_metadata.get('format', 'standard')
                        }
                        
                        vap3_data.filtered_sheets[sheet_name] = sheet_info
                        debug_print(f"Loaded sheet: {sheet_name} ({len(sheet_data)} rows)")
                        
                    except Exception as sheet_error:
                        debug_print(f"Error loading sheet {sheet_file}: {sheet_error}")
                        continue
                
                # Load images if present
                image_files = [name for name in archive.namelist() if name.startswith('images/')]
                if image_files:
                    temp_image_dir = tempfile.mkdtemp()
                    self.temp_files.append(temp_image_dir)
                    
                    for img_file in image_files:
                        try:
                            # Extract image to temp location
                            temp_img_path = os.path.join(temp_image_dir, os.path.basename(img_file))
                            with open(temp_img_path, 'wb') as temp_file:
                                temp_file.write(archive.read(img_file))
                            
                            # Parse path structure: images/file_name/sheet_name/image_x.ext
                            path_parts = img_file.split('/')
                            if len(path_parts) >= 4:
                                file_name = path_parts[1]
                                sheet_name = path_parts[2]
                                
                                if file_name not in vap3_data.sheet_images:
                                    vap3_data.sheet_images[file_name] = {}
                                if sheet_name not in vap3_data.sheet_images[file_name]:
                                    vap3_data.sheet_images[file_name][sheet_name] = []
                                
                                vap3_data.sheet_images[file_name][sheet_name].append(temp_img_path)
                        
                        except Exception as img_error:
                            debug_print(f"Error loading image {img_file}: {img_error}")
                            continue
                
                # Load sample images if present
                sample_image_files = [name for name in archive.namelist() if name.startswith('sample_images/') and not name.endswith('.json')]
                if sample_image_files:
                    if not temp_image_dir:
                        temp_image_dir = tempfile.mkdtemp()
                        self.temp_files.append(temp_image_dir)
                    
                    for sample_img_file in sample_image_files:
                        try:
                            # Extract to temp location
                            temp_img_path = os.path.join(temp_image_dir, f"sample_{os.path.basename(sample_img_file)}")
                            with open(temp_img_path, 'wb') as temp_file:
                                temp_file.write(archive.read(sample_img_file))
                            
                            # Parse path: sample_images/sample_id/image_x.ext
                            path_parts = sample_img_file.split('/')
                            if len(path_parts) >= 3:
                                sample_id = path_parts[1]
                                
                                if sample_id not in vap3_data.sample_images:
                                    vap3_data.sample_images[sample_id] = []
                                
                                vap3_data.sample_images[sample_id].append(temp_img_path)
                        
                        except Exception as sample_img_error:
                            debug_print(f"Error loading sample image {sample_img_file}: {sample_img_error}")
                            continue
                
                # Load various settings
                settings_files = {
                    'image_crop_states.json': 'image_crop_states',
                    'sample_images/crop_states.json': 'sample_image_crop_states',
                    'plot_settings.json': 'plot_settings',
                    'sample_notes.json': 'sample_notes_data',
                    'sample_images/metadata.json': 'sample_header_data'
                }
                
                for file_name, attr_name in settings_files.items():
                    if file_name in archive.namelist():
                        try:
                            content = archive.read(file_name)
                            data = json.loads(content)
                            
                            if attr_name == 'sample_header_data':
                                # Extract header_data from sample images metadata
                                data = data.get('header_data', {})
                            
                            setattr(vap3_data, attr_name, data)
                            debug_print(f"Loaded {attr_name}")
                        except Exception as settings_error:
                            debug_print(f"Error loading {file_name}: {settings_error}")
            
            debug_print(f"Successfully loaded VAP3 file with {len(vap3_data.filtered_sheets)} sheets")
            debug_print(f"Sample images: {len(vap3_data.sample_images)} groups")
            return vap3_data
            
        except Exception as e:
            debug_print(f"Error loading VAP3 file: {e}")
            return None
    
    # ===================== BATCH FILE OPERATIONS =====================
    
    def scan_folder_for_excel_files(self, folder_path: str) -> List[str]:
        """Recursively scan folder for Excel files."""
        debug_print(f"Starting recursive scan of folder: {folder_path}")
        
        excel_files = []
        
        try:
            for root, dirs, files in os.walk(folder_path):
                debug_print(f"Scanning directory: {root}")
                
                for file in files:
                    if file.lower().endswith(('.xlsx', '.xls')):
                        full_path = os.path.join(root, file)
                        
                        # Skip temporary Excel files
                        if file.startswith('~$'):
                            debug_print(f"Skipping temporary file: {file}")
                            continue
                        
                        # Check if file is accessible
                        try:
                            if os.access(full_path, os.R_OK):
                                excel_files.append(full_path)
                                debug_print(f"Added Excel file: {file}")
                            else:
                                debug_print(f"Skipping inaccessible file: {file}")
                        except Exception as e:
                            debug_print(f"Error checking file access for {file}: {e}")
                            continue
            
            debug_print(f"Completed folder scan, found {len(excel_files)} Excel files")
            return excel_files
            
        except Exception as e:
            debug_print(f"Failed to scan folder: {e}")
            raise
    
    def filter_files_by_test_names(self, excel_files: List[str], test_names: List[str]) -> List[str]:
        """Filter Excel files by checking if filename contains any test names."""
        debug_print(f"Filtering {len(excel_files)} files against {len(test_names)} test names")
        
        matching_files = []
        
        for file_path in excel_files:
            filename = os.path.basename(file_path).lower()
            debug_print(f"Checking file: {filename}")
            
            # Check if any test name is contained in the filename
            for test_name in test_names:
                if test_name.lower() in filename:
                    matching_files.append(file_path)
                    debug_print(f"MATCH - '{filename}' contains '{test_name}'")
                    break
        
        debug_print(f"Filtered to {len(matching_files)} matching files")
        return matching_files
    
    # ===================== FILE UTILITIES =====================
    
    def get_file_hash(self, filepath: str) -> Optional[str]:
        """Calculate MD5 hash of a file."""
        try:
            hasher = hashlib.md5()
            with open(filepath, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception as e:
            debug_print(f"Error calculating hash for {filepath}: {e}")
            return None
    
    def ensure_directory_exists(self, directory: str):
        """Ensure directory exists, create if necessary."""
        try:
            Path(directory).mkdir(parents=True, exist_ok=True)
            debug_print(f"Directory ensured: {directory}")
        except Exception as e:
            debug_print(f"Error creating directory {directory}: {e}")
            raise
    
    def cleanup_temp_files(self):
        """Clean up temporary files created during operations."""
        for temp_path in self.temp_files:
            try:
                if os.path.isfile(temp_path):
                    os.remove(temp_path)
                elif os.path.isdir(temp_path):
                    shutil.rmtree(temp_path)
                debug_print(f"Cleaned up temporary file: {temp_path}")
            except Exception as e:
                debug_print(f"Error cleaning up {temp_path}: {e}")
        
        self.temp_files.clear()
    
    def open_file_in_excel(self, filepath: str):
        """Open a file in Excel for editing."""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(filepath)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.call(['open', filepath])
            else:
                debug_print(f"Unsupported OS for opening Excel: {os.name}")
                return False
            
            debug_print(f"Opened file in Excel: {filepath}")
            return True
            
        except Exception as e:
            debug_print(f"Error opening file in Excel: {e}")
            return False
    
    # ===================== CALCULATION METHODS =====================
    
    def calculate_tpm_from_weights(self, puffs: pd.Series, before_weights: pd.Series, 
                                   after_weights: pd.Series) -> pd.Series:
        """Calculate TPM from weight differences and puffing intervals."""
        try:
            # Convert to numeric, handling any string values
            puffs_numeric = pd.to_numeric(puffs, errors='coerce')
            before_numeric = pd.to_numeric(before_weights, errors='coerce')
            after_numeric = pd.to_numeric(after_weights, errors='coerce')
            
            # Calculate weight differences
            weight_diff = before_numeric - after_numeric
            
            # Calculate TPM (weight difference / number of puffs)
            # Handle division by zero
            tpm = weight_diff / puffs_numeric.replace(0, np.nan)
            
            debug_print(f"Calculated TPM from weights - samples: {len(tpm)}")
            debug_print(f"TPM range: {tpm.min():.3f} to {tpm.max():.3f}")
            
            return tpm.round(self.precision_digits)
            
        except Exception as e:
            debug_print(f"Error calculating TPM from weights: {e}")
            return pd.Series(dtype=float)
    
    def calculate_statistics(self, data: pd.Series) -> Dict[str, float]:
        """Calculate basic statistics for a data series."""
        try:
            clean_data = pd.to_numeric(data, errors='coerce').dropna()
            
            if len(clean_data) < self.statistical_threshold:
                debug_print(f"Insufficient data for statistics: {len(clean_data)} points")
                return {}
            
            stats = {
                'count': len(clean_data),
                'mean': round_values(clean_data.mean(), self.precision_digits),
                'std': round_values(clean_data.std(), self.precision_digits),
                'min': round_values(clean_data.min(), self.precision_digits),
                'max': round_values(clean_data.max(), self.precision_digits),
                'median': round_values(clean_data.median(), self.precision_digits)
            }
            
            debug_print(f"Calculated statistics for {len(clean_data)} data points")
            return stats
            
        except Exception as e:
            debug_print(f"Error calculating statistics: {e}")
            return {}
    
    # ===================== UTILITY METHODS =====================
    
    def get_model_stats(self) -> Dict[str, Any]:
        """Get comprehensive statistics about the file model."""
        return {
            'total_file_states': len(self.file_states),
            'total_caches': len(self.file_caches),
            'loaded_files_count': len(self.loaded_files_cache),
            'stored_files_count': len(self.stored_files_cache),
            'cache_duration_hours': self.cache_duration_hours,
            'supported_formats': self.supported_formats,
            'temp_files_count': len(self.temp_files),
            'file_states_by_status': self._get_files_by_status(),
            'cache_info': self._get_cache_statistics(),
            'database_connected': self.database_service is not None
        }
    
    def _get_files_by_status(self) -> Dict[str, int]:
        """Get file counts by status."""
        status_counts = {}
        for file_state in self.file_states.values():
            status = file_state.load_status
            status_counts[status] = status_counts.get(status, 0) + 1
        return status_counts
    
    def _get_cache_statistics(self) -> Dict[str, Any]:
        """Get cache statistics."""
        total_cache_size = 0
        total_cache_items = 0
        
        for cache in self.file_caches.values():
            cache_info = cache.get_cache_info()
            total_cache_size += cache_info['cache_size_bytes']
            total_cache_items += cache_info['item_count']
        
        return {
            'total_caches': len(self.file_caches),
            'total_cache_items': total_cache_items,
            'total_cache_size_bytes': total_cache_size,
            'total_cache_size_mb': round(total_cache_size / (1024 * 1024), 2)
        }
    
    def export_file_states(self) -> Dict[str, Any]:
        """Export file states for debugging or persistence."""
        export_data = {
            'export_timestamp': datetime.now().isoformat(),
            'file_states': {},
            'cache_info': self._get_cache_statistics(),
            'model_stats': self.get_model_stats()
        }
        
        for filepath, file_state in self.file_states.items():
            export_data['file_states'][filepath] = {
                'filename': file_state.filename,
                'file_size': file_state.file_size,
                'file_hash': file_state.file_hash,
                'last_modified': file_state.last_modified.isoformat(),
                'load_status': file_state.load_status,
                'processing_status': file_state.processing_status,
                'error_message': file_state.error_message,
                'load_time_seconds': file_state.load_time_seconds,
                'is_cached': file_state.is_cached,
                'file_type': file_state.file_type,
                'legacy_mode': file_state.legacy_mode,
                'sheet_count': file_state.sheet_count,
                'database_stored': file_state.database_stored
            }
        
        return export_data


# Export the main classes
__all__ = ['FileModel', 'FileState', 'FileCache', 'VAP3FileData']

# Debug output for model initialization
debug_print("FileModel module loaded successfully")
debug_print("Available classes: FileModel, FileState, FileCache, VAP3FileData")
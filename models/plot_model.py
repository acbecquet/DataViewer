"""
models/file_model.py
File management models for state tracking, caching, and file operations.
These models will replace file state management currently in file_manager.py.
"""

from typing import Dict, List, Optional, Set, Any
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
import hashlib


@dataclass
class FileState:
    """Model for tracking file loading and processing state."""
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
    
    def __post_init__(self):
        """Post-initialization processing."""
        self._calculate_file_hash()
        print(f"DEBUG: Created FileState for '{self.filename}'")
        print(f"DEBUG: File size: {self.file_size} bytes, Hash: {self.file_hash[:8]}...")
    
    def _calculate_file_hash(self):
        """Calculate MD5 hash of the file for change detection."""
        try:
            if Path(self.filepath).exists():
                hasher = hashlib.md5()
                with open(self.filepath, 'rb') as f:
                    for chunk in iter(lambda: f.read(4096), b""):
                        hasher.update(chunk)
                self.file_hash = hasher.hexdigest()
                print(f"DEBUG: Calculated hash for '{self.filename}': {self.file_hash[:8]}...")
            else:
                print(f"WARNING: File not found for hash calculation: {self.filepath}")
        except Exception as e:
            print(f"ERROR: Failed to calculate hash for '{self.filename}': {e}")
            self.file_hash = None
    
    def set_loading(self):
        """Mark file as currently loading."""
        self.load_status = "loading"
        print(f"DEBUG: File '{self.filename}' status: loading")
    
    def set_loaded(self, load_time: float):
        """Mark file as successfully loaded."""
        self.load_status = "loaded"
        self.load_time_seconds = load_time
        print(f"DEBUG: File '{self.filename}' loaded successfully in {load_time:.2f}s")
    
    def set_error(self, error_message: str):
        """Mark file as having an error."""
        self.load_status = "error"
        self.error_message = error_message
        print(f"ERROR: File '{self.filename}' failed to load: {error_message}")
    
    def set_processing(self):
        """Mark file as currently processing."""
        self.processing_status = "processing"
        print(f"DEBUG: File '{self.filename}' status: processing")
    
    def set_processed(self):
        """Mark file as successfully processed."""
        self.processing_status = "processed"
        print(f"DEBUG: File '{self.filename}' processed successfully")
    
    def set_cached(self, cache_duration_hours: int = 24):
        """Mark file as cached with expiry time."""
        self.is_cached = True
        self.cache_expiry = datetime.now() + timedelta(hours=cache_duration_hours)
        print(f"DEBUG: File '{self.filename}' cached until {self.cache_expiry}")
    
    def is_cache_valid(self) -> bool:
        """Check if the cached data is still valid."""
        if not self.is_cached or not self.cache_expiry:
            return False
        
        is_valid = datetime.now() < self.cache_expiry
        if not is_valid:
            print(f"DEBUG: Cache expired for '{self.filename}'")
        return is_valid
    
    def has_file_changed(self) -> bool:
        """Check if the file has changed since last hash calculation."""
        if not self.file_hash:
            return True
        
        try:
            if not Path(self.filepath).exists():
                print(f"WARNING: File no longer exists: {self.filepath}")
                return True
            
            current_stat = Path(self.filepath).stat()
            if current_stat.st_size != self.file_size:
                print(f"DEBUG: File size changed for '{self.filename}'")
                return True
            
            # Quick check: if modification time is newer, file likely changed
            current_mtime = datetime.fromtimestamp(current_stat.st_mtime)
            if current_mtime > self.last_modified:
                print(f"DEBUG: File modification time changed for '{self.filename}'")
                return True
            
            return False
        except Exception as e:
            print(f"ERROR: Failed to check file changes for '{self.filename}': {e}")
            return True


@dataclass
class FileCache:
    """Model for caching file data and processing results."""
    cache_data: Dict[str, Any] = field(default_factory=dict)
    created_at: datetime = field(default_factory=datetime.now)
    access_count: int = 0
    last_accessed: Optional[datetime] = None
    cache_size_bytes: int = 0
    
    def __post_init__(self):
        """Post-initialization processing."""
        print(f"DEBUG: Created FileCache with {len(self.cache_data)} items")
    
    def store_data(self, key: str, data: Any):
        """Store data in the cache."""
        self.cache_data[key] = data
        self.last_accessed = datetime.now()
        
        # Estimate cache size (rough approximation)
        try:
            import sys
            self.cache_size_bytes = sys.getsizeof(self.cache_data)
        except:
            pass
        
        print(f"DEBUG: Stored cache data for key '{key}' - {len(self.cache_data)} total items")
    
    def get_data(self, key: str) -> Optional[Any]:
        """Retrieve data from the cache."""
        if key in self.cache_data:
            self.access_count += 1
            self.last_accessed = datetime.now()
            print(f"DEBUG: Retrieved cache data for key '{key}' (access #{self.access_count})")
            return self.cache_data[key]
        
        print(f"DEBUG: Cache miss for key '{key}'")
        return None
    
    def has_data(self, key: str) -> bool:
        """Check if cache has data for a specific key."""
        return key in self.cache_data
    
    def remove_data(self, key: str) -> bool:
        """Remove data from cache."""
        if key in self.cache_data:
            del self.cache_data[key]
            print(f"DEBUG: Removed cache data for key '{key}'")
            return True
        return False
    
    def clear_cache(self):
        """Clear all cached data."""
        item_count = len(self.cache_data)
        self.cache_data.clear()
        self.access_count = 0
        self.cache_size_bytes = 0
        print(f"DEBUG: Cleared cache ({item_count} items removed)")
    
    def get_cache_info(self) -> Dict[str, Any]:
        """Get information about the cache."""
        return {
            'item_count': len(self.cache_data),
            'created_at': self.created_at,
            'last_accessed': self.last_accessed,
            'access_count': self.access_count,
            'cache_size_bytes': self.cache_size_bytes,
            'keys': list(self.cache_data.keys())
        }


class FileModel:
    """Main file model that manages file states and caching."""
    
    def __init__(self, cache_duration_hours: int = 24):
        """Initialize the file model."""
        self.file_states: Dict[str, FileState] = {}  # filepath -> FileState
        self.file_caches: Dict[str, FileCache] = {}  # filepath -> FileCache
        self.loaded_files_cache: Set[str] = set()  # filepaths of loaded files
        self.stored_files_cache: Set[str] = set()  # filepaths of database-stored files
        self.cache_duration_hours = cache_duration_hours
        
        print("DEBUG: FileModel initialized")
        print(f"DEBUG: Default cache duration: {cache_duration_hours} hours")
    
    def add_file_state(self, filepath: str, filename: str) -> FileState:
        """Add or update file state tracking."""
        try:
            file_path = Path(filepath)
            file_size = file_path.stat().st_size if file_path.exists() else 0
            
            file_state = FileState(
                filepath=filepath,
                filename=filename,
                file_size=file_size
            )
            
            self.file_states[filepath] = file_state
            print(f"DEBUG: Added file state for '{filename}' at '{filepath}'")
            return file_state
            
        except Exception as e:
            print(f"ERROR: Failed to add file state for '{filename}': {e}")
            # Create a minimal file state even if we can't read the file
            file_state = FileState(
                filepath=filepath,
                filename=filename,
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
    
    def create_cache(self, filepath: str) -> FileCache:
        """Create a cache for a specific file."""
        cache = FileCache()
        self.file_caches[filepath] = cache
        print(f"DEBUG: Created cache for file '{filepath}'")
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
        print(f"DEBUG: Added '{filepath}' to loaded files cache")
    
    def is_in_loaded_cache(self, filepath: str) -> bool:
        """Check if file is in loaded cache."""
        return filepath in self.loaded_files_cache
    
    def add_to_stored_cache(self, filepath: str):
        """Mark a file as stored in database cache."""
        self.stored_files_cache.add(filepath)
        print(f"DEBUG: Added '{filepath}' to stored files cache")
    
    def is_in_stored_cache(self, filepath: str) -> bool:
        """Check if file is in database stored cache."""
        return filepath in self.stored_files_cache
    
    def should_reload_file(self, filepath: str) -> bool:
        """Determine if a file should be reloaded."""
        file_state = self.get_file_state(filepath)
        if not file_state:
            print(f"DEBUG: No file state found for '{filepath}' - should reload")
            return True
        
        if file_state.load_status == "error":
            print(f"DEBUG: Previous error for '{filepath}' - should reload")
            return True
        
        if file_state.has_file_changed():
            print(f"DEBUG: File changed '{filepath}' - should reload")
            return True
        
        if file_state.is_cached and not file_state.is_cache_valid():
            print(f"DEBUG: Cache expired for '{filepath}' - should reload")
            return True
        
        print(f"DEBUG: File '{filepath}' is up to date - no reload needed")
        return False
    
    def clear_file_state(self, filepath: str):
        """Clear state and cache for a specific file."""
        if filepath in self.file_states:
            del self.file_states[filepath]
            print(f"DEBUG: Cleared file state for '{filepath}'")
        
        if filepath in self.file_caches:
            del self.file_caches[filepath]
            print(f"DEBUG: Cleared file cache for '{filepath}'")
        
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
        
        print(f"DEBUG: Cleared all caches - {state_count} states, {cache_count} caches")
    
    def get_model_stats(self) -> Dict[str, Any]:
        """Get statistics about the file model."""
        return {
            'total_file_states': len(self.file_states),
            'total_caches': len(self.file_caches),
            'loaded_files_count': len(self.loaded_files_cache),
            'stored_files_count': len(self.stored_files_cache),
            'cache_duration_hours': self.cache_duration_hours,
            'file_states_list': list(self.file_states.keys()),
            'cached_files_list': list(self.file_caches.keys())
        }
    
    def cleanup_expired_caches(self):
        """Remove expired caches to free memory."""
        expired_files = []
        
        for filepath, file_state in self.file_states.items():
            if file_state.is_cached and not file_state.is_cache_valid():
                expired_files.append(filepath)
        
        for filepath in expired_files:
            self.clear_file_state(filepath)
        
        if expired_files:
            print(f"DEBUG: Cleaned up {len(expired_files)} expired file caches")
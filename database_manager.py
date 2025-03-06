"""
database_manager.py
Developed By Charlie Becquet.
Manages the SQL database for the DataViewer Application.

This module provides functions to store and retrieve .vap3 files and their meta_data
in an SQL database, enabling persistent storage and sharing of test data.
"""

import os
import datetime
import json
import base64
import io
import logging
from typing import Dict, List, Any, Optional, Tuple

import sqlalchemy
from sqlalchemy import create_engine, Column, Integer, String, DateTime, LargeBinary, Text, ForeignKey, Boolean, Table
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship, Session

# Set up logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Create a base class for declarative class definitions
Base = declarative_base()

# Define the association table for the many-to-many relationship between Files and Images
file_image_association = Table(
    'file_image_association', 
    Base.metadata,
    Column('file_id', Integer, ForeignKey('files.id')),
    Column('image_id', Integer, ForeignKey('images.id'))
)

class File(Base):
    """Represents a .vap3 file in the database."""
    __tablename__ = 'files'
    
    id = Column(Integer, primary_key=True)
    filename = Column(String(255), nullable=False)
    original_path = Column(String(512))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    file_content = Column(LargeBinary, nullable=False)  # The actual .vap3 file content
    meta_data = Column(Text)  # JSON string containing meta_data
    
    # Relationships
    sheets = relationship("Sheet", back_populates="file", cascade="all, delete-orphan")
    images = relationship("Image", secondary=file_image_association, back_populates="files")
    
    def __repr__(self):
        return f"<File(id={self.id}, filename='{self.filename}', created_at='{self.created_at}')>"

class Sheet(Base):
    """Represents a sheet within a .vap3 file."""
    __tablename__ = 'sheets'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('files.id'))
    name = Column(String(255), nullable=False)
    is_plotting = Column(Boolean, default=False)
    is_empty = Column(Boolean, default=False)
    
    # Relationships
    file = relationship("File", back_populates="sheets")
    
    def __repr__(self):
        return f"<Sheet(id={self.id}, name='{self.name}', is_plotting={self.is_plotting})>"

class Image(Base):
    """Represents an image stored in the database."""
    __tablename__ = 'images'
    
    id = Column(Integer, primary_key=True)
    filename = Column(String(255), nullable=False)
    content = Column(LargeBinary, nullable=False)
    mime_type = Column(String(100), default="image/jpeg")
    sheet_name = Column(String(255))  # The sheet this image is associated with
    crop_enabled = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.datetime.now)
    
    # Relationships
    files = relationship("File", secondary=file_image_association, back_populates="images")
    
    def __repr__(self):
        return f"<Image(id={self.id}, filename='{self.filename}', sheet_name='{self.sheet_name}')>"

class DatabaseManager:
    """Manages database connections and operations for DataViewer."""
    
    def __init__(self, db_path: str = None):
        """
        Initialize the DatabaseManager.
        
        Args:
            db_path: Path to the database file. If None, uses a default path.
        """
        if db_path is None:
            # Create database in user's home directory
            home_dir = os.path.expanduser("~")
            app_dir = os.path.join(home_dir, ".dataviewer")
            os.makedirs(app_dir, exist_ok=True)
            db_path = os.path.join(app_dir, "dataviewer.db")
        
        self.db_path = db_path
        self.engine = create_engine(f'sqlite:///{db_path}')
        self.Session = sessionmaker(bind=self.engine)
        
        # Create tables if they don't exist
        Base.metadata.create_all(self.engine)
        logger.info(f"Database initialized at {db_path}")
    
    def store_vap3_file(self, filepath: str, meta_data: Dict = None) -> int:
        """
        Store a .vap3 file in the database.
        
        Args:
            filepath: Path to the .vap3 file
            meta_data: Dictionary of meta_data about the file
            
        Returns:
            int: ID of the created file record
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        
        filename = os.path.basename(filepath)
        
        # Read the file content
        with open(filepath, 'rb') as f:
            file_content = f.read()
        
        # Serialize meta_data to JSON
        if meta_data:
            meta_data_json = json.dumps(meta_data)
        else:
            meta_data_json = None
        
        with self.Session() as session:
            # Check if file already exists
            existing_file = session.query(File).filter_by(filename=filename).first()
            
            if existing_file:
                # Update existing file
                existing_file.file_content = file_content
                existing_file.meta_data = meta_data_json
                existing_file.updated_at = datetime.datetime.now()
                file_id = existing_file.id
                logger.info(f"Updated existing file: {filename} (ID: {file_id})")
            else:
                # Create new file record
                new_file = File(
                    filename=filename,
                    original_path=filepath,
                    file_content=file_content,
                    meta_data=meta_data_json
                )
                session.add(new_file)
                session.flush()  # Flush to get the ID
                file_id = new_file.id
                logger.info(f"Stored new file: {filename} (ID: {file_id})")
            
            session.commit()
            return file_id
    
    def store_sheet_info(self, file_id: int, sheet_name: str, is_plotting: bool, is_empty: bool) -> int:
        """
        Store information about a sheet associated with a file.
        
        Args:
            file_id: The ID of the parent file
            sheet_name: Name of the sheet
            is_plotting: Whether this sheet is a plotting sheet
            is_empty: Whether this sheet is empty
            
        Returns:
            int: ID of the created sheet record
        """
        with self.Session() as session:
            # Check if sheet already exists for this file
            existing_sheet = session.query(Sheet).filter_by(
                file_id=file_id, name=sheet_name).first()
            
            if existing_sheet:
                # Update existing sheet
                existing_sheet.is_plotting = is_plotting
                existing_sheet.is_empty = is_empty
                sheet_id = existing_sheet.id
                logger.info(f"Updated sheet: {sheet_name} for file ID {file_id}")
            else:
                # Create new sheet record
                new_sheet = Sheet(
                    file_id=file_id,
                    name=sheet_name,
                    is_plotting=is_plotting,
                    is_empty=is_empty
                )
                session.add(new_sheet)
                session.flush()  # Flush to get the ID
                sheet_id = new_sheet.id
                logger.info(f"Stored new sheet: {sheet_name} for file ID {file_id}")
            
            session.commit()
            return sheet_id
    
    def store_image(self, file_id: int, image_path: str, sheet_name: str, crop_enabled: bool) -> int:
        """
        Store an image associated with a file.
        
        Args:
            file_id: The ID of the parent file
            image_path: Path to the image file
            sheet_name: Name of the sheet this image is associated with
            crop_enabled: Whether auto-crop is enabled for this image
            
        Returns:
            int: ID of the created image record
        """
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Image not found: {image_path}")
        
        filename = os.path.basename(image_path)
        
        # Read the image content
        with open(image_path, 'rb') as f:
            image_content = f.read()
        
        # Determine MIME type based on file extension
        _, ext = os.path.splitext(filename)
        mime_type = {
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.pdf': 'application/pdf'
        }.get(ext.lower(), 'application/octet-stream')
        
        with self.Session() as session:
            # Check if file exists
            file = session.query(File).filter_by(id=file_id).first()
            if not file:
                raise ValueError(f"File with ID {file_id} not found")
            
            # Check if this image already exists for this file and sheet
            existing_image = None
            for img in file.images:
                if img.filename == filename and img.sheet_name == sheet_name:
                    existing_image = img
                    break
            
            if existing_image:
                # Update existing image
                existing_image.content = image_content
                existing_image.crop_enabled = crop_enabled
                existing_image.mime_type = mime_type
                image_id = existing_image.id
                logger.info(f"Updated image: {filename} for sheet {sheet_name}")
            else:
                # Create new image record
                new_image = Image(
                    filename=filename,
                    content=image_content,
                    mime_type=mime_type,
                    sheet_name=sheet_name,
                    crop_enabled=crop_enabled
                )
                session.add(new_image)
                session.flush()  # Flush to get the ID
                
                # Associate image with file
                file.images.append(new_image)
                
                image_id = new_image.id
                logger.info(f"Stored new image: {filename} for sheet {sheet_name}")
            
            session.commit()
            return image_id
    
    def get_file_by_id(self, file_id: int) -> Optional[Dict]:
        """
        Retrieve a file from the database by ID.
        
        Args:
            file_id: The ID of the file to retrieve
            
        Returns:
            dict: Dictionary containing file data, or None if not found
        """
        with self.Session() as session:
            file = session.query(File).filter_by(id=file_id).first()
            if not file:
                return None
            
            # Build result dictionary
            result = {
                'id': file.id,
                'filename': file.filename,
                'original_path': file.original_path,
                'created_at': file.created_at,
                'updated_at': file.updated_at,
                'file_content': file.file_content,
                'meta_data': json.loads(file.meta_data) if file.meta_data else {},
                'sheets': []
            }
            
            # Add sheets
            for sheet in file.sheets:
                result['sheets'].append({
                    'id': sheet.id,
                    'name': sheet.name,
                    'is_plotting': sheet.is_plotting,
                    'is_empty': sheet.is_empty
                })
            
            return result
    
    def get_file_by_name(self, filename: str) -> Optional[Dict]:
        """
        Retrieve a file from the database by filename.
        
        Args:
            filename: The name of the file to retrieve
            
        Returns:
            dict: Dictionary containing file data, or None if not found
        """
        with self.Session() as session:
            file = session.query(File).filter_by(filename=filename).first()
            if not file:
                return None
            
            return self.get_file_by_id(file.id)
    
    def save_vap3_to_disk(self, file_id: int, output_path: str) -> bool:
        """
        Save a .vap3 file from the database to disk.
        
        Args:
            file_id: The ID of the file to save
            output_path: Path where the file should be saved
            
        Returns:
            bool: True if successful, False otherwise
        """
        file_data = self.get_file_by_id(file_id)
        if not file_data:
            logger.error(f"File with ID {file_id} not found")
            return False
        
        try:
            with open(output_path, 'wb') as f:
                f.write(file_data['file_content'])
            logger.info(f"File saved to {output_path}")
            return True
        except Exception as e:
            logger.error(f"Error saving file to disk: {e}")
            return False
    
    def list_files(self, limit: int = 100) -> List[Dict]:
        """
        List files in the database.
        
        Args:
            limit: Maximum number of files to return
            
        Returns:
            list: List of dictionaries containing file information
        """
        with self.Session() as session:
            files = session.query(File).order_by(File.updated_at.desc()).limit(limit).all()
            
            result = []
            for file in files:
                sheet_count = len(file.sheets)
                image_count = len(file.images)
                
                result.append({
                    'id': file.id,
                    'filename': file.filename,
                    'created_at': file.created_at,
                    'updated_at': file.updated_at,
                    'sheet_count': sheet_count,
                    'image_count': image_count
                })
            
            return result
    
    def get_images_for_sheet(self, file_id: int, sheet_name: str) -> List[Dict]:
        """
        Get all images associated with a specific sheet in a file.
        
        Args:
            file_id: The ID of the file
            sheet_name: The name of the sheet
            
        Returns:
            list: List of dictionaries containing image information
        """
        with self.Session() as session:
            file = session.query(File).filter_by(id=file_id).first()
            if not file:
                return []
            
            result = []
            for image in file.images:
                if image.sheet_name == sheet_name:
                    result.append({
                        'id': image.id,
                        'filename': image.filename,
                        'mime_type': image.mime_type,
                        'crop_enabled': image.crop_enabled,
                        'created_at': image.created_at,
                        'content': image.content  # Binary content
                    })
            
            return result
    
    def save_image_to_disk(self, image_id: int, output_path: str) -> bool:
        """
        Save an image from the database to disk.
        
        Args:
            image_id: The ID of the image to save
            output_path: Path where the image should be saved
            
        Returns:
            bool: True if successful, False otherwise
        """
        with self.Session() as session:
            image = session.query(Image).filter_by(id=image_id).first()
            if not image:
                logger.error(f"Image with ID {image_id} not found")
                return False
            
            try:
                with open(output_path, 'wb') as f:
                    f.write(image.content)
                logger.info(f"Image saved to {output_path}")
                return True
            except Exception as e:
                logger.error(f"Error saving image to disk: {e}")
                return False
    
    def delete_file(self, file_id: int) -> bool:
        """
        Delete a file from the database.
        
        Args:
            file_id: The ID of the file to delete
            
        Returns:
            bool: True if successful, False otherwise
        """
        with self.Session() as session:
            file = session.query(File).filter_by(id=file_id).first()
            if not file:
                logger.error(f"File with ID {file_id} not found")
                return False
            
            try:
                session.delete(file)
                session.commit()
                logger.info(f"Deleted file with ID {file_id}")
                return True
            except Exception as e:
                session.rollback()
                logger.error(f"Error deleting file: {e}")
                return False
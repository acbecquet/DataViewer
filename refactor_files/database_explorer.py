"""
database_explorer.py - Comprehensive Database Content Explorer

This tool helps you explore and analyze the contents of your TestingGUI database,
including file sizes, storage statistics, and content analysis.

Usage: python database_explorer.py
"""

import os
import sys
import json
import sqlite3
from datetime import datetime, timedelta
from typing import List, Tuple, Dict, Any

# Add current directory to path so we can import from the main application
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

try:
    from database_manager import DatabaseManager
    from utils import debug_print
except ImportError as e:
    print(f"❌ Error importing required modules: {e}")
    print("Make sure you're running this from the TestingGUI directory")
    sys.exit(1)


class DatabaseExplorer:
    """Comprehensive database exploration and analysis tool"""
    
    def __init__(self):
        """Initialize the database explorer"""
        try:
            self.db_manager = DatabaseManager()
            print("✅ Connected to database successfully")
        except Exception as e:
            print(f"❌ Failed to connect to database: {e}")
            sys.exit(1)
    
    def format_size(self, size_bytes: int) -> str:
        """Format bytes into human readable size"""
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
    
    def get_database_overview(self) -> Dict[str, Any]:
        """Get basic database statistics"""
        cursor = self.db_manager.conn.cursor()
        
        # File statistics
        cursor.execute("""
            SELECT 
                COUNT(*) as file_count,
                SUM(LENGTH(file_content)) as total_size,
                AVG(LENGTH(file_content)) as avg_size,
                MAX(LENGTH(file_content)) as max_size,
                MIN(LENGTH(file_content)) as min_size
            FROM files
        """)
        
        file_stats = cursor.fetchone()
        
        # Recent activity (last 7 days)
        cursor.execute("""
            SELECT COUNT(*) 
            FROM files 
            WHERE created_at > datetime('now', '-7 days')
        """)
        recent_files = cursor.fetchone()[0]
        
        # Sheet and image counts
        cursor.execute("SELECT COUNT(*) FROM sheets")
        sheet_count = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(*) FROM images")
        image_count = cursor.fetchone()[0]
        
        return {
            'file_count': file_stats[0] or 0,
            'total_size': file_stats[1] or 0,
            'avg_size': file_stats[2] or 0,
            'max_size': file_stats[3] or 0,
            'min_size': file_stats[4] or 0,
            'recent_files': recent_files,
            'sheet_count': sheet_count,
            'image_count': image_count
        }
    
    def show_overview(self):
        """Display database overview"""
        print("=" * 80)
        print("📊 DATABASE OVERVIEW")
        print("=" * 80)
        
        try:
            stats = self.get_database_overview()
            
            if stats['file_count'] == 0:
                print("📭 Database is empty - no files stored yet")
                return
            
            print(f"📁 Total Files: {stats['file_count']:,}")
            print(f"💾 Total Size: {self.format_size(stats['total_size'])}")
            print(f"📏 Average File Size: {self.format_size(int(stats['avg_size']))}")
            print(f"📈 Largest File: {self.format_size(stats['max_size'])}")
            print(f"📉 Smallest File: {self.format_size(stats['min_size'])}")
            print(f"🕒 Files Added (Last 7 Days): {stats['recent_files']}")
            print(f"📋 Sheet Records: {stats['sheet_count']}")
            print(f"🖼️  Image Records: {stats['image_count']}")
            
        except Exception as e:
            print(f"❌ Error getting overview: {e}")
    
    def show_file_list(self, limit: int = None, sort_by: str = "size"):
        """Show detailed file list"""
        print("=" * 80)
        print("📄 DETAILED FILE LIST")
        print("=" * 80)
        
        try:
            cursor = self.db_manager.conn.cursor()
            
            # Build query based on sort preference
            if sort_by == "size":
                order_clause = "ORDER BY LENGTH(file_content) DESC"
                print("Sorted by: File Size (Largest First)")
            elif sort_by == "date":
                order_clause = "ORDER BY created_at DESC"
                print("Sorted by: Date Created (Newest First)")
            elif sort_by == "name":
                order_clause = "ORDER BY filename ASC"
                print("Sorted by: Filename (A-Z)")
            else:
                order_clause = "ORDER BY LENGTH(file_content) DESC"
                print("Sorted by: File Size (Largest First)")
            
            query = f"""
                SELECT 
                    id,
                    filename,
                    LENGTH(file_content) as size_bytes,
                    created_at,
                    meta_data
                FROM files 
                {order_clause}
            """
            
            if limit:
                query += f" LIMIT {limit}"
            
            cursor.execute(query)
            files = cursor.fetchall()
            
            if not files:
                print("📭 No files found in database")
                return
            
            print()
            print(f"{'ID':<4} {'Size':<12} {'Filename':<45} {'Created':<20}")
            print("-" * 85)
            
            for file_id, filename, size_bytes, created_at, meta_data in files:
                size_str = self.format_size(size_bytes)
                
                # Truncate filename if too long
                display_filename = filename[:42] + "..." if len(filename) > 45 else filename
                
                # Format date
                try:
                    if created_at:
                        if isinstance(created_at, str):
                            date_obj = datetime.fromisoformat(created_at.replace('Z', '+00:00'))
                        else:
                            date_obj = created_at
                        date_str = date_obj.strftime("%Y-%m-%d %H:%M")
                    else:
                        date_str = "Unknown"
                except:
                    date_str = str(created_at)[:16] if created_at else "Unknown"
                
                print(f"{file_id:<4} {size_str:<12} {display_filename:<45} {date_str:<20}")
            
            print("-" * 85)
            print(f"Showing {len(files)} files")
            
        except Exception as e:
            print(f"❌ Error listing files: {e}")
    
    def show_size_distribution(self):
        """Show file size distribution"""
        print("=" * 80)
        print("📊 FILE SIZE DISTRIBUTION")
        print("=" * 80)
        
        try:
            cursor = self.db_manager.conn.cursor()
            
            cursor.execute("""
                SELECT 
                    CASE 
                        WHEN LENGTH(file_content) < 1024 THEN 'Under 1 KB'
                        WHEN LENGTH(file_content) < 1024*1024 THEN '1 KB - 1 MB'
                        WHEN LENGTH(file_content) < 10*1024*1024 THEN '1 MB - 10 MB'
                        WHEN LENGTH(file_content) < 100*1024*1024 THEN '10 MB - 100 MB'
                        ELSE 'Over 100 MB'
                    END as size_category,
                    COUNT(*) as count,
                    SUM(LENGTH(file_content)) as total_bytes
                FROM files
                GROUP BY size_category
                ORDER BY 
                    CASE size_category
                        WHEN 'Under 1 KB' THEN 1
                        WHEN '1 KB - 1 MB' THEN 2
                        WHEN '1 MB - 10 MB' THEN 3
                        WHEN '10 MB - 100 MB' THEN 4
                        WHEN 'Over 100 MB' THEN 5
                    END
            """)
            
            results = cursor.fetchall()
            
            if not results:
                print("📭 No files to analyze")
                return
            
            print(f"{'Size Range':<20} {'Files':<8} {'Total Size':<15} {'Percentage':<12}")
            print("-" * 60)
            
            total_files = sum(count for _, count, _ in results)
            total_size = sum(size for _, _, size in results)
            
            for category, count, size_bytes in results:
                size_str = self.format_size(size_bytes)
                file_percent = (count / total_files * 100) if total_files > 0 else 0
                size_percent = (size_bytes / total_size * 100) if total_size > 0 else 0
                
                print(f"{category:<20} {count:<8} {size_str:<15} {file_percent:>5.1f}% files")
            
            print("-" * 60)
            print(f"{'TOTAL':<20} {total_files:<8} {self.format_size(total_size):<15}")
            
        except Exception as e:
            print(f"❌ Error analyzing size distribution: {e}")
    
    def show_recent_activity(self, days: int = 7):
        """Show recent file activity"""
        print("=" * 80)
        print(f"🕒 RECENT ACTIVITY (Last {days} Days)")
        print("=" * 80)
        
        try:
            cursor = self.db_manager.conn.cursor()
            
            cursor.execute("""
                SELECT 
                    filename,
                    LENGTH(file_content) as size_bytes,
                    created_at
                FROM files 
                WHERE created_at > datetime('now', '-{} days')
                ORDER BY created_at DESC
            """.format(days))
            
            recent_files = cursor.fetchall()
            
            if not recent_files:
                print(f"📭 No files added in the last {days} days")
                return
            
            print(f"Found {len(recent_files)} recent files:")
            print()
            print(f"{'Filename':<50} {'Size':<12} {'Added':<20}")
            print("-" * 85)
            
            for filename, size_bytes, created_at in recent_files:
                size_str = self.format_size(size_bytes)
                display_filename = filename[:47] + "..." if len(filename) > 50 else filename
                
                try:
                    if isinstance(created_at, str):
                        date_obj = datetime.fromisoformat(created_at.replace('Z', '+00:00'))
                    else:
                        date_obj = created_at
                    date_str = date_obj.strftime("%Y-%m-%d %H:%M")
                except:
                    date_str = str(created_at)[:16] if created_at else "Unknown"
                
                print(f"{display_filename:<50} {size_str:<12} {date_str:<20}")
            
        except Exception as e:
            print(f"❌ Error showing recent activity: {e}")
    
    def find_largest_files(self, top_n: int = 10):
        """Find the largest files in the database"""
        print("=" * 80)
        print(f"🔍 TOP {top_n} LARGEST FILES")
        print("=" * 80)
        
        try:
            cursor = self.db_manager.conn.cursor()
            
            cursor.execute("""
                SELECT filename, LENGTH(file_content) as size_bytes, created_at
                FROM files 
                ORDER BY size_bytes DESC
                LIMIT ?
            """, (top_n,))
            
            largest_files = cursor.fetchall()
            
            if not largest_files:
                print("📭 No files found")
                return
            
            for i, (filename, size_bytes, created_at) in enumerate(largest_files, 1):
                size_str = self.format_size(size_bytes)
                try:
                    if isinstance(created_at, str):
                        date_obj = datetime.fromisoformat(created_at.replace('Z', '+00:00'))
                    else:
                        date_obj = created_at
                    date_str = date_obj.strftime("%Y-%m-%d")
                except:
                    date_str = "Unknown"
                
                print(f"{i:>2}. {filename} ({size_str}) - {date_str}")
            
        except Exception as e:
            print(f"❌ Error finding largest files: {e}")
    
    def search_files(self, search_term: str):
        """Search for files by name"""
        print("=" * 80)
        print(f"🔍 SEARCH RESULTS FOR: '{search_term}'")
        print("=" * 80)
        
        try:
            cursor = self.db_manager.conn.cursor()
            
            cursor.execute("""
                SELECT 
                    id, filename, LENGTH(file_content) as size_bytes, created_at
                FROM files 
                WHERE filename LIKE ?
                ORDER BY created_at DESC
            """, (f"%{search_term}%",))
            
            results = cursor.fetchall()
            
            if not results:
                print(f"📭 No files found matching '{search_term}'")
                return
            
            print(f"Found {len(results)} matching files:")
            print()
            print(f"{'ID':<4} {'Filename':<50} {'Size':<12} {'Created':<20}")
            print("-" * 90)
            
            for file_id, filename, size_bytes, created_at in results:
                size_str = self.format_size(size_bytes)
                display_filename = filename[:47] + "..." if len(filename) > 50 else filename
                
                try:
                    if isinstance(created_at, str):
                        date_obj = datetime.fromisoformat(created_at.replace('Z', '+00:00'))
                    else:
                        date_obj = created_at
                    date_str = date_obj.strftime("%Y-%m-%d %H:%M")
                except:
                    date_str = str(created_at)[:16] if created_at else "Unknown"
                
                print(f"{file_id:<4} {display_filename:<50} {size_str:<12} {date_str:<20}")
            
        except Exception as e:
            print(f"❌ Error searching files: {e}")
    
    def show_database_path(self):
        """Show current database path and information"""
        print("=" * 80)
        print("🗂️  DATABASE INFORMATION")
        print("=" * 80)
        
        try:
            # Get database path
            db_path = getattr(self.db_manager, 'db_path', 'Unknown')
            print(f"Database Path: {db_path}")
            
            # Check if file exists and get size
            if os.path.exists(db_path):
                db_file_size = os.path.getsize(db_path)
                print(f"Database File Size: {self.format_size(db_file_size)}")
                print(f"Database File Exists: ✅ Yes")
                
                # Get file modification time
                mod_time = datetime.fromtimestamp(os.path.getmtime(db_path))
                print(f"Last Modified: {mod_time.strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                print(f"Database File Exists: ❌ No")
            
            # Get SQLite version
            cursor = self.db_manager.conn.cursor()
            cursor.execute("SELECT sqlite_version()")
            sqlite_version = cursor.fetchone()[0]
            print(f"SQLite Version: {sqlite_version}")
            
            # Get table information
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [row[0] for row in cursor.fetchall()]
            print(f"Tables: {', '.join(tables)}")
            
        except Exception as e:
            print(f"❌ Error getting database information: {e}")
    
    def close(self):
        """Close database connection"""
        if self.db_manager:
            self.db_manager.close()


def show_menu():
    """Display the main menu"""
    print("\n" + "=" * 80)
    print("🔍 TESTINGGUI DATABASE EXPLORER")
    print("=" * 80)
    print("1.  📊 Database Overview")
    print("2.  📄 File List (by size)")
    print("3.  📄 File List (by date)")
    print("4.  📄 File List (by name)")
    print("5.  📊 Size Distribution")
    print("6.  🕒 Recent Activity (7 days)")
    print("7.  🔍 Top 10 Largest Files")
    print("8.  🔍 Search Files")
    print("9.  🗂️  Database Info")
    print("10. 🚀 Full Report (All Above)")
    print("0.  ❌ Exit")
    print("=" * 80)


def main():
    """Main application loop"""
    print("🔍 TestingGUI Database Explorer")
    print("Connecting to database...")
    
    explorer = None
    try:
        explorer = DatabaseExplorer()
        
        while True:
            show_menu()
            choice = input("\nEnter your choice (0-10): ").strip()
            
            if choice == "0":
                print("\n👋 Goodbye!")
                break
            elif choice == "1":
                explorer.show_overview()
            elif choice == "2":
                explorer.show_file_list(sort_by="size")
            elif choice == "3":
                explorer.show_file_list(sort_by="date")
            elif choice == "4":
                explorer.show_file_list(sort_by="name")
            elif choice == "5":
                explorer.show_size_distribution()
            elif choice == "6":
                explorer.show_recent_activity()
            elif choice == "7":
                explorer.find_largest_files()
            elif choice == "8":
                search_term = input("Enter search term: ").strip()
                if search_term:
                    explorer.search_files(search_term)
                else:
                    print("❌ Please enter a search term")
            elif choice == "9":
                explorer.show_database_path()
            elif choice == "10":
                print("\n🚀 GENERATING FULL REPORT...")
                explorer.show_overview()
                print("\n")
                explorer.show_file_list(limit=20, sort_by="size")
                print("\n")
                explorer.show_size_distribution()
                print("\n")
                explorer.show_recent_activity()
                print("\n")
                explorer.find_largest_files()
                print("\n")
                explorer.show_database_path()
                print("\n✅ Full report complete!")
            else:
                print("❌ Invalid choice. Please enter 0-10.")
            
            if choice != "0":
                input("\nPress Enter to continue...")
    
    except KeyboardInterrupt:
        print("\n\n👋 Interrupted by user. Goodbye!")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
    finally:
        if explorer:
            explorer.close()


if __name__ == "__main__":
    main()
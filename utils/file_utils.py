"""
File utilities for the H2O data visualization tool
"""

import os
import shutil
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox


class FileUtils:
    """Utility class for file operations"""
    
    @staticmethod
    def get_file_info(file_path):
        """Get file information"""
        if not os.path.exists(file_path):
            return None
            
        stat = os.stat(file_path)
        
        return {
            'name': os.path.basename(file_path),
            'path': file_path,
            'size': stat.st_size,
            'size_mb': stat.st_size / (1024 * 1024),
            'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
            'created': datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
        }
        
    @staticmethod
    def validate_excel_file(file_path):
        """Validate if file is a valid Excel file"""
        if not os.path.exists(file_path):
            return False, "File does not exist"
            
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            return False, "File is not an Excel file"
            
        try:
            import pandas as pd
            pd.read_excel(file_path, nrows=1)  # Try to read first row
            return True, "Valid Excel file"
        except Exception as e:
            return False, f"Invalid Excel file: {str(e)}"
            
    @staticmethod
    def backup_file(file_path, backup_dir=None):
        """Create backup of file"""
        try:
            if not os.path.exists(file_path):
                return False, "Source file does not exist"
                
            if backup_dir is None:
                backup_dir = os.path.dirname(file_path)
                
            os.makedirs(backup_dir, exist_ok=True)
            
            # Create backup filename with timestamp
            basename = os.path.basename(file_path)
            name, ext = os.path.splitext(basename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_name = f"{name}_backup_{timestamp}{ext}"
            backup_path = os.path.join(backup_dir, backup_name)
            
            # Copy file
            shutil.copy2(file_path, backup_path)
            
            return True, backup_path
            
        except Exception as e:
            return False, f"Backup failed: {str(e)}"
            
    @staticmethod
    def ensure_directory(directory):
        """Ensure directory exists"""
        try:
            os.makedirs(directory, exist_ok=True)
            return True
        except Exception as e:
            print(f"Failed to create directory {directory}: {str(e)}")
            return False
            
    @staticmethod
    def get_safe_filename(filename):
        """Get safe filename by removing invalid characters"""
        # Remove or replace invalid characters
        invalid_chars = '<>:"/\\|?*'
        safe_name = filename
        
        for char in invalid_chars:
            safe_name = safe_name.replace(char, '_')
            
        # Remove multiple underscores
        while '__' in safe_name:
            safe_name = safe_name.replace('__', '_')
            
        # Remove leading/trailing underscores and spaces
        safe_name = safe_name.strip('_ ')
        
        return safe_name
        
    @staticmethod
    def format_file_size(size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
            
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024
            i += 1
            
        return f"{size_bytes:.1f} {size_names[i]}"
        
    @staticmethod
    def check_disk_space(path, required_mb=100):
        """Check if there's enough disk space"""
        try:
            stat = shutil.disk_usage(path)
            free_mb = stat.free / (1024 * 1024)
            return free_mb >= required_mb, free_mb
        except Exception as e:
            print(f"Failed to check disk space: {str(e)}")
            return True, 0  # Assume OK if can't check
            
    @staticmethod
    def clean_temp_files(temp_dir, max_age_hours=24):
        """Clean temporary files older than specified hours"""
        try:
            if not os.path.exists(temp_dir):
                return 0
                
            current_time = datetime.now().timestamp()
            max_age_seconds = max_age_hours * 3600
            cleaned_count = 0
            
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getmtime(file_path)
                    
                    if file_age > max_age_seconds:
                        try:
                            os.remove(file_path)
                            cleaned_count += 1
                        except:
                            pass  # Skip files that can't be deleted
                            
            return cleaned_count
            
        except Exception as e:
            print(f"Failed to clean temp files: {str(e)}")
            return 0
            
    @staticmethod
    def show_file_error(parent, error_message, file_path=None):
        """Show standardized file error message"""
        title = "File Error"
        message = error_message
        
        if file_path:
            message += f"\n\nFile: {file_path}"
            
        QMessageBox.critical(parent, title, message)
        
    @staticmethod
    def show_file_success(parent, success_message, file_path=None):
        """Show standardized file success message"""
        title = "Success"
        message = success_message
        
        if file_path:
            message += f"\n\nFile: {file_path}"
            
        QMessageBox.information(parent, title, message) 
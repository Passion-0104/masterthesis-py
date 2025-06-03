"""
Data loader module for loading and preprocessing Excel files
"""

import pandas as pd
import os
from PyQt5.QtWidgets import QMessageBox


class DataLoader:
    """Class for loading and preprocessing data from Excel files"""
    
    def __init__(self):
        self.data = None
        self.file_path = None
        
    def load_file(self, file_path):
        """Load Excel file and preprocess data"""
        try:
            # Read Excel file
            self.data = pd.read_excel(file_path)
            self.file_path = file_path
            
            # Preprocess data
            self._preprocess_data()
            
            return self.data
            
        except Exception as e:
            raise Exception(f"Failed to load file: {str(e)}")
            
    def _preprocess_data(self):
        """Preprocess loaded data"""
        if self.data is None:
            return
            
        # Convert all non-time columns to numeric
        for col in self.data.columns:
            try:
                self.data[col] = pd.to_numeric(self.data[col], errors='coerce')
            except:
                pass  # Keep non-numeric columns as is
                
    def get_columns(self):
        """Get list of column names"""
        if self.data is None:
            return []
        return list(self.data.columns)
        
    def get_time_columns(self):
        """Identify potential time columns"""
        if self.data is None:
            return []
            
        time_columns = []
        for col in self.data.columns:
            # Check if column contains datetime-like data
            try:
                pd.to_datetime(self.data[col].dropna().head(), errors='raise')
                time_columns.append(col)
            except:
                continue
                
        return time_columns
        
    def auto_match_pairs(self):
        """Automatically match moisture-pressure column pairs"""
        if self.data is None:
            return []
            
        columns = self.get_columns()
        
        # Keywords for identifying moisture and pressure columns
        moisture_keywords = ['h2o', 'water', 'humid', 'moisture', 'moistre', 'ppm', 'ppb']
        pressure_keywords = ['pressure', 'press', 'bar', 'pa', 'mpa']
        
        moisture_columns = []
        pressure_columns = []
        
        for col in columns:
            col_lower = str(col).lower()
            
            # Score based on keyword matches
            moisture_score = sum(1 for kw in moisture_keywords if kw in col_lower)
            pressure_score = sum(1 for kw in pressure_keywords if kw in col_lower)
            
            if moisture_score > pressure_score:
                moisture_columns.append(col)
            elif pressure_score > moisture_score:
                pressure_columns.append(col)
                
        # Match pairs based on numerical suffixes or order
        pairs = []
        
        # Try to match by numerical suffixes
        for moisture_col in moisture_columns[:]:
            for pressure_col in pressure_columns[:]:
                # Extract digits from column names
                moisture_digits = ''.join([c for c in str(moisture_col) if c.isdigit()])
                pressure_digits = ''.join([c for c in str(pressure_col) if c.isdigit()])
                
                if moisture_digits and moisture_digits == pressure_digits:
                    pairs.append((moisture_col, pressure_col))
                    moisture_columns.remove(moisture_col)
                    pressure_columns.remove(pressure_col)
                    break
                    
        # Match remaining columns in order
        remaining_pairs = min(len(moisture_columns), len(pressure_columns))
        for i in range(remaining_pairs):
            pairs.append((moisture_columns[i], pressure_columns[i]))
            
        return pairs
        
    def prepare_plot_data(self, time_col, selected_columns, time_range, start_time=0, end_time=2):
        """Prepare data for plotting"""
        if self.data is None:
            return None
            
        # Create working copy
        plot_df = self.data.copy()
        
        # Process time column
        if time_col and time_col in plot_df.columns:
            plot_df[time_col] = pd.to_datetime(plot_df[time_col], errors='coerce')
            plot_df = plot_df.dropna(subset=[time_col])
            
            # Calculate relative time in hours
            if not plot_df.empty:
                start_datetime = plot_df[time_col].min()
                plot_df['relative_time'] = (plot_df[time_col] - start_datetime).dt.total_seconds() / 3600
            else:
                plot_df['relative_time'] = 0
        else:
            # Create dummy time column
            plot_df['relative_time'] = range(len(plot_df))
            
        # Convert selected columns to numeric
        for col in selected_columns:
            if col in plot_df.columns:
                plot_df[col] = pd.to_numeric(plot_df[col], errors='coerce')
                
        # Apply time range filtering
        if time_range == 1:  # First 2 hours
            plot_df = plot_df[plot_df['relative_time'] <= 2]
        elif time_range == 2:  # Custom range
            plot_df = plot_df[(plot_df['relative_time'] >= start_time) & 
                            (plot_df['relative_time'] <= end_time)]
            
        return plot_df
        
    def export_data(self, data, parent_widget):
        """Export data to Excel file"""
        from PyQt5.QtWidgets import QFileDialog
        
        file_path, _ = QFileDialog.getSaveFileName(
            parent_widget, "Export Data", "", "Excel files (*.xlsx)"
        )
        
        if file_path:
            try:
                data.to_excel(file_path, index=False)
                QMessageBox.information(parent_widget, "Success", f"Data exported to {file_path}")
            except Exception as e:
                QMessageBox.critical(parent_widget, "Error", f"Export failed: {str(e)}")
                
    def is_data_loaded(self):
        """Check if data is loaded"""
        return self.data is not None
        
    def get_file_path(self):
        """Get current file path"""
        return self.file_path 
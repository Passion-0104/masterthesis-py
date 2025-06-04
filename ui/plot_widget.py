"""
Plot widget for displaying data visualization in independent windows
"""

import numpy as np
import matplotlib.pyplot as plt
try:
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
except ImportError:
    # Fallback for older versions
    try:
        from matplotlib.backends.backend_qt5agg import FigureCanvasQT5Agg as FigureCanvas
        from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
    except ImportError:
        # Last resort fallback
        from matplotlib.backends.backend_qt5agg import FigureCanvas
        from matplotlib.backends.backend_qt5agg import NavigationToolbar
        
from matplotlib.figure import Figure
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                           QMainWindow, QApplication, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import pandas as pd


class IndependentPlotWindow(QMainWindow):
    """Independent window for displaying charts"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the plot window UI"""
        self.setWindowTitle("Data Visualization Chart")
        self.setGeometry(200, 200, 1000, 700)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create layout
        layout = QVBoxLayout(central_widget)
        
        # Create matplotlib figure and canvas
        self.figure = Figure(figsize=(12, 8))
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        
        # Add widgets to layout
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        
        # Create button layout
        button_layout = QHBoxLayout()
        
        # Save chart button
        self.save_button = QPushButton("Save Chart")
        self.save_button.clicked.connect(self.save_chart)
        button_layout.addWidget(self.save_button)
        
        # Copy to clipboard button
        self.copy_button = QPushButton("Copy to Clipboard")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        button_layout.addWidget(self.copy_button)
        
        button_layout.addStretch()
        
        # Close button
        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.close)
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)
        
    def plot_data(self, data, selected_columns, plot_settings):
        """Plot data in the independent window"""
        try:
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            # Configure matplotlib for academic paper standards
            plt.rcParams['font.family'] = ['Arial', 'DejaVu Sans', 'Liberation Sans', 'Helvetica']
            plt.rcParams['mathtext.fontset'] = 'dejavusans'
            plt.rcParams['axes.unicode_minus'] = False
            plt.rcParams['font.size'] = 10  # Academic standard base font size
            
            # Get time column
            time_col = plot_settings.get('time_column', 'time')
            
            # Apply time range filtering
            plot_df = self._apply_time_filter(data, plot_settings)
            
            # Get calibration settings
            enable_calibration = plot_settings.get('enable_calibration', False)
            show_original = plot_settings.get('show_original', True)
            show_error = plot_settings.get('show_error', False)
            error_value = plot_settings.get('error_value', 10.0)
            
            # Get calibration data if enabled
            calibrated_data = {}
            if enable_calibration:
                calibrated_data = self._calculate_calibration(plot_df, plot_settings)
            
            # Plot data
            colors = plt.cm.tab10.colors
            calib_colors = ['#ff7f0e', '#d62728', '#9467bd', '#8c564b', 
                          '#e377c2', '#bcbd22', '#17becf', '#f0027f']
            
            for i, col in enumerate(selected_columns):
                if col not in plot_df.columns:
                    continue
                    
                color_idx = i % len(colors)
                
                # Check if this is a moisture column with calibration
                is_moisture = self._is_moisture_column(col, plot_settings)
                
                if is_moisture and enable_calibration:
                    # Plot original data if requested
                    if show_original:
                        valid_data = plot_df[['relative_time', col]].dropna()
                        if not valid_data.empty:
                            ax.plot(valid_data['relative_time'], valid_data[col],
                                  '-', color=colors[color_idx],
                                  label=f"{col} (original)",
                                  linewidth=2.0, alpha=0.7)
                    
                    # Plot calibrated data
                    calib_col = f"{col}_calib"
                    if calib_col in calibrated_data:
                        times = calibrated_data[calib_col]['times']
                        values = calibrated_data[calib_col]['values']
                        
                        calib_color = calib_colors[color_idx % len(calib_colors)]
                        ax.plot(times, values, '--', color=calib_color,
                              label=f"{col} (calibrated)",
                              linewidth=2.5, marker='o', markersize=2,
                              markevery=len(times)//20)
                        
                        # Add error range if requested
                        if show_error:
                            upper_bound = values + error_value
                            lower_bound = values - error_value
                            ax.fill_between(times, lower_bound, upper_bound,
                                          color=calib_color, alpha=0.2,
                                          label=f"Error Range (±{error_value:.1f} ppm)")
                else:
                    # Plot regular data
                    valid_data = plot_df[['relative_time', col]].dropna()
                    if not valid_data.empty:
                        ax.plot(valid_data['relative_time'], valid_data[col],
                              '-', color=colors[color_idx], label=col,
                              linewidth=2.0)
            
            # Set labels and title - academic paper standards
            ax.set_xlabel("Time (hours)", fontsize=11, fontweight='normal')
            ax.set_ylabel("Water Concentration (ppm)", fontsize=11, fontweight='normal')
            
            # Set up legend - smaller for academic papers
            legend = ax.legend(loc='best', frameon=True, fontsize=9, markerscale=1.5)
            for line in legend.get_lines():
                line.set_linewidth(2.0)
            
            # Set grid
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # Set tick parameters - academic standard
            ax.tick_params(axis='both', which='major', labelsize=10)
            
            # Add time markers if showing specific range
            time_range = plot_settings.get('time_range', 0)  # 0=all, 1=2hours, 2=custom
            if time_range == 1:  # 2 hours
                time_ticks = [0, 0.5, 1, 1.5, 2]
                time_labels = ['0', '30 min', '1 hour', '1.5 hours', '2 hours']
                ax.set_xticks(time_ticks)
                ax.set_xticklabels(time_labels, fontsize=10)
            
            # Mark 30min line if difference calculation is enabled
            if plot_settings.get('enable_30min_diff', False):
                ax.axvline(x=0.5, color='gray', linestyle='--', alpha=0.7, linewidth=1.5)
                ax.text(0.5, ax.get_ylim()[0], '30min', 
                      horizontalalignment='center', verticalalignment='bottom',
                      fontsize=10, fontweight='normal')
            
            plt.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            QMessageBox.critical(self, "Plot Error", f"Error creating plot: {str(e)}")
            
    def _apply_time_filter(self, data, plot_settings):
        """Apply time range filtering to data"""
        plot_df = data.copy()
        
        # Convert time column to relative time if needed
        time_col = plot_settings.get('time_column')
        if time_col and time_col in plot_df.columns:
            plot_df[time_col] = pd.to_datetime(plot_df[time_col], errors='coerce')
            plot_df = plot_df.dropna(subset=[time_col])
            
            # Calculate relative time
            if not plot_df.empty:
                start_time = plot_df[time_col].min()
                plot_df['relative_time'] = (plot_df[time_col] - start_time).dt.total_seconds() / 3600
            else:
                plot_df['relative_time'] = 0
        else:
            # Create dummy relative time if no time column
            plot_df['relative_time'] = range(len(plot_df))
            
        # Apply time range filter
        time_range = plot_settings.get('time_range', 0)
        if time_range == 1:  # First 2 hours
            plot_df = plot_df[plot_df['relative_time'] <= 2]
        elif time_range == 2:  # Custom range
            start_time = plot_settings.get('start_time', 0)
            end_time = plot_settings.get('end_time', 2)
            plot_df = plot_df[(plot_df['relative_time'] >= start_time) & 
                            (plot_df['relative_time'] <= end_time)]
        
        return plot_df
        
    def _is_moisture_column(self, col, plot_settings):
        """Check if a column is configured as a moisture column"""
        pairs = plot_settings.get('moisture_pressure_pairs', [])
        for moisture_col, pressure_col in pairs:
            if col == moisture_col:
                return True
        return False
        
    def _calculate_calibration(self, plot_df, plot_settings):
        """Calculate calibration data"""
        calibrated_data = {}
        
        pairs = plot_settings.get('moisture_pressure_pairs', [])
        if not pairs:
            return calibrated_data
            
        f1 = plot_settings.get('f1', 0.196798)
        f2 = plot_settings.get('f2', 0.419073)
        p_ref = plot_settings.get('p_ref', 1.0)
        
        for moisture_col, pressure_col in pairs:
            if moisture_col in plot_df.columns and pressure_col in plot_df.columns:
                # Ensure numeric data
                plot_df[moisture_col] = pd.to_numeric(plot_df[moisture_col], errors='coerce')
                plot_df[pressure_col] = pd.to_numeric(plot_df[pressure_col], errors='coerce')
                
                # Remove invalid data
                valid_mask = ~(pd.isna(plot_df[moisture_col]) | 
                             pd.isna(plot_df[pressure_col]) | 
                             (plot_df[pressure_col] <= 0))
                
                valid_df = plot_df[valid_mask].copy()
                
                if not valid_df.empty:
                    # Apply calibration formula
                    ratio = p_ref / valid_df[pressure_col]
                    exponent = f1 * np.log(ratio) + f2
                    exponent = np.clip(exponent, -10, 10)  # Prevent numerical explosion
                    
                    calibrated_values = valid_df[moisture_col] * (ratio ** exponent)
                    
                    calib_col = f"{moisture_col}_calib"
                    calibrated_data[calib_col] = {
                        'times': valid_df['relative_time'].values,
                        'values': calibrated_values.values,
                        'column': moisture_col,
                        'valid_df': valid_df
                    }
        
        return calibrated_data
        
    def save_chart(self):
        """Save chart to file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Chart", "", 
            "PNG files (*.png);;PDF files (*.pdf);;SVG files (*.svg)"
        )
        
        if file_path:
            try:
                self.figure.savefig(file_path, dpi=300, bbox_inches='tight')
                QMessageBox.information(self, "Success", f"Chart saved to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save chart: {str(e)}")
                
    def copy_to_clipboard(self):
        """Copy chart to clipboard"""
        try:
            # Get the clipboard
            clipboard = QApplication.clipboard()
            
            # Convert figure to image and copy to clipboard
            import io
            buf = io.BytesIO()
            self.figure.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            buf.seek(0)
            
            from PyQt5.QtGui import QPixmap
            pixmap = QPixmap()
            pixmap.loadFromData(buf.getvalue())
            clipboard.setPixmap(pixmap)
            
            QMessageBox.information(self, "Success", "Chart copied to clipboard")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to copy chart: {str(e)}")


class PlotWidget(QWidget):
    """Legacy plot widget for compatibility"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the plot widget UI"""
        layout = QVBoxLayout(self)
        
        # Create matplotlib figure and canvas
        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        
        # Add widgets to layout
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        
    def plot_data(self, data, columns):
        """Plot data in the widget"""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        for col in columns:
            if col in data.columns:
                ax.plot(data.index, data[col], label=col)
                
        ax.legend()
        ax.grid(True)
        self.canvas.draw() 
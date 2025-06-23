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
from data_processing.slope_calculator import SlopeCalculator


class IndependentPlotWindow(QMainWindow):
    """Independent window for displaying charts"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.slope_calculator = SlopeCalculator()
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
            
            # Mark 20min interval lines if 20min interval difference calculation is enabled
            if plot_settings.get('enable_20min_interval_diff', False):
                # Calculate time range for marking
                time_range = plot_settings.get('time_range', 0)
                if time_range == 1:  # 2 hours
                    max_time = 2.0
                elif time_range == 2:  # Custom range
                    max_time = plot_settings.get('end_time', 2.0)
                else:  # All data
                    max_time = plot_df['relative_time'].max() if not plot_df.empty else 2.0
                
                # Mark every 20 minutes (1/3 hour)
                interval = 1/3  # 20 minutes in hours
                for i in range(1, int(max_time / interval) + 1):
                    time_point = i * interval
                    if time_point <= max_time:
                        ax.axvline(x=time_point, color='lightgray', linestyle=':', alpha=0.5, linewidth=1.0)
                        ax.text(time_point, ax.get_ylim()[0], f'{int(time_point*60)}min', 
                              horizontalalignment='center', verticalalignment='bottom',
                              fontsize=8, fontweight='normal', alpha=0.7)
            
            # Mark 20/40/60min lines if multi-time difference calculation is enabled
            if plot_settings.get('enable_multi_time_diff', False):
                time_points = [20/60, 40/60, 60/60]  # 20min, 40min, 60min in hours
                for time_point in time_points:
                    ax.axvline(x=time_point, color='orange', linestyle='--', alpha=0.7, linewidth=1.5)
                    ax.text(time_point, ax.get_ylim()[0], f'{int(time_point*60)}min', 
                          horizontalalignment='center', verticalalignment='bottom',
                          fontsize=10, fontweight='normal', color='orange')
            
            plt.tight_layout()
            self.canvas.draw()
            
            # Add selected columns to plot settings for difference calculation
            plot_settings['selected_columns'] = selected_columns
            
            # Calculate and display slope results if enabled
            self._calculate_and_display_slopes(plot_df, plot_settings)
            
            # Calculate and display difference results if enabled
            self._calculate_and_display_differences(plot_df, calibrated_data, plot_settings)
            
        except Exception as e:
            QMessageBox.critical(self, "Plot Error", f"Error creating plot: {str(e)}")
            
    def _apply_time_filter(self, data, plot_settings):
        """Apply time range filtering to data"""
        plot_df = data.copy()
        
        # Check if relative_time already exists and is valid
        if 'relative_time' in plot_df.columns:
            relative_time_col = plot_df['relative_time']
            # Check if it's numeric and has reasonable values
            if pd.api.types.is_numeric_dtype(relative_time_col) and not relative_time_col.isna().all():
                max_time = relative_time_col.max()
                min_time = relative_time_col.min()
                # If it looks like valid relative time (hours), use it
                if max_time > min_time and max_time <= 24:  # reasonable hour range
                    print(f"Debug: Using existing relative_time column, range: {min_time:.3f} - {max_time:.3f}h")
                    # Apply time range filter and return
                    time_range = plot_settings.get('time_range', 0)
                    if time_range == 1:  # First 2 hours
                        plot_df = plot_df[plot_df['relative_time'] <= 2]
                    elif time_range == 2:  # Custom range
                        start_time = plot_settings.get('start_time', 0)
                        end_time = plot_settings.get('end_time', 2)
                        plot_df = plot_df[(plot_df['relative_time'] >= start_time) & 
                                        (plot_df['relative_time'] <= end_time)]
                    return plot_df
        
        # Convert time column to relative time if needed
        time_col = plot_settings.get('time_column')
        if time_col and time_col in plot_df.columns:
            print(f"Debug: Converting time column {time_col} to relative_time")
            plot_df[time_col] = pd.to_datetime(plot_df[time_col], errors='coerce')
            plot_df = plot_df.dropna(subset=[time_col])
            
            # Calculate relative time
            if not plot_df.empty:
                start_time = plot_df[time_col].min()
                plot_df['relative_time'] = (plot_df[time_col] - start_time).dt.total_seconds() / 3600
                print(f"Debug: Converted to relative_time, range: {plot_df['relative_time'].min():.3f} - {plot_df['relative_time'].max():.3f}h")
            else:
                plot_df['relative_time'] = 0
                print("Debug: Empty dataframe after time conversion")
        else:
            # Create dummy relative time if no time column
            plot_df['relative_time'] = np.arange(len(plot_df)) / 60.0  # Assume 1 minute intervals
            print(f"Debug: Created dummy relative_time, range: {plot_df['relative_time'].min():.3f} - {plot_df['relative_time'].max():.3f}h")
            
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
        
        print(f"Debug: Calibration parameters - f1: {f1}, f2: {f2}, p_ref: {p_ref}")
        print(f"Debug: Data shape: {plot_df.shape}, columns: {list(plot_df.columns)}")
        
        for moisture_col, pressure_col in pairs:
            if moisture_col in plot_df.columns and pressure_col in plot_df.columns:
                print(f"Debug: Calibrating pair {moisture_col} - {pressure_col}")
                
                # Ensure numeric data
                moisture_data = pd.to_numeric(plot_df[moisture_col], errors='coerce')
                pressure_data = pd.to_numeric(plot_df[pressure_col], errors='coerce')
                
                print(f"Debug: Original {moisture_col} range: {moisture_data.min():.6f} - {moisture_data.max():.6f}")
                print(f"Debug: Original {pressure_col} range: {pressure_data.min():.6f} - {pressure_data.max():.6f}")
                
                # Create a working dataframe
                work_df = plot_df[['relative_time']].copy()
                work_df[moisture_col] = moisture_data
                work_df[pressure_col] = pressure_data
                
                # Remove invalid data and extreme values - 适应小数数据
                valid_mask = ~(pd.isna(work_df[moisture_col]) | 
                             pd.isna(work_df[pressure_col]) | 
                             (work_df[pressure_col] <= 0) |
                             (work_df[pressure_col] > 2) |    # 适应更小的压力范围
                             (work_df[moisture_col] < 0) |    # 避免负湿度值
                             (work_df[moisture_col] > 100))   # 适应小数湿度值范围
                
                valid_df = work_df[valid_mask].copy()
                
                if not valid_df.empty:
                    print(f"Debug: Valid data points for {moisture_col}: {len(valid_df)} out of {len(work_df)}")
                    
                    # Apply calibration formula with careful validation
                    try:
                        ratio = p_ref / valid_df[pressure_col]
                        
                        # 检查比率的合理性，适应小数数据
                        ratio = ratio.clip(0.5, 20.0)  # 适应压力比率范围
                        
                        # Calculate exponent
                        log_ratio = np.log(ratio)
                        exponent = f1 * log_ratio + f2
                        
                        # 严格限制指数值
                        exponent = np.clip(exponent, -3, 3)  # 适中的限制
                        
                        # Calculate calibrated values
                        power_term = ratio ** exponent
                        
                        # 检查幂次项的合理性
                        power_term = np.clip(power_term, 0.1, 10.0)  # 适应小数数据
                        
                        calibrated_values = valid_df[moisture_col] * power_term
                        
                        # 最终检查校准后的值 - 适应小数数据
                        valid_calibrated_mask = (calibrated_values > 0) & (calibrated_values < 100) & np.isfinite(calibrated_values)
                        
                        if valid_calibrated_mask.any():
                            final_valid_df = valid_df[valid_calibrated_mask].copy()
                            final_calibrated_values = calibrated_values[valid_calibrated_mask]
                            
                            calib_col = f"{moisture_col}_calib"
                            calibrated_data[calib_col] = {
                                'times': final_valid_df['relative_time'].values,
                                'values': final_calibrated_values.values,
                                'column': moisture_col,
                                'valid_df': final_valid_df
                            }
                            
                            print(f"Debug: Calibrated {moisture_col}: {len(final_calibrated_values)} valid points")
                            print(f"Debug: Calibrated value range: {final_calibrated_values.min():.6f} - {final_calibrated_values.max():.6f}")
                            print(f"Debug: Time range for calibrated data: {final_valid_df['relative_time'].min():.3f} - {final_valid_df['relative_time'].max():.3f}h")
                        else:
                            print(f"Debug: No valid calibrated values for {moisture_col}")
                            
                    except Exception as e:
                        print(f"Debug: Error in calibration calculation for {moisture_col}: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        
                else:
                    print(f"Debug: No valid data for calibration pair {moisture_col} - {pressure_col}")
            else:
                missing_cols = []
                if moisture_col not in plot_df.columns:
                    missing_cols.append(moisture_col)
                if pressure_col not in plot_df.columns:
                    missing_cols.append(pressure_col)
                print(f"Debug: Missing columns for calibration: {missing_cols}")
        
        print(f"Debug: Total calibrated columns: {len(calibrated_data)}")
        for calib_col in calibrated_data.keys():
            data_len = len(calibrated_data[calib_col]['times'])
            time_range = f"{calibrated_data[calib_col]['times'].min():.3f} - {calibrated_data[calib_col]['times'].max():.3f}h"
            print(f"Debug: {calib_col}: {data_len} points, time range: {time_range}")
        
        return calibrated_data
        
    def _calculate_and_display_differences(self, plot_df, calibrated_data, plot_settings):
        """Calculate and display difference results"""
        try:
            # Check if any difference calculation is enabled
            enable_30min = plot_settings.get('enable_30min_diff', False)
            enable_multi = plot_settings.get('enable_multi_time_diff', False)
            enable_20min_interval = plot_settings.get('enable_20min_interval_diff', False)
            
            print(f"Debug: Difference calculation enabled - 30min: {enable_30min}, multi: {enable_multi}, 20min_interval: {enable_20min_interval}")
            
            if not any([enable_30min, enable_multi, enable_20min_interval]):
                print("Debug: No difference calculation enabled")
                return
                
            # Get reference columns
            ref_col1 = plot_settings.get('reference_column', '')
            ref_col2 = plot_settings.get('reference2_column', '')
            time_window = plot_settings.get('time_window', 5.0) / 60  # Convert to hours
            
            print(f"Debug: Reference columns - ref1: {ref_col1}, ref2: {ref_col2}")
            
            if not ref_col1 and not ref_col2:
                print("Debug: No reference columns selected")
                return
                
            # Get selected columns for comparison
            selected_columns = plot_settings.get('selected_columns', [])
            print(f"Debug: Selected columns: {selected_columns}")
            
            # Calculate differences based on enabled options
            results = []
            
            if enable_30min:
                if ref_col1:
                    result_30min = self._calculate_difference_at_time(plot_df, calibrated_data, ref_col1, 0.5, time_window, selected_columns)
                    if result_30min:
                        results.append(f"=== 30min Difference (Reference: {ref_col1}) ===\n{result_30min}")
                    else:
                        results.append(f"=== 30min Difference (Reference: {ref_col1}) ===\nNo data available for calculation")
                        
            if enable_multi:
                if ref_col1:
                    result_multi1 = self._calculate_multi_time_differences(plot_df, calibrated_data, ref_col1, time_window, selected_columns)
                    if result_multi1:
                        results.append(f"=== Multi-time Differences (Reference: {ref_col1}) ===\n{result_multi1}")
                    else:
                        results.append(f"=== Multi-time Differences (Reference: {ref_col1}) ===\nNo data available for calculation")
                if ref_col2 and ref_col2 != ref_col1:
                    result_multi2 = self._calculate_multi_time_differences(plot_df, calibrated_data, ref_col2, time_window, selected_columns)
                    if result_multi2:
                        results.append(f"=== Multi-time Differences (Reference: {ref_col2}) ===\n{result_multi2}")
                    else:
                        results.append(f"=== Multi-time Differences (Reference: {ref_col2}) ===\nNo data available for calculation")
                        
            if enable_20min_interval:
                if ref_col1:
                    result_interval1 = self._calculate_20min_interval_differences(plot_df, calibrated_data, ref_col1, time_window, plot_settings, selected_columns)
                    if result_interval1:
                        results.append(f"=== 20min Interval Differences (Reference: {ref_col1}) ===\n{result_interval1}")
                    else:
                        results.append(f"=== 20min Interval Differences (Reference: {ref_col1}) ===\nNo data available for calculation")
                if ref_col2 and ref_col2 != ref_col1:
                    result_interval2 = self._calculate_20min_interval_differences(plot_df, calibrated_data, ref_col2, time_window, plot_settings, selected_columns)
                    if result_interval2:
                        results.append(f"=== 20min Interval Differences (Reference: {ref_col2}) ===\n{result_interval2}")
                    else:
                        results.append(f"=== 20min Interval Differences (Reference: {ref_col2}) ===\nNo data available for calculation")
            
            # Display results if any
            if results:
                print(f"Debug: Showing difference results with {len(results)} sections")
                self._show_difference_results("\n\n".join(results))
            else:
                print("Debug: No results to display")
                
        except Exception as e:
            print(f"Debug: Exception in difference calculation: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.warning(self, "Difference Calculation Warning", f"Error calculating differences: {str(e)}")
            
    def _calculate_difference_at_time(self, plot_df, calibrated_data, ref_col, target_time, window, selected_columns):
        """Calculate difference at a specific time point"""
        try:
            print(f"Debug: Calculating difference at {target_time}h for reference {ref_col}")
            print(f"Debug: Available columns in plot_df: {list(plot_df.columns)}")
            print(f"Debug: Available calibrated data: {list(calibrated_data.keys())}")
            
            # 检查参考列是否是时间列，如果是则跳过
            if ref_col in ['Zeit', 'Time', 'time', 'zeit'] or 'time' in ref_col.lower():
                print(f"Debug: Reference column '{ref_col}' appears to be a time column, cannot use for difference calculation")
                return None
            
            # Get reference value (try calibrated first, then original)
            ref_value = None
            ref_col_calib = f"{ref_col}_calib"
            has_calibrated_ref = ref_col_calib in calibrated_data
            
            if has_calibrated_ref:
                ref_times = calibrated_data[ref_col_calib]['times']
                ref_values = calibrated_data[ref_col_calib]['values']
                print(f"Debug: Calibrated reference data range: {len(ref_times)} points, time range: {ref_times.min():.3f} - {ref_times.max():.3f}h")
                
                ref_indices = [i for i, t in enumerate(ref_times) if 
                             (t >= (target_time - window) and t <= (target_time + window))]
                if ref_indices:
                    ref_window_values = [ref_values[i] for i in ref_indices]
                    # 检查数值有效性，适应小数数据
                    valid_ref_values = [v for v in ref_window_values if np.isfinite(v) and abs(v) < 1e4]
                    if valid_ref_values:
                        ref_value = np.mean(valid_ref_values)
                        print(f"Debug: Using calibrated reference value: {ref_value:.6f} (from {len(valid_ref_values)} points)")
            
            if ref_value is None and ref_col in plot_df.columns:
                time_mask = (plot_df['relative_time'] >= (target_time - window)) & \
                           (plot_df['relative_time'] <= (target_time + window))
                ref_data = plot_df.loc[time_mask, ref_col]
                print(f"Debug: Original reference data in time window: {len(ref_data)} points")
                
                if not ref_data.empty:
                    # 确保数据是数值型且有效，适应小数数据
                    ref_data_numeric = pd.to_numeric(ref_data, errors='coerce')
                    valid_ref_data = ref_data_numeric.dropna()
                    # 调整过滤范围以适应小数数据
                    valid_ref_data = valid_ref_data[(valid_ref_data > -1e3) & (valid_ref_data < 1e3)]
                    if not valid_ref_data.empty:
                        ref_value = valid_ref_data.mean()
                        print(f"Debug: Using original reference value: {ref_value:.6f} (from {len(valid_ref_data)} points)")
                        print(f"Debug: Reference data range: {valid_ref_data.min():.6f} - {valid_ref_data.max():.6f}")
            
            if ref_value is None or not np.isfinite(ref_value):
                print(f"Debug: No valid reference value found for {ref_col} at {target_time}h")
                return None
            
            # Calculate differences for all selected columns
            results = []
            total_diff = 0
            count = 0
            
            for col in selected_columns:
                if col == ref_col:
                    continue
                    
                print(f"Debug: Processing column {col}")
                
                # Try calibrated data first
                calib_col = f"{col}_calib"
                if calib_col in calibrated_data:
                    times = calibrated_data[calib_col]['times']
                    values = calibrated_data[calib_col]['values']
                    column_name = f"{col} (calibrated)"
                    
                    print(f"Debug: Calibrated data for {col}: {len(times)} points, time range: {times.min():.3f} - {times.max():.3f}h")
                    
                    time_indices = [i for i, t in enumerate(times) if 
                                  (t >= (target_time - window) and t <= (target_time + window))]
                    
                    if time_indices:
                        window_values = [values[i] for i in time_indices]
                        print(f"Debug: Found {len(window_values)} calibrated values in time window for {col}")
                        
                        # 检查数值有效性，适应小数数据
                        valid_values = [v for v in window_values if np.isfinite(v) and abs(v) < 1e4]
                        if valid_values:
                            calib_value = np.mean(valid_values)
                            if np.isfinite(calib_value) and abs(calib_value) < 1e4:
                                difference = calib_value - ref_value
                                if np.isfinite(difference) and abs(difference) < 1e4:
                                    total_diff += difference
                                    count += 1
                                    
                                    results.append(f"{column_name}: {calib_value:.6f} ppm (diff: {difference:.6f} ppm)")
                                    print(f"Debug: Calibrated {col}: {calib_value:.6f} ppm (diff: {difference:.6f} ppm)")
                                else:
                                    print(f"Debug: Invalid difference for calibrated {col}: {difference}")
                            else:
                                print(f"Debug: Invalid calibrated value for {col}: {calib_value}")
                        else:
                            print(f"Debug: No valid calibrated values for {col}")
                    else:
                        print(f"Debug: No calibrated data points in time window for {col}")
                
                # Try original data if no calibrated data or if calibration is not enabled
                elif col in plot_df.columns:
                    time_mask = (plot_df['relative_time'] >= (target_time - window)) & \
                               (plot_df['relative_time'] <= (target_time + window))
                    col_data = plot_df.loc[time_mask, col]
                    
                    print(f"Debug: Original data for {col} in time window: {len(col_data)} points")
                    
                    if not col_data.empty:
                        # 确保数据是数值型且有效，适应小数数据
                        col_data_numeric = pd.to_numeric(col_data, errors='coerce')
                        valid_col_data = col_data_numeric.dropna()
                        # 调整过滤范围以适应小数数据
                        valid_col_data = valid_col_data[(valid_col_data > -1e3) & (valid_col_data < 1e3)]
                        
                        if not valid_col_data.empty:
                            orig_value = valid_col_data.mean()
                            print(f"Debug: Original {col} data range: {valid_col_data.min():.6f} - {valid_col_data.max():.6f}")
                            
                            if np.isfinite(orig_value) and abs(orig_value) < 1e3:
                                difference = orig_value - ref_value
                                if np.isfinite(difference) and abs(difference) < 1e3:
                                    total_diff += difference
                                    count += 1
                                    
                                    results.append(f"{col} (original): {orig_value:.6f} ppm (diff: {difference:.6f} ppm)")
                                    print(f"Debug: Original {col}: {orig_value:.6f} ppm (diff: {difference:.6f} ppm)")
                                else:
                                    print(f"Debug: Invalid difference for original {col}: {difference}")
                            else:
                                print(f"Debug: Invalid original value for {col}: {orig_value}")
                        else:
                            print(f"Debug: No valid original values for {col}")
                    else:
                        print(f"Debug: No original data points in time window for {col}")
                else:
                    print(f"Debug: Column {col} not found in data")
            
            if results:
                avg_diff = total_diff / count if count > 0 else 0
                if np.isfinite(avg_diff) and abs(avg_diff) < 1e3:
                    results.append(f"\nAverage Difference: {avg_diff:.6f} ppm")
                    results.append(f"Reference Value: {ref_value:.6f} ppm")
                    results.append(f"Time Window: ±{window*60:.1f} minutes around {target_time*60:.0f}min")
                    results.append(f"Data points processed: {count}")
                    print(f"Debug: Average difference: {avg_diff:.6f} ppm from {count} columns")
                    return "\n".join(results)
                else:
                    print(f"Debug: Invalid average difference: {avg_diff}")
                    return None
            else:
                print(f"Debug: No valid data found for difference calculation")
                return None
                
        except Exception as e:
            print(f"Debug: Exception in _calculate_difference_at_time: {str(e)}")
            import traceback
            traceback.print_exc()
            return f"Error calculating difference: {str(e)}"
            
    def _calculate_multi_time_differences(self, plot_df, calibrated_data, ref_col, window, selected_columns):
        """Calculate differences at 20, 40, 60 minutes"""
        try:
            target_times = [20/60, 40/60, 60/60]  # 20min, 40min, 60min
            results = []
            
            for target_time in target_times:
                result = self._calculate_difference_at_time(plot_df, calibrated_data, ref_col, target_time, window, selected_columns)
                if result:
                    time_label = f"{int(target_time*60)}min"
                    results.append(f"--- {time_label} ---\n{result}")
            
            return "\n\n".join(results) if results else None
            
        except Exception as e:
            return f"Error calculating multi-time differences: {str(e)}"
            
    def _calculate_20min_interval_differences(self, plot_df, calibrated_data, ref_col, window, plot_settings, selected_columns):
        """Calculate differences every 20 minutes"""
        try:
            print(f"Debug: Starting 20min interval calculation for ref_col: {ref_col}")
            print(f"Debug: plot_df shape: {plot_df.shape}")
            print(f"Debug: selected_columns: {selected_columns}")
            
            # Determine time range
            time_range = plot_settings.get('time_range', 0)
            if time_range == 1:  # 2 hours
                max_time = 2.0
            elif time_range == 2:  # Custom range
                max_time = plot_settings.get('end_time', 2.0)
            else:  # All data
                max_time = plot_df['relative_time'].max() if not plot_df.empty else 2.0
            
            print(f"Debug: Time range setting: {time_range}, max_time: {max_time:.3f}h")
            
            # Calculate intervals (every 20 minutes)
            interval = 1/3  # 20 minutes in hours
            target_times = []
            for i in range(1, int(max_time / interval) + 1):
                time_point = i * interval
                if time_point <= max_time:
                    target_times.append(time_point)
            
            print(f"Debug: Target times: {[f'{t:.3f}h ({t*60:.0f}min)' for t in target_times]}")
            
            results = []
            all_differences = []
            
            for target_time in target_times:
                print(f"Debug: Calculating for time point {target_time:.3f}h ({target_time*60:.0f}min)")
                result = self._calculate_difference_at_time(plot_df, calibrated_data, ref_col, target_time, window, selected_columns)
                
                if result:
                    time_label = f"{int(target_time*60)}min"
                    results.append(f"--- {time_label} ---\n{result}")
                    print(f"Debug: ✓ Got result for {time_label}")
                    
                    # Extract average difference for statistics
                    lines = result.split('\n')
                    for line in lines:
                        if "Average Difference:" in line:
                            try:
                                avg_diff = float(line.split(':')[1].strip().split()[0])
                                all_differences.append(avg_diff)
                                print(f"Debug: Extracted average difference: {avg_diff:.6f}")
                            except Exception as e:
                                print(f"Debug: Failed to extract average difference from '{line}': {e}")
                else:
                    print(f"Debug: ✗ No result for {target_time:.3f}h ({target_time*60:.0f}min)")
            
            print(f"Debug: Total results collected: {len(results)}")
            print(f"Debug: All differences: {all_differences}")
            
            # Add overall statistics
            if all_differences:
                overall_avg = np.mean(all_differences)
                overall_std = np.std(all_differences)
                results.append(f"\n=== Overall Statistics ===")
                results.append(f"Number of time points: {len(all_differences)}")
                results.append(f"Overall average difference: {overall_avg:.6f} ppm")
                results.append(f"Standard deviation: {overall_std:.6f} ppm")
                results.append(f"Min difference: {min(all_differences):.6f} ppm")
                results.append(f"Max difference: {max(all_differences):.6f} ppm")
                print(f"Debug: Added overall statistics for {len(all_differences)} points")
            
            final_result = "\n\n".join(results) if results else None
            print(f"Debug: Returning result: {final_result is not None}")
            if final_result:
                print(f"Debug: Result preview: {final_result[:200]}...")
            
            return final_result
            
        except Exception as e:
            print(f"Debug: Exception in _calculate_20min_interval_differences: {str(e)}")
            import traceback
            traceback.print_exc()
            return f"Error calculating 20min interval differences: {str(e)}"
            
    def _show_difference_results(self, results):
        """Show difference calculation results in a dialog"""
        try:
            print(f"Debug: Showing difference results dialog with {len(results)} characters")
            print(f"Debug: Results preview: {results[:200]}...")
            
            from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton, QHBoxLayout
            
            dialog = QDialog(self)
            dialog.setWindowTitle("Difference Calculation Results")
            dialog.setGeometry(200, 200, 600, 500)
            
            layout = QVBoxLayout(dialog)
            
            # Text area for results
            text_edit = QTextEdit()
            text_edit.setPlainText(results)
            text_edit.setReadOnly(True)
            layout.addWidget(text_edit)
            
            # Buttons
            button_layout = QHBoxLayout()
            
            save_btn = QPushButton("Save Results")
            save_btn.clicked.connect(lambda: self._save_difference_results(results))
            button_layout.addWidget(save_btn)
            
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(dialog.accept)
            button_layout.addWidget(close_btn)
            
            layout.addLayout(button_layout)
            
            print("Debug: About to show dialog")
            dialog.exec_()
            print("Debug: Dialog closed")
            
        except Exception as e:
            print(f"Debug: Exception in _show_difference_results: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"Failed to show difference results: {str(e)}")
        
    def _save_difference_results(self, results):
        """Save difference results to file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Difference Results", "", 
            "Text files (*.txt);;All files (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(results)
                QMessageBox.information(self, "Success", f"Results saved to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save results: {str(e)}")
        
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
    
    def _calculate_and_display_slopes(self, plot_df, plot_settings):
        """计算并显示斜率结果"""
        if not plot_settings.get('enable_slope_calc', False):
            return
        
        try:
            selected_columns = plot_settings.get('selected_columns', [])
            if not selected_columns:
                return
            
            # 获取斜率计算设置
            interval_minutes = plot_settings.get('slope_interval', 15.0)
            time_column = plot_settings.get('time_column', 'time')
            
            print(f"Debug: Starting interval-based slope calculation - interval: {interval_minutes} minutes")
            
            # 计算斜率
            slope_results = self.slope_calculator.calculate_slopes(
                plot_df, selected_columns, time_column, interval_minutes
            )
            
            if not slope_results:
                print("Debug: 没有计算出斜率结果")
                QMessageBox.warning(self, "斜率计算", "没有计算出斜率结果，请检查数据和设置")
                return
            
            print(f"Debug: 计算出 {len(slope_results)} 列的斜率")
            
            # 创建独立的斜率图表窗口
            self._create_slope_chart_window(slope_results, plot_settings)
            
            # 显示斜率统计信息
            self._show_slope_statistics(slope_results)
            
        except Exception as e:
            print(f"Debug: 斜率计算异常: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "斜率计算错误", f"计算斜率时出错: {str(e)}")
    
    def _create_slope_chart_window(self, slope_results, plot_settings):
        """创建斜率图表窗口"""
        try:
            from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout
            
            # 创建新窗口
            slope_window = QMainWindow(self)
            slope_window.setWindowTitle("Slope Chart")
            slope_window.setGeometry(300, 300, 1000, 600)
            
            # 创建中心部件
            central_widget = QWidget()
            slope_window.setCentralWidget(central_widget)
            layout = QVBoxLayout(central_widget)
            
            # 添加顶部工具栏
            toolbar_layout = QHBoxLayout()
            save_data_btn = QPushButton("Save Data")
            save_data_btn.setFont(QFont("Arial", 10))
            save_data_btn.clicked.connect(lambda: self._manual_save_slope_data(slope_results, plot_settings))
            
            toolbar_layout.addWidget(save_data_btn)
            toolbar_layout.addStretch()
            layout.addLayout(toolbar_layout)
            
            # 创建绘图
            figure = Figure(figsize=(12, 8))
            canvas = FigureCanvas(figure)
            toolbar = NavigationToolbar(canvas, slope_window)
            
            layout.addWidget(toolbar)
            layout.addWidget(canvas)
            
            # 绘制斜率图
            ax = figure.add_subplot(111)
            
            # 配置matplotlib
            plt.rcParams['font.family'] = ['Arial', 'DejaVu Sans', 'Liberation Sans', 'Helvetica']
            plt.rcParams['font.size'] = 10
            
            colors = plt.cm.tab10.colors
            
            for i, (col, data) in enumerate(slope_results.items()):
                color = colors[i % len(colors)]
                times = data['times']
                slopes = data['slopes']
                
                # Academic paper style: smooth line without markers
                ax.plot(times, slopes, '-', color=color, linewidth=1.5, alpha=0.9, 
                       label=f"{col} Slope")
                
                # Add zero reference line only once
                if i == 0:
                    ax.axhline(y=0, color='black', linestyle='-', alpha=0.3, linewidth=0.8)
            
            # Academic paper style labels and formatting
            ax.set_xlabel("Time (hours)", fontsize=12, fontfamily='serif')
            ax.set_ylabel(f"Slope ({slope_results[list(slope_results.keys())[0]]['units']})", fontsize=12, fontfamily='serif')
            
            # Clean title for academic papers
            interval_min = slope_results[list(slope_results.keys())[0]]['interval_minutes']
            ax.set_title(f"Slope vs Time ({interval_min}-minute intervals)", 
                        fontsize=14, fontfamily='serif', pad=20)
            
            # Academic style legend
            legend = ax.legend(loc='best', frameon=False, fontsize=10)
            for text in legend.get_texts():
                text.set_fontfamily('serif')
            
            # Clean grid for academic papers
            ax.grid(True, linestyle='-', alpha=0.2, linewidth=0.5)
            
            # Academic style ticks
            ax.tick_params(axis='both', which='major', labelsize=10, direction='out', length=4, width=0.8)
            ax.tick_params(axis='both', which='minor', labelsize=8, direction='out', length=2, width=0.5)
            
            # Remove top and right spines for cleaner look
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_linewidth(0.8)
            ax.spines['bottom'].set_linewidth(0.8)
            
            plt.tight_layout()
            canvas.draw()
            
            # 显示窗口
            slope_window.show()
            
            # 保存窗口引用（防止被垃圾回收）
            if not hasattr(self, 'slope_windows'):
                self.slope_windows = []
            self.slope_windows.append(slope_window)
            
        except Exception as e:
            print(f"Debug: 创建斜率图表窗口异常: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"创建斜率图表失败: {str(e)}")
    
    def _show_slope_statistics(self, slope_results):
        """显示斜率统计信息"""
        try:
            from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton, QHBoxLayout
            
            # 获取统计信息
            stats = self.slope_calculator.get_slope_statistics(slope_results)
            
            if not stats:
                return
            
            # 创建对话框
            dialog = QDialog(self)
            dialog.setWindowTitle("斜率统计信息")
            dialog.setGeometry(400, 400, 500, 400)
            
            layout = QVBoxLayout(dialog)
            
            # 构建统计信息文本
            stats_text = "=== 斜率统计信息 ===\n\n"
            
            for col, col_stats in stats.items():
                stats_text += f"--- {col} ---\n"
                stats_text += f"平均斜率: {col_stats['mean_slope']:.6f} ppm/小时\n"
                stats_text += f"标准差: {col_stats['std_slope']:.6f} ppm/小时\n"
                stats_text += f"最大斜率: {col_stats['max_slope']:.6f} ppm/小时\n"
                stats_text += f"最小斜率: {col_stats['min_slope']:.6f} ppm/小时\n"
                stats_text += f"中位数斜率: {col_stats['median_slope']:.6f} ppm/小时\n"

                stats_text += f"数据点数: {col_stats['num_points']}\n\n"
            
            # 文本显示区域
            text_edit = QTextEdit()
            text_edit.setPlainText(stats_text)
            text_edit.setReadOnly(True)
            layout.addWidget(text_edit)
            
            # 按钮
            button_layout = QHBoxLayout()
            
            # 保存按钮
            save_btn = QPushButton("保存统计信息")
            save_btn.clicked.connect(lambda: self._save_slope_statistics(stats_text))
            button_layout.addWidget(save_btn)
            
            # 导出数据按钮
            export_btn = QPushButton("导出斜率数据")
            export_btn.clicked.connect(lambda: self._export_slope_data(slope_results))
            button_layout.addWidget(export_btn)
            
            # 关闭按钮
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(dialog.accept)
            button_layout.addWidget(close_btn)
            
            layout.addLayout(button_layout)
            
            # 显示对话框
            dialog.exec_()
            
        except Exception as e:
            print(f"Debug: 显示斜率统计异常: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"显示斜率统计失败: {str(e)}")
    
    def _save_slope_statistics(self, stats_text):
        """保存斜率统计信息"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存斜率统计信息", "", 
            "文本文件 (*.txt);;所有文件 (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(stats_text)
                QMessageBox.information(self, "成功", f"统计信息已保存到 {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存失败: {str(e)}")
    
    def _export_slope_data(self, slope_results):
        """导出斜率数据"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出斜率数据", "", 
            "Excel文件 (*.xlsx);;CSV文件 (*.csv);;所有文件 (*)"
        )
        
        if file_path:
            try:
                # 导出为DataFrame
                slope_df = self.slope_calculator.export_slope_data(slope_results)
                
                if file_path.endswith('.xlsx'):
                    slope_df.to_excel(file_path, index=False)
                elif file_path.endswith('.csv'):
                    slope_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    # 默认保存为Excel
                    slope_df.to_excel(file_path, index=False)
                
                QMessageBox.information(self, "成功", f"斜率数据已导出到 {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
    
    def _manual_save_slope_data(self, slope_results, plot_settings):
        """Manual save slope data with user selection"""
        try:
            from datetime import datetime
            
            # Ask user to choose save location and filename
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Slope Calculation Results", 
                f"slope_calculation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "Excel files (*.xlsx);;CSV files (*.csv);;All files (*)"
            )
            
            if not file_path:
                return  # User cancelled
            
            # Export slope data
            slope_df = self.slope_calculator.export_slope_data(slope_results)
            
            if file_path.endswith('.xlsx'):
                slope_df.to_excel(file_path, index=False)
            elif file_path.endswith('.csv'):
                slope_df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                # Default to Excel
                slope_df.to_excel(file_path + '.xlsx', index=False)
                file_path += '.xlsx'
            
            print(f"Debug: Slope data saved to: {file_path}")
            
            # Also save detailed information
            base_path = file_path.rsplit('.', 1)[0]
            details_path = f"{base_path}_details.txt"
            
            stats = self.slope_calculator.get_slope_statistics(slope_results)
            
            with open(details_path, 'w', encoding='utf-8') as f:
                f.write("=== Slope Calculation Details ===\n\n")
                f.write(f"Calculation Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Time Interval: {plot_settings.get('slope_interval', 15.0)} minutes\n")
                f.write(f"Calculation Method: Interval-based slope calculation\n\n")
                
                f.write("=== Statistical Information ===\n\n")
                for col, col_stats in stats.items():
                    f.write(f"--- {col} ---\n")
                    f.write(f"Mean Slope: {col_stats['mean_slope']:.6f} ppm/hour\n")
                    f.write(f"Standard Deviation: {col_stats['std_slope']:.6f} ppm/hour\n")
                    f.write(f"Maximum Slope: {col_stats['max_slope']:.6f} ppm/hour\n")
                    f.write(f"Minimum Slope: {col_stats['min_slope']:.6f} ppm/hour\n")
                    f.write(f"Median Slope: {col_stats['median_slope']:.6f} ppm/hour\n")
                    f.write(f"Number of Points: {col_stats['num_points']}\n\n")
                
                # Add detailed calculation point information
                f.write("=== Detailed Calculation Points ===\n\n")
                for col, data in slope_results.items():
                    f.write(f"--- {col} ---\n")
                    if 'used_points' in data:
                        for i, point_info in enumerate(data['used_points'][:10]):  # Show first 10 points
                            f.write(f"Point {i+1}: Time={point_info['calc_time']:.3f}h, "
                                   f"Interval=[{point_info['point1_time']:.3f}h, {point_info['point2_time']:.3f}h], "
                                   f"Values=[{point_info['value1']:.6f}, {point_info['value2']:.6f}], "
                                   f"Slope={point_info['slope']:.6f}\n")
                        if len(data['used_points']) > 10:
                            f.write(f"... and {len(data['used_points']) - 10} more calculation points\n")
                    f.write("\n")
            
            print(f"Debug: Detailed information saved to: {details_path}")
            
            # Show success message
            QMessageBox.information(self, "Save Complete", 
                                   f"Slope calculation results saved:\n"
                                   f"Data file: {file_path}\n"
                                   f"Details file: {details_path}")
            
        except Exception as e:
            print(f"Debug: Manual save failed: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Save Error", f"Failed to save slope data: {str(e)}")


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
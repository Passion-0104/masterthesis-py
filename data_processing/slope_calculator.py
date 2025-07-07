"""
斜率计算模块
用于计算数据的时间斜率（变化率）
"""

import numpy as np
import pandas as pd
from scipy import stats
from scipy.signal import savgol_filter
from typing import Dict, List, Optional, Tuple


class SlopeCalculator:
    """斜率计算器"""
    
    def __init__(self):
        self.slope_data = {}
        
    def calculate_slopes(self, data: pd.DataFrame, 
                        selected_columns: List[str],
                        time_column: str,
                        interval_minutes: float = None,
                        method: str = 'interval_regression',
                        window_minutes: float = None,
                        left_window_minutes: float = None,
                        right_window_minutes: float = None,
                        calculation_interval_seconds: float = 30.0,
                        smoothing: bool = False,
                        smooth_window: int = 15,
                        smooth_order: int = 2) -> Dict:
        """
        计算斜率
        
        Args:
            data: 包含数据的DataFrame
            selected_columns: 要计算斜率的列名列表
            time_column: 时间列名
            interval_minutes: 计算间隔（分钟）- 用于interval_based方法
            method: 计算方法 ('interval_regression', 'continuous_regression', 'moving_regression', 'interval_based')
            window_minutes: 滑动窗口大小（分钟）- 用于moving_regression方法
            left_window_minutes: 左窗口大小（分钟）- 用于continuous_regression方法
            right_window_minutes: 右窗口大小（分钟）- 用于continuous_regression方法
            calculation_interval_seconds: 计算间隔（秒）- 用于interval_regression方法
            smoothing: 是否应用Savitzky-Golay平滑
            smooth_window: 平滑窗口大小
            smooth_order: 平滑多项式阶数
            
        Returns:
            包含斜率数据的字典
        """
        if data.empty or not selected_columns:
            return {}
        
        # Choose calculation method
        if method == 'interval_regression':
            results = self._calculate_slopes_interval_regression(
                data, selected_columns, time_column, 
                calculation_interval_seconds, left_window_minutes, right_window_minutes
            )
        elif method == 'continuous_regression':
            results = self._calculate_slopes_continuous_regression(
                data, selected_columns, time_column, left_window_minutes, right_window_minutes
            )
        elif method == 'moving_regression':
            if window_minutes is None:
                window_minutes = interval_minutes * 2 if interval_minutes else 30.0
            results = self._calculate_slopes_moving_regression(
                data, selected_columns, time_column, interval_minutes, window_minutes
            )
        elif method == 'interval_based':
            if interval_minutes is None:
                interval_minutes = 30.0
            results = self._calculate_slopes_interval_based(
                data, selected_columns, time_column, interval_minutes
            )
        else:
            raise ValueError(f"Unknown method: {method}")
        
        # Apply Savitzky-Golay smoothing if enabled
        if smoothing and results:
            results = self._apply_savgol_smoothing(results, smooth_window, smooth_order)
        
        return results
    
    def _calculate_slopes_moving_regression(self, data: pd.DataFrame, 
                                          selected_columns: List[str],
                                          time_column: str,
                                          interval_minutes: float,
                                          window_minutes: float) -> Dict:
        """
        使用滑动线性拟合方法计算斜率
        
        Args:
            data: DataFrame containing the data
            selected_columns: List of columns to calculate slopes for
            time_column: Name of the time column
            interval_minutes: Time interval in minutes (calculate slope every X minutes)
            window_minutes: Window size for moving regression in minutes
            
        Returns:
            Dictionary containing slope calculation results
        """
        results = {}
        
        # Convert time intervals to hours
        interval_hours = interval_minutes / 60.0
        window_hours = window_minutes / 60.0
        half_window = window_hours / 2.0
        
        print(f"Debug: Using moving linear regression method")
        print(f"Debug: Calculation interval: {interval_minutes} minutes ({interval_hours:.3f} hours)")
        print(f"Debug: Sliding window: {window_minutes} minutes ({window_hours:.3f} hours)")
        
        # 准备时间数据
        if 'relative_time' in data.columns:
            time_data = data['relative_time'].copy()
        else:
            # 创建相对时间
            if time_column in data.columns:
                time_data = pd.to_datetime(data[time_column], errors='coerce')
                time_data = (time_data - time_data.min()).dt.total_seconds() / 3600
            else:
                time_data = pd.Series(range(len(data)), dtype=float)
        
        # 获取时间范围
        min_time = time_data.min()
        max_time = time_data.max()
        
        if pd.isna(min_time) or pd.isna(max_time):
            return {}
        
        # 生成计算时间点：从窗口的一半开始，到数据结束减去窗口的一半
        start_time = min_time + half_window
        end_time = max_time - half_window
        
        if start_time >= end_time:
            print(f"Debug: Data range too small for moving regression analysis")
            return {}
        
        calc_times = np.arange(start_time, end_time + 0.001, interval_hours)
        
        print(f"Debug: Valid calculation time range: {start_time:.3f}h to {end_time:.3f}h")
        print(f"Debug: Number of calculation time points: {len(calc_times)}")
        
        # 对每个选定的列计算斜率
        for col in selected_columns:
            if col not in data.columns:
                continue
                
            # 获取有效数据
            valid_mask = pd.notna(data[col]) & pd.notna(time_data)
            if valid_mask.sum() < 3:  # 至少需要3个点才能进行线性拟合
                print(f"Debug: {col} insufficient data points, skipping")
                continue
                
            col_data = data[col][valid_mask]
            col_time = time_data[valid_mask]
            
            slopes = []
            slope_times = []
            r_squared_values = []  # 存储拟合优度
            n_points_used = []  # 存储每次拟合使用的点数
            
            for calc_time in calc_times:
                # 定义滑动窗口
                window_start = calc_time - half_window
                window_end = calc_time + half_window
                
                # 选择窗口内的数据点
                window_mask = (col_time >= window_start) & (col_time <= window_end)
                window_times = col_time[window_mask]
                window_values = col_data[window_mask]
                
                if len(window_times) < 3:  # 至少需要3个点进行线性拟合
                    continue
                
                try:
                    # 执行线性拟合 y = ax + b
                    slope, intercept, r_value, p_value, std_err = stats.linregress(window_times, window_values)
                    
                    # 检查拟合质量
                    if np.isfinite(slope) and np.isfinite(r_value):
                        slopes.append(slope)
                        slope_times.append(calc_time)
                        r_squared_values.append(r_value ** 2)
                        n_points_used.append(len(window_times))
                    
                except Exception as e:
                    print(f"Debug: Linear fitting failed at {calc_time:.3f}h: {e}")
                    continue
            
            if slopes:
                print(f"Debug: {col} calculated {len(slopes)} slope points")
                print(f"Debug: Average R² = {np.mean(r_squared_values):.4f}")
                print(f"Debug: Average points used = {np.mean(n_points_used):.1f}")
                
                results[col] = {
                    'times': np.array(slope_times),
                    'slopes': np.array(slopes),
                    'r_squared': np.array(r_squared_values),
                    'n_points': np.array(n_points_used),
                    'original_column': col,
                    'interval_minutes': interval_minutes,
                    'window_minutes': window_minutes,
                    'units': 'ppm/hour',
                    'calculation_method': 'moving_regression'
                }
        
        return results
    
    def _calculate_slopes_interval_regression(self, data: pd.DataFrame, 
                                            selected_columns: List[str],
                                            time_column: str,
                                            calculation_interval_seconds: float,
                                            left_window_minutes: float = None,
                                            right_window_minutes: float = None) -> Dict:
        """
        使用间隔回归方法计算斜率：每隔指定秒数计算一次，使用左右窗口进行线性拟合
        
        Args:
            data: 包含数据的DataFrame
            selected_columns: 要计算斜率的列名列表
            time_column: 时间列名
            calculation_interval_seconds: 计算间隔（秒）
            left_window_minutes: 左窗口大小（分钟）
            right_window_minutes: 右窗口大小（分钟）
            
        Returns:
            包含斜率数据的字典
        """
        if data.empty or not selected_columns:
            return {}
        
        # Set default window sizes if not provided
        if left_window_minutes is None:
            left_window_minutes = 15.0  # Default 15 minutes left
        if right_window_minutes is None:
            right_window_minutes = 15.0  # Default 15 minutes right
        
        # Convert calculation interval to hours
        calculation_interval_hours = calculation_interval_seconds / 3600.0
        
        print(f"Debug: Starting slope calculation - method: interval_regression")
        print(f"Debug: Calculation interval: {calculation_interval_seconds} seconds ({calculation_interval_hours:.4f} hours)")
        print(f"Debug: Left window: {left_window_minutes} minutes, Right window: {right_window_minutes} minutes")
        
        # Prepare time data
        if 'relative_time' in data.columns:
            time_data = data['relative_time'].copy()
        else:
            # Create relative time
            if time_column in data.columns:
                time_data = pd.to_datetime(data[time_column], errors='coerce')
                time_data = (time_data - time_data.min()).dt.total_seconds() / 3600
            else:
                time_data = pd.Series(range(len(data)), dtype=float)
        
        # Get time range
        min_time = time_data.min()
        max_time = time_data.max()
        
        if pd.isna(min_time) or pd.isna(max_time):
            return {}
        
        # Generate calculation time points: every calculation_interval_seconds
        calc_times = np.arange(min_time, max_time + 0.001, calculation_interval_hours)
        
        print(f"Debug: Data time range: {min_time:.3f}h to {max_time:.3f}h")
        print(f"Debug: Number of calculation time points: {len(calc_times)}")
        
        results = {}
        
        # Calculate slopes for each selected column
        for col in selected_columns:
            if col not in data.columns:
                continue
                
            # Get valid data
            valid_mask = pd.notna(data[col]) & pd.notna(time_data)
            if valid_mask.sum() < 3:  # Need at least 3 points for regression
                continue
                
            col_data = data[col][valid_mask]
            col_time = time_data[valid_mask]
            
            slopes = []
            slope_times = []
            r_squared_values = []
            n_points_used = []
            valid_calculations = 0
            total_points = len(calc_times)
            
            for idx, calc_time in enumerate(calc_times):
                # Define window boundaries
                left_boundary = calc_time - (left_window_minutes / 60.0)
                right_boundary = calc_time + (right_window_minutes / 60.0)
                
                # 从原始数据中选择窗口内的数据点（重新检查NaN）
                window_mask = (time_data >= left_boundary) & (time_data <= right_boundary)
                window_times = time_data[window_mask]
                window_values = data[col][window_mask]
                
                # 在窗口内重新过滤NaN值
                valid_in_window = pd.notna(window_times) & pd.notna(window_values)
                window_times_clean = window_times[valid_in_window]
                window_values_clean = window_values[valid_in_window]
                
                # 设置最小数据点要求
                min_required_points = max(10, int(len(calc_times) * 0.05))  # 至少10个点，或总点数的5%
                
                # 如果双窗口数据不足，尝试只使用左窗口
                if len(window_times_clean) < min_required_points:
                    # 对于数据末尾的点，如果双窗口数据不足，尝试只用左窗口
                    left_mask = (time_data >= left_boundary) & (time_data <= calc_time)
                    left_times = time_data[left_mask]
                    left_values = data[col][left_mask]
                    
                    # 在左窗口内重新过滤NaN值
                    valid_in_left = pd.notna(left_times) & pd.notna(left_values)
                    left_times_clean = left_times[valid_in_left]
                    left_values_clean = left_values[valid_in_left]
                    
                    if len(left_times_clean) >= min_required_points:
                        window_times_clean = left_times_clean
                        window_values_clean = left_values_clean
                        print(f"Debug: Using left window only for calc_time {calc_time:.3f}h ({len(left_times_clean)} points)")
                    else:
                        print(f"Debug: Skipping calc_time {calc_time:.3f}h - insufficient valid data ({len(left_times_clean)} < {min_required_points})")
                        continue  # 数据不足，跳过该点
                
                if len(window_times_clean) < min_required_points:
                    print(f"Debug: Skipping calc_time {calc_time:.3f}h - insufficient valid data ({len(window_times_clean)} < {min_required_points})")
                    continue
                try:
                    # Perform linear regression y = ax + b with cleaned data
                    slope, intercept, r_value, p_value, std_err = stats.linregress(window_times_clean, window_values_clean)
                    
                    # 额外验证：检查结果是否合理
                    if (np.isfinite(slope) and np.isfinite(r_value) and 
                        np.isfinite(intercept) and abs(slope) < 1000):  # 防止异常大的斜率值
                        slopes.append(slope)
                        slope_times.append(calc_time)
                        r_squared_values.append(r_value ** 2)
                        n_points_used.append(len(window_times_clean))
                        valid_calculations += 1
                        
                        # 如果R²值很低，给出警告
                        if r_value ** 2 < 0.1:
                            print(f"Debug: Low R² ({r_value**2:.3f}) at calc_time {calc_time:.3f}h with {len(window_times_clean)} points")
                    else:
                        print(f"Debug: Invalid regression result at calc_time {calc_time:.3f}h (slope={slope:.3f}, r_value={r_value:.3f})")
                        
                except Exception as e:
                    print(f"Debug: Regression failed at calc_time {calc_time:.3f}h: {str(e)}")
                    continue
            
            if slopes:
                print(f"Debug: {col} calculated {len(slopes)} slope points from {total_points} calculation points ({valid_calculations/total_points*100:.1f}% coverage)")
                print(f"Debug: Average R² = {np.mean(r_squared_values):.4f}")
                print(f"Debug: Average points used = {np.mean(n_points_used):.1f}")
                
                results[col] = {
                    'times': np.array(slope_times),
                    'slopes': np.array(slopes),
                    'r_squared': np.array(r_squared_values),
                    'n_points': np.array(n_points_used),
                    'original_column': col,
                    'calculation_interval_seconds': calculation_interval_seconds,
                    'left_window_minutes': left_window_minutes,
                    'right_window_minutes': right_window_minutes,
                    'total_window_minutes': left_window_minutes + right_window_minutes,
                    'units': 'ppm/hour',
                    'calculation_method': 'interval_regression'
                }
        
        print(f"Debug: Interval regression completed, {len(results)} columns processed")
        return results
    
    def _calculate_slopes_continuous_regression(self, data: pd.DataFrame, 
                                              selected_columns: List[str],
                                              time_column: str,
                                              left_window_minutes: float,
                                              right_window_minutes: float) -> Dict:
        """
        连续斜率计算：对每个数据点都计算斜率
        
        Args:
            data: DataFrame containing the data
            selected_columns: List of columns to calculate slopes for
            time_column: Name of the time column
            left_window_minutes: 向左取多少分钟的数据
            right_window_minutes: 向右取多少分钟的数据
            
        Returns:
            Dictionary containing slope calculation results
        """
        results = {}
        
        # Convert time intervals to hours
        left_window_hours = left_window_minutes / 60.0
        right_window_hours = right_window_minutes / 60.0
        total_window_hours = left_window_hours + right_window_hours
        
        print(f"Debug: Using continuous linear regression method")
        print(f"Debug: Left window: {left_window_minutes} minutes ({left_window_hours:.3f} hours)")
        print(f"Debug: Right window: {right_window_minutes} minutes ({right_window_hours:.3f} hours)")
        print(f"Debug: Total window: {left_window_minutes + right_window_minutes} minutes ({total_window_hours:.3f} hours)")
        
        # 准备时间数据
        if 'relative_time' in data.columns:
            time_data = data['relative_time'].copy()
        else:
            # 创建相对时间
            if time_column in data.columns:
                time_data = pd.to_datetime(data[time_column], errors='coerce')
                time_data = (time_data - time_data.min()).dt.total_seconds() / 3600
            else:
                time_data = pd.Series(range(len(data)), dtype=float)
        
        # 获取时间范围
        min_time = time_data.min()
        max_time = time_data.max()
        
        if pd.isna(min_time) or pd.isna(max_time):
            return {}
        
        print(f"Debug: Data time range: {min_time:.3f}h to {max_time:.3f}h")
        
        # 对每个选定的列计算斜率
        for col in selected_columns:
            if col not in data.columns:
                continue
                
            # 获取有效数据
            valid_mask = pd.notna(data[col]) & pd.notna(time_data)
            if valid_mask.sum() < 3:  # 至少需要3个点才能进行线性拟合
                print(f"Debug: {col} insufficient data points, skipping")
                continue
                
            col_data = data[col][valid_mask]
            col_time = time_data[valid_mask]
            
            slopes = []
            slope_times = []
            r_squared_values = []  # 存储拟合优度
            n_points_used = []  # 存储每次拟合使用的点数
            
            # 对每个数据点都计算斜率
            total_points = len(col_time)
            valid_calculations = 0
            
            for i, calc_time in enumerate(col_time):
                # 定义当前点的左右窗口
                window_start = calc_time - left_window_hours
                window_end = calc_time + right_window_hours
                
                # 从原始数据中选择窗口内的数据点（重新检查NaN）
                window_mask = (time_data >= window_start) & (time_data <= window_end)
                window_times = time_data[window_mask]
                window_values = data[col][window_mask]
                
                # 在窗口内重新过滤NaN值
                valid_in_window = pd.notna(window_times) & pd.notna(window_values)
                window_times_clean = window_times[valid_in_window]
                window_values_clean = window_values[valid_in_window]
                
                # 设置最小数据点要求（连续回归需要更少的点）
                min_required_points = max(5, int(total_points * 0.01))  # 至少5个点，或总点数的1%
                
                if len(window_times_clean) < min_required_points:
                    continue
                
                try:
                    # 执行线性拟合 y = ax + b with cleaned data
                    slope, intercept, r_value, p_value, std_err = stats.linregress(window_times_clean, window_values_clean)
                    
                    # 检查拟合质量和结果合理性
                    if (np.isfinite(slope) and np.isfinite(r_value) and 
                        np.isfinite(intercept) and abs(slope) < 1000):
                        slopes.append(slope)
                        slope_times.append(calc_time)
                        r_squared_values.append(r_value ** 2)
                        n_points_used.append(len(window_times_clean))
                        valid_calculations += 1
                    
                except Exception as e:
                    # 静默跳过拟合失败的点，避免过多输出
                    continue
            
            if slopes:
                print(f"Debug: {col} calculated {len(slopes)} slope points from {total_points} data points ({valid_calculations/total_points*100:.1f}% coverage)")
                print(f"Debug: Average R² = {np.mean(r_squared_values):.4f}")
                print(f"Debug: Average points used = {np.mean(n_points_used):.1f}")
                
                results[col] = {
                    'times': np.array(slope_times),
                    'slopes': np.array(slopes),
                    'r_squared': np.array(r_squared_values),
                    'n_points': np.array(n_points_used),
                    'original_column': col,
                    'left_window_minutes': left_window_minutes,
                    'right_window_minutes': right_window_minutes,
                    'total_window_minutes': left_window_minutes + right_window_minutes,
                    'units': 'ppm/hour',
                    'calculation_method': 'continuous_regression'
                }
        
        print(f"Debug: Continuous regression completed, {len(results)} columns processed")
        return results
    
    def _calculate_slopes_interval_based(self, data: pd.DataFrame, 
                                       selected_columns: List[str],
                                       time_column: str,
                                       interval_minutes: float) -> Dict:
        """
        使用原有的基于间隔的方法计算斜率
        """
        results = {}
        
        # Convert time interval to hours
        interval_hours = interval_minutes / 60.0
        
        print(f"Debug: Using interval-based method")
        print(f"Debug: Calculation interval: {interval_minutes} minutes ({interval_hours:.3f} hours)")
        
        # 准备时间数据
        if 'relative_time' in data.columns:
            time_data = data['relative_time'].copy()
        else:
            # 创建相对时间
            if time_column in data.columns:
                time_data = pd.to_datetime(data[time_column], errors='coerce')
                time_data = (time_data - time_data.min()).dt.total_seconds() / 3600
            else:
                time_data = pd.Series(range(len(data)), dtype=float)
        
        # 获取时间范围
        min_time = time_data.min()
        max_time = time_data.max()
        
        if pd.isna(min_time) or pd.isna(max_time):
            return {}
        
        # 生成计算时间点：从0开始，每隔interval_hours一个点，直到数据结束
        calc_times = np.arange(min_time, max_time + 0.001, interval_hours)
        
        print(f"Debug: Calculation time range: {min_time:.3f}h to {max_time:.3f}h")
        print(f"Debug: Number of calculation time points: {len(calc_times)}")
        
        # 对每个选定的列计算斜率
        for col in selected_columns:
            if col not in data.columns:
                continue
                
            # 获取有效数据
            valid_mask = pd.notna(data[col]) & pd.notna(time_data)
            if valid_mask.sum() < 2:  # 至少需要2个点才能计算斜率
                continue
                
            col_data = data[col][valid_mask]
            col_time = time_data[valid_mask]
            
            # 创建插值函数以获取任意时间点的值
            from scipy.interpolate import interp1d
            try:
                interp_func = interp1d(col_time, col_data, kind='linear', 
                                     bounds_error=False, fill_value='extrapolate')
            except Exception as e:
                print(f"Debug: Interpolation function creation failed for {col}: {e}")
                continue
            
            slopes = []
            slope_times = []
            used_points = []  # 记录使用的点对信息
            
            for i, calc_time in enumerate(calc_times):
                # Calculate using X-minute intervals: every X minutes, calculate slope over X minutes
                # For each time point, use that point and the next interval point
                if i >= len(calc_times) - 1:  # Skip last point as there's no next point
                    continue
                    
                point1_time = calc_time
                point2_time = calc_times[i + 1]
                
                # 确保时间点在数据范围内
                point1_time = max(point1_time, min_time)
                point2_time = min(point2_time, max_time)
                
                if point2_time <= point1_time:
                    continue
                
                # 获取两个时间点的值
                try:
                    value1 = float(interp_func(point1_time))
                    value2 = float(interp_func(point2_time))
                    
                    if np.isnan(value1) or np.isnan(value2):
                        continue
                    
                    # 计算斜率 (ppm/hour)
                    time_diff = point2_time - point1_time
                    if time_diff > 0:
                        slope = (value2 - value1) / time_diff
                        slopes.append(slope)
                        slope_times.append(calc_time)
                        used_points.append({
                            'calc_time': calc_time,
                            'point1_time': point1_time,
                            'point2_time': point2_time,
                            'value1': value1,
                            'value2': value2,
                            'slope': slope
                        })
                        
                except Exception as e:
                    print(f"Debug: Slope calculation failed at {calc_time:.3f}h: {e}")
                    continue
            
            if slopes:
                print(f"Debug: {col} calculated {len(slopes)} slope points")
                results[col] = {
                    'times': np.array(slope_times),
                    'slopes': np.array(slopes),
                    'original_column': col,
                    'interval_minutes': interval_minutes,
                    'units': 'ppm/hour',
                    'calculation_method': 'interval_based',
                    'used_points': used_points  # Debug information
                }
        
        return results
    
    def get_slope_statistics(self, slope_results: Dict) -> Dict:
        """
        获取斜率统计信息
        
        Args:
            slope_results: calculate_slopes返回的结果
            
        Returns:
            统计信息字典
        """
        stats_dict = {}
        
        for col, data in slope_results.items():
            slopes = data['slopes']
            if len(slopes) == 0:
                continue
                
            stats_dict[col] = {
                'mean_slope': np.mean(slopes),
                'std_slope': np.std(slopes),
                'max_slope': np.max(slopes),
                'min_slope': np.min(slopes),
                'median_slope': np.median(slopes),
                'num_points': len(slopes)
            }
        
        return stats_dict
    
    def export_slope_data(self, slope_results: Dict) -> pd.DataFrame:
        """
        将斜率数据导出为DataFrame
        
        Args:
            slope_results: calculate_slopes返回的结果
            
        Returns:
            包含斜率数据的DataFrame
        """
        if not slope_results:
            return pd.DataFrame()
        
        # 找到最长的时间序列作为基准
        max_length = 0
        base_times = None
        
        for col, data in slope_results.items():
            if len(data['times']) > max_length:
                max_length = len(data['times'])
                base_times = data['times']
        
        if base_times is None:
            return pd.DataFrame()
        
        # 创建结果DataFrame
        result_df = pd.DataFrame()
        result_df['time_hours'] = base_times
        
        # 添加每列的斜率数据
        for col, data in slope_results.items():
            # 使用插值将数据对齐到基准时间
            if len(data['times']) == len(base_times) and np.allclose(data['times'], base_times):
                # 时间完全匹配
                result_df[f'{col}_slope'] = data['slopes']
            else:
                # 需要插值
                interp_slopes = np.interp(base_times, data['times'], data['slopes'], 
                                        left=np.nan, right=np.nan)
                result_df[f'{col}_slope'] = interp_slopes
        
        return result_df
    
    def smooth_slopes(self, slope_results: Dict, window_size: int = 3) -> Dict:
        """
        对斜率数据进行平滑处理
        
        Args:
            slope_results: 原始斜率结果
            window_size: 平滑窗口大小
            
        Returns:
            平滑后的斜率结果
        """
        smoothed_results = {}
        
        for col, data in slope_results.items():
            slopes = data['slopes']
            
            if len(slopes) < window_size:
                # 数据点太少，不进行平滑
                smoothed_results[col] = data.copy()
                continue
            
            # 使用移动平均进行平滑
            smoothed_slopes = np.convolve(slopes, np.ones(window_size)/window_size, mode='same')
            
            # 复制原始数据并替换斜率
            smoothed_data = data.copy()
            smoothed_data['slopes'] = smoothed_slopes
            smoothed_data['smoothed'] = True
            smoothed_data['smooth_window'] = window_size
            
            smoothed_results[col] = smoothed_data
        
        return smoothed_results 
    
    def _apply_savgol_smoothing(self, slope_results: Dict, window_length: int, polyorder: int) -> Dict:
        """
        Apply Savitzky-Golay smoothing to slope results
        
        Args:
            slope_results: Original slope calculation results
            window_length: Length of the filter window (must be odd)
            polyorder: Order of the polynomial used for fitting
            
        Returns:
            Smoothed slope results
        """
        smoothed_results = {}
        
        # Ensure window_length is odd
        if window_length % 2 == 0:
            window_length += 1
        
        print(f"Debug: Applying Savitzky-Golay smoothing (window={window_length}, order={polyorder})")
        
        for col, data in slope_results.items():
            slopes = data['slopes']
            
            if len(slopes) < window_length:
                # Not enough points for smoothing, keep original
                print(f"Debug: {col} - insufficient points for smoothing ({len(slopes)} < {window_length}), keeping original")
                smoothed_results[col] = data.copy()
                continue
            
            try:
                # Apply Savitzky-Golay filter
                smoothed_slopes = savgol_filter(slopes, window_length, polyorder)
                
                # Calculate smoothing statistics
                original_std = np.std(slopes)
                smoothed_std = np.std(smoothed_slopes)
                noise_reduction = (1 - smoothed_std / original_std) * 100 if original_std > 0 else 0
                
                print(f"Debug: {col} - smoothed {len(slopes)} points, noise reduction: {noise_reduction:.1f}%")
                
                # Copy original data and replace slopes
                smoothed_data = data.copy()
                smoothed_data['slopes'] = smoothed_slopes
                smoothed_data['original_slopes'] = slopes  # Keep original for comparison
                smoothed_data['smoothed'] = True
                smoothed_data['smooth_method'] = 'savgol'
                smoothed_data['smooth_window'] = window_length
                smoothed_data['smooth_order'] = polyorder
                smoothed_data['noise_reduction_percent'] = noise_reduction
                
                smoothed_results[col] = smoothed_data
                
            except Exception as e:
                print(f"Debug: Savitzky-Golay smoothing failed for {col}: {e}")
                # Keep original if smoothing fails
                smoothed_results[col] = data.copy()
        
        return smoothed_results
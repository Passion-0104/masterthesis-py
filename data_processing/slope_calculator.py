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
                        interval_minutes: float,
                        method: str = 'moving_regression',
                        window_minutes: float = None,
                        smoothing: bool = False,
                        smooth_window: int = 15,
                        smooth_order: int = 2) -> Dict:
        """
        Calculate slopes for specified columns using different methods
        
        Args:
            data: DataFrame containing the data
            selected_columns: List of columns to calculate slopes for
            time_column: Name of the time column
            interval_minutes: Time interval in minutes (calculate slope every X minutes)
            method: Calculation method ('moving_regression' or 'interval_based')
            window_minutes: Window size for moving regression (default: 2 * interval_minutes)
            smoothing: Enable Savitzky-Golay post-processing smoothing
            smooth_window: Window length for Savitzky-Golay filter (must be odd)
            smooth_order: Polynomial order for Savitzky-Golay filter
            
        Returns:
            Dictionary containing slope calculation results
        """
        if data.empty or not selected_columns:
            return {}
        
        # Set default window size for moving regression
        if window_minutes is None:
            window_minutes = 2 * interval_minutes
        
        # Choose calculation method
        if method == 'moving_regression':
            results = self._calculate_slopes_moving_regression(
                data, selected_columns, time_column, interval_minutes, window_minutes)
        elif method == 'interval_based':
            results = self._calculate_slopes_interval_based(
                data, selected_columns, time_column, interval_minutes)
        else:
            raise ValueError(f"Unsupported method: {method}. Use 'moving_regression' or 'interval_based'")
        
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
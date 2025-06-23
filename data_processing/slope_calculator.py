"""
斜率计算模块
用于计算数据的时间斜率（变化率）
"""

import numpy as np
import pandas as pd
from scipy import stats
from typing import Dict, List, Optional, Tuple


class SlopeCalculator:
    """斜率计算器"""
    
    def __init__(self):
        self.slope_data = {}
        
    def calculate_slopes(self, data: pd.DataFrame, 
                        selected_columns: List[str],
                        time_column: str,
                        interval_minutes: float) -> Dict:
        """
        Calculate slopes for specified columns using X-minute intervals
        
        Args:
            data: DataFrame containing the data
            selected_columns: List of columns to calculate slopes for
            time_column: Name of the time column
            interval_minutes: Time interval in minutes (calculate slope every X minutes using X-minute intervals)
            
        Returns:
            Dictionary containing slope calculation results
        """
        if data.empty or not selected_columns:
            return {}
            
        results = {}
        
        # Convert time interval to hours
        interval_hours = interval_minutes / 60.0
        
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
        
        print(f"Debug: 计算时间点范围: {min_time:.3f}h 到 {max_time:.3f}h")
        print(f"Debug: 计算间隔: {interval_hours:.3f}h ({interval_minutes}分钟)")
        print(f"Debug: 计算时间点: {[f'{t:.3f}h({t*60:.0f}min)' for t in calc_times[:10]]}...")
        
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
                print(f"Debug: 插值函数创建失败 for {col}: {e}")
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
                    print(f"Debug: 计算斜率失败 at {calc_time:.3f}h: {e}")
                    continue
            
            if slopes:
                print(f"Debug: {col} 计算了 {len(slopes)} 个斜率点")
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
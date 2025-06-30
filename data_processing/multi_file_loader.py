"""
多文件数据加载器模块，用于加载和组合多个Excel文件的数据
"""

import pandas as pd
import os
from PyQt5.QtWidgets import QMessageBox
from data_processing.data_loader import DataLoader


class MultiFileLoader:
    """用于加载和组合多个Excel文件数据的类"""
    
    def __init__(self):
        self.file_loaders = []  # 存储各个文件的DataLoader实例
        self.file_paths = []    # 存储文件路径
        self.file_data = []     # 存储每个文件的数据
        self.combined_data = None  # 组合后的数据
        
    def load_files(self, file_paths):
        """加载多个Excel文件"""
        try:
            self.file_paths = file_paths
            self.file_loaders = []
            self.file_data = []
            
            for file_path in file_paths:
                loader = DataLoader()
                data = loader.load_file(file_path)
                self.file_loaders.append(loader)
                self.file_data.append(data)
                
            return True
            
        except Exception as e:
            raise Exception(f"加载文件失败: {str(e)}")
            
    def get_file_columns(self, file_index):
        """获取指定文件的列名"""
        if 0 <= file_index < len(self.file_loaders):
            return self.file_loaders[file_index].get_columns()
        return []
        
    def get_file_time_range(self, file_index, time_column):
        """获取指定文件指定时间列的时间范围"""
        if 0 <= file_index < len(self.file_data):
            try:
                data = self.file_data[file_index]
                if time_column in data.columns:
                    time_data = pd.to_datetime(data[time_column], errors='coerce')
                    time_data = time_data.dropna()
                    
                    if not time_data.empty:
                        start_time = time_data.min()
                        end_time = time_data.max()
                        
                        # 计算相对时间范围（小时）
                        duration_hours = (end_time - start_time).total_seconds() / 3600
                        
                        return {
                            'start_time': 0,  # 相对开始时间
                            'end_time': duration_hours,  # 相对结束时间（小时）
                            'total_duration': duration_hours
                        }
            except Exception as e:
                print(f"获取文件时间范围失败: {e}")
                
        return {'start_time': 0, 'end_time': 2, 'total_duration': 2}  # 默认值
        
    def get_all_files_info(self):
        """获取所有文件的信息"""
        files_info = []
        for i, file_path in enumerate(self.file_paths):
            info = {
                'index': i,
                'filename': os.path.basename(file_path),
                'path': file_path,
                'columns': self.get_file_columns(i)
            }
            files_info.append(info)
        return files_info
        
    def combine_data_segments(self, segments_config):
        """
        组合多个文件的数据段
        segments_config: 列表，每个元素包含：
        {
            'file_index': 文件索引,
            'column': 列名,
            'start_time': 开始时间(小时),
            'end_time': 结束时间(小时),
            'time_column': 时间列名,
            'label': 自定义标注名称
        }
        """
        try:
            combined_time = []
            combined_values = []
            combined_filenames = []
            time_offset = 0  # 时间偏移，确保数据连续
            
            for i, segment in enumerate(segments_config):
                file_index = segment['file_index']
                column = segment['column']
                start_time = segment.get('start_time', 0)
                end_time = segment.get('end_time', None)
                time_column = segment.get('time_column', None)
                label = segment.get('label', f"数据段{i + 1}")
                
                if file_index >= len(self.file_data):
                    continue
                    
                data = self.file_data[file_index].copy()
                
                # 处理时间列
                if time_column and time_column in data.columns:
                    data[time_column] = pd.to_datetime(data[time_column], errors='coerce')
                    data = data.dropna(subset=[time_column])
                    
                    if not data.empty:
                        start_datetime = data[time_column].min()
                        data['relative_time'] = (data[time_column] - start_datetime).dt.total_seconds() / 3600
                    else:
                        continue
                else:
                    # 创建虚拟时间列
                    data['relative_time'] = range(len(data))
                    
                # 应用时间范围过滤
                if start_time is not None:
                    data = data[data['relative_time'] >= start_time]
                if end_time is not None:
                    data = data[data['relative_time'] <= end_time]
                    
                # 检查列是否存在
                if column not in data.columns:
                    continue
                    
                # 转换数据为数值型
                data[column] = pd.to_numeric(data[column], errors='coerce')
                valid_data = data[['relative_time', column]].dropna()
                
                if valid_data.empty:
                    continue
                    
                # 调整时间，使数据连续
                adjusted_time = valid_data['relative_time'].values + time_offset
                values = valid_data[column].values
                
                combined_time.extend(adjusted_time)
                combined_values.extend(values)
                
                # 为每个数据点添加自定义标注
                combined_filenames.extend([label] * len(adjusted_time))
                
                # 更新时间偏移
                if len(adjusted_time) > 0:
                    time_offset = max(adjusted_time) + 0.1  # 添加小间隔
                    
            # 创建组合数据DataFrame
            if combined_time:
                self.combined_data = pd.DataFrame({
                    'relative_time': combined_time,
                    'combined_value': combined_values,
                    'source': combined_filenames
                })
                
                # 按时间排序
                self.combined_data = self.combined_data.sort_values('relative_time').reset_index(drop=True)
                
                return self.combined_data
            else:
                raise Exception("没有有效的数据可以组合")
                
        except Exception as e:
            raise Exception(f"组合数据失败: {str(e)}")
            
    def get_combined_data(self):
        """获取组合后的数据"""
        return self.combined_data
        
    def is_data_loaded(self):
        """检查是否有数据加载"""
        return len(self.file_data) > 0
        
    def get_file_count(self):
        """获取加载的文件数量"""
        return len(self.file_paths)
        
    def export_combined_data(self, file_path):
        """导出组合数据"""
        if self.combined_data is not None:
            self.combined_data.to_excel(file_path, index=False)
        else:
            raise Exception("没有组合数据可以导出") 
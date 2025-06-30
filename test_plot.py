#!/usr/bin/env python3
"""
测试绘图功能的独立脚本
"""

import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def create_test_data():
    """创建测试数据"""
    print("创建测试数据...")
    
    # 创建4个数据段模拟温度加热过程
    segments = []
    
    # 段1: 20-90°C (0-24小时)
    times1 = pd.date_range(start='2024-01-01 00:00:00', periods=1440, freq='1min')  # 24小时，每分钟一个点
    temp1 = np.linspace(20, 90, 1440) + np.random.normal(0, 2, 1440)  # 加点噪声
    data1 = pd.DataFrame({
        'time': times1,
        'temperature': temp1,
        'relative_time': np.arange(1440) / 60.0  # 转换为小时
    })
    segments.append(('20-90°C', data1))
    
    # 段2: 90-290°C (24-48小时)
    times2 = pd.date_range(start='2024-01-02 00:00:00', periods=1440, freq='1min')
    temp2 = np.linspace(90, 290, 1440) + np.random.normal(0, 3, 1440)
    data2 = pd.DataFrame({
        'time': times2,
        'temperature': temp2,
        'relative_time': np.arange(1440) / 60.0 + 24  # 偏移24小时
    })
    segments.append(('90-290°C', data2))
    
    # 段3: 290-500°C (48-72小时)
    times3 = pd.date_range(start='2024-01-03 00:00:00', periods=1440, freq='1min')
    temp3 = np.linspace(290, 500, 1440) + np.random.normal(0, 5, 1440)
    data3 = pd.DataFrame({
        'time': times3,
        'temperature': temp3,
        'relative_time': np.arange(1440) / 60.0 + 48  # 偏移48小时
    })
    segments.append(('290-500°C', data3))
    
    # 段4: 500-550°C降温 (72-96小时)
    times4 = pd.date_range(start='2024-01-04 00:00:00', periods=1440, freq='1min')
    temp4 = np.linspace(550, 50, 1440) + np.random.normal(0, 4, 1440)  # 降温过程
    data4 = pd.DataFrame({
        'time': times4,
        'temperature': temp4,
        'relative_time': np.arange(1440) / 60.0 + 72  # 偏移72小时
    })
    segments.append(('500-550°C', data4))
    
    return segments

def create_combined_data(segments):
    """模拟多文件加载器的组合数据功能"""
    print("组合测试数据...")
    
    combined_time = []
    combined_values = []
    combined_sources = []
    
    for label, data in segments:
        combined_time.extend(data['relative_time'].values)
        combined_values.extend(data['temperature'].values)
        combined_sources.extend([label] * len(data))
    
    combined_df = pd.DataFrame({
        'relative_time': combined_time,
        'combined_value': combined_values,
        'source': combined_sources,
        'time_unit': ['hours'] * len(combined_time)
    })
    
    print(f"组合数据完成: {len(combined_df)} 个数据点")
    print(f"时间范围: {combined_df['relative_time'].min():.3f} - {combined_df['relative_time'].max():.3f} 小时")
    
    return combined_df

def test_plot_widget():
    """测试绘图组件"""
    try:
        from PyQt5.QtWidgets import QApplication
        from ui.plot_widget import IndependentPlotWindow
        
        # 创建测试数据
        segments = create_test_data()
        combined_data = create_combined_data(segments)
        
        # 创建Qt应用
        app = QApplication(sys.argv)
        
        # 创建绘图窗口
        plot_window = IndependentPlotWindow()
        
        # 设置绘图参数
        plot_settings = {
            'is_multi_file_mode': True,
            'custom_ylabel': 'Temperature (°C)',
            'time_range': 0,  # 显示全部数据
        }
        
        # 选择的列（这里是组合数据，所以用combined_value）
        selected_columns = ['combined_value']
        
        print("开始绘图测试...")
        
        # 调用绘图方法
        plot_window.plot_data(combined_data, selected_columns, plot_settings)
        
        # 显示窗口
        plot_window.show()
        
        print("绘图窗口已打开，请检查结果")
        
        # 运行应用
        sys.exit(app.exec_())
        
    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc()

def debug_data_structure():
    """调试数据结构"""
    print("=== 调试数据结构 ===")
    
    segments = create_test_data()
    combined_data = create_combined_data(segments)
    
    print("\n数据基本信息:")
    print(f"数据形状: {combined_data.shape}")
    print(f"列名: {list(combined_data.columns)}")
    
    print("\n时间数据统计:")
    print(f"相对时间范围: {combined_data['relative_time'].min():.6f} - {combined_data['relative_time'].max():.6f}")
    print(f"相对时间数据类型: {combined_data['relative_time'].dtype}")
    print(f"时间数据前10个值: {combined_data['relative_time'].head(10).values}")
    
    print("\n温度数据统计:")
    print(f"温度范围: {combined_data['combined_value'].min():.2f} - {combined_data['combined_value'].max():.2f}")
    
    print("\n数据源统计:")
    print(f"数据源: {combined_data['source'].unique()}")
    
    print("\n前几行数据:")
    print(combined_data.head())
    
    print("\n后几行数据:")
    print(combined_data.tail())

if __name__ == "__main__":
    print("=== 绘图功能测试 ===")
    
    # 先调试数据结构
    debug_data_structure()
    
    print("\n" + "="*50)
    print("是否启动GUI测试? (y/n): ", end="")
    
    # 简化：直接启动GUI测试
    response = input().strip().lower()
    if response in ['y', 'yes', '']:
        test_plot_widget()
    else:
        print("仅运行数据结构测试") 
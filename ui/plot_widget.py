"""
Plot widget for displaying data visualization in independent windows
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
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
                           QMainWindow, QApplication, QMessageBox, QFileDialog,
                           QLabel, QLineEdit, QComboBox, QDoubleSpinBox, QCheckBox,
                           QGroupBox, QSpinBox, QColorDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor
import pandas as pd
from data_processing.slope_calculator import SlopeCalculator

# 导入插值模块用于平滑曲线
try:
    from scipy.interpolate import interp1d, make_interp_spline
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False
    print("Warning: scipy not available, using linear interpolation for smoothing")


class IndependentPlotWindow(QMainWindow):
    """Independent window for displaying charts"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.slope_calculator = SlopeCalculator()
        
        # 初始化垂直线设置
        self.vertical_lines = []  # 存储添加的垂直线
        self.vline_settings = {
            'position': 70.0,
            'color': '#ff0000',  # 红色
            'linewidth': 2.0,
            'label': 'Open the valve',
            'enabled': False
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the plot window UI"""
        self.setWindowTitle("Data Visualization Chart")
        self.setGeometry(200, 200, 1200, 700)  # 稍微增加宽度以容纳右侧面板
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main horizontal layout
        main_layout = QHBoxLayout(central_widget)
        
        # Create left layout for chart
        left_layout = QVBoxLayout()
        
        # Create matplotlib figure and canvas
        self.figure = Figure(figsize=(12, 8))
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        
        # Add widgets to left layout
        left_layout.addWidget(self.toolbar)
        left_layout.addWidget(self.canvas)
        
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
        
        left_layout.addLayout(button_layout)
        
        # Add left layout to main layout
        main_layout.addLayout(left_layout, stretch=4)  # 占主要空间
        
        # Create right layout for vertical line control panel
        self._create_vertical_line_panel_right(main_layout)
        
    def _create_vertical_line_panel_right(self, main_layout):
        """创建右侧小巧的垂直线控制面板"""
        # 创建右侧面板
        right_widget = QWidget()
        right_widget.setMaximumWidth(250)  # 限制最大宽度
        right_widget.setMinimumWidth(230)  # 设置最小宽度
        right_layout = QVBoxLayout(right_widget)
        
        # 创建垂直线设置组框
        vline_group = QGroupBox("垂直线标记")
        vline_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        vline_layout = QVBoxLayout(vline_group)
        vline_layout.setSpacing(8)  # 减少间距
        
        # 启用复选框
        self.enable_vline_checkbox = QCheckBox("启用标记")
        self.enable_vline_checkbox.setChecked(self.vline_settings['enabled'])
        self.enable_vline_checkbox.stateChanged.connect(self._on_vline_enabled_changed)
        vline_layout.addWidget(self.enable_vline_checkbox)
        
        # 标签文本
        label_layout = QVBoxLayout()
        label_layout.addWidget(QLabel("标注文本:"))
        self.vline_label_input = QLineEdit(self.vline_settings['label'])
        self.vline_label_input.setMaximumHeight(25)
        self.vline_label_input.textChanged.connect(self._on_vline_label_changed)
        label_layout.addWidget(self.vline_label_input)
        vline_layout.addLayout(label_layout)
        
        # 位置设置
        position_layout = QVBoxLayout()
        position_layout.addWidget(QLabel("位置 (小时):"))
        self.vline_position_spin = QDoubleSpinBox()
        self.vline_position_spin.setMinimum(0.0)
        self.vline_position_spin.setMaximum(999.0)
        self.vline_position_spin.setValue(self.vline_settings['position'])
        self.vline_position_spin.setDecimals(1)
        self.vline_position_spin.setMaximumHeight(25)
        self.vline_position_spin.valueChanged.connect(self._on_vline_position_changed)
        position_layout.addWidget(self.vline_position_spin)
        vline_layout.addLayout(position_layout)
        
        # 线宽和颜色设置
        style_layout = QHBoxLayout()
        
        # 线宽
        width_layout = QVBoxLayout()
        width_layout.addWidget(QLabel("线宽:"))
        self.vline_width_spin = QDoubleSpinBox()
        self.vline_width_spin.setMinimum(0.5)
        self.vline_width_spin.setMaximum(10.0)
        self.vline_width_spin.setValue(self.vline_settings['linewidth'])
        self.vline_width_spin.setDecimals(1)
        self.vline_width_spin.setSingleStep(0.5)
        self.vline_width_spin.setMaximumWidth(70)
        self.vline_width_spin.setMaximumHeight(25)
        self.vline_width_spin.valueChanged.connect(self._on_vline_linewidth_changed)
        width_layout.addWidget(self.vline_width_spin)
        style_layout.addLayout(width_layout)
        
        # 颜色
        color_layout = QVBoxLayout()
        color_layout.addWidget(QLabel("颜色:"))
        self.vline_color_button = QPushButton()
        self.vline_color_button.setMaximumWidth(50)
        self.vline_color_button.setMaximumHeight(25)
        self._update_color_button_style()
        self.vline_color_button.clicked.connect(self._choose_vline_color)
        color_layout.addWidget(self.vline_color_button)
        style_layout.addLayout(color_layout)
        
        vline_layout.addLayout(style_layout)
        
        # 应用按钮
        self.apply_vline_button = QPushButton("应用")
        self.apply_vline_button.clicked.connect(self._apply_vertical_line)
        self.apply_vline_button.setEnabled(False)  # 初始禁用
        self.apply_vline_button.setMaximumHeight(30)
        self.apply_vline_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        vline_layout.addWidget(self.apply_vline_button)
        
        # 添加弹性空间
        vline_layout.addStretch()
        
        # 将组框添加到右侧布局
        right_layout.addWidget(vline_group)
        right_layout.addStretch()  # 底部添加弹性空间
        
        # 将右侧面板添加到主布局
        main_layout.addWidget(right_widget, stretch=1)  # 占较小空间
        
    def _update_color_button_style(self):
        """更新颜色按钮的样式"""
        color = self.vline_settings['color']
        self.vline_color_button.setStyleSheet(f"background-color: {color}; border: 1px solid #ccc;")
        
    def _on_vline_enabled_changed(self, state):
        """垂直线启用状态改变"""
        self.vline_settings['enabled'] = state == Qt.Checked
        self.apply_vline_button.setEnabled(self.vline_settings['enabled'])
        
    def _on_vline_label_changed(self, text):
        """垂直线标签改变"""
        self.vline_settings['label'] = text
        
    def _on_vline_position_changed(self, value):
        """垂直线位置改变"""
        self.vline_settings['position'] = value
        
    def _on_vline_linewidth_changed(self, value):
        """垂直线线宽改变"""
        self.vline_settings['linewidth'] = value
        
    def _choose_vline_color(self):
        """选择垂直线颜色"""
        current_color = QColor(self.vline_settings['color'])
        color = QColorDialog.getColor(current_color, self, "选择垂直线颜色")
        
        if color.isValid():
            self.vline_settings['color'] = color.name()
            self._update_color_button_style()
            
    def _apply_vertical_line(self):
        """应用垂直线到当前图表"""
        if hasattr(self, 'figure') and hasattr(self, 'canvas'):
            try:
                ax = self.figure.gca()  # 获取当前轴
                self._add_vertical_line_to_plot(ax)
                self.canvas.draw()
                QMessageBox.information(self, "成功", "垂直线已添加到图表中")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"添加垂直线时出错: {str(e)}")
    
    def _add_vertical_line_to_plot(self, ax):
        """在图表中添加垂直线"""
        if not self.vline_settings['enabled']:
            return
            
        try:
            # 清除之前的垂直线
            for line in self.vertical_lines:
                if hasattr(line, 'remove'):
                    line.remove()
            self.vertical_lines.clear()
            
            # 获取当前绘图设置（如果存在）
            plot_settings = getattr(self, 'current_plot_settings', {})
            is_multi_file_mode = plot_settings.get('is_multi_file_mode', False)
            
            # 根据时间单位调整位置
            position = self.vline_settings['position']
            time_unit_label = "小时"
            
            # 检查是否为多文件模式且有时间单位信息
            if is_multi_file_mode and hasattr(self, 'current_plot_data'):
                plot_df = self.current_plot_data
                if 'time_unit' in plot_df.columns and len(plot_df) > 0:
                    time_unit = plot_df['time_unit'].iloc[0]
                    if time_unit == 'minutes':
                        time_unit_label = "分钟"
                        # 如果当前图表使用分钟单位，但设置是小时，需要转换
                        # 这里position默认按小时设置，需要转换为分钟
                        position = position * 60  # 小时转分钟
            
            color = self.vline_settings['color']
            linewidth = self.vline_settings['linewidth']
            label = self.vline_settings['label']
            
            # 绘制垂直虚线
            vline = ax.axvline(x=position, color=color, linestyle='--', 
                             linewidth=linewidth, alpha=0.8, zorder=10)
            self.vertical_lines.append(vline)
            
            # 添加文本标注
            ylim = ax.get_ylim()
            text_y = ylim[1] * 0.95  # 在图的顶部95%位置
            
            # 添加时间单位到标签
            label_with_unit = f"{label}\n({position:.1f} {time_unit_label})"
            
            text = ax.text(position, text_y, label_with_unit, 
                         horizontalalignment='center', verticalalignment='top',
                         fontsize=10, fontweight='bold', color=color,
                         bbox=dict(boxstyle='round,pad=0.4', facecolor='white', 
                                 edgecolor=color, alpha=0.9),
                         zorder=11)
            
            # 将文本也添加到垂直线列表中以便清理
            self.vertical_lines.append(text)
            
        except Exception as e:
            print(f"绘制垂直线时出错: {e}")
    
    def _create_smooth_curve(self, x_data, y_data, smooth_factor=300):
        """创建平滑曲线数据"""
        try:
            # 确保数据是排序的
            sorted_indices = np.argsort(x_data)
            x_sorted = x_data[sorted_indices]
            y_sorted = y_data[sorted_indices]
            
            # 移除重复的x值
            unique_mask = np.diff(x_sorted, prepend=x_sorted[0] - 1) != 0
            x_unique = x_sorted[unique_mask]
            y_unique = y_sorted[unique_mask]
            
            if len(x_unique) < 3:
                # 数据点太少，返回原始数据
                return x_unique, y_unique
            
            # 创建更密集的x点用于平滑
            x_smooth = np.linspace(x_unique.min(), x_unique.max(), smooth_factor)
            
            if SCIPY_AVAILABLE:
                # 使用样条插值创建平滑曲线
                if len(x_unique) >= 4:
                    # 使用三次样条插值
                    try:
                        spline = make_interp_spline(x_unique, y_unique, k=3)
                        y_smooth = spline(x_smooth)
                    except:
                        # 如果三次样条失败，使用线性插值
                        interp_func = interp1d(x_unique, y_unique, kind='linear', 
                                             bounds_error=False, fill_value='extrapolate')
                        y_smooth = interp_func(x_smooth)
                else:
                    # 数据点太少，使用线性插值
                    interp_func = interp1d(x_unique, y_unique, kind='linear',
                                         bounds_error=False, fill_value='extrapolate')
                    y_smooth = interp_func(x_smooth)
            else:
                # 没有scipy，使用numpy的线性插值
                y_smooth = np.interp(x_smooth, x_unique, y_unique)
            
            return x_smooth, y_smooth
            
        except Exception as e:
            print(f"创建平滑曲线时出错: {e}")
            # 返回原始数据
            return x_data, y_data
    
    def _format_time_axis(self, ax, time_unit='hours'):
        """格式化时间轴显示"""
        try:
            if time_unit == 'minutes':
                # 分钟格式化函数
                def minutes_formatter(x, pos):
                    """将数值转换为时间格式"""
                    total_minutes = int(x)
                    hours = total_minutes // 60
                    minutes = total_minutes % 60
                    
                    if hours > 0:
                        return f"{hours}h{minutes:02d}m"
                    else:
                        return f"{minutes}m"
                
                # 应用格式化
                ax.xaxis.set_major_formatter(FuncFormatter(minutes_formatter))
                
                # 设置智能刻度
                x_min, x_max = ax.get_xlim()
                if x_max <= 120:  # 2小时以内
                    tick_interval = 15  # 每15分钟
                elif x_max <= 300:  # 5小时以内
                    tick_interval = 30  # 每30分钟
                else:
                    tick_interval = 60  # 每1小时
                    
                ticks = np.arange(0, x_max + tick_interval, tick_interval)
                ticks = ticks[ticks <= x_max]
                if len(ticks) > 0:
                    ax.set_xticks(ticks)
                
            else:  # hours
                # 小时格式化函数
                def hours_formatter(x, pos):
                    """将数值转换为时间格式"""
                    if x == 0:
                        return "0h"
                    
                    total_hours = float(x)
                    hours = int(total_hours)
                    minutes = int((total_hours - hours) * 60)
                    
                    if hours > 0:
                        if minutes > 0:
                            return f"{hours}h{minutes:02d}m"
                        else:
                            return f"{hours}h"
                    else:
                        if minutes > 0:
                            return f"{minutes}m"
                        else:
                            return "0h"
                
                # 应用格式化
                ax.xaxis.set_major_formatter(FuncFormatter(hours_formatter))
                
                # 设置智能刻度 - 针对长数据优化
                x_min, x_max = ax.get_xlim()
                print(f"Debug: x轴范围: {x_min:.3f} - {x_max:.3f} 小时")
                
                # 更智能的刻度设置，确保长数据也能正确显示
                if x_max <= 2:  # 2小时以内
                    tick_interval = 0.25  # 每15分钟
                    max_ticks = 10
                elif x_max <= 5:  # 5小时以内
                    tick_interval = 0.5   # 每30分钟
                    max_ticks = 12
                elif x_max <= 12:  # 12小时以内
                    tick_interval = 1     # 每1小时
                    max_ticks = 15
                elif x_max <= 24:  # 24小时以内
                    tick_interval = 2     # 每2小时
                    max_ticks = 15
                elif x_max <= 48:  # 48小时以内
                    tick_interval = 4     # 每4小时
                    max_ticks = 15
                elif x_max <= 72:  # 72小时以内
                    tick_interval = 6     # 每6小时
                    max_ticks = 15
                elif x_max <= 168:  # 1周以内
                    tick_interval = 12    # 每12小时
                    max_ticks = 15
                else:  # 超过1周的数据
                    tick_interval = 24    # 每24小时
                    max_ticks = 15
                    
                # 生成刻度
                ticks = np.arange(0, x_max + tick_interval, tick_interval)
                ticks = ticks[ticks <= x_max]
                
                # 如果刻度太多，减少数量
                if len(ticks) > max_ticks:
                    # 使用更稀疏的刻度
                    step = len(ticks) // max_ticks + 1
                    ticks = ticks[::step]
                
                # 确保始终包含起点和终点
                if len(ticks) > 0:
                    if ticks[0] > 0:
                        ticks = np.insert(ticks, 0, 0)
                    if ticks[-1] < x_max * 0.95:  # 如果最后一个刻度离结束太远
                        ticks = np.append(ticks, x_max)
                
                # 确保至少有几个刻度
                if len(ticks) < 3:
                    num_ticks = min(8, max(3, int(x_max) + 1))
                    ticks = np.linspace(0, x_max, num_ticks)
                
                print(f"Debug: 设置刻度数量: {len(ticks)}, 刻度值: {ticks}")
                
                # 强制设置刻度
                if len(ticks) > 0:
                    ax.set_xticks(ticks)
                    # 强制显示所有标签
                    ax.tick_params(axis='x', which='major', labelsize=9)
                    
                    # 禁用matplotlib的自动标签隐藏
                    ax.xaxis.set_tick_params(which='major', pad=5)
                    
                    # 对于长数据，使用更大的旋转角度避免重叠
                    if x_max > 24:
                        rotation_angle = 60
                    elif x_max > 12:
                        rotation_angle = 45
                    else:
                        rotation_angle = 30
                    
                    ax.tick_params(axis='x', rotation=rotation_angle, labelsize=9)
            
            # 确保x轴标签可见
            plt.setp(ax.xaxis.get_majorticklabels(), visible=True)
            
            # 强制更新图形
            ax.figure.canvas.draw_idle()
            
        except Exception as e:
            print(f"格式化时间轴时出错: {e}")
            # 如果格式化失败，使用默认格式
            ax.ticklabel_format(style='plain', axis='x')
     
    def plot_data(self, data, selected_columns, plot_settings):
        """Plot data in the independent window"""
        try:
            # 存储当前绘图设置和数据供垂直线功能使用
            self.current_plot_settings = plot_settings
            self.current_plot_data = data
            
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            # Configure matplotlib for academic paper standards
            plt.rcParams['font.family'] = ['Arial', 'DejaVu Sans', 'Helvetica', 'sans-serif']
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
            
            # 检查是否为多文件模式
            is_multi_file_mode = plot_settings.get('is_multi_file_mode', False)
            
            # Plot data
            colors = plt.cm.tab10.colors
            calib_colors = ['#ff7f0e', '#d62728', '#9467bd', '#8c564b', 
                          '#e377c2', '#bcbd22', '#17becf', '#f0027f']
            
            if is_multi_file_mode:
                # 多文件模式：绘制组合数据
                if 'combined_value' in plot_df.columns and 'relative_time' in plot_df.columns:
                    valid_data = plot_df[['relative_time', 'combined_value']].dropna()
                    if not valid_data.empty:
                        # 如果有source列，显示不同段的颜色
                        if 'source' in plot_df.columns:
                            sources = plot_df['source'].unique()
                            for i, source in enumerate(sources):
                                source_data = plot_df[plot_df['source'] == source][['relative_time', 'combined_value']].dropna()
                                if not source_data.empty:
                                    color_idx = i % len(colors)
                                    
                                    # 创建平滑曲线
                                    x_smooth, y_smooth = self._create_smooth_curve(
                                        source_data['relative_time'].values, 
                                        source_data['combined_value'].values
                                    )
                                    
                                    ax.plot(x_smooth, y_smooth,
                                          '-', color=colors[color_idx], label=source,
                                          linewidth=2.5, alpha=0.9)
                        else:
                            # 没有source列，绘制整体数据
                            x_smooth, y_smooth = self._create_smooth_curve(
                                valid_data['relative_time'].values, 
                                valid_data['combined_value'].values
                            )
                            
                            ax.plot(x_smooth, y_smooth,
                                  '-', color=colors[0], label="组合数据",
                                  linewidth=2.5, alpha=0.9)
            else:
                # 单文件模式：原有逻辑
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
                                # 创建平滑曲线
                                x_smooth, y_smooth = self._create_smooth_curve(
                                    valid_data['relative_time'].values, 
                                    valid_data[col].values
                                )
                                
                                ax.plot(x_smooth, y_smooth,
                                      '-', color=colors[color_idx],
                                      label=f"{col} (original)",
                                      linewidth=2.0, alpha=0.7)
                        
                        # Plot calibrated data
                        calib_col = f"{col}_calib"
                        if calib_col in calibrated_data:
                            times = calibrated_data[calib_col]['times']
                            values = calibrated_data[calib_col]['values']
                            
                            calib_color = calib_colors[color_idx % len(calib_colors)]
                            
                            # 创建校准数据的平滑曲线
                            x_smooth, y_smooth = self._create_smooth_curve(
                                np.array(times), np.array(values)
                            )
                            
                            ax.plot(x_smooth, y_smooth, '--', color=calib_color,
                                  label=f"{col} (calibrated)",
                                  linewidth=2.5, alpha=0.9)
                            
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
                            # 创建平滑曲线
                            x_smooth, y_smooth = self._create_smooth_curve(
                                valid_data['relative_time'].values, 
                                valid_data[col].values
                            )
                            
                            ax.plot(x_smooth, y_smooth,
                                  '-', color=colors[color_idx], label=col,
                                  linewidth=2.5, alpha=0.9)
            
            # Set labels and title - academic paper standards
            # 根据时间单位设置x轴标签
            if is_multi_file_mode and 'time_unit' in plot_df.columns:
                time_unit = plot_df['time_unit'].iloc[0] if len(plot_df) > 0 else 'hours'
                ax.set_xlabel("Time", fontsize=11, fontweight='normal')
            else:
                time_unit = 'hours'  # 默认单位
                ax.set_xlabel("Time", fontsize=11, fontweight='normal')
            
            # Use custom Y-axis label if provided, otherwise use default
            custom_ylabel = plot_settings.get('custom_ylabel', 'Water Concentration (ppm)')
            if not custom_ylabel or custom_ylabel.strip() == '':
                custom_ylabel = 'Water Concentration (ppm)'
            ax.set_ylabel(custom_ylabel, fontsize=11, fontweight='normal')
            
            # Set up legend - smaller for academic papers
            legend = ax.legend(loc='best', frameon=True, fontsize=9, markerscale=1.5)
            for line in legend.get_lines():
                line.set_linewidth(2.0)
            
            # Store legend reference for updates
            self.current_legend = legend
            
            # Enable legend editing
            legend.set_picker(True)
            
            # Add legend update functionality
            self._setup_legend_editing(ax)
            
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
            
            # 添加用户自定义的垂直线标记
            self._add_vertical_line_to_plot(ax)
            
            # 应用时间格式化（在所有绘图完成后）
            # 检查是否需要使用预设的2小时标签
            time_range_setting = plot_settings.get('time_range', 0)
            use_custom_labels = False
            
            # 只有在明确设置了2小时范围且数据确实在2小时以内时才使用预设标签
            if time_range_setting == 1 and not plot_df.empty:
                max_time = plot_df['relative_time'].max()
                if max_time <= 2.5:  # 给一点容差
                    use_custom_labels = True
                    print("Debug: 使用2小时范围预设标签")
            
            if not use_custom_labels:
                # 应用自动时间格式化 - 对所有其他情况
                print(f"Debug: 应用自动时间格式化，时间单位: {time_unit}, time_range_setting: {time_range_setting}")
                self._format_time_axis(ax, time_unit)
            else:
                # 即使使用预设标签，也要确保标签可见
                ax.tick_params(axis='x', rotation=0, labelsize=10)
                plt.setp(ax.xaxis.get_majorticklabels(), visible=True)
            
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
    
    def _setup_legend_editing(self, ax):
        """Setup legend editing functionality"""
        try:
            # Store axis reference
            self.current_ax = ax
            
            # Connect canvas event to legend update
            self.canvas.mpl_connect('button_release_event', self._on_canvas_click)
            
            # Connect to draw event to update legend after matplotlib operations
            self.canvas.mpl_connect('draw_event', self._check_legend_update)
            
        except Exception as e:
            print(f"Warning: Could not setup legend editing: {e}")
    
    def _on_canvas_click(self, event):
        """Handle canvas click events"""
        try:
            # Check if we need to update legend after user interaction
            if hasattr(self, 'current_ax'):
                # Small delay to allow matplotlib to process the event
                from PyQt5.QtCore import QTimer
                QTimer.singleShot(100, self._update_legend_if_needed)
        except Exception as e:
            print(f"Debug: Canvas click event error: {e}")
    
    def _check_legend_update(self, event):
        """Check if legend needs to be updated after drawing"""
        try:
            self._update_legend_if_needed()
        except Exception as e:
            print(f"Debug: Legend update check error: {e}")
    
    def _update_legend_if_needed(self):
        """Update legend if line labels have changed"""
        try:
            if not hasattr(self, 'current_ax') or not hasattr(self, 'current_legend'):
                return
            
            ax = self.current_ax
            
            # Get current line labels
            current_labels = []
            for line in ax.get_lines():
                label = line.get_label()
                if label and not label.startswith('_'):  # Skip hidden lines
                    current_labels.append(label)
            
            # Get legend labels
            if self.current_legend:
                legend_labels = [text.get_text() for text in self.current_legend.get_texts()]
                
                # Check if labels have changed
                if current_labels != legend_labels:
                    print(f"Debug: Updating legend - old: {legend_labels}, new: {current_labels}")
                    self._refresh_legend(ax)
                    
        except Exception as e:
            print(f"Debug: Legend update error: {e}")
    
    def _refresh_legend(self, ax):
        """Refresh the legend with current line labels"""
        try:
            # Remove old legend
            if hasattr(self, 'current_legend') and self.current_legend:
                self.current_legend.remove()
            
            # Create new legend with updated labels
            legend = ax.legend(loc='best', frameon=True, fontsize=9, markerscale=1.5)
            for line in legend.get_lines():
                line.set_linewidth(2.0)
            
            # Store new legend reference
            self.current_legend = legend
            legend.set_picker(True)
            
            # Redraw canvas
            self.canvas.draw_idle()
            
            print("Debug: Legend refreshed successfully")
            
        except Exception as e:
            print(f"Debug: Legend refresh error: {e}")
    
    def _setup_slope_legend_editing(self, canvas, ax, legend):
        """Setup legend editing functionality for slope charts"""
        try:
            # Store references
            canvas.current_ax = ax
            canvas.current_legend = legend
            
            # Enable legend editing
            legend.set_picker(True)
            
            # Connect canvas events
            canvas.mpl_connect('button_release_event', 
                             lambda event: self._on_slope_canvas_click(event, canvas))
            canvas.mpl_connect('draw_event', 
                             lambda event: self._check_slope_legend_update(event, canvas))
            
        except Exception as e:
            print(f"Warning: Could not setup slope legend editing: {e}")
    
    def _on_slope_canvas_click(self, event, canvas):
        """Handle canvas click events for slope charts"""
        try:
            # Check if we need to update legend after user interaction
            if hasattr(canvas, 'current_ax'):
                # Small delay to allow matplotlib to process the event
                from PyQt5.QtCore import QTimer
                QTimer.singleShot(100, lambda: self._update_slope_legend_if_needed(canvas))
        except Exception as e:
            print(f"Debug: Slope canvas click event error: {e}")
    
    def _check_slope_legend_update(self, event, canvas):
        """Check if slope legend needs to be updated after drawing"""
        try:
            self._update_slope_legend_if_needed(canvas)
        except Exception as e:
            print(f"Debug: Slope legend update check error: {e}")
    
    def _update_slope_legend_if_needed(self, canvas):
        """Update slope legend if line labels have changed"""
        try:
            if not hasattr(canvas, 'current_ax') or not hasattr(canvas, 'current_legend'):
                return
            
            ax = canvas.current_ax
            
            # Get current line labels
            current_labels = []
            for line in ax.get_lines():
                label = line.get_label()
                if label and not label.startswith('_'):  # Skip hidden lines
                    current_labels.append(label)
            
            # Get legend labels
            if canvas.current_legend:
                legend_labels = [text.get_text() for text in canvas.current_legend.get_texts()]
                
                # Check if labels have changed
                if current_labels != legend_labels:
                    print(f"Debug: Updating slope legend - old: {legend_labels}, new: {current_labels}")
                    self._refresh_slope_legend(canvas, ax)
                    
        except Exception as e:
            print(f"Debug: Slope legend update error: {e}")
    
    def _refresh_slope_legend(self, canvas, ax):
        """Refresh the slope legend with current line labels"""
        try:
            # Remove old legend
            if hasattr(canvas, 'current_legend') and canvas.current_legend:
                canvas.current_legend.remove()
            
            # Create new legend with updated labels
            legend = ax.legend(loc='best', frameon=False, fontsize=10)
            for text in legend.get_texts():
                text.set_fontfamily('serif')
            
            # Store new legend reference
            canvas.current_legend = legend
            legend.set_picker(True)
            
            # Redraw canvas
            canvas.draw_idle()
            
            print("Debug: Slope legend refreshed successfully")
            
        except Exception as e:
            print(f"Debug: Slope legend refresh error: {e}")
            
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
                # Check if it's valid relative time (non-negative, reasonable range)
                if (max_time >= min_time and 
                    min_time >= 0 and 
                    max_time <= 8760 and  # Up to 1 year in hours
                    np.isfinite(max_time) and np.isfinite(min_time)):
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
            # Check if the time column is already relative_time (avoid reprocessing)
            if time_col == 'relative_time':
                print(f"Debug: Time column is already relative_time, skipping conversion")
                # No need to convert, just validate and use
                if 'relative_time' not in plot_df.columns:
                    plot_df['relative_time'] = np.arange(len(plot_df)) / 60.0
            else:
                print(f"Debug: Converting time column {time_col} to relative_time")
                try:
                    original_col = plot_df[time_col].copy()
                    plot_df[time_col] = pd.to_datetime(plot_df[time_col], errors='coerce')
                    valid_datetime_data = plot_df.dropna(subset=[time_col])
                    
                    # Calculate relative time
                    if not valid_datetime_data.empty:
                        start_time = valid_datetime_data[time_col].min()
                        plot_df['relative_time'] = (plot_df[time_col] - start_time).dt.total_seconds() / 3600
                        print(f"Debug: Converted to relative_time, range: {plot_df['relative_time'].min():.3f} - {plot_df['relative_time'].max():.3f}h")
                    else:
                        print("Debug: Failed to convert datetime, using original data as relative time")
                        # Fallback: try to use original data as numeric time
                        original_numeric = pd.to_numeric(original_col, errors='coerce')
                        if not original_numeric.isna().all():
                            plot_df['relative_time'] = original_numeric
                            print(f"Debug: Used original numeric data as relative_time, range: {plot_df['relative_time'].min():.3f} - {plot_df['relative_time'].max():.3f}h")
                        else:
                            plot_df['relative_time'] = np.arange(len(plot_df)) / 60.0
                            print("Debug: Created dummy relative_time due to conversion failure")
                except Exception as e:
                    print(f"Debug: Error converting time column: {e}")
                    # Create dummy time if conversion fails
                    plot_df['relative_time'] = np.arange(len(plot_df)) / 60.0
                    print(f"Debug: Created dummy relative_time due to exception, range: {plot_df['relative_time'].min():.3f} - {plot_df['relative_time'].max():.3f}h")
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
            calculation_method = plot_settings.get('slope_method', 'interval_regression')  # 默认使用间隔回归
            window_minutes = plot_settings.get('slope_window', None)  # None = 自动设置为2倍间隔
            left_window_minutes = plot_settings.get('slope_left_window', 15.0)
            right_window_minutes = plot_settings.get('slope_right_window', 15.0)
            calculation_interval_seconds = plot_settings.get('slope_calculation_interval_seconds', 30.0)
            smoothing_enabled = plot_settings.get('slope_smoothing', False)
            smooth_window = plot_settings.get('slope_smooth_window', 15)
            smooth_order = plot_settings.get('slope_smooth_order', 2)
            
            print(f"Debug: Plot settings - calculation_interval_seconds: {calculation_interval_seconds}")
            print(f"Debug: Plot settings - left_window_minutes: {left_window_minutes}")
            print(f"Debug: Plot settings - right_window_minutes: {right_window_minutes}")
            
            print(f"Debug: Starting slope calculation - method: {calculation_method}")
            
            if calculation_method == 'interval_regression':
                print(f"Debug: Calculation interval: {calculation_interval_seconds} seconds")
                print(f"Debug: Left window: {left_window_minutes} minutes, Right window: {right_window_minutes} minutes")
            elif calculation_method == 'continuous_regression':
                print(f"Debug: Left window: {left_window_minutes} minutes, Right window: {right_window_minutes} minutes")
            elif calculation_method == 'moving_regression':
                print(f"Debug: Interval: {interval_minutes} minutes")
                if window_minutes:
                    print(f"Debug: Window size: {window_minutes} minutes")
                else:
                    print(f"Debug: Window size: auto (2 × {interval_minutes} = {2 * interval_minutes} minutes)")
            else:
                print(f"Debug: Interval: {interval_minutes} minutes")
            
            if smoothing_enabled:
                print(f"Debug: Savitzky-Golay smoothing enabled (window={smooth_window}, order={smooth_order})")
            else:
                print(f"Debug: Savitzky-Golay smoothing disabled")
            
            # 计算斜率
            slope_results = self.slope_calculator.calculate_slopes(
                plot_df, selected_columns, time_column, 
                interval_minutes=interval_minutes,
                method=calculation_method, 
                window_minutes=window_minutes,
                left_window_minutes=left_window_minutes,
                right_window_minutes=right_window_minutes,
                calculation_interval_seconds=calculation_interval_seconds,
                smoothing=smoothing_enabled, 
                smooth_window=smooth_window, 
                smooth_order=smooth_order
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
            plt.rcParams['font.family'] = ['Arial', 'DejaVu Sans', 'Helvetica', 'sans-serif']
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
            
            # Clean title for academic papers - handle different calculation methods
            first_result = slope_results[list(slope_results.keys())[0]]
            if 'interval_minutes' in first_result:
                # For interval-based methods
                interval_min = first_result['interval_minutes']
                ax.set_title(f"Slope vs Time ({interval_min}-min intervals)", 
                            fontsize=14, fontfamily='serif', pad=20)
            elif 'calculation_interval_seconds' in first_result:
                # For new interval regression method
                interval_sec = first_result['calculation_interval_seconds']
                ax.set_title(f"Slope vs Time ({interval_sec}-sec intervals)", 
                            fontsize=14, fontfamily='serif', pad=20)
            else:
                # For continuous methods
                ax.set_title(f"Slope vs Time", 
                            fontsize=14, fontfamily='serif', pad=20)
            
            # Academic style legend
            legend = ax.legend(loc='best', frameon=False, fontsize=10)
            for text in legend.get_texts():
                text.set_fontfamily('serif')
            
            # Store legend reference and setup editing for slope chart
            self._setup_slope_legend_editing(canvas, ax, legend)
            
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
                stats_text += f"数据点数: {col_stats['num_points']}\n"
                
                # Add smoothing information if available
                if col in slope_results:
                    slope_data = slope_results[col]
                    if slope_data.get('smoothed', False):
                        noise_reduction = slope_data.get('noise_reduction_percent', 0)
                        smooth_method = slope_data.get('smooth_method', 'unknown')
                        smooth_window = slope_data.get('smooth_window', 'unknown')
                        smooth_order = slope_data.get('smooth_order', 'unknown')
                        stats_text += f"平滑处理: {smooth_method} (窗口={smooth_window}, 阶数={smooth_order})\n"
                        stats_text += f"噪声降低: {noise_reduction:.1f}%\n"
                
                stats_text += "\n"
            
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
                interval_min = plot_settings.get('slope_interval', 15.0)
                method = plot_settings.get('slope_method', 'continuous_regression')
                window_min = plot_settings.get('slope_window', None)
                left_window_min = plot_settings.get('slope_left_window', 15.0)
                right_window_min = plot_settings.get('slope_right_window', 15.0)
                smoothing_enabled = plot_settings.get('slope_smoothing', False)
                smooth_window = plot_settings.get('slope_smooth_window', 15)
                smooth_order = plot_settings.get('slope_smooth_order', 2)
                
                f.write(f"Calculation Method: {method}\n")
                
                if method == 'continuous_regression':
                    f.write(f"Left Window: {left_window_min} minutes\n")
                    f.write(f"Right Window: {right_window_min} minutes\n")
                elif method == 'moving_regression':
                    f.write(f"Time Interval: {interval_min} minutes\n")
                    if window_min:
                        f.write(f"Sliding Window: {window_min} minutes\n")
                    else:
                        f.write(f"Sliding Window: {2 * interval_min} minutes (auto)\n")
                else:
                    f.write(f"Time Interval: {interval_min} minutes\n")
                
                f.write(f"Savitzky-Golay Smoothing: {'Enabled' if smoothing_enabled else 'Disabled'}\n")
                if smoothing_enabled:
                    f.write(f"Smoothing Window: {smooth_window}\n")
                    f.write(f"Polynomial Order: {smooth_order}\n")
                f.write("\n")
                
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
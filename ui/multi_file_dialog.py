"""
多文件选择和配置对话框
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                           QLabel, QComboBox, QListWidget, QGroupBox, QTableWidget,
                           QTableWidgetItem, QSpinBox, QDoubleSpinBox, QFileDialog,
                           QMessageBox, QHeaderView, QAbstractItemView, QCheckBox, QLineEdit)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from data_processing.multi_file_loader import MultiFileLoader
import os
import functools


class MultiFileDialog(QDialog):
    """多文件选择和配置对话框"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.multi_file_loader = MultiFileLoader()
        self.segments_config = []  # 存储数据段配置
        self.setup_ui()
        
    def setup_ui(self):
        """设置对话框UI"""
        self.setWindowTitle("多文件数据组合配置")
        self.setGeometry(200, 200, 800, 600)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        
        # 文件选择区域
        self.create_file_selection_section(main_layout)
        
        # 文件信息展示区域
        self.create_file_info_section(main_layout)
        
        # 数据段配置区域
        self.create_segments_config_section(main_layout)
        
        # 按钮区域
        self.create_buttons_section(main_layout)
        
    def create_file_selection_section(self, parent_layout):
        """创建文件选择区域"""
        file_group = QGroupBox("文件选择")
        file_group.setFont(QFont("Arial", 12))
        file_layout = QVBoxLayout(file_group)
        
        # 选择文件按钮
        button_layout = QHBoxLayout()
        
        self.select_files_btn = QPushButton("选择多个Excel文件")
        self.select_files_btn.setFont(QFont("Arial", 11))
        self.select_files_btn.setMinimumHeight(32)
        self.select_files_btn.clicked.connect(self.select_files)
        
        self.clear_files_btn = QPushButton("清空文件")
        self.clear_files_btn.setFont(QFont("Arial", 11))
        self.clear_files_btn.setMinimumHeight(32)
        self.clear_files_btn.clicked.connect(self.clear_files)
        
        button_layout.addWidget(self.select_files_btn)
        button_layout.addWidget(self.clear_files_btn)
        button_layout.addStretch()
        
        file_layout.addLayout(button_layout)
        
        # 已选文件列表
        self.files_list = QListWidget()
        self.files_list.setFont(QFont("Arial", 10))
        self.files_list.setMaximumHeight(100)
        file_layout.addWidget(self.files_list)
        
        parent_layout.addWidget(file_group)
        
    def create_file_info_section(self, parent_layout):
        """创建文件信息展示区域"""
        info_group = QGroupBox("文件信息")
        info_group.setFont(QFont("Arial", 12))
        info_layout = QVBoxLayout(info_group)
        
        # 文件详细信息表格
        self.file_info_table = QTableWidget()
        self.file_info_table.setFont(QFont("Arial", 10))
        self.file_info_table.setColumnCount(3)
        self.file_info_table.setHorizontalHeaderLabels(["文件名", "列数", "可用列"])
        
        # 设置表格属性
        header = self.file_info_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        
        self.file_info_table.setMaximumHeight(150)
        self.file_info_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        
        info_layout.addWidget(self.file_info_table)
        parent_layout.addWidget(info_group)
        
    def create_segments_config_section(self, parent_layout):
        """创建数据段配置区域"""
        config_group = QGroupBox("数据段配置")
        config_group.setFont(QFont("Arial", 12))
        config_layout = QVBoxLayout(config_group)
        
        # 配置说明
        instruction_label = QLabel("配置要组合的数据段（按顺序添加）：")
        instruction_label.setFont(QFont("Arial", 10))
        config_layout.addWidget(instruction_label)
        
        # 时间单位选项
        time_unit_layout = QHBoxLayout()
        time_unit_label = QLabel("时间显示单位：")
        time_unit_label.setFont(QFont("Arial", 10))
        time_unit_layout.addWidget(time_unit_label)
        
        self.time_unit_combo = QComboBox()
        self.time_unit_combo.setFont(QFont("Arial", 10))
        self.time_unit_combo.addItem("小时 (h)", "hours")
        self.time_unit_combo.addItem("分钟 (min)", "minutes")
        self.time_unit_combo.setToolTip("选择横坐标时间的显示单位")
        time_unit_layout.addWidget(self.time_unit_combo)
        
        time_unit_layout.addStretch()
        config_layout.addLayout(time_unit_layout)
        
        # 添加数据段按钮和刷新按钮
        button_row = QHBoxLayout()
        
        add_segment_btn = QPushButton("添加数据段")
        add_segment_btn.setFont(QFont("Arial", 11))
        add_segment_btn.setMinimumHeight(32)
        add_segment_btn.clicked.connect(self.add_segment)
        button_row.addWidget(add_segment_btn)
        
        refresh_btn = QPushButton("刷新列选择")
        refresh_btn.setFont(QFont("Arial", 11))
        refresh_btn.setMinimumHeight(32)
        refresh_btn.clicked.connect(self.refresh_all_columns)
        button_row.addWidget(refresh_btn)
        
        button_row.addStretch()
        config_layout.addLayout(button_row)
        
        # 数据段配置表格
        self.segments_table = QTableWidget()
        self.segments_table.setFont(QFont("Arial", 10))
        self.segments_table.setColumnCount(9)
        self.segments_table.setHorizontalHeaderLabels([
            "顺序", "文件", "列名", "时间列", "标注名称", "开始时间(h)", "结束时间(h)", "时间操作", "删除"
        ])
        
        # 设置表格属性
        header = self.segments_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.setSectionResizeMode(4, QHeaderView.Interactive)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.setSectionResizeMode(7, QHeaderView.Fixed)
        header.setSectionResizeMode(8, QHeaderView.Fixed)
        
        self.segments_table.setColumnWidth(0, 60)
        self.segments_table.setColumnWidth(4, 120)  # 标注名称列
        self.segments_table.setColumnWidth(5, 100)  # 开始时间
        self.segments_table.setColumnWidth(6, 100)  # 结束时间
        self.segments_table.setColumnWidth(7, 80)   # 时间操作
        self.segments_table.setColumnWidth(8, 80)   # 删除
        
        config_layout.addWidget(self.segments_table)
        parent_layout.addWidget(config_group)
        
    def create_buttons_section(self, parent_layout):
        """创建按钮区域"""
        button_layout = QHBoxLayout()
        
        self.preview_btn = QPushButton("预览组合结果")
        self.preview_btn.setFont(QFont("Arial", 11))
        self.preview_btn.setMinimumHeight(32)
        self.preview_btn.clicked.connect(self.preview_combination)
        
        self.ok_btn = QPushButton("确定")
        self.ok_btn.setFont(QFont("Arial", 11, QFont.Bold))
        self.ok_btn.setMinimumHeight(32)
        self.ok_btn.clicked.connect(self.accept)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setFont(QFont("Arial", 11))
        self.cancel_btn.setMinimumHeight(32)
        self.cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(self.preview_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        
        parent_layout.addLayout(button_layout)
        
    def select_files(self):
        """选择多个Excel文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择多个Excel文件", "", 
            "Excel files (*.xlsx *.xls)"
        )
        
        if file_paths:
            try:
                # 加载文件
                self.multi_file_loader.load_files(file_paths)
                
                # 更新UI
                self.update_files_list()
                self.update_file_info_table()
                
                QMessageBox.information(self, "成功", f"成功加载 {len(file_paths)} 个文件！")
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载文件失败: {str(e)}")
                
    def clear_files(self):
        """清空已选文件"""
        self.multi_file_loader = MultiFileLoader()
        self.segments_config = []
        self.update_files_list()
        self.update_file_info_table()
        self.update_segments_table()
        
    def update_files_list(self):
        """更新文件列表显示"""
        self.files_list.clear()
        
        if self.multi_file_loader.is_data_loaded():
            files_info = self.multi_file_loader.get_all_files_info()
            for file_info in files_info:
                self.files_list.addItem(f"{file_info['index']}: {file_info['filename']}")
                
    def update_file_info_table(self):
        """更新文件信息表格"""
        self.file_info_table.setRowCount(0)
        
        if self.multi_file_loader.is_data_loaded():
            files_info = self.multi_file_loader.get_all_files_info()
            self.file_info_table.setRowCount(len(files_info))
            
            for i, file_info in enumerate(files_info):
                # 文件名
                self.file_info_table.setItem(i, 0, QTableWidgetItem(file_info['filename']))
                
                # 列数
                col_count = len(file_info['columns'])
                self.file_info_table.setItem(i, 1, QTableWidgetItem(str(col_count)))
                
                # 可用列（显示前几个列名）
                columns_text = ", ".join(file_info['columns'][:5])
                if len(file_info['columns']) > 5:
                    columns_text += "..."
                self.file_info_table.setItem(i, 2, QTableWidgetItem(columns_text))
                
    def add_segment(self):
        """添加数据段配置"""
        if not self.multi_file_loader.is_data_loaded():
            QMessageBox.warning(self, "警告", "请先选择Excel文件！")
            return
            
        # 添加新行
        row_count = self.segments_table.rowCount()
        self.segments_table.setRowCount(row_count + 1)
        
        # 顺序号
        self.segments_table.setItem(row_count, 0, QTableWidgetItem(str(row_count + 1)))
        
        # 文件选择下拉框
        file_combo = QComboBox()
        files_info = self.multi_file_loader.get_all_files_info()
        for file_info in files_info:
            file_combo.addItem(f"{file_info['index']}: {file_info['filename']}", file_info['index'])
        # 使用functools.partial来正确绑定行索引
        file_combo.currentIndexChanged.connect(functools.partial(self.update_column_combo, row_count))
        self.segments_table.setCellWidget(row_count, 1, file_combo)
        
        # 列名选择下拉框
        column_combo = QComboBox()
        self.segments_table.setCellWidget(row_count, 2, column_combo)
        
        # 时间列选择下拉框
        time_combo = QComboBox()
        self.segments_table.setCellWidget(row_count, 3, time_combo)
        
        # 标注名称输入框
        label_input = QLineEdit()
        label_input.setPlaceholderText("输入数据标注名称...")
        # 设置默认标注名称
        if files_info:
            filename = files_info[0]['filename']
            default_label = f"数据段{row_count + 1}"
            label_input.setText(default_label)
        self.segments_table.setCellWidget(row_count, 4, label_input)
        
        # 开始时间
        start_time_spin = QDoubleSpinBox()
        start_time_spin.setRange(0, 9999)
        start_time_spin.setValue(0)
        start_time_spin.setSuffix(" h")
        self.segments_table.setCellWidget(row_count, 5, start_time_spin)
        
        # 结束时间
        end_time_spin = QDoubleSpinBox()
        end_time_spin.setRange(0, 9999)
        end_time_spin.setValue(2)
        end_time_spin.setSuffix(" h")
        self.segments_table.setCellWidget(row_count, 6, end_time_spin)
        
        # 全选时间按钮
        full_time_btn = QPushButton("全选时间")
        full_time_btn.setFont(QFont("Arial", 9))
        full_time_btn.clicked.connect(functools.partial(self.set_full_time_range, row_count))
        self.segments_table.setCellWidget(row_count, 7, full_time_btn)
        
        # 删除按钮
        delete_btn = QPushButton("删除")
        delete_btn.setFont(QFont("Arial", 9))
        delete_btn.clicked.connect(functools.partial(self.delete_segment, row_count))
        self.segments_table.setCellWidget(row_count, 8, delete_btn)
        
        # 立即手动触发列选择更新
        if files_info:
            file_combo.setCurrentIndex(0)  # 设置默认选择第一个文件
            # 立即手动更新列选择
            self.manual_update_column_combo(row_count)
            
            # 也设置延迟更新作为备份
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(100, lambda: self.manual_update_column_combo(row_count))
        
    def update_column_combo(self, row):
        """更新指定行的列选择下拉框"""
        try:
            file_combo = self.segments_table.cellWidget(row, 1)
            column_combo = self.segments_table.cellWidget(row, 2)
            time_combo = self.segments_table.cellWidget(row, 3)
            
            if file_combo and column_combo and time_combo:
                file_index = file_combo.currentData()
                print(f"Debug: Row {row}, file_index = {file_index}")
                
                if file_index is not None:
                    columns = self.multi_file_loader.get_file_columns(file_index)
                    print(f"Debug: Found {len(columns)} columns: {columns}")
                    
                    # 更新列选择
                    column_combo.clear()
                    column_combo.addItems(columns)
                    
                    # 更新时间列选择
                    time_combo.clear()
                    time_combo.addItems(columns)
                    
                    # 自动选择合理的默认值
                    if columns:
                        # 尝试找到温度相关的列
                        temp_keywords = ['temperatur', 'temp', '温度', '°c', 'celsius']
                        time_keywords = ['zeit', 'time', 'timestamp', 'datetime', '时间']
                        
                        # 自动选择温度列
                        for col in columns:
                            col_lower = col.lower()
                            for keyword in temp_keywords:
                                if keyword in col_lower:
                                    column_combo.setCurrentText(col)
                                    print(f"Debug: Auto-selected temperature column: {col}")
                                    break
                            else:
                                continue
                            break
                        
                        # 自动选择时间列
                        for col in columns:
                            col_lower = col.lower()
                            for keyword in time_keywords:
                                if keyword in col_lower:
                                    time_combo.setCurrentText(col)
                                    print(f"Debug: Auto-selected time column: {col}")
                                    break
                            else:
                                continue
                            break
                else:
                    print(f"Debug: file_index is None for row {row}")
            else:
                print(f"Debug: Missing widgets for row {row}")
        except Exception as e:
            print(f"Debug: Error in update_column_combo: {e}")
            QMessageBox.critical(self, "错误", f"更新列选择失败: {str(e)}")
    
    def manual_update_column_combo(self, row):
        """手动更新指定行的列选择下拉框"""
        try:
            # 获取UI组件
            file_combo = self.segments_table.cellWidget(row, 1)
            column_combo = self.segments_table.cellWidget(row, 2)
            time_combo = self.segments_table.cellWidget(row, 3)
            
            # 检查组件是否有效 - 使用更精确的检查
            if file_combo is None or column_combo is None or time_combo is None:
                return
            
            # 检查组件是否有addItems方法
            if not (hasattr(column_combo, 'addItems') and hasattr(time_combo, 'addItems')):
                return
            
            # 获取文件索引
            file_index = file_combo.currentData()
            if file_index is None:
                file_index = file_combo.currentIndex()
            
            if file_index is None or file_index < 0:
                return
            
            # 获取列信息
            columns = self.multi_file_loader.get_file_columns(file_index)
            if not columns:
                return
            
            # 清空并填充下拉框
            column_combo.clear()
            time_combo.clear()
            column_combo.addItems(columns)
            time_combo.addItems(columns)
            
            # 智能选择合适的默认值
            temp_keywords = ['temperatur', 'temp', '温度', '°c', 'celsius']
            time_keywords = ['zeit', 'time', 'timestamp', 'datetime', '时间', 'datum']
            
            # 自动选择温度列
            for col in columns:
                col_lower = col.lower()
                for keyword in temp_keywords:
                    if keyword in col_lower:
                        column_combo.setCurrentText(col)
                        break
                else:
                    continue
                break
            
            # 自动选择时间列
            for col in columns:
                col_lower = col.lower()
                for keyword in time_keywords:
                    if keyword in col_lower:
                        time_combo.setCurrentText(col)
                        break
                else:
                    continue
                break
            
        except Exception as e:
            print(f"Error updating column combo: {e}")
    
    def refresh_all_columns(self):
        """刷新所有行的列选择"""
        try:
            for row in range(self.segments_table.rowCount()):
                self.manual_update_column_combo(row)
            QMessageBox.information(self, "成功", "已刷新所有列选择！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"刷新列选择失败: {str(e)}")
                 
    def set_full_time_range(self, row):
        """设置指定行为全时间范围"""
        try:
            file_combo = self.segments_table.cellWidget(row, 1)
            time_combo = self.segments_table.cellWidget(row, 3)
            start_time_spin = self.segments_table.cellWidget(row, 5)
            end_time_spin = self.segments_table.cellWidget(row, 6)
            
            if all([file_combo, time_combo, start_time_spin, end_time_spin]):
                file_index = file_combo.currentData()
                time_column = time_combo.currentText()
                
                if file_index is not None and time_column:
                    time_range = self.multi_file_loader.get_file_time_range(file_index, time_column)
                    
                    start_time_spin.setValue(time_range['start_time'])
                    end_time_spin.setValue(time_range['end_time'])
                    
                    QMessageBox.information(self, "成功", 
                                          f"已设置为全时间范围：{time_range['start_time']:.2f} - {time_range['end_time']:.2f} 小时")
                else:
                    QMessageBox.warning(self, "警告", "请先选择文件和时间列！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"设置时间范围失败: {str(e)}")
            
    def delete_segment(self, row):
        """删除指定行的数据段"""
        # 确认删除
        reply = QMessageBox.question(self, "确认删除", "确定要删除这个数据段吗？", 
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.segments_table.removeRow(row)
            self.update_segment_order()
            self.rebind_all_signals()
        
    def update_segment_order(self):
        """更新数据段顺序号"""
        for i in range(self.segments_table.rowCount()):
            self.segments_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            
    def rebind_all_signals(self):
        """重新绑定所有信号"""
        for row in range(self.segments_table.rowCount()):
            file_combo = self.segments_table.cellWidget(row, 1)
            full_time_btn = self.segments_table.cellWidget(row, 7)
            delete_btn = self.segments_table.cellWidget(row, 8)
            
            if file_combo:
                # 断开旧连接
                file_combo.currentIndexChanged.disconnect()
                # 重新连接
                file_combo.currentIndexChanged.connect(functools.partial(self.update_column_combo, row))
                
            if full_time_btn:
                # 断开旧连接
                full_time_btn.clicked.disconnect()
                # 重新连接
                full_time_btn.clicked.connect(functools.partial(self.set_full_time_range, row))
                
            if delete_btn:
                # 断开旧连接  
                delete_btn.clicked.disconnect()
                # 重新连接
                delete_btn.clicked.connect(functools.partial(self.delete_segment, row))
            
    def update_segments_table(self):
        """更新数据段表格"""
        self.segments_table.setRowCount(0)
        
    def preview_combination(self):
        """预览组合结果"""
        try:
            segments_config = self.get_segments_config()
            if not segments_config:
                QMessageBox.warning(self, "警告", "请至少添加一个数据段！")
                return
                
            # 组合数据
            combined_data = self.multi_file_loader.combine_data_segments(segments_config)
            
            # 显示预览信息
            info_text = f"组合成功！\n"
            info_text += f"数据点总数: {len(combined_data)}\n"
            info_text += f"时间范围: {combined_data['relative_time'].min():.2f} - {combined_data['relative_time'].max():.2f} 小时\n"
            info_text += f"数据段数: {len(segments_config)}\n\n"
            
            # 显示各段信息
            sources = combined_data['source'].unique()
            for source in sources:
                source_data = combined_data[combined_data['source'] == source]
                info_text += f"{source}: {len(source_data)} 个数据点\n"
                
            QMessageBox.information(self, "预览结果", info_text)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预览失败: {str(e)}")
            
    def get_segments_config(self):
        """获取数据段配置"""
        segments_config = []
        
        for row in range(self.segments_table.rowCount()):
            file_combo = self.segments_table.cellWidget(row, 1)
            column_combo = self.segments_table.cellWidget(row, 2)
            time_combo = self.segments_table.cellWidget(row, 3)
            label_input = self.segments_table.cellWidget(row, 4)
            start_time_spin = self.segments_table.cellWidget(row, 5)
            end_time_spin = self.segments_table.cellWidget(row, 6)
            
            if all([file_combo, column_combo, time_combo, label_input, start_time_spin, end_time_spin]):
                config = {
                    'file_index': file_combo.currentData(),
                    'column': column_combo.currentText(),
                    'time_column': time_combo.currentText(),
                    'label': label_input.text().strip() or f"数据段{row + 1}",  # 如果标注为空，使用默认名称
                    'start_time': start_time_spin.value(),
                    'end_time': end_time_spin.value()
                }
                segments_config.append(config)
                
        return segments_config
        
    def get_multi_file_loader(self):
        """获取多文件加载器"""
        return self.multi_file_loader
        
    def get_combined_data(self):
        """获取组合数据"""
        try:
            segments_config = self.get_segments_config()
            if segments_config:
                time_unit = self.time_unit_combo.currentData()  # "hours" 或 "minutes"
                return self.multi_file_loader.combine_data_segments(segments_config, time_unit)
            return None
        except Exception as e:
            raise e 
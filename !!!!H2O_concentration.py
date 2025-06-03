import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from tkinter import Tk, filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import json
import os

class DataVisualizer:
    def __init__(self):
        # 设置全局字体
        plt.rcParams['font.family'] = 'Arial'
        plt.rcParams['mathtext.fontset'] = 'dejavusans'
        plt.rcParams['axes.unicode_minus'] = False
        plt.rcParams['font.weight'] = 'bold'
        
        # 创建主窗口
        self.root = Tk()
        self.root.title("Data Visualization and Calibration Tool")
        self.root.geometry("900x700")
        
        # 最近的配对设置保存路径
        self.recent_pairs_file = os.path.join(os.path.expanduser("~"), "h2o_calibration_recent_pairs.json")
        
        # 创建整体容器框架
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建固定顶部框架
        self.top_frame = ttk.Frame(main_container, padding="10")
        self.top_frame.pack(fill=tk.X, side=tk.TOP)
        
        # 文件选择按钮放在固定顶部框架
        ttk.Button(self.top_frame, text="Select Excel File", command=self.select_file).pack(side=tk.LEFT, padx=10)
        self.file_label = ttk.Label(self.top_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # 创建滚动画布区域
        scroll_container = ttk.Frame(main_container)
        scroll_container.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(scroll_container, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # 设置滚动区域响应大小变化
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        # 适应宽度填充
        self.scrollable_frame.bind("<Configure>", 
                                  lambda e: self.canvas.configure(width=e.width))
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 布局滚动区域
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定鼠标滚轮事件
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)  # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)    # Linux
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)    # Linux
        
        # 创建主框架 - 放在滚动画布内
        self.main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建列表框用于显示列名，带有滚动条
        list_frame = ttk.Frame(self.main_frame)
        list_frame.grid(row=0, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(list_frame, text="Select Columns to Plot:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        
        # 添加操作提示标签
        tips_label = ttk.Label(list_frame, 
                              text="(Click to select, Double-click to toggle, Ctrl+Click for multiple)", 
                              font=('Arial', 8, 'italic'))
        tips_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        # 添加取消选择按钮
        ttk.Button(list_frame, text="Clear Selection", 
                  command=self.clear_selection).grid(row=0, column=2, padx=5, pady=2, sticky=tk.E)
        
        # 列表框和滚动条
        list_subframe = ttk.Frame(list_frame)
        list_subframe.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.column_listbox = tk.Listbox(list_subframe, selectmode=tk.MULTIPLE, width=50, height=8)
        self.column_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        list_scrollbar = ttk.Scrollbar(list_subframe, orient=tk.VERTICAL, command=self.column_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.column_listbox.config(yscrollcommand=list_scrollbar.set)
        
        # 绑定双击事件，用于取消选择
        self.column_listbox.bind('<Double-1>', self.toggle_selection)
        
        # 时间列选择
        ttk.Label(self.main_frame, text="Select Time Column:").grid(row=1, column=0, sticky=tk.W)
        self.time_column_var = tk.StringVar()
        self.time_column_combo = ttk.Combobox(self.main_frame, textvariable=self.time_column_var)
        self.time_column_combo.grid(row=1, column=1, pady=5, sticky=tk.W)
        
        # 校准参数框架
        self.calib_frame = ttk.LabelFrame(self.main_frame, text="Moisture Calibration Settings")
        self.calib_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # 参考列和差值计算设置 - 移动到更靠前的位置
        self.diff_frame = ttk.LabelFrame(self.main_frame, text="Difference Calculation")
        self.diff_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # 创建左侧的复选框框架
        self.diff_options_frame = ttk.Frame(self.diff_frame)
        self.diff_options_frame.grid(row=0, column=0, rowspan=3, padx=5, pady=5, sticky=(tk.W, tk.N))
        
        # 启用差值计算 - 放在左侧框架中
        self.enable_diff_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.diff_options_frame, text="Calculate Difference at 30min", 
                       variable=self.enable_diff_var).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 启用多时间点差值计算 - 放在左侧框架中，更加明显
        self.enable_multi_time_diff_var = tk.BooleanVar(value=False)
        
        # 应用样式让它更明显
        style = ttk.Style()
        style.configure('Bold.TCheckbutton', font=('Arial', 10, 'bold'))
        
        # 使用较大字体和加粗显示，使用英文
        multi_time_check = ttk.Checkbutton(self.diff_options_frame, 
                                         text="Calculate at 20/40/60min", 
                                         style='Bold.TCheckbutton',
                                         variable=self.enable_multi_time_diff_var)
        multi_time_check.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 添加说明文本标签，英文
        explanation_text = "Automatically calculate the difference at 20min, 40min and 60min, and calculate the average"
        ttk.Label(self.diff_options_frame, text=explanation_text, 
                wraplength=200, justify=tk.LEFT, 
                font=('Arial', 8, 'italic')).grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        
        # 创建右侧的设置框架
        self.diff_settings_frame = ttk.Frame(self.diff_frame)
        self.diff_settings_frame.grid(row=0, column=1, rowspan=3, padx=5, pady=5, sticky=(tk.W, tk.N))
        
        # 参考列选择 - 放在右侧框架中
        ttk.Label(self.diff_settings_frame, text="Reference Column:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.reference_col_var = tk.StringVar()
        self.reference_combo = ttk.Combobox(self.diff_settings_frame, textvariable=self.reference_col_var, width=15)
        self.reference_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        # 新增第二个reference选择
        ttk.Label(self.diff_settings_frame, text="Second Reference:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.reference_col2_var = tk.StringVar()
        self.reference_combo2 = ttk.Combobox(self.diff_settings_frame, textvariable=self.reference_col2_var, width=15)
        self.reference_combo2.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        
        # 添加"从选择中设置"按钮
        ttk.Button(self.diff_settings_frame, text="Set from Selection", 
                  command=self.set_reference_from_selection).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 时间窗口设置 - 放在右侧框架中
        ttk.Label(self.diff_settings_frame, text="Time Window (min):").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.time_window_var = tk.DoubleVar(value=5.0)
        ttk.Entry(self.diff_settings_frame, textvariable=self.time_window_var, width=6).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 使用校准功能的复选框
        self.use_calibration_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.calib_frame, text="Enable Moisture Calibration", 
                      variable=self.use_calibration_var,
                      command=self.toggle_calibration).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 显示原始数据的复选框
        self.show_original_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.calib_frame, text="Show Original Data", 
                      variable=self.show_original_var).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 误差显示设置
        self.show_error_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.calib_frame, text="Show Error Range", 
                      variable=self.show_error_var).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 误差值设置
        ttk.Label(self.calib_frame, text="Error Value (ppm):").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.error_value_var = tk.DoubleVar(value=10.0)
        self.error_entry = ttk.Entry(self.calib_frame, textvariable=self.error_value_var, width=6)
        self.error_entry.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        
        # 校准组框架（用于湿度-压力对的选择）
        self.calib_pairs_frame = ttk.LabelFrame(self.calib_frame, text="Moisture-Pressure Pairs")
        self.calib_pairs_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # 创建湿度-压力对选择的框架
        self.pairs_frame = ttk.Frame(self.calib_pairs_frame)
        self.pairs_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 表头
        ttk.Label(self.pairs_frame, text="No.", width=5).grid(row=0, column=0, padx=2, pady=2)
        ttk.Label(self.pairs_frame, text="Moisture Column", width=15).grid(row=0, column=1, padx=2, pady=2)
        ttk.Label(self.pairs_frame, text="Pressure Column", width=15).grid(row=0, column=2, padx=2, pady=2)
        ttk.Label(self.pairs_frame, text="Action", width=15).grid(row=0, column=3, padx=2, pady=2)
        
        # 湿度-压力对的最大数量
        self.max_pairs = 6
        self.pair_vars = []
        
        # 创建多组湿度-压力选择对
        for i in range(self.max_pairs):
            ttk.Label(self.pairs_frame, text=f"{i+1}", width=5).grid(row=i+1, column=0, padx=2, pady=2)
            
            # 湿度列选择
            moisture_var = tk.StringVar()
            moisture_combo = ttk.Combobox(self.pairs_frame, textvariable=moisture_var, width=15, state='disabled')
            moisture_combo.grid(row=i+1, column=1, padx=2, pady=2)
            
            # 压力列选择
            pressure_var = tk.StringVar()
            pressure_combo = ttk.Combobox(self.pairs_frame, textvariable=pressure_var, width=15, state='disabled')
            pressure_combo.grid(row=i+1, column=2, padx=2, pady=2)
            
            # 添加快速选择按钮
            action_frame = ttk.Frame(self.pairs_frame)
            action_frame.grid(row=i+1, column=3, padx=2, pady=2)
            
            # 从选择设置按钮
            ttk.Button(action_frame, text="Set", width=4, 
                     command=lambda idx=i: self.set_pair_from_selection(idx)).pack(side=tk.LEFT, padx=2)
            
            # 清除按钮
            ttk.Button(action_frame, text="Clear", width=5,
                     command=lambda idx=i: self.clear_pair(idx)).pack(side=tk.LEFT, padx=2)
            
            self.pair_vars.append((moisture_var, pressure_var, moisture_combo, pressure_combo))
        
        # 添加自动匹配按钮
        auto_match_frame = ttk.Frame(self.calib_pairs_frame)
        auto_match_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(auto_match_frame, text="Auto-Match Pairs", 
                 command=self.auto_match_pairs).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(auto_match_frame, text="Clear All Pairs", 
                 command=self.clear_all_pairs).pack(side=tk.LEFT, padx=5)
        
        # 添加保存/加载配对按钮
        ttk.Button(auto_match_frame, text="Save Pairs", 
                 command=self.save_pairs).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(auto_match_frame, text="Load Pairs", 
                 command=self.load_pairs).pack(side=tk.LEFT, padx=5)
        
        # 校准参数
        ttk.Label(self.calib_frame, text="f1:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.f1_var = tk.DoubleVar(value=0.196798)
        self.f1_entry = ttk.Entry(self.calib_frame, textvariable=self.f1_var, width=10, state='disabled')
        self.f1_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.calib_frame, text="f2:").grid(row=3, column=2, padx=5, pady=5, sticky=tk.W)
        self.f2_var = tk.DoubleVar(value=0.419073)
        self.f2_entry = ttk.Entry(self.calib_frame, textvariable=self.f2_var, width=10, state='disabled')
        self.f2_entry.grid(row=3, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.calib_frame, text="p_ref:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.p_ref_var = tk.DoubleVar(value=1.0)
        self.p_ref_entry = ttk.Entry(self.calib_frame, textvariable=self.p_ref_var, width=10, state='disabled')
        self.p_ref_entry.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 参数导入导出按钮
        self.param_buttons_frame = ttk.Frame(self.calib_frame)
        self.param_buttons_frame.grid(row=5, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)
        
        ttk.Button(self.param_buttons_frame, text="Import Parameters from File", 
                  command=self.import_parameters).grid(row=0, column=0, padx=10, pady=5)
        
        ttk.Button(self.param_buttons_frame, text="Export Parameters to File", 
                  command=self.export_parameters).grid(row=0, column=1, padx=10, pady=5)
        
        # 时间范围设置
        ttk.Label(self.main_frame, text="Time Range Settings:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.time_range_frame = ttk.Frame(self.main_frame)
        self.time_range_frame.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.time_range_var = tk.StringVar(value="2hours")
        ttk.Radiobutton(self.time_range_frame, text="Show All Data", 
                       variable=self.time_range_var, value="all").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(self.time_range_frame, text="First 2 Hours Only", 
                       variable=self.time_range_var, value="2hours").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(self.time_range_frame, text="Custom Range", 
                       variable=self.time_range_var, value="custom").grid(row=0, column=2, padx=5)
        
        self.custom_range_frame = ttk.Frame(self.main_frame)
        self.custom_range_frame.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Label(self.custom_range_frame, text="Start Time (h):").grid(row=0, column=0, padx=5)
        self.start_time_var = tk.DoubleVar(value=0.0)
        ttk.Entry(self.custom_range_frame, textvariable=self.start_time_var, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(self.custom_range_frame, text="End Time (h):").grid(row=0, column=2, padx=5)
        self.end_time_var = tk.DoubleVar(value=2.0)
        ttk.Entry(self.custom_range_frame, textvariable=self.end_time_var, width=8).grid(row=0, column=3, padx=5)
        
        # 绘图按钮
        ttk.Button(self.main_frame, text="Generate Chart", command=self.plot_data).grid(row=7, column=0, pady=10)
        ttk.Button(self.main_frame, text="Export Data", command=self.export_data).grid(row=7, column=1, pady=10)
        
        # 初始化变量
        self.df = None
        self.file_path = None

    def toggle_calibration(self):
        """Enable or disable calibration related widgets"""
        state = 'normal' if self.use_calibration_var.get() else 'disabled'
        self.f1_entry.config(state=state)
        self.f2_entry.config(state=state)
        self.p_ref_entry.config(state=state)
        
        # 更新所有湿度-压力对的状态
        for _, _, moisture_combo, pressure_combo in self.pair_vars:
            try:
                moisture_combo.config(state=state)
                pressure_combo.config(state=state)
            except Exception as e:
                print(f"Error updating moisture-pressure pair state: {str(e)}")

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_path:
            try:
                self.df = pd.read_excel(self.file_path)
                # 预处理所有非时间列为数值
                for col in self.df.columns:
                    # 尝试转换为数值，保留原始列名
                    try:
                        self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                    except:
                        pass  # 忽略无法转换的列
                
                # 更新文件标签
                self.file_label.config(text=os.path.basename(self.file_path))
                
                self.update_column_list()
                
                # 尝试加载与此文件关联的最近配对
                self.load_recent_pairs()
                
                messagebox.showinfo("Success", "File loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Error loading file: {str(e)}")

    def update_column_list(self):
        if self.df is not None:
            # 清空列表框
            self.column_listbox.delete(0, tk.END)
            column_values = []
            
            # 添加列名到列表框和下拉框
            for col in self.df.columns:
                self.column_listbox.insert(tk.END, col)
                column_values.append(col)
            
            # 更新所有下拉框的选项
            self.time_column_combo['values'] = column_values
            self.reference_combo['values'] = column_values
            self.reference_combo2['values'] = column_values  # 新增，保证第二个reference可选
            
            # 更新所有湿度-压力对下拉框的选项
            for _, _, moisture_combo, pressure_combo in self.pair_vars:
                if moisture_combo['state'] != 'disabled':
                    moisture_combo['values'] = column_values
                if pressure_combo['state'] != 'disabled':
                    pressure_combo['values'] = column_values
            
            # 默认选择第一列作为时间列
            if len(self.df.columns) > 0:
                self.time_column_var.set(self.df.columns[0])

    def auto_match_pairs(self):
        """自动匹配湿度-压力对"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a file first")
            return
        
        columns = self.df.columns.tolist()
        
        # 清除现有对
        self.clear_all_pairs()
        
        # 更精确的关键词匹配
        moisture_keywords = ['h2o', 'water', 'humid', 'moisture', 'moistre', 'ppm', 'ppb']
        pressure_keywords = ['pressure', 'press', 'bar', 'pa', 'mpa']
        
        # 找到可能的湿度和压力列
        moisture_columns = []
        pressure_columns = []
        ambiguous_columns = []  # 可能同时包含湿度和压力关键词的列
        
        for col in columns:
            col_lower = str(col).lower()
            
            # 计算湿度关键词匹配分数
            moisture_score = 0
            for kw in moisture_keywords:
                if kw in col_lower:
                    moisture_score += 1
            
            # 计算压力关键词匹配分数
            pressure_score = 0
            for kw in pressure_keywords:
                if kw in col_lower:
                    pressure_score += 1
            
            # 根据得分对列进行分类
            if moisture_score > 0 and pressure_score > 0:
                if moisture_score > pressure_score:
                    moisture_columns.append(col)
                elif pressure_score > moisture_score:
                    pressure_columns.append(col)
                else:
                    ambiguous_columns.append(col)
            elif moisture_score > 0:
                moisture_columns.append(col)
            elif pressure_score > 0:
                pressure_columns.append(col)
        
        # 尝试匹配湿度和压力列（通过名称中的数字或位置）
        pairs = []
        
        # 首先尝试通过数字匹配
        for moisture_col in moisture_columns[:]:
            for pressure_col in pressure_columns[:]:
                # 从列名中提取数字
                moisture_digits = ''.join([c for c in str(moisture_col) if c.isdigit()])
                pressure_digits = ''.join([c for c in str(pressure_col) if c.isdigit()])
                
                # 如果两个列名中包含相同的数字，它们可能属于同一组
                if moisture_digits and moisture_digits == pressure_digits:
                    pairs.append((moisture_col, pressure_col))
                    if moisture_col in moisture_columns:
                        moisture_columns.remove(moisture_col)
                    if pressure_col in pressure_columns:
                        pressure_columns.remove(pressure_col)
                    break
        
        # 如果仍有未匹配的列，尝试一对一匹配
        remaining_pairs = min(len(moisture_columns), len(pressure_columns))
        for i in range(remaining_pairs):
            pairs.append((moisture_columns[i], pressure_columns[i]))
        
        # 处理剩下的列（可能需要进一步猜测）
        remaining_moisture = moisture_columns[remaining_pairs:]
        remaining_pressure = pressure_columns[remaining_pairs:]
        
        # 尝试利用其他未分类的列
        other_columns = [col for col in columns if col not in moisture_columns and col not in pressure_columns and col not in ambiguous_columns]
        
        # 如果还有未匹配的湿度列，尝试与其他列配对
        for moisture_col in remaining_moisture:
            if other_columns:
                pressure_col = other_columns.pop(0)
                pairs.append((moisture_col, pressure_col))
        
        # 如果还有未匹配的压力列，尝试与其他列配对
        for pressure_col in remaining_pressure:
            if other_columns:
                moisture_col = other_columns.pop(0)
                pairs.append((moisture_col, pressure_col))
        
        # 将匹配的对设置到界面上
        for i, (moisture_col, pressure_col) in enumerate(pairs):
            if i < self.max_pairs:
                moisture_var, pressure_var, _, _ = self.pair_vars[i]
                moisture_var.set(moisture_col)
                pressure_var.set(pressure_col)
        
        # 如果有任何配对被设置，保存最近的配对
        if pairs:
            self.save_recent_pairs()
            messagebox.showinfo("Auto-Match", f"Successfully matched {len(pairs)} moisture-pressure pairs")
        else:
            messagebox.showwarning("Auto-Match", "No valid moisture-pressure pairs found. Try selecting them manually.")
    
    def clear_all_pairs(self):
        """清除所有湿度-压力对"""
        for moisture_var, pressure_var, _, _ in self.pair_vars:
            moisture_var.set("")
            pressure_var.set("")
        
        # 如果有文件路径，也清除保存的记录
        if hasattr(self, 'file_path') and self.file_path:
            try:
                if os.path.exists(self.recent_pairs_file):
                    with open(self.recent_pairs_file, 'r') as f:
                        pairs_dict = json.load(f)
                    
                    # 检查当前文件是否有保存的配对
                    file_key = os.path.basename(self.file_path)
                    if file_key in pairs_dict:
                        # 清除此文件的配对
                        pairs_dict[file_key]["pairs"] = {}
                        
                        # 保存回文件
                        with open(self.recent_pairs_file, 'w') as f:
                            json.dump(pairs_dict, f, indent=4)
                        
                        print(f"已清除文件 {file_key} 的保存配对")
            except Exception as e:
                print(f"清除保存的配对记录失败: {str(e)}")

    def calibrate(self, ppm, p, f1, f2, p_ref):
        """Calculate calibrated ppm value based on calibration formula
        ppm_calibrated = concentration × (p_ref/p)^(f1·ln(p_ref/p)+f2)
        """
        # 完全采用与Moisture sensor correction.py相同的实现
        ratio = p_ref / p
        exponent = f1 * np.log(ratio) + f2
        
        # 限制指数范围，避免数值爆炸
        exponent = np.clip(exponent, -10, 10)
        
        # 直接返回计算结果
        return ppm * (ratio ** exponent)
    
    def import_parameters(self):
        """导入校准参数"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select Parameter File",
                filetypes=[("JSON files", "*.json")]
            )
            
            if not file_path:
                return
            
            with open(file_path, 'r') as f:
                params = json.load(f)
            
            # 验证参数格式
            required_keys = ['f1', 'f2', 'p_ref']
            if not all(key in params for key in required_keys):
                messagebox.showerror("Error", "Invalid parameter file format")
                return
            
            # 更新参数
            self.f1_var.set(params['f1'])
            self.f2_var.set(params['f2'])
            self.p_ref_var.set(params['p_ref'])
            
            # 显示参数详情
            details = f"Imported parameters:\n\nf1 = {params['f1']:.6f}\nf2 = {params['f2']:.6f}\np_ref = {params['p_ref']:.2f}"
            if 'date' in params:
                details += f"\n\nCreated date: {params['date']}"
                
            messagebox.showinfo("Parameters Imported", details)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import parameters: {str(e)}")

    def export_parameters(self):
        """导出校准参数"""
        try:
            # 获取当前参数
            f1_value = self.f1_var.get()
            f2_value = self.f2_var.get()
            p_ref_value = self.p_ref_var.get()
            
            # 创建参数字典
            params = {
                'f1': f1_value,
                'f2': f2_value,
                'p_ref': p_ref_value,
                'date': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 请求保存位置
            file_path = filedialog.asksaveasfilename(
                title="Save Parameters",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialfile="moisture_calibration_params.json"
            )
            
            if not file_path:
                return
            
            # 保存参数到文件
            with open(file_path, 'w') as f:
                json.dump(params, f, indent=4)
            
            messagebox.showinfo("Success", f"Parameters saved to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export parameters: {str(e)}")

    def calculate_difference_at_30min(self, plot_df, calibrated_data, ref_col):
        """Calculate the difference between calibrated data and reference at 30min mark"""
        # 目标时间：30分钟 = 0.5小时
        target_time = 0.5
        window_minutes = self.time_window_var.get() / 60  # 从分钟转为小时
        
        # 检查参考列是否有对应的校准数据
        ref_col_calib = f"{ref_col}_calib"
        has_calibrated_ref = self.use_calibration_var.get() and ref_col_calib in calibrated_data
        
        # 如果有校准后的参考数据，使用校准后的值
        if has_calibrated_ref:
            # 使用校准后的参考数据
            ref_times = calibrated_data[ref_col_calib]['times']
            ref_values = calibrated_data[ref_col_calib]['values']
            
            # 过滤在窗口内的数据点
            ref_indices = [i for i, t in enumerate(ref_times) if 
                         (t >= (target_time - window_minutes) and 
                          t <= (target_time + window_minutes))]
            
            # 确保有参考数据
            if not ref_indices:
                return f"No calibrated reference data found for {ref_col} at 30min ±{self.time_window_var.get():.1f}min"
            
            # 计算校准后参考值的平均值
            ref_window_values = [ref_values[i] for i in ref_indices]
            ref_value = np.mean(ref_window_values)
            ref_type = "calibrated"
        else:
            # 使用原始参考数据
            # 在时间窗口内查找数据点
            time_mask = (plot_df['relative_time'] >= (target_time - window_minutes)) & \
                        (plot_df['relative_time'] <= (target_time + window_minutes))
            
            # 获取窗口内的参考列数据
            ref_data = plot_df.loc[time_mask, ref_col]
            
            # 确保有参考数据
            if ref_data.empty:
                return f"No reference data found for {ref_col} at 30min ±{self.time_window_var.get():.1f}min"
            
            # 计算参考值的平均值(在30分钟左右的窗口内)
            ref_value = ref_data.mean()
            ref_type = "original"
        
        # 初始化结果变量
        results = []
        total_difference = 0.0
        valid_difference_count = 0
        
        # 报告参考值类型和值
        results.append(f"Reference Column: {ref_col} ({ref_type})")
        results.append(f"Reference Value at 30min: {ref_value:.2f} ppm")
        results.append("-" * 30)
        
        # 对每个校准后的数据集计算差值
        for data_key, data_dict in calibrated_data.items():
            # 跳过参考列自身的校准数据
            if data_key == ref_col_calib:
                continue
                
            # 使用时间作为索引
            times = data_dict['times']
            values = data_dict['values']
            column_name = data_dict['column']
            
            # 过滤在窗口内的数据点
            time_indices = [i for i, t in enumerate(times) if 
                          (t >= (target_time - window_minutes) and 
                           t <= (target_time + window_minutes))]
            
            if not time_indices:
                results.append(f"{column_name}: No data in time window")
                continue
            
            # 计算平均值
            window_values = [values[i] for i in time_indices]
            calib_value = np.mean(window_values)
            
            # 计算差值
            difference = calib_value - ref_value
            
            # 累加差值
            total_difference += difference
            valid_difference_count += 1
            
            # 保存结果
            results.append(f"{column_name} at 30min: {calib_value:.2f} ppm")
            results.append(f"Difference: {difference:.2f} ppm")
            results.append("-" * 20)
        
        # 添加总差值
        if valid_difference_count > 0:
            results.append("\n----- Total -----")
            results.append(f"Total Difference: {total_difference:.2f} ppm")
            results.append(f"Average Difference: {total_difference/valid_difference_count:.2f} ppm")
        
        return "\n".join(results)
        
    def calculate_time_point_differences(self, plot_df, calibrated_data, ref_col, export_to_excel=False, ref_label=None):
        """计算20、40、60分钟和1h、1h30min、2h时与参考值的差值，并计算其平均值，并可导出为Excel。ref_label用于区分reference"""
        # 目标时间点：20、40、60分钟和1h、1h30min、2h（全部转换为小时）
        target_times = [20/60, 40/60, 60/60, 1, 1.5, 2]
        window_minutes = self.time_window_var.get() / 60
        ref_col_calib = f"{ref_col}_calib"
        has_calibrated_ref = self.use_calibration_var.get() and ref_col_calib in calibrated_data
        results = []
        if ref_label is None:
            ref_label = ref_col
        results.append(f"Reference Column: {ref_label}")
        results.append("=" * 40)
        export_rows = []
        all_time_differences = {}
        average_differences = {}
        for data_key, data_dict in calibrated_data.items():
            if data_key == ref_col_calib:
                continue
            column_name = data_dict['column']
            all_time_differences[column_name] = []
            for target_time in target_times:
                if has_calibrated_ref:
                    ref_times = calibrated_data[ref_col_calib]['times']
                    ref_values = calibrated_data[ref_col_calib]['values']
                    ref_indices = [i for i, t in enumerate(ref_times) if (t >= (target_time - window_minutes) and (t <= (target_time + window_minutes)))]
                    if not ref_indices:
                        results.append(f"At {self._format_time_label(target_time)}: No calibrated reference data found")
                        export_rows.append({
                            'Reference Used': ref_label,
                            'Column': column_name,
                            'Time Point': self._format_time_label(target_time),
                            'Reference Value': 'N/A',
                            'Calibrated Value': 'N/A',
                            'Difference': 'N/A',
                        })
                        continue
                    ref_window_values = [ref_values[i] for i in ref_indices]
                    ref_value = np.mean(ref_window_values)
                else:
                    time_mask = (plot_df['relative_time'] >= (target_time - window_minutes)) & (plot_df['relative_time'] <= (target_time + window_minutes))
                    ref_data = plot_df.loc[time_mask, ref_col]
                    if ref_data.empty:
                        results.append(f"At {self._format_time_label(target_time)}: No reference data found")
                        export_rows.append({
                            'Reference Used': ref_label,
                            'Column': column_name,
                            'Time Point': self._format_time_label(target_time),
                            'Reference Value': 'N/A',
                            'Calibrated Value': 'N/A',
                            'Difference': 'N/A',
                        })
                        continue
                    ref_value = ref_data.mean()
                times = data_dict['times']
                values = data_dict['values']
                time_indices = [i for i, t in enumerate(times) if (t >= (target_time - window_minutes) and (t <= (target_time + window_minutes)))]
                if not time_indices:
                    results.append(f"{column_name} at {self._format_time_label(target_time)}: No data in time window")
                    export_rows.append({
                        'Reference Used': ref_label,
                        'Column': column_name,
                        'Time Point': self._format_time_label(target_time),
                        'Reference Value': ref_value if 'ref_value' in locals() else 'N/A',
                        'Calibrated Value': 'N/A',
                        'Difference': 'N/A',
                    })
                    continue
                window_values = [values[i] for i in time_indices]
                calib_value = np.mean(window_values)
                difference = calib_value - ref_value
                all_time_differences[column_name].append(difference)
                if target_time not in average_differences:
                    average_differences[target_time] = []
                average_differences[target_time].append(difference)
                export_rows.append({
                    'Reference Used': ref_label,
                    'Column': column_name,
                    'Time Point': self._format_time_label(target_time),
                    'Reference Value': ref_value,
                    'Calibrated Value': calib_value,
                    'Difference': difference,
                })
            if all_time_differences[column_name]:
                col_avg_diff = np.mean(all_time_differences[column_name])
                results.append(f"{column_name}")
                for i, diff in enumerate(all_time_differences[column_name]):
                    results.append(f"   Difference at {self._format_time_label(target_times[i])}: {diff:.2f} ppm")
                results.append(f"   Average Difference: {col_avg_diff:.2f} ppm")
                results.append("-" * 30)
        results.append("\n=== Time Point Statistics ===")
        total_avg = []
        for time_point in target_times:
            if time_point in average_differences and average_differences[time_point]:
                point_avg = np.mean(average_differences[time_point])
                total_avg.append(point_avg)
                results.append(f"Average Difference at {self._format_time_label(time_point)}: {point_avg:.2f} ppm")
        if total_avg:
            grand_avg = np.mean(total_avg)
            results.append("\n=== Total ===")
            results.append(f"Average Difference across all time points: {grand_avg:.2f} ppm")
        if export_to_excel:
            return export_rows, "\n".join(results)
        return "\n".join(results)

    def _format_time_label(self, time_in_hours):
        """辅助函数：将小时数格式化为友好的标签"""
        minutes = int(round(time_in_hours * 60))
        if minutes == 20:
            return "20min"
        elif minutes == 40:
            return "40min"
        elif minutes == 60:
            return "60min (1h)"
        elif minutes == 90:
            return "90min (1h30min)"
        elif minutes == 120:
            return "120min (2h)"
        elif minutes == 30:
            return "30min"
        else:
            return f"{minutes}min"

    def plot_data(self):
        if self.df is None:
            messagebox.showerror("Error", "Please select a file first!")
            return
        
        time_col = self.time_column_var.get()
        if not time_col:
            messagebox.showerror("Error", "Please select a time column!")
            return
        
        selected_columns = [self.column_listbox.get(i) for i in self.column_listbox.curselection()]
        if not selected_columns:
            messagebox.showerror("Error", "Please select columns to plot!")
            return
        
        # 检查差值计算
        calculate_diff = self.enable_diff_var.get() or self.enable_multi_time_diff_var.get()
        if calculate_diff:
            ref_col = self.reference_col_var.get()
            if not ref_col:
                messagebox.showerror("Error", "Please select a reference column for difference calculation!")
                return
            if ref_col not in selected_columns:
                messagebox.showwarning("Warning", "Reference column should be included in selected columns")
        
        try:
            # 创建工作副本
            plot_df = self.df.copy()
            
            # 处理时间列
            plot_df[time_col] = pd.to_datetime(plot_df[time_col], errors='coerce')
            plot_df = plot_df.dropna(subset=[time_col])
            
            # 计算相对时间
            start_time = plot_df[time_col].min()
            plot_df['relative_time'] = (plot_df[time_col] - start_time).dt.total_seconds() / 3600
            
            # 应用时间范围过滤
            if self.time_range_var.get() == "2hours":
                plot_df = plot_df[plot_df['relative_time'] <= 2]
                print(f"Only showing data from the first 2 hours: {len(plot_df)} rows")
            elif self.time_range_var.get() == "custom":
                plot_df = plot_df[(plot_df['relative_time'] >= self.start_time_var.get()) & 
                             (plot_df['relative_time'] <= self.end_time_var.get())]
                print(f"Showing data from custom time range: {len(plot_df)} rows")
            else:
                print(f"Showing all data: {len(plot_df)} rows")
            
            # 转换所有选中的列为数值类型
            for col in selected_columns:
                if col in plot_df.columns:
                    plot_df[col] = pd.to_numeric(plot_df[col], errors='coerce')
            
            # 创建图表
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # 定义颜色
            normal_colors = plt.cm.tab10.colors  # 常规颜色
            calib_colors = [
                '#ff7f0e',  # 橙色
                '#d62728',  # 红色
                '#9467bd',  # 紫色
                '#8c564b',  # 棕色
                '#e377c2',  # 粉色
                '#bcbd22',  # 黄绿色
                '#17becf',  # 青色
                '#f0027f'   # 品红色
            ]
            
            # 如果启用校准，先计算校准数据
            calibrated_columns = {}  # 存储校准后的列名与原列名的映射
            
            # 用于收集校准后的数据（用于差值计算）
            calibrated_data = {}
            
            if self.use_calibration_var.get():
                # 获取校准参数
                f1 = self.f1_var.get()
                f2 = self.f2_var.get()
                p_ref = self.p_ref_var.get()
                
                # 确保压力列为数值
                for _, pressure_var, _, _ in self.pair_vars:
                    pressure_col = pressure_var.get()
                    if pressure_col and pressure_col in plot_df.columns:
                        plot_df[pressure_col] = pd.to_numeric(plot_df[pressure_col], errors='coerce')
                
                # 为每个湿度列添加校准数据
                for moisture_var, pressure_var, _, _ in self.pair_vars:
                    moisture_col = moisture_var.get()
                    pressure_col = pressure_var.get()
                    
                    if moisture_col and pressure_col and moisture_col in selected_columns:
                        if moisture_col in plot_df.columns and pressure_col in plot_df.columns:
                            # 删除NaN值和非正压力值 - 匹配Moisture程序处理
                            valid_mask = ~(pd.isna(plot_df[moisture_col]) | 
                                         pd.isna(plot_df[pressure_col]) | 
                                         (plot_df[pressure_col] <= 0))
                            
                            valid_df = plot_df[valid_mask].copy()
                            
                            if not valid_df.empty:
                                # 创建校准结果列名
                                calib_col = f"{moisture_col}_calib"
                                
                                # 使用校准公式
                                valid_df[calib_col] = self.calibrate(
                                    valid_df[moisture_col],
                                    valid_df[pressure_col],
                                    f1, f2, p_ref
                                )
                                
                                # 收集可视化所需数据
                                calib_times = valid_df['relative_time'].values
                                calib_values = valid_df[calib_col].values
                                
                                # 保存校准后的数据，用于差值计算和绘图
                                calibrated_data[calib_col] = {
                                    'times': calib_times,
                                    'values': calib_values,
                                    'column': moisture_col,
                                    'valid_df': valid_df
                                }
                                
                                # 记录校准列
                                calibrated_columns[calib_col] = moisture_col
            
            # 绘制所有选中的列（非校准数据）
            for i, col in enumerate(selected_columns):
                if col in plot_df.columns:
                    # 判断是否是湿度列
                    is_moisture = any(col == moisture_var.get() for moisture_var, _, _, _ in self.pair_vars)
                    
                    if is_moisture and self.use_calibration_var.get():
                        # 如果是湿度列并启用了校准，使用特殊颜色
                        color_idx = i % len(normal_colors)
                        
                        # 只有当show_original_var为True时才绘制原始数据
                        if self.show_original_var.get():
                            valid_data = plot_df[['relative_time', col]].dropna()
                            if not valid_data.empty:
                                line_orig = ax.plot(valid_data['relative_time'], valid_data[col],
                                      '-', color=normal_colors[color_idx],
                                      label=f"{col} (original)",
                                      linewidth=2.5,  # 增大线宽
                                      alpha=0.7)  # 增加不透明度
                                
                                # 在线附近添加标签
                                if len(valid_data) > 10:
                                    # 在曲线中间位置添加标签
                                    mid_idx = len(valid_data) // 2
                                    mid_time = valid_data['relative_time'].iloc[mid_idx]
                                    mid_val = valid_data[col].iloc[mid_idx]
                                    ax.annotate(f"{col}",
                                               xy=(mid_time, mid_val),
                                               xytext=(5, 5),
                                               textcoords='offset points',
                                               fontsize=11,  # 增大字体
                                               color=normal_colors[color_idx],
                                               alpha=0.9,
                                               weight='bold',  # 加粗
                                               bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="none", alpha=0.7))
                        
                        # 绘制对应的校准数据
                        calib_col = f"{col}_calib"
                        if calib_col in calibrated_data:
                            times = calibrated_data[calib_col]['times']
                            values = calibrated_data[calib_col]['values']
                            
                            # 使用不同的校准颜色
                            calib_color = calib_colors[color_idx % len(calib_colors)]
                            
                            # 绘制校准数据，使用粗线并添加标记
                            line_calib = ax.plot(times, values,
                                  '--', color=calib_color,
                                  label=f"{col} (calibrated)",
                                  linewidth=3.0,  # 增大线宽
                                  marker='o',        # 添加标记
                                  markersize=5,      # 增大标记
                                  markevery=len(times)//10)  # 每10个点一个标记
                            
                            # 在线附近添加标签
                            if len(times) > 10:
                                # 在曲线中间位置添加标签
                                mid_idx = len(times) // 2
                                mid_time = times[mid_idx]
                                mid_val = values[mid_idx]
                                ax.annotate(f"{col}(calibrated)",
                                           xy=(mid_time, mid_val),
                                           xytext=(5, 5),
                                           textcoords='offset points',
                                           fontsize=11,  # 增大字体
                                           color=calib_color,
                                           weight='bold',
                                           bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="none", alpha=0.8))
                            
                            # 如果启用误差范围显示
                            if self.show_error_var.get():
                                # 获取误差值
                                error_value = self.error_value_var.get()
                                
                                # 计算上下误差范围
                                upper_bound = values + error_value
                                lower_bound = values - error_value
                                
                                # 绘制误差范围区域
                                ax.fill_between(
                                    times, 
                                    lower_bound, 
                                    upper_bound,
                                    color=calib_color,
                                    alpha=0.2,
                                    label=f"{col} Error Range (±{error_value:.2f} ppm)"
                                )
                    else:
                        # 对于非湿度列或未启用校准，使用常规颜色
                        valid_data = plot_df[['relative_time', col]].dropna()
                        if not valid_data.empty:
                            line = ax.plot(valid_data['relative_time'], valid_data[col],
                                  '-', color=normal_colors[i % len(normal_colors)],
                                  label=col,
                                  linewidth=2.5)  # 增大线宽
                            
                            # 在线附近添加标签
                            if len(valid_data) > 10:
                                # 在曲线中间位置添加标签
                                mid_idx = len(valid_data) // 2
                                mid_time = valid_data['relative_time'].iloc[mid_idx]
                                mid_val = valid_data[col].iloc[mid_idx]
                                ax.annotate(f"{col}",
                                           xy=(mid_time, mid_val),
                                           xytext=(5, 5),
                                           textcoords='offset points',
                                           fontsize=11,  # 增大字体
                                           color=normal_colors[i % len(normal_colors)],
                                           weight='bold',  # 加粗
                                           bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="none", alpha=0.7))
            
            # 计算30分钟差值
            diff_results = ""
            if self.enable_diff_var.get() and self.reference_col_var.get() and calibrated_data:
                diff_results = self.calculate_difference_at_30min(
                    plot_df, calibrated_data, self.reference_col_var.get())
            
            # 计算多时间点差值（如果启用） - 新增
            if self.enable_multi_time_diff_var.get():
                ref_col1 = self.reference_col_var.get()
                ref_col2 = self.reference_col2_var.get()
                export_rows_all = []
                multi_diff_result = ""
                if ref_col1:
                    res1 = self.calculate_time_point_differences(plot_df, calibrated_data, ref_col1, export_to_excel=True, ref_label=ref_col1)
                    if isinstance(res1, tuple):
                        export_rows_all.extend(res1[0])
                        multi_diff_result += f"[Reference: {ref_col1}]\n" + res1[1] + "\n\n"
                    else:
                        multi_diff_result += f"[Reference: {ref_col1}]\n" + res1 + "\n\n"
                if ref_col2 and ref_col2 != ref_col1:
                    res2 = self.calculate_time_point_differences(plot_df, calibrated_data, ref_col2, export_to_excel=True, ref_label=ref_col2)
                    if isinstance(res2, tuple):
                        export_rows_all.extend(res2[0])
                        multi_diff_result += f"[Reference: {ref_col2}]\n" + res2[1] + "\n\n"
                    else:
                        multi_diff_result += f"[Reference: {ref_col2}]\n" + res2 + "\n\n"
                # 导出Excel
                if export_rows_all:
                    df_export = pd.DataFrame(export_rows_all)
                    file_path = filedialog.asksaveasfilename(
                        title="Save Difference Analysis",
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")],
                        initialfile="difference_analysis.xlsx"
                    )
                    if file_path:
                        df_export.to_excel(file_path, index=False)
                # 显示多时间点差值计算结果
                multi_diff_window = tk.Toplevel(self.root)
                multi_diff_window.title("Multi-Reference Difference Analysis")
                multi_diff_window.geometry("650x600")
                text_widget = tk.Text(multi_diff_window, wrap=tk.WORD)
                text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
                text_widget.insert(tk.END, multi_diff_result)
                text_widget.config(state=tk.DISABLED)
                tk.Button(multi_diff_window, text="Close", command=multi_diff_window.destroy).pack(pady=10)
            
            # 轴标签和标题设置 - 修改后
            ax.set_xlabel("Time (hours)", fontsize=14, fontweight='bold')
            ax.set_ylabel("Water Concentration (ppm)", fontsize=14, fontweight='bold')
            
            # 移除标题，仅在图例中显示校准参数（如果需要）
            if self.use_calibration_var.get():
                f1 = self.f1_var.get()
                f2 = self.f2_var.get()
                p_ref = self.p_ref_var.get()
                # 可以考虑添加参数信息到图例而不是标题
                # 以下行已注释掉，不再设置标题
                # ax.set_title(title, fontsize=16, fontweight='bold')
            
            # 设置网格
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # 设置图例 - 修改后
            # 设置图例，增大字体和线宽
            legend = ax.legend(loc='best', frameon=True, fontsize=12, markerscale=2)
            
            # 增大图例中线条的粗细
            for line in legend.get_lines():
                line.set_linewidth(3.0)
                
            # 设置刻度，增大标签大小
            ax.tick_params(axis='both', which='major', labelsize=11)
            
            # 如果显示特定时间范围，设置合适的刻度
            if self.time_range_var.get() == "2hours":
                time_ticks = [0, 0.5, 1, 1.5, 2]
                time_labels = ['0', '30 minutes', '1 hour', '1 hour 30 minutes', '2 hours']
                ax.set_xticks(time_ticks)
                ax.set_xticklabels(time_labels)
            
            # 标记30分钟位置（如果启用差值计算）
            if self.enable_diff_var.get():
                ax.axvline(x=0.5, color='gray', linestyle='--', alpha=0.7, linewidth=2.0)  # 增加线宽
                ax.text(0.5, ax.get_ylim()[0], '30min', 
                      horizontalalignment='center', verticalalignment='bottom',
                      fontsize=12, fontweight='bold')  # 增大字体并加粗
            
            # 调整布局
            plt.tight_layout()
            
            # 显示图表
            plt.show()
            
            # 如果有差值结果，显示在消息框中
            if diff_results:
                messagebox.showinfo("30min Difference Results", diff_results)
            
        except Exception as e:
            # 绘图出错详细信息
            import traceback
            print("Error details when plotting:")
            traceback.print_exc()
            messagebox.showerror("Error", f"Error while plotting: {str(e)}")

    def export_data(self):
        """Export original data and calibrated data"""
        if self.df is None:
            messagebox.showerror("Error", "Please select a file first!")
            return
        
        time_col = self.time_column_var.get()
        if not time_col:
            messagebox.showerror("Error", "Please select a time column!")
            return
        
        try:
            # 创建工作副本
            export_df = self.df.copy()
            
            # 预处理：清理所有数据列，确保为数值类型
            for col in export_df.columns:
                if col != time_col:  # 不处理时间列
                    try:
                        export_df[col] = pd.to_numeric(export_df[col], errors='coerce')
                    except Exception as e:
                        print(f"Error converting column '{col}' to numeric: {str(e)}")
            
            # 处理时间列
            export_df[time_col] = pd.to_datetime(export_df[time_col], errors='coerce')
            export_df = export_df.dropna(subset=[time_col])
            
            # 计算相对时间
            start_time = export_df[time_col].min()
            export_df['relative_time'] = (export_df[time_col] - start_time).dt.total_seconds() / 3600
            
            # 应用时间范围过滤
            if self.time_range_var.get() == "2hours":
                export_df = export_df[export_df['relative_time'] <= 2]
                print(f"Exporting data from the first 2 hours: {len(export_df)} rows")
            elif self.time_range_var.get() == "custom":
                export_df = export_df[(export_df['relative_time'] >= self.start_time_var.get()) & 
                                  (export_df['relative_time'] <= self.end_time_var.get())]
                print(f"Exporting data from custom time range: {len(export_df)} rows")
            else:
                print(f"Exporting all data: {len(export_df)} rows")
            
            # 应用校准功能
            if self.use_calibration_var.get():
                # 获取校准参数
                f1 = self.f1_var.get()
                f2 = self.f2_var.get()
                p_ref = self.p_ref_var.get()
                
                # 处理所有有效的湿度-压力对
                for moisture_var, pressure_var, _, _ in self.pair_vars:
                    moisture_col = moisture_var.get()
                    pressure_col = pressure_var.get()
                    
                    if moisture_col and pressure_col:
                        if moisture_col in export_df.columns and pressure_col in export_df.columns:
                            print(f"Calibrated data column: {moisture_col}")
                            
                            # 移除NaN和非正压力值 - 匹配Moisture程序处理方式
                            valid_mask = ~(pd.isna(export_df[moisture_col]) | 
                                         pd.isna(export_df[pressure_col]) | 
                                         (export_df[pressure_col] <= 0))
                            
                            # 仅保留有效数据行
                            valid_df = export_df[valid_mask].copy()
                            
                            if not valid_df.empty:
                                # 添加校准后的湿度列
                                calibrated_col = f"{moisture_col}_calibrated"
                                
                                # 应用校准 - 与Moisture程序完全一致
                                valid_df[calibrated_col] = self.calibrate(
                                    valid_df[moisture_col], 
                                    valid_df[pressure_col], 
                                    f1, f2, p_ref
                                )
                                
                                # 将结果合并回导出数据框
                                export_df = pd.merge(
                                    export_df.drop(columns=[calibrated_col] if calibrated_col in export_df.columns else []),
                                    valid_df[['relative_time', calibrated_col]],
                                    on='relative_time',
                                    how='left'
                                )
                                
                                # 如果需要显示误差范围，添加上下限列
                                if self.show_error_var.get():
                                    error_value = self.error_value_var.get()
                                    export_df[f"{calibrated_col}_upper"] = export_df[calibrated_col] + error_value
                                    export_df[f"{calibrated_col}_lower"] = export_df[calibrated_col] - error_value
            
            # 导出到Excel
            file_path = filedialog.asksaveasfilename(
                title="Export Data",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="data_with_calibration.xlsx"
            )
            
            if not file_path:
                return
            
            export_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Data exported to {file_path}")
            
        except Exception as e:
            import traceback
            print("Error details when exporting data:")
            traceback.print_exc()
            messagebox.showerror("Error", f"Error exporting data: {str(e)}")

    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        try:
            # 区分事件类型
            if hasattr(event, 'num') and event.num in (4, 5):  # Linux
                if event.num == 5:
                    self.canvas.yview_scroll(1, "units")
                elif event.num == 4:
                    self.canvas.yview_scroll(-1, "units")
            elif hasattr(event, 'delta'):  # Windows
                # Windows中delta是正负值
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception as e:
            print(f"鼠标滚轮事件处理错误: {str(e)}")

    def clear_selection(self):
        """清除列表框中的所有选择"""
        self.column_listbox.selection_clear(0, tk.END)
    
    def toggle_selection(self, event):
        """切换选择状态 - 双击取消特定项目的选择"""
        # 获取点击的索引
        index = self.column_listbox.nearest(event.y)
        
        # 如果该项已选中，则取消选择；否则选中
        if index in self.column_listbox.curselection():
            self.column_listbox.selection_clear(index)
        else:
            self.column_listbox.selection_set(index)

    def set_reference_from_selection(self):
        """从列表框的选择中设置参考列"""
        selected = self.column_listbox.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a column first")
            return
        
        # 使用第一个选中的列作为参考列
        index = selected[0]
        col_name = self.column_listbox.get(index)
        self.reference_col_var.set(col_name)
        messagebox.showinfo("Reference Column", f"Reference column set to: {col_name}")
        
        # 确保在计算差值时启用该列的显示
        if col_name not in [self.column_listbox.get(i) for i in self.column_listbox.curselection()]:
            messagebox.showwarning("Warning", "Remember to include the reference column in your plot selection")

    def set_pair_from_selection(self, pair_index):
        """从列表选择中设置湿度-压力对"""
        selected = self.column_listbox.curselection()
        if len(selected) < 2:
            messagebox.showwarning("Warning", "Please select at least two columns (moisture and pressure)")
            return
        
        # 获取选中的列名
        selected_cols = [self.column_listbox.get(i) for i in selected]
        
        # 获取湿度和压力变量
        moisture_var, pressure_var, _, _ = self.pair_vars[pair_index]
        
        # 尝试猜测哪个是湿度列，哪个是压力列
        moisture_col = None
        pressure_col = None
        
        # 更精确的关键词匹配
        moisture_keywords = ['h2o', 'water', 'humid', 'moisture', 'moistre', 'ppm', 'ppb']
        pressure_keywords = ['pressure', 'press', 'bar', 'pa', 'mpa']
        
        # 首先，精确检查每个列名是否包含明确的关键词
        moisture_scores = {}
        pressure_scores = {}
        
        for col in selected_cols:
            col_lower = col.lower()
            
            # 计算湿度关键词匹配分数
            moisture_score = 0
            for kw in moisture_keywords:
                if kw in col_lower:
                    moisture_score += 1
            
            # 计算压力关键词匹配分数
            pressure_score = 0
            for kw in pressure_keywords:
                if kw in col_lower:
                    pressure_score += 1
            
            # 如果一个列同时包含湿度和压力关键词，使用得分高的那个类别
            if moisture_score > 0 and pressure_score > 0:
                if moisture_score > pressure_score:
                    moisture_scores[col] = moisture_score
                else:
                    pressure_scores[col] = pressure_score
            elif moisture_score > 0:
                moisture_scores[col] = moisture_score
            elif pressure_score > 0:
                pressure_scores[col] = pressure_score
        
        # 选择得分最高的列作为湿度列和压力列
        if moisture_scores:
            moisture_col = max(moisture_scores.items(), key=lambda x: x[1])[0]
        
        if pressure_scores:
            pressure_col = max(pressure_scores.items(), key=lambda x: x[1])[0]
        
        # 如果某一类型未找到，则从剩余的列中选择
        remaining_cols = [col for col in selected_cols if col != moisture_col and col != pressure_col]
        
        if not moisture_col and not pressure_col:
            # 如果两者都未找到，使用前两个选择
            if len(selected_cols) >= 2:
                moisture_col = selected_cols[0]
                pressure_col = selected_cols[1]
        elif not moisture_col and pressure_col:
            # 如果只找到了压力列，从剩余列中选择第一个作为湿度列
            if remaining_cols:
                moisture_col = remaining_cols[0]
        elif moisture_col and not pressure_col:
            # 如果只找到了湿度列，从剩余列中选择第一个作为压力列
            if remaining_cols:
                pressure_col = remaining_cols[0]
        
        # 最后的检查，确保同时有湿度和压力列
        if not moisture_col or not pressure_col:
            messagebox.showerror("Error", "Could not determine moisture and pressure columns")
            return
        
        # 设置值
        moisture_var.set(moisture_col)
        pressure_var.set(pressure_col)
        
        # 保存最近的配对设置
        self.save_recent_pairs()
        
        # 显示更详细的配对信息
        messagebox.showinfo("Pair Set", f"Pair {pair_index+1} set to:\nMoisture: {moisture_col}\nPressure: {pressure_col}")
        
        # 如果湿度列名包含"pressure"或压力列名包含"moisture"，显示警告
        if any(kw in moisture_col.lower() for kw in pressure_keywords) and any(kw in pressure_col.lower() for kw in moisture_keywords):
            if messagebox.askyesno("Swap Columns?", 
                                 f"It seems the columns might be reversed:\n\n"
                                 f"Moisture: {moisture_col}\n"
                                 f"Pressure: {pressure_col}\n\n"
                                 f"Would you like to swap them?"):
                moisture_var.set(pressure_col)
                pressure_var.set(moisture_col)
                
                # 再次保存最新配对
                self.save_recent_pairs()
                
                messagebox.showinfo("Columns Swapped", 
                                   f"New arrangement:\n"
                                   f"Moisture: {pressure_col}\n"
                                   f"Pressure: {moisture_col}")
        
        # 如果一个列名包含典型的另一个类型的关键词，也显示询问
        elif any(kw in moisture_col.lower() for kw in pressure_keywords):
            if messagebox.askyesno("Swap Columns?", 
                                 f"The moisture column contains pressure keywords:\n\n"
                                 f"Moisture: {moisture_col}\n"
                                 f"Pressure: {pressure_col}\n\n"
                                 f"Would you like to swap them?"):
                moisture_var.set(pressure_col)
                pressure_var.set(moisture_col)
                
                # 再次保存最新配对
                self.save_recent_pairs()
                
                messagebox.showinfo("Columns Swapped", 
                                   f"New arrangement:\n"
                                   f"Moisture: {pressure_col}\n"
                                   f"Pressure: {moisture_col}")
        elif any(kw in pressure_col.lower() for kw in moisture_keywords):
            if messagebox.askyesno("Swap Columns?", 
                                 f"The pressure column contains moisture keywords:\n\n"
                                 f"Moisture: {moisture_col}\n"
                                 f"Pressure: {pressure_col}\n\n"
                                 f"Would you like to swap them?"):
                moisture_var.set(pressure_col)
                pressure_var.set(moisture_col)
                
                # 再次保存最新配对
                self.save_recent_pairs()
                
                messagebox.showinfo("Columns Swapped", 
                                   f"New arrangement:\n"
                                   f"Moisture: {pressure_col}\n"
                                   f"Pressure: {moisture_col}")

    def clear_pair(self, pair_index):
        """清除特定的湿度-压力对"""
        moisture_var, pressure_var, _, _ = self.pair_vars[pair_index]
        moisture_var.set("")
        pressure_var.set("")
    
    def save_pairs(self):
        """保存当前的湿度-压力配对设置到文件"""
        if not any(moisture_var.get() for moisture_var, _, _, _ in self.pair_vars):
            messagebox.showwarning("Warning", "No pairs to save")
            return
            
        try:
            # 创建一个包含所有配对的字典
            pairs_dict = {}
            for i, (moisture_var, pressure_var, _, _) in enumerate(self.pair_vars):
                moisture_col = moisture_var.get()
                pressure_col = pressure_var.get()
                
                if moisture_col and pressure_col:
                    pairs_dict[f"pair_{i+1}"] = {
                        "moisture": moisture_col,
                        "pressure": pressure_col
                    }
            
            # 添加时间戳
            pairs_dict["saved_on"] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # 添加可能的文件标识
            if self.file_path:
                pairs_dict["source_file"] = os.path.basename(self.file_path)
            
            # 请求保存位置
            file_path = filedialog.asksaveasfilename(
                title="Save Pairs Configuration",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialfile="moisture_pressure_pairs.json"
            )
            
            if not file_path:
                return
            
            # 保存到JSON文件
            with open(file_path, 'w') as f:
                json.dump(pairs_dict, f, indent=4)
            
            messagebox.showinfo("Success", f"Pairs saved to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save pairs: {str(e)}")
    
    def load_pairs(self):
        """从文件加载湿度-压力配对设置"""
        try:
            # 请求文件位置
            file_path = filedialog.askopenfilename(
                title="Load Pairs Configuration",
                filetypes=[("JSON files", "*.json")]
            )
            
            if not file_path:
                return
            
            # 从JSON文件加载
            with open(file_path, 'r') as f:
                pairs_dict = json.load(f)
            
            # 清除现有配对
            self.clear_all_pairs()
            
            # 设置加载的配对
            loaded_pairs = []
            for i in range(self.max_pairs):
                pair_key = f"pair_{i+1}"
                if pair_key in pairs_dict:
                    pair_data = pairs_dict[pair_key]
                    moisture_col = pair_data.get("moisture", "")
                    pressure_col = pair_data.get("pressure", "")
                    
                    if moisture_col and pressure_col:
                        moisture_var, pressure_var, _, _ = self.pair_vars[i]
                        moisture_var.set(moisture_col)
                        pressure_var.set(pressure_col)
                        loaded_pairs.append(f"Pair {i+1}: {moisture_col} - {pressure_col}")
            
            # 显示加载信息
            if loaded_pairs:
                info_text = "Loaded pairs:\n" + "\n".join(loaded_pairs)
                if "source_file" in pairs_dict:
                    info_text += f"\n\nOriginal source file: {pairs_dict['source_file']}"
                if "saved_on" in pairs_dict:
                    info_text += f"\nSaved on: {pairs_dict['saved_on']}"
                    
                messagebox.showinfo("Pairs Loaded", info_text)
            else:
                messagebox.showwarning("Warning", "No valid pairs found in the file")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load pairs: {str(e)}")

    def save_recent_pairs(self):
        """自动保存当前配对设置，以便下次打开相同文件时使用"""
        try:
            if not self.file_path:
                return
                
            # 创建一个包含所有配对的字典
            pairs_dict = {}
            
            # 添加文件标识作为键
            file_key = os.path.basename(self.file_path)
            
            pairs_dict[file_key] = {
                "pairs": {},
                "saved_on": pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 保存当前的配对
            for i, (moisture_var, pressure_var, _, _) in enumerate(self.pair_vars):
                moisture_col = moisture_var.get()
                pressure_col = pressure_var.get()
                
                if moisture_col and pressure_col:
                    pairs_dict[file_key]["pairs"][f"pair_{i+1}"] = {
                        "moisture": moisture_col,
                        "pressure": pressure_col
                    }
            
            # 如果存在，加载已有的配对文件
            existing_pairs = {}
            if os.path.exists(self.recent_pairs_file):
                try:
                    with open(self.recent_pairs_file, 'r') as f:
                        existing_pairs = json.load(f)
                except:
                    pass
            
            # 更新已有配对文件，保留其他文件的配对设置
            existing_pairs.update(pairs_dict)
            
            # 保存到JSON文件
            with open(self.recent_pairs_file, 'w') as f:
                json.dump(existing_pairs, f, indent=4)
            
            print(f"自动保存了当前文件的配对设置")
            
        except Exception as e:
            print(f"保存最近配对失败: {str(e)}")
    
    def load_recent_pairs(self):
        """尝试加载与当前文件关联的最近配对设置"""
        try:
            if not self.file_path or not os.path.exists(self.recent_pairs_file):
                return False
                
            # 加载配对文件
            with open(self.recent_pairs_file, 'r') as f:
                all_pairs = json.load(f)
            
            # 检查当前文件是否有保存的配对
            file_key = os.path.basename(self.file_path)
            if file_key not in all_pairs:
                return False
            
            # 获取此文件的配对设置
            file_pairs = all_pairs[file_key].get("pairs", {})
            if not file_pairs:
                return False
            
            # 清除现有配对
            self.clear_all_pairs()
            
            # 设置保存的配对
            loaded_count = 0
            for i in range(self.max_pairs):
                pair_key = f"pair_{i+1}"
                if pair_key in file_pairs:
                    pair_data = file_pairs[pair_key]
                    moisture_col = pair_data.get("moisture", "")
                    pressure_col = pair_data.get("pressure", "")
                    
                    if moisture_col and pressure_col:
                        # 验证列名在当前文件中存在
                        if moisture_col in self.df.columns and pressure_col in self.df.columns:
                            moisture_var, pressure_var, _, _ = self.pair_vars[i]
                            moisture_var.set(moisture_col)
                            pressure_var.set(pressure_col)
                            loaded_count += 1
            
            if loaded_count > 0:
                messagebox.showinfo("Previous Pairs Loaded", 
                                  f"Loaded {loaded_count} pairs from your previous settings for this file.\n\n"
                                  f"Last saved on: {all_pairs[file_key].get('saved_on', 'unknown')}")
                return True
            
            return False
            
        except Exception as e:
            print(f"加载最近配对失败: {str(e)}")
            return False

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DataVisualizer()
    app.run() 
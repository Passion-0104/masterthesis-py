import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import os
import traceback  # 添加这一行，用于更详细的错误跟踪
from scipy.optimize import minimize

# 检查matplotlib版本 - 修正版本检查方式
print(f"Matplotlib 版本: {matplotlib.__version__}")

class MoistureSensorCalibration:
    def __init__(self, root):
        self.root = root
        self.root.title("Moisture Sensor Calibration Tool")
        self.root.geometry("1000x800")
        
        # 校准参数 - 修改为可调整
        self.f1 = tk.DoubleVar(value=0.196798)
        self.f2 = tk.DoubleVar(value=0.419073)
        self.p_ref = tk.DoubleVar(value=1.0)  # 目标归一化压力（bar）
        
        # 数据点范围选择
        self.data_range = tk.StringVar(value="全部")
        
        # 数据变量
        self.input_file = None
        self.df = None
        self.output_file = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # 顶部控制面板
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(pady=10, fill=tk.X)
        
        # 文件选择按钮
        self.file_button = tk.Button(self.control_frame, text="Select Excel File", command=self.load_file)
        self.file_button.pack(side=tk.LEFT, padx=10)
        
        # 文件名显示标签
        self.file_label = tk.Label(self.control_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # 参数设置框架
        self.param_frame = tk.LabelFrame(self.root, text="Calibration Parameters")
        self.param_frame.pack(pady=5, fill=tk.X, padx=10)
        
        # f1参数
        self.f1_label = tk.Label(self.param_frame, text="f1 Value:")
        self.f1_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.f1_entry = tk.Entry(self.param_frame, textvariable=self.f1, width=10)
        self.f1_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # f2参数
        self.f2_label = tk.Label(self.param_frame, text="f2 Value:")
        self.f2_label.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.f2_entry = tk.Entry(self.param_frame, textvariable=self.f2, width=10)
        self.f2_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # p_ref参数
        self.p_ref_label = tk.Label(self.param_frame, text="Reference Pressure (bar):")
        self.p_ref_label.grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.p_ref_entry = tk.Entry(self.param_frame, textvariable=self.p_ref, width=10)
        self.p_ref_entry.grid(row=0, column=5, padx=5, pady=5, sticky="w")
        
        # 参数说明
        formula_label = tk.Label(self.param_frame, 
                               text="Calibration Formula: ppm_calibrated = concentration × (p_ref/p)^(f1·ln(p_ref/p)+f2)", 
                               font=("Arial", 9))
        formula_label.grid(row=1, column=0, columnspan=6, padx=5, pady=5, sticky="w")
        
        # 第一组数据框架 (用LabelFrame标识)
        self.group1_frame = tk.LabelFrame(self.root, text="Data Group 1")
        self.group1_frame.pack(pady=5, fill=tk.X, padx=10)
        
        # 第一组压力列选择
        self.pressure1_label = tk.Label(self.group1_frame, text="Pressure Column:")
        self.pressure1_label.pack(side=tk.LEFT, padx=5)
        self.pressure1_combobox = ttk.Combobox(self.group1_frame, state="readonly", width=15)
        self.pressure1_combobox.pack(side=tk.LEFT, padx=5)
        
        # 第一组PPM列选择
        self.ppm1_label = tk.Label(self.group1_frame, text="PPM Column:")
        self.ppm1_label.pack(side=tk.LEFT, padx=5)
        self.ppm1_combobox = ttk.Combobox(self.group1_frame, state="readonly", width=15)
        self.ppm1_combobox.pack(side=tk.LEFT, padx=5)
        
        # 数据点范围选择框架
        self.range_frame = tk.LabelFrame(self.root, text="Data Points Range", padx=10, pady=5)
        self.range_frame.pack(pady=5, fill=tk.X, padx=10)
        
        # 单选按钮
        self.range_all = tk.Radiobutton(self.range_frame, text="Process All Data Points", 
                                       variable=self.data_range, value="全部",
                                       font=("Arial", 10))
        self.range_all.pack(side=tk.LEFT, padx=20, pady=5)
        
        self.range_1000 = tk.Radiobutton(self.range_frame, text="Process Only First 1000 Points", 
                                        variable=self.data_range, value="1000",
                                        font=("Arial", 10))
        self.range_1000.pack(side=tk.LEFT, padx=20, pady=5)
        
        self.range_10000 = tk.Radiobutton(self.range_frame, text="Process Only First 10000 Points", 
                                         variable=self.data_range, value="10000",
                                         font=("Arial", 10))
        self.range_10000.pack(side=tk.LEFT, padx=20, pady=5)
        
        # 第二组数据框架 (用LabelFrame标识)
        self.group2_frame = tk.LabelFrame(self.root, text="Data Group 2")
        self.group2_frame.pack(pady=5, fill=tk.X, padx=10)
        
        # 第二组压力列选择
        self.pressure2_label = tk.Label(self.group2_frame, text="Pressure Column:")
        self.pressure2_label.pack(side=tk.LEFT, padx=5)
        self.pressure2_combobox = ttk.Combobox(self.group2_frame, state="readonly", width=15)
        self.pressure2_combobox.pack(side=tk.LEFT, padx=5)
        
        # 第二组PPM列选择
        self.ppm2_label = tk.Label(self.group2_frame, text="PPM Column:")
        self.ppm2_label.pack(side=tk.LEFT, padx=5)
        self.ppm2_combobox = ttk.Combobox(self.group2_frame, state="readonly", width=15)
        self.ppm2_combobox.pack(side=tk.LEFT, padx=5)
        
        # 启用/禁用第二组数据的复选框
        self.enable_group2_var = tk.BooleanVar(value=True)
        self.enable_group2 = tk.Checkbutton(self.group2_frame, text="Enable Group 2", 
                                          variable=self.enable_group2_var,
                                          command=self.toggle_group2)
        self.enable_group2.pack(side=tk.LEFT, padx=10)
        
        # 按钮框架
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=10, fill=tk.X)
        
        # 校准按钮
        self.calibrate_button = tk.Button(self.button_frame, text="Calibrate Data", command=self.run_calibration)
        self.calibrate_button.pack(side=tk.LEFT, padx=10)
        self.calibrate_button.config(state=tk.DISABLED)  # 初始禁用
        
        # 自动优化按钮
        self.optimize_button = tk.Button(self.button_frame, text="Find Optimal Parameters", command=self.optimize_parameters)
        self.optimize_button.pack(side=tk.LEFT, padx=10)
        self.optimize_button.config(state=tk.DISABLED)  # 初始禁用
        
        # 保存按钮
        self.save_button = tk.Button(self.button_frame, text="Save Results", command=self.save_results)
        self.save_button.pack(side=tk.LEFT, padx=10)
        self.save_button.config(state=tk.DISABLED)  # 初始禁用
        
        # 参数框架
        self.param_buttons_frame = tk.Frame(self.root)
        self.param_buttons_frame.pack(pady=5, fill=tk.X)
        
        # 保存参数按钮
        self.save_params_button = tk.Button(self.param_buttons_frame, text="Export Parameters", command=self.export_parameters)
        self.save_params_button.pack(side=tk.LEFT, padx=10)
        
        # 加载参数按钮
        self.load_params_button = tk.Button(self.param_buttons_frame, text="Import Parameters", command=self.import_parameters)
        self.load_params_button.pack(side=tk.LEFT, padx=10)
        
        # 图表显示区域
        self.plot_frame = tk.Frame(self.root)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)
        
        self.figure = plt.Figure(figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.plot_frame)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.plot_frame)
        self.toolbar.update()
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
    def toggle_group2(self):
        """启用或禁用第二组数据输入"""
        state = "readonly" if self.enable_group2_var.get() else "disabled"
        self.pressure2_combobox.config(state=state)
        self.ppm2_combobox.config(state=state)
        
    def load_file(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if file_path:
                self.input_file = file_path
                self.file_label.config(text=os.path.basename(file_path))
                
                try:
                    # 尝试读取Excel文件
                    self.df = pd.read_excel(file_path)
                    print(f"Successfully loaded file: {file_path}")
                    print(f"Data shape: {self.df.shape}")
                    print(f"Column names: {self.df.columns.tolist()}")
                    
                    # 更新列选择下拉框
                    columns = self.df.columns.tolist()
                    self.pressure1_combobox['values'] = columns
                    self.ppm1_combobox['values'] = columns
                    self.pressure2_combobox['values'] = columns
                    self.ppm2_combobox['values'] = columns
                    
                    # 尝试自动选择最可能的列
                    pressure_cols = []
                    ppm_cols = []
                    
                    for col in columns:
                        if 'press' in str(col).lower():
                            pressure_cols.append(col)
                        elif 'ppm' in str(col).lower() or 'humid' in str(col).lower() or 'h2o' in str(col).lower():
                            ppm_cols.append(col)
                    
                    # 第一组数据列
                    if pressure_cols and len(pressure_cols) > 0:
                        self.pressure1_combobox.set(pressure_cols[0])
                    elif len(columns) > 0:
                        self.pressure1_combobox.set(columns[0])
                        
                    if ppm_cols and len(ppm_cols) > 0:
                        self.ppm1_combobox.set(ppm_cols[0])
                    elif len(columns) > 1:
                        self.ppm1_combobox.set(columns[1])
                    
                    # 第二组数据列
                    if pressure_cols and len(pressure_cols) > 1:
                        self.pressure2_combobox.set(pressure_cols[1])
                    elif len(columns) > 2:
                        self.pressure2_combobox.set(columns[2])
                        
                    if ppm_cols and len(ppm_cols) > 1:
                        self.ppm2_combobox.set(ppm_cols[1])
                    elif len(columns) > 3:
                        self.ppm2_combobox.set(columns[3])
                    
                    # 启用校准按钮和优化按钮
                    self.calibrate_button.config(state=tk.NORMAL)
                    self.optimize_button.config(state=tk.NORMAL)
                    messagebox.showinfo("Success", "File loaded successfully, please select the columns to calibrate")
                except Exception as e:
                    print(f"File reading error: {str(e)}")
                    print(traceback.format_exc())  # 打印详细错误堆栈
                    messagebox.showerror("Error", f"Failed to read file: {str(e)}")
        except Exception as e:
            print(f"File selection error: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Error selecting file: {str(e)}")
                
    def calibrate(self, ppm, p, f1, f2):
        """根据校准公式计算校准后的ppm值
        ppm_calibrated = concentration × (p_ref/p)^(f1·ln(p_ref/p)+f2)
        """
        ratio = self.p_ref.get() / p
        exponent = f1 * np.log(ratio) + f2
        
        # 限制指数范围，避免数值爆炸
        exponent = np.clip(exponent, -10, 10)
        
        return ppm * (ratio ** exponent)
                
    def run_calibration(self):
        try:
            if self.df is None:
                messagebox.showwarning("Warning", "Please select a data file first")
                return
            
            # 获取第一组数据列
            pressure1_col = self.pressure1_combobox.get()
            ppm1_col = self.ppm1_combobox.get()
            
            if not pressure1_col or not ppm1_col:
                messagebox.showwarning("Warning", "Please select at least pressure and PPM columns for Group 1")
                return
                
            try:
                # 获取当前参数值
                f1_value = self.f1.get()
                f2_value = self.f2.get()
                p_ref_value = self.p_ref.get()
                
                # 清除图表
                self.figure.clear()
                ax = self.figure.add_subplot(111)
                
                # 显示数据信息
                print(f"Pressure column: {pressure1_col}, Type: {self.df[pressure1_col].dtype}")
                print(f"PPM column: {ppm1_col}, Type: {self.df[ppm1_col].dtype}")
                
                # 确保数据为数值类型
                self.df[pressure1_col] = pd.to_numeric(self.df[pressure1_col], errors='coerce')
                self.df[ppm1_col] = pd.to_numeric(self.df[ppm1_col], errors='coerce')
                
                # 删除任何NaN值
                self.df = self.df.dropna(subset=[pressure1_col, ppm1_col])
                
                # 根据选择的数据点范围处理数据
                range_value = self.data_range.get()
                original_data_points = len(self.df)
                print(f"原始数据点数量: {original_data_points}")
                
                if range_value == "1000" and original_data_points > 1000:
                    self.df = self.df.iloc[:1000]
                    print(f"数据点超过1000个，仅使用前1000个数据点")
                elif range_value == "10000" and original_data_points > 10000:
                    self.df = self.df.iloc[:10000]
                    print(f"数据点超过10000个，仅使用前10000个数据点")
                
                print(f"实际使用数据点数量: {len(self.df)}")
                
                # 校准第一组数据
                self.df['ppm1_calibrated'] = self.calibrate(
                    self.df[ppm1_col], 
                    self.df[pressure1_col], 
                    f1_value, 
                    f2_value
                )
    
                # 创建点索引作为X轴
                points1 = list(range(1, len(self.df) + 1))
                
                # 绘制第一组数据 - 点对点比较
                ax.plot(points1, self.df[ppm1_col], 'o-', color='blue', 
                       label=f'Original ppm (Group 1)')
                ax.plot(points1, self.df['ppm1_calibrated'], 's--', color='green',
                       label=f'Calibrated ppm (Group 1)')
                
                # 处理第二组数据 (如果启用)
                if self.enable_group2_var.get():
                    pressure2_col = self.pressure2_combobox.get()
                    ppm2_col = self.ppm2_combobox.get()
                    
                    if pressure2_col and ppm2_col:
                        # 确保数据为数值类型
                        self.df[pressure2_col] = pd.to_numeric(self.df[pressure2_col], errors='coerce')
                        self.df[ppm2_col] = pd.to_numeric(self.df[ppm2_col], errors='coerce')
                        
                        self.df['ppm2_calibrated'] = self.calibrate(
                            self.df[ppm2_col], 
                            self.df[pressure2_col], 
                            f1_value, 
                            f2_value
                        )
                        
                        # 绘制第二组数据 - 点对点比较
                        ax.plot(points1, self.df[ppm2_col], 'o-', color='red',
                               label=f'Original ppm (Group 2)')
                        ax.plot(points1, self.df['ppm2_calibrated'], 's--', color='purple',
                               label=f'Calibrated ppm (Group 2)')
                
                # 显示简化的数据表格
                print("Calibration Results Summary:")
                for i, (_, row) in enumerate(self.df.head().iterrows(), 1):
                    print(f"Point {i}: Pressure={row[pressure1_col]:.2f}, Original PPM={row[ppm1_col]:.2f}, Calibrated={row['ppm1_calibrated']:.2f}")
                
                # 显示当前参数值
                param_text = f'Parameters: f1={f1_value:.6f}, f2={f2_value:.6f}, p_ref={p_ref_value:.2f}'
                ax.text(0.05, 0.95, param_text, transform=ax.transAxes, fontsize=9,
                       verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
                
                ax.set_xlabel('Data Point Number')
                ax.set_ylabel('Moisture (ppm)')
                ax.set_title('Point-by-Point Moisture Data Calibration')
                ax.legend()
                ax.grid(True)
                
                # 设置X轴为整数刻度
                ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
                
                # 使用简单布局
                self.figure.tight_layout()
                self.canvas.draw()
                
                # 添加详细数据显示
                print("=" * 50)
                print("Detailed Calibration Results:")
                print(self.df[[pressure1_col, ppm1_col, 'ppm1_calibrated']].head(10).to_string())
                print("=" * 50)
                
                # 启用保存按钮
                self.save_button.config(state=tk.NORMAL)
                
                # 计算校准后的最大误差
                if self.enable_group2_var.get() and 'ppm2_calibrated' in self.df.columns:
                    # 计算两组校准数据之间的绝对误差
                    abs_errors = np.abs(self.df['ppm1_calibrated'] - self.df['ppm2_calibrated'])
                    max_abs_error = np.max(abs_errors)
                    
                    # 在图表上标注最大误差信息
                    max_error_text = f'最大误差: {max_abs_error:.2f} ppm'
                    ax.text(0.05, 0.88, max_error_text, transform=ax.transAxes, fontsize=9,
                           verticalalignment='top', bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.5))
                    
                    # 显示误差消息框
                    messagebox.showinfo("校准误差分析", f"最大绝对误差: {max_abs_error:.2f} ppm")
                
                messagebox.showinfo("Success", "Calibration completed")
                
            except Exception as e:
                print(f"Calibration process error: {str(e)}")
                print(traceback.format_exc())
                messagebox.showerror("Error", f"Calibration failed: {str(e)}")
                
        except Exception as e:
            print(f"Overall calibration error: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def save_results(self):
        if self.df is None:
            messagebox.showwarning("Warning", "No data to save")
            return
            
        if 'ppm1_calibrated' not in self.df.columns:
            messagebox.showwarning("Warning", "No calibration results to save")
            return
            
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="humidity_calibrated.xlsx"
        )
        
        if output_file:
            try:
                self.df.to_excel(output_file, index=False)
                messagebox.showinfo("Success", f"Results saved to: {output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Save failed: {str(e)}")

    def optimize_parameters(self):
        """Find optimal f1 and f2 parameters to minimize the difference between two calibrated curves"""
        try:
            if self.df is None:
                messagebox.showwarning("Warning", "Please select a data file first")
                return
            
            # Get data columns
            pressure1_col = self.pressure1_combobox.get()
            ppm1_col = self.ppm1_combobox.get()
            
            if not pressure1_col or not ppm1_col:
                messagebox.showwarning("Warning", "Please select at least pressure and PPM columns for Group 1")
                return
            
            # Check if Group 2 is enabled and has columns selected
            if not self.enable_group2_var.get():
                messagebox.showwarning("Warning", "Group 2 must be enabled for optimization")
                return
                
            pressure2_col = self.pressure2_combobox.get()
            ppm2_col = self.ppm2_combobox.get()
            
            if not pressure2_col or not ppm2_col:
                messagebox.showwarning("Warning", "Please select pressure and PPM columns for Group 2")
                return
            
            # Ensure data types are numeric
            self.df[pressure1_col] = pd.to_numeric(self.df[pressure1_col], errors='coerce')
            self.df[ppm1_col] = pd.to_numeric(self.df[ppm1_col], errors='coerce')
            self.df[pressure2_col] = pd.to_numeric(self.df[pressure2_col], errors='coerce')
            self.df[ppm2_col] = pd.to_numeric(self.df[ppm2_col], errors='coerce')
            
            # Remove NaN values
            self.df = self.df.dropna(subset=[pressure1_col, ppm1_col, pressure2_col, ppm2_col])
            
            if len(self.df) == 0:
                messagebox.showerror("Error", "No valid data points after removing NaN values")
                return
            
            # 根据选择的数据点范围处理数据
            range_value = self.data_range.get()
            original_data_points = len(self.df)
            print(f"原始数据点数量: {original_data_points}")
            
            if range_value == "1000" and original_data_points > 1000:
                self.df = self.df.iloc[:1000]
                print(f"数据点超过1000个，仅使用前1000个数据点进行优化")
            elif range_value == "10000" and original_data_points > 10000:
                self.df = self.df.iloc[:10000]
                print(f"数据点超过10000个，仅使用前10000个数据点进行优化")
            
            print(f"优化使用数据点数量: {len(self.df)}")
            
            # Current p_ref value
            p_ref_value = self.p_ref.get()
            
            # Define improved objective function for optimization
            def objective_function(params):
                f1, f2 = params
                
                # Calculate calibrated values for both groups
                cal1 = self.calibrate(self.df[ppm1_col], self.df[pressure1_col], f1, f2)
                cal2 = self.calibrate(self.df[ppm2_col], self.df[pressure2_col], f1, f2)
                
                # 目标是让两组的均值一致，同时自身波动不大
                mse = (np.mean(cal1) - np.mean(cal2))**2
                smoothness = np.var(cal1) + np.var(cal2)
                
                return mse + 0.1 * smoothness  # 可调节权重
            
            # Initial guess based on current values
            initial_guess = [0.2, 0.4]  # 合理的起始点
            
            # Run optimization
            print("Starting parameter optimization...")
            print(f"Initial parameters: f1={initial_guess[0]}, f2={initial_guess[1]}")
            
            # Show wait message
            messagebox.showinfo("Optimization", "Finding optimal parameters... This may take a moment.")
            
            # Perform the optimization with boundary constraints
            result = minimize(
                objective_function,
                initial_guess,
                method='L-BFGS-B',
                bounds=[(0.0, 2.0), (-2.0, 2.0)],  # 根据经验设置合理的参数范围
                options={'maxiter': 1000, 'disp': True}
            )
            
            # Get optimized parameters
            opt_f1, opt_f2 = result.x
            
            # 计算最大误差
            cal1 = self.calibrate(self.df[ppm1_col], self.df[pressure1_col], opt_f1, opt_f2)
            cal2 = self.calibrate(self.df[ppm2_col], self.df[pressure2_col], opt_f1, opt_f2)
            
            # Calculate the absolute error between two calibrated groups
            abs_errors = np.abs(cal1 - cal2)
            max_abs_error = np.max(abs_errors)
            
            # Print the maximum absolute error
            print("=" * 50)
            print("Optimization Result Analysis:")
            print(f"Maximum Absolute Error: {max_abs_error:.2f} ppm")
            print("=" * 50)
            
            # Update parameters in the UI
            self.f1.set(opt_f1)
            self.f2.set(opt_f2)
            
            print(f"Optimization completed")
            print(f"Optimal parameters: f1={opt_f1:.6f}, f2={opt_f2:.6f}")
            print(f"Final error: {result.fun:.6f}")
            
            # Run calibration with the new parameters
            self.run_calibration()
            
            # Show optimization results and maximum error
            messagebox.showinfo("Optimization Complete", 
                              f"优化参数:\nf1 = {opt_f1:.6f}\nf2 = {opt_f2:.6f}\n\nmax error: {max_abs_error:.2f} ppm")
            
        except Exception as e:
            print(f"Optimization error: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Optimization failed: {str(e)}")

    def export_parameters(self):
        """Export the current calibration parameters to a JSON file"""
        try:
            # Get current parameters
            f1_value = self.f1.get()
            f2_value = self.f2.get()
            p_ref_value = self.p_ref.get()
            
            # Create parameters dictionary
            params = {
                'f1': f1_value,
                'f2': f2_value,
                'p_ref': p_ref_value,
                'date': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialfile="moisture_calibration_params.json"
            )
            
            if not file_path:
                return
            
            # Save parameters to file
            import json
            with open(file_path, 'w') as f:
                json.dump(params, f, indent=4)
            
            messagebox.showinfo("Success", f"Parameters exported to {file_path}")
            
        except Exception as e:
            print(f"Error exporting parameters: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Failed to export parameters: {str(e)}")
    
    def import_parameters(self):
        """Import calibration parameters from a JSON file"""
        try:
            # Ask user for file location
            file_path = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json")]
            )
            
            if not file_path:
                return
            
            # Load parameters from file
            import json
            with open(file_path, 'r') as f:
                params = json.load(f)
            
            # Verify parameters
            required_keys = ['f1', 'f2', 'p_ref']
            if not all(key in params for key in required_keys):
                messagebox.showerror("Error", "Invalid parameter file format")
                return
            
            # Update parameters in UI
            self.f1.set(params['f1'])
            self.f2.set(params['f2'])
            self.p_ref.set(params['p_ref'])
            
            # Show parameter details
            details = f"Imported parameters:\n\nf1 = {params['f1']:.6f}\nf2 = {params['f2']:.6f}\np_ref = {params['p_ref']:.2f}"
            if 'date' in params:
                details += f"\n\nCreated on: {params['date']}"
                
            messagebox.showinfo("Parameters Imported", details)
            
            # If data is loaded, apply calibration with new parameters
            if self.df is not None and self.calibrate_button['state'] != 'disabled':
                self.run_calibration()
                
        except Exception as e:
            print(f"Error importing parameters: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Failed to import parameters: {str(e)}")

    def calculate_time_differences(self):
        """计算第20分钟、40分钟和60分钟时的差值，并计算平均值"""
        try:
            if self.df is None or 'ppm1_calibrated' not in self.df.columns:
                messagebox.showwarning("Warning", "请先进行校准")
                return
                
            if not self.enable_group2_var.get() or 'ppm2_calibrated' not in self.df.columns:
                messagebox.showwarning("Warning", "需要启用两组数据进行差值计算")
                return
                
            # 选择参考组 (默认使用第一组)
            reference_values = self.df['ppm1_calibrated']
            compare_values = self.df['ppm2_calibrated']
            
            # 创建相对时间列（假设数据点是均匀分布的）
            # 转换为分钟
            total_points = len(self.df)
            if total_points < 10:
                messagebox.showwarning("Warning", "数据点数量不足")
                return
                
            # 估计每个点的时间间隔（假设2小时的数据）
            time_interval = 120 / total_points  # 单位：分钟
            time_points = [i * time_interval for i in range(total_points)]
            
            # 目标时间点（分钟）
            target_times = [20, 40, 60]
            
            # 查找最接近的时间点
            results = []
            differences = []
            
            for target_time in target_times:
                # 找到最接近目标时间的点
                closest_idx = min(range(len(time_points)), key=lambda i: abs(time_points[i] - target_time))
                closest_time = time_points[closest_idx]
                
                # 获取该时间点的值
                ref_value = reference_values.iloc[closest_idx]
                comp_value = compare_values.iloc[closest_idx]
                
                # 计算差值
                difference = abs(comp_value - ref_value)
                differences.append(difference)
                
                # 添加到结果中
                results.append(f"{target_time}分钟处 - 参考值: {ref_value:.2f}, 比较值: {comp_value:.2f}, 差值: {difference:.2f} ppm")
            
            # 计算平均差值
            avg_difference = sum(differences) / len(differences)
            results.append(f"\n平均差值: {avg_difference:.2f} ppm")
            
            # 显示结果
            messagebox.showinfo("时间点差值计算", "\n".join(results))
            
        except Exception as e:
            print(f"计算时间差值错误: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"计算差值失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MoistureSensorCalibration(root)
    root.mainloop()

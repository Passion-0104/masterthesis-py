import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from tkinter import Tk, filedialog, messagebox
import tkinter as tk
from tkinter import ttk

class DataVisualizer:
    def __init__(self):
        # 设置全局字体
        plt.rcParams['font.family'] = 'Arial'
        plt.rcParams['mathtext.fontset'] = 'dejavusans'
        plt.rcParams['axes.unicode_minus'] = False
        plt.rcParams['font.weight'] = 'bold'
        
        # 创建主窗口
        self.root = Tk()
        self.root.title("数据可视化工具")
        self.root.geometry("800x600")
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择按钮
        ttk.Button(self.main_frame, text="选择Excel文件", command=self.select_file).grid(row=0, column=0, pady=10)
        
        # 创建列表框用于显示列名
        self.column_listbox = tk.Listbox(self.main_frame, selectmode=tk.MULTIPLE, width=50, height=10)
        self.column_listbox.grid(row=1, column=0, pady=10)
        
        # 时间列选择
        ttk.Label(self.main_frame, text="选择时间列:").grid(row=2, column=0)
        self.time_column_var = tk.StringVar()
        self.time_column_combo = ttk.Combobox(self.main_frame, textvariable=self.time_column_var)
        self.time_column_combo.grid(row=3, column=0, pady=5)
        
        # 添加时间范围选择
        self.time_range_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.main_frame, text="只显示前2小时数据", variable=self.time_range_var).grid(row=4, column=0, pady=5)
        
        # 绘图按钮
        ttk.Button(self.main_frame, text="生成图表", command=self.plot_data).grid(row=5, column=0, pady=10)
        
        # 初始化变量
        self.df = None
        self.file_path = None

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_path:
            try:
                self.df = pd.read_excel(self.file_path)
                self.update_column_list()
                messagebox.showinfo("成功", "文件加载成功！")
            except Exception as e:
                messagebox.showerror("错误", f"加载文件时出错：{str(e)}")

    def update_column_list(self):
        if self.df is not None:
            # 清空列表框
            self.column_listbox.delete(0, tk.END)
            self.time_column_combo['values'] = []
            
            # 添加列名到列表框和下拉框
            for col in self.df.columns:
                self.column_listbox.insert(tk.END, col)
                self.time_column_combo['values'] = (*self.time_column_combo['values'], col)
            
            # 默认选择第一列作为时间列
            if len(self.df.columns) > 0:
                self.time_column_var.set(self.df.columns[0])

    def plot_data(self):
        if self.df is None:
            messagebox.showerror("错误", "请先选择文件！")
            return
        
        time_col = self.time_column_var.get()
        if not time_col:
            messagebox.showerror("错误", "请选择时间列！")
            return
        
        selected_columns = [self.column_listbox.get(i) for i in self.column_listbox.curselection()]
        if not selected_columns:
            messagebox.showerror("错误", "请选择要绘制的列！")
            return
        
        try:
            # 处理时间列
            self.df[time_col] = pd.to_datetime(self.df[time_col], errors='coerce')
            self.df = self.df.dropna(subset=[time_col])
            
            # 计算相对时间
            start_time = self.df[time_col].min()
            self.df['relative_time'] = (self.df[time_col] - start_time).dt.total_seconds() / 3600
            
            # 如果选择了只显示前2小时数据
            if self.time_range_var.get():
                self.df = self.df[self.df['relative_time'] <= 2]
            
            # 创建图表
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # 设置颜色
            colors = plt.cm.tab10(np.linspace(0, 1, len(selected_columns)))
            
            # 绘制数据
            for i, col in enumerate(selected_columns):
                ax.plot(self.df['relative_time'], self.df[col], 
                       color=colors[i],
                       label=col,
                       linewidth=1.5)
            
            # 设置图表样式
            ax.set_xlabel("时间 (小时)", fontsize=12, fontweight='bold')
            ax.set_ylabel("数值", fontsize=12, fontweight='bold')
            ax.set_title("数据可视化", fontsize=14, fontweight='bold')
            
            # 设置网格
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # 设置图例
            ax.legend(loc='best', frameon=True)
            
            # 设置刻度
            ax.tick_params(axis='both', which='major', labelsize=10)
            
            # 如果选择了只显示前2小时数据，设置x轴刻度
            if self.time_range_var.get():
                time_ticks = [0, 0.5, 1, 1.5, 2]
                time_labels = ['0', '30min', '1h', '1h30min', '2h']
                ax.set_xticks(time_ticks)
                ax.set_xticklabels(time_labels)
            
            # 调整布局
            plt.tight_layout()
            
            # 显示图表
            plt.show()
            
        except Exception as e:
            messagebox.showerror("错误", f"绘图时出错：{str(e)}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DataVisualizer()
    app.run() 
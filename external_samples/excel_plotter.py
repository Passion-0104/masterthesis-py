import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
import matplotlib

matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 显示中文
matplotlib.rcParams['axes.unicode_minus'] = False    # 显示负号

class ExcelPlotter:
    def __init__(self, root):
        self.root = root
        self.root.title("多文件Excel数据可视化工具")
        self.root.geometry("1000x800")
        
        self.file_paths = []
        self.dfs = []
        self.columns = []
        self.time_column = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # 文件选择按钮
        self.file_button = tk.Button(self.root, text="选择多个Excel文件", command=self.load_files)
        self.file_button.pack(pady=10)
        
        # 列选择框架
        self.column_frame = tk.Frame(self.root)
        self.column_frame.pack(pady=10)
        
        # 时间列选择
        self.time_label = tk.Label(self.column_frame, text="时间列:")
        self.time_label.pack(side=tk.LEFT, padx=5)
        self.time_combobox = ttk.Combobox(self.column_frame, state="readonly")
        self.time_combobox.pack(side=tk.LEFT, padx=5)
        
        # Y轴列选择
        self.y_label = tk.Label(self.column_frame, text="Y轴:")
        self.y_label.pack(side=tk.LEFT, padx=5)
        self.y_combobox = ttk.Combobox(self.column_frame, state="readonly")
        self.y_combobox.pack(side=tk.LEFT, padx=5)
        
        # 文件列表显示
        self.file_list_frame = tk.Frame(self.root)
        self.file_list_frame.pack(pady=10)
        self.file_list_label = tk.Label(self.file_list_frame, text="已选择的文件:")
        self.file_list_label.pack()
        self.file_list = tk.Listbox(self.file_list_frame, width=80, height=5)
        self.file_list.pack()
        
        # 绘图按钮
        self.plot_button = tk.Button(self.root, text="绘制图表", command=self.plot_data)
        self.plot_button.pack(pady=10)
        
        # 图表显示区域
        self.figure = plt.Figure(figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
    def load_files(self):
        self.file_paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_paths:
            self.file_list.delete(0, tk.END)
            self.dfs = []
            for file_path in self.file_paths:
                self.file_list.insert(tk.END, file_path)
                try:
                    df = pd.read_excel(file_path)
                    self.dfs.append(df)
                except Exception as e:
                    messagebox.showerror("错误", f"无法读取文件 {file_path}: {str(e)}")
            
            if self.dfs:
                # 获取所有文件的列名
                all_columns = set()
                for df in self.dfs:
                    all_columns.update(df.columns)
                self.columns = list(all_columns)
                
                # 自动识别时间列
                time_columns = []
                for col in self.columns:
                    if any(df[col].dtype == 'datetime64[ns]' for df in self.dfs):
                        time_columns.append(col)
                
                self.time_combobox['values'] = self.columns
                self.y_combobox['values'] = self.columns
                if len(self.columns) >= 1:
                    self.time_combobox.set(self.columns[0])
                    self.y_combobox.set(self.columns[0])
                
    def plot_data(self):
        if not self.dfs:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        time_col = self.time_combobox.get()
        y_col = self.y_combobox.get()
        
        if not time_col or not y_col:
            messagebox.showwarning("警告", "请选择时间列和Y轴列")
            return
            
        try:
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            # 为每个文件绘制一条线
            for i, (df, file_path) in enumerate(zip(self.dfs, self.file_paths)):
                if time_col in df.columns and y_col in df.columns:
                    # 尝试解析时间列
                    try:
                        df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
                    except Exception as e:
                        messagebox.showerror("错误", f"{file_path} 的时间列解析失败: {str(e)}")
                        continue

                    # 过滤掉无法解析的时间
                    df = df.dropna(subset=[time_col, y_col])

                    # 打印调试信息
                    print(f"{file_path} 时间列前5行：", df[time_col].head())

                    if df.empty:
                        messagebox.showwarning("警告", f"{file_path} 没有有效的时间数据，已跳过。")
                        continue

                    df = df.sort_values(by=time_col)
                    ax.plot(df[time_col], df[y_col], label=f'文件 {i+1}')
            
            ax.set_xlabel(time_col)
            ax.set_ylabel(y_col)
            ax.set_title(f"{y_col} vs {time_col}")
            ax.legend()
            ax.grid(True)
            
            # 自动调整X轴标签角度
            plt.xticks(rotation=45)
            
            # 调整布局
            self.figure.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            messagebox.showerror("错误", f"绘图时出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelPlotter(root)
    root.mainloop() 
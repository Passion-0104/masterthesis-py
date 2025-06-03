import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 1. 读取Excel文件
file_path = r"D:\masterarbeit\raw data\correction H2O reference chamber1 and 0.05bar.xlsx"
df = pd.read_excel(file_path)

# 2. 检查列名
print("表头：", df.columns.tolist())

# 3. 只保留需要的两列，并去除缺失值
columns_to_plot = ["reference in chamber 1", "reference in 0.05bar","reference in chamber 1(raw data)","reference in 0.05bar(raw data)"]
df = df[columns_to_plot]
df = df.apply(pd.to_numeric, errors='coerce')  # 关键：将#WERT!等替换为NaN
df = df.dropna()

# 4. 绘图
plt.figure(figsize=(8, 5))
for col in columns_to_plot:
    plt.plot(df.index, df[col], label=col, linewidth=4)

plt.xlabel("数据点序号", fontname='Arial', fontsize=12, fontweight='bold')
plt.ylabel("数值", fontname='Arial', fontsize=12, fontweight='bold')
plt.title("参考值对比", fontname='Arial', fontsize=14, fontweight='bold')
plt.legend(fontsize=11)
plt.grid(True, linestyle='--', alpha=0.5)
plt.tight_layout()
plt.show() 
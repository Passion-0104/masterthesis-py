import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 读取Excel文件
print("正在读取数据文件...")
file_path = r"C:\Users\12404\PycharmProjects\pythonProject1\data.xlsx"
df = pd.read_excel(file_path)

# 打印数据预览
print("\n数据预览：")
print(df.head())
print("\n列名：", df.columns.tolist())

# 创建图表
print("\n正在创建图表...")
plt.figure(figsize=(10, 6))
plt.scatter(df['pressure'], df['ppb'], alpha=0.5, color='blue')
plt.xlabel('压力 (mbar)', fontsize=12)
plt.ylabel('PPB', fontsize=12)
plt.title('压力与PPB的关系图', fontsize=14)
plt.grid(True, linestyle='--', alpha=0.7)

# 添加趋势线
z = np.polyfit(df['pressure'], df['ppb'], 1)
p = np.poly1d(z)
plt.plot(df['pressure'], p(df['pressure']), "r--", alpha=0.8, 
         label=f'趋势线: y = {z[0]:.2f}x + {z[1]:.2f}')

# 计算相关系数
correlation = df['pressure'].corr(df['ppb'])
plt.text(0.05, 0.95, f'相关系数: {correlation:.3f}', 
         transform=plt.gca().transAxes, fontsize=10,
         bbox=dict(facecolor='white', alpha=0.8, edgecolor='none'))

plt.legend(loc='best', framealpha=0.8)
plt.tight_layout()

# 保存图表
plt.savefig('pressure_vs_ppb_from_data.png', dpi=300, bbox_inches='tight')
print("\n图表已保存为 'pressure_vs_ppb_from_data.png'")

# 显示统计信息
print("\n数据统计信息：")
print(f"压力范围: {df['pressure'].min():.2f} 到 {df['pressure'].max():.2f} mbar")
print(f"PPB范围: {df['ppb'].min():.2f} 到 {df['ppb'].max():.2f} ppb")
print(f"相关系数: {correlation:.3f}")

# 显示图表
plt.show() 
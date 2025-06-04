import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
plt.style.use('seaborn')  # 使用更现代的绘图风格

# 读取两个Excel文件
file1_path = r"D:\masterarbeit\raw data\Bronkhorst FlowSuite Export 2025_04_24 19_29_2.xlsx"
file2_path = r"D:\masterarbeit\raw data\DataLogger 2025-04-24 ambient pressure.xlsx"

print("=== 步骤1：读取数据 ===")
# 读取第一个文件
print("读取文件1 (Bronkhorst)...")
df1 = pd.read_excel(file1_path)
print("文件1原始列名:", df1.columns.tolist())
print("文件1前5行:\n", df1.head())

# 读取第二个文件
print("\n读取文件2 (DataLogger)...")
df2 = pd.read_excel(file2_path)
print("文件2原始列名:", df2.columns.tolist())
print("文件2前5行:\n", df2.head())

print("\n=== 步骤2：数据预处理 ===")
# 处理文件1 (Bronkhorst)
print("处理文件1...")
# 假设第一列是时间，第二列是数值，具体列名需要根据实际数据调整
time_col1 = df1.columns[0]  # 第一列作为时间
value_col1 = df1.columns[1]  # 第二列作为数值
df1 = df1[[time_col1, value_col1]].copy()
df1.columns = ['Time', 'Value']

# 处理文件2 (DataLogger)
print("处理文件2...")
df2 = df2[['Zeit', 'S 25007 O']].copy()  # 使用已知的列名
df2.columns = ['Time', 'Value']

# 确保时间列是datetime类型，并且都是tz-naive
df1['Time'] = pd.to_datetime(df1['Time']).dt.tz_localize(None)
df2['Time'] = pd.to_datetime(df2['Time']).dt.tz_localize(None)

# 确保Value列是数值类型
df1['Value'] = pd.to_numeric(df1['Value'], errors='coerce')
df2['Value'] = pd.to_numeric(df2['Value'], errors='coerce')

# 删除任何可能的NaN值
df1 = df1.dropna()
df2 = df2.dropna()

print("\n处理后的数据预览：")
print("文件1:\n", df1.head())
print("\n文件2:\n", df2.head())
print("\n数据类型信息：")
print("文件1的数据类型:\n", df1.dtypes)
print("\n文件2的数据类型:\n", df2.dtypes)

print("\n=== 步骤3：数据对齐 ===")
# 计算时间间隔
time_diff1 = df1['Time'].diff().median()
time_diff2 = df2['Time'].diff().median()
print(f"文件1的时间间隔: {time_diff1}")
print(f"文件2的时间间隔: {time_diff2}")

# 对齐数据
def align_data(df_long, df_short, max_time_diff='1s'):
    """
    将较长的数据集与较短的数据集对齐
    df_long: 数据点较多的DataFrame
    df_short: 数据点较少的DataFrame
    max_time_diff: 最大允许的时间差
    """
    aligned_indices = []
    aligned_times = []
    
    # 将时间转换为时间戳以加快计算
    df_long_timestamps = df_long['Time'].astype(np.int64).values
    df_short_timestamps = df_short['Time'].astype(np.int64).values
    
    print(f"开始对齐，总计需要处理 {len(df_short)} 个点...")
    
    for i, target_time in enumerate(df_short_timestamps):
        if i % 100 == 0:
            print(f"进度: {i}/{len(df_short)} ({i/len(df_short)*100:.1f}%)")
        
        # 找到最接近的时间点
        time_diffs = np.abs(df_long_timestamps - target_time)
        closest_idx = np.argmin(time_diffs)
        min_diff = time_diffs[closest_idx]
        
        # 如果时间差在允许范围内，则保存这个点
        if pd.Timedelta(min_diff) <= pd.Timedelta(max_time_diff):
            aligned_indices.append(closest_idx)
            aligned_times.append(df_short['Time'].iloc[i])
    
    return df_long.iloc[aligned_indices].copy()

# 执行对齐
print("正在对齐数据...")
# 确定哪个数据集更长
if len(df1) > len(df2):
    df1_aligned = align_data(df1, df2, max_time_diff='1s')
    df2_aligned = df2
else:
    df2_aligned = align_data(df2, df1, max_time_diff='1s')
    df1_aligned = df1

print("\n=== 步骤4：验证结果 ===")
print("对齐后的数据形状：")
print(f"文件1: {df1_aligned.shape}")
print(f"文件2: {df2_aligned.shape}")

print("\n对齐后的前5行数据：")
print("文件1对齐后:\n", df1_aligned.head())
print("\n文件2对齐后:\n", df2_aligned.head())

# 验证时间对齐质量
time_diffs = pd.Timedelta(seconds=0)  # 初始化为0
for t1, t2 in zip(df1_aligned['Time'], df2_aligned['Time']):
    time_diffs = max(time_diffs, abs(pd.Timedelta(t1 - t2)))

print("\n时间对齐质量：")
print(f"最大时间差: {time_diffs}")

print("\n=== 步骤5：保存结果 ===")
# 保存对齐后的数据
output_path1 = "aligned_bronkhorst.xlsx"
output_path2 = "aligned_datalogger.xlsx"

df1_aligned.to_excel(output_path1, index=False)
df2_aligned.to_excel(output_path2, index=False)

print(f"数据已保存到：")
print(f"文件1 (Bronkhorst): {output_path1}")
print(f"文件2 (DataLogger): {output_path2}")

# 额外的数据验证
print("\n=== 最终验证 ===")
print("时间范围：")
print(f"文件1: {df1_aligned['Time'].min()} 到 {df1_aligned['Time'].max()}")
print(f"文件2: {df2_aligned['Time'].min()} 到 {df2_aligned['Time'].max()}")
print("\n数值范围：")
print(f"文件1值的类型: {df1_aligned['Value'].dtype}")
print(f"文件2值的类型: {df2_aligned['Value'].dtype}")
print(f"文件1: {df1_aligned['Value'].min():.3f} 到 {df1_aligned['Value'].max():.3f}")
print(f"文件2: {df2_aligned['Value'].min():.3f} 到 {df2_aligned['Value'].max():.3f}")

# 显示时间重叠情况
print("\n时间重叠情况：")
overlap_start = max(df1_aligned['Time'].min(), df2_aligned['Time'].min())
overlap_end = min(df1_aligned['Time'].max(), df2_aligned['Time'].max())
print(f"重叠时间范围: {overlap_start} 到 {overlap_end}")

# 在保存完数据之后，添加绘图代码
print("\n=== 步骤6：创建压力-PPB关系图 ===")
plt.figure(figsize=(10, 6))
plt.scatter(df2_aligned['Value'], df1_aligned['Value'], alpha=0.5)
plt.xlabel('压力 (mbar)', fontsize=12)
plt.ylabel('PPB', fontsize=12)
plt.title('压力与PPB的关系图', fontsize=14)
plt.grid(True)

# 添加趋势线
z = np.polyfit(df2_aligned['Value'], df1_aligned['Value'], 1)
p = np.poly1d(z)
plt.plot(df2_aligned['Value'], p(df2_aligned['Value']), "r--", alpha=0.8, label=f'趋势线: y = {z[0]:.2f}x + {z[1]:.2f}')

# 计算相关系数
correlation = df2_aligned['Value'].corr(df1_aligned['Value'])
plt.text(0.05, 0.95, f'相关系数: {correlation:.3f}', 
         transform=plt.gca().transAxes, fontsize=10)

plt.legend()
plt.tight_layout()

# 保存图表
plt.savefig('pressure_vs_ppb.png', dpi=300, bbox_inches='tight')
print("图表已保存为 'pressure_vs_ppb.png'")

# 显示一些基本的统计信息
print("\n数据统计信息：")
print(f"压力范围: {df2_aligned['Value'].min():.2f} 到 {df2_aligned['Value'].max():.2f} mbar")
print(f"PPB范围: {df1_aligned['Value'].min():.2f} 到 {df1_aligned['Value'].max():.2f} ppb")
print(f"相关系数: {correlation:.3f}")

plt.show() 
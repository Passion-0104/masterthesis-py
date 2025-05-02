import matplotlib.pyplot as plt
import numpy as np

# 数据定义
x_positions = [1, 2, 3]  # 用于绘图的x轴位置（等距）
pressures = [1, 2, 2.5]  # 实际压力值（用于标签）

# exp1 数据 (1h)
exp1_total = [8.17E-04, 8.43E-04, 8.45E-04]  # Total Volume
exp1_chamber1 = [7.33E-04, 7.48E-04, 7.49E-04]  # Volume Chamber 1
exp1_autoclave = [8.37E-05, 9.57E-05, 9.61E-05]  # Autoclave Inner Volume

# exp2 数据
exp2_total = [8.22E-04, 8.40E-04, 8.42E-04]  # Total Volume
exp2_chamber1 = [7.27E-04, 7.43E-04, 7.45E-04]  # Volume Chamber 1
exp2_autoclave = [9.51E-05, 9.68E-05, 9.66E-05]  # Autoclave Inner Volume

# 创建图形
plt.figure(figsize=(8, 6))  # 调整图形大小，使其更适合论文

# 设置字体
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.weight'] = 'bold'
plt.rcParams['axes.linewidth'] = 1.5  # 增加坐标轴线条粗细

# 绘制exp1的数据（实线）
plt.plot(x_positions, exp1_total, 'o-', color='#1f77b4', label='Total Volume (Pumping 1h)', linewidth=2.5, markersize=8)
plt.plot(x_positions, exp1_chamber1, 's-', color='#ff7f0e', label='Chamber 1 volume(Pumping 1h)', linewidth=2.5, markersize=8)
plt.plot(x_positions, exp1_autoclave, '^-', color='#2ca02c', label='Autoclave Inner Volume', linewidth=2.5, markersize=8)

# 绘制exp2的数据（虚线）
plt.plot(x_positions, exp2_total, 'o--', color='#1f77b4', label='Total Volume (Pumping 19h)', linewidth=2.5, dashes=(5, 5), markersize=8)
plt.plot(x_positions, exp2_chamber1, 's--', color='#ff7f0e', label='Chamber 1 volume(Pumping 66h)', linewidth=2.5, dashes=(5, 5), markersize=8)
plt.plot(x_positions, exp2_autoclave, '^--', color='#2ca02c', label='Autoclave Inner Volume', linewidth=2.5, dashes=(5, 5), markersize=8)

# 设置坐标轴
plt.xlabel('Pressure [bar]', fontsize=14, fontweight='bold')
plt.ylabel('Volume [m³]', fontsize=14, fontweight='bold')
plt.title('Volume Changes with Pressure at Different Pumping Durations', fontsize=16, fontweight='bold', pad=15)

# 设置刻度
plt.tick_params(axis='both', which='major', labelsize=12, width=1.5, length=6)
plt.tick_params(axis='both', which='minor', width=1.5, length=3)

# 设置x轴刻度为等距
plt.xticks(x_positions, ['1', '2', '2.5'])

# 使用科学计数法表示y轴
plt.ticklabel_format(axis='y', style='sci', scilimits=(0,0))

# 设置x轴范围
plt.xlim(0.8, 3.2)

# 添加图例，分两组显示
handles, labels = plt.gca().get_legend_handles_labels()
# 创建两个图例，分别显示1h和19h的数据
legend1 = plt.legend(handles[:3], labels[:3], 
                    bbox_to_anchor=(0.5, -0.15),
                    loc='upper center',
                    ncol=3,
                    fontsize=12,
                    title='EXP1',
                    title_fontsize=12,
                    frameon=False,  # 移除图例边框
                    handletextpad=0.5,  # 增加图例标记和文本之间的距离
                    borderaxespad=0)
plt.gca().add_artist(legend1)

plt.legend(handles[3:], labels[3:],
          bbox_to_anchor=(0.5, -0.25),
          loc='upper center',
          ncol=3,
          fontsize=12,
          title='EXP2',
          title_fontsize=12,
          frameon=False,  # 移除图例边框
          handletextpad=0.5,  # 增加图例标记和文本之间的距离
          borderaxespad=0)

# 设置图形边框粗细
ax = plt.gca()
for spine in ax.spines.values():
    spine.set_linewidth(1.5)

# 调整布局
plt.tight_layout()
plt.subplots_adjust(bottom=0.25)  # 增加底部空间以容纳两个图例

# 显示图形
plt.show() 
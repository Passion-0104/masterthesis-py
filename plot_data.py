import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager

# 设置字体为Arial，并调大字号
plt.rcParams['font.sans-serif'] = ['Arial']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 16  # 全局字体大小
plt.rcParams['axes.titlesize'] = 20  # 标题字体大小
plt.rcParams['axes.labelsize'] = 18  # 坐标轴标签字体大小
plt.rcParams['xtick.labelsize'] = 16  # x轴刻度字体大小
plt.rcParams['ytick.labelsize'] = 16  # y轴刻度字体大小
plt.rcParams['legend.fontsize'] = 16  # 图例字体大小

# 新数据
# 区间
dist_sections = ['18–90', '90–290', '290–500', '500–550']
# 参考值
reference_min = np.array([0.0401, 0.0153, 0.0077, 0.0056])
reference_max = np.array([0.135734635, 0.1709, 0.1470, 0.1158])
# 实测值
measured_min = np.array([0.1354, 0.1264, 0.1174, 0.0834])
measured_max = np.array([0.1416, 0.1316, 0.1226, 0.0886])

x = np.arange(len(dist_sections))

plt.figure(figsize=(10, 6))

# 参考值 min-max 竖线和端点
plt.vlines(x-0.1, reference_min, reference_max, color='blue', linewidth=8, label='Reference', alpha=0.7)

# 实测值 min-max 竖线和端点
plt.vlines(x+0.1, measured_min, measured_max, color='orange', linewidth=8, label='Measured', alpha=0.7)

plt.xticks(x, dist_sections)
plt.xlabel('Temperature Section')
plt.ylabel('Measured Value')
plt.title('Pressure at different temperature sections')
plt.legend()
plt.grid(axis='y', linestyle='--', alpha=0.6)
plt.tight_layout()
plt.show() 
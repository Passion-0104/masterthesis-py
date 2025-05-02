import matplotlib.pyplot as plt
import numpy as np

# 设置字体
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 使用微软雅黑字体
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# Data definition
pre_vacuum_time = [1, 2, 18]  # Pre-vacuum time (hours)
measured_O2 = [1.55, 1.90, 1.64]  # O2 concentration (ppm)
error = [0.0775, 0.095, 0.082]  # ±5% error

# Create figure
plt.figure(figsize=(8, 5))  # 减小图表尺寸

# Create subplot with a box around it
ax = plt.gca()
ax.spines['top'].set_visible(True)
ax.spines['right'].set_visible(True)
ax.spines['left'].set_visible(True)
ax.spines['bottom'].set_visible(True)
for spine in ax.spines.values():
    spine.set_linewidth(1.0)

# Plot main data (points and connecting line)
x_positions = [1, 2, 3]  # Positions on x-axis (equidistant)
plt.plot(x_positions, measured_O2, 
        'o-', color='black',
        linewidth=1.5, 
        markersize=8,
        label='Measured O2 (ppm)')

# Add error bars (dashed style)
plt.errorbar(x_positions, measured_O2, yerr=error,
            fmt='none',  # No markers or lines
            ecolor='red',
            capsize=5,
            linestyle='--',  # Dashed error lines
            elinewidth=1,
            capthick=1,
            label='Relative error range')

# Axis and title settings
plt.xticks(x_positions, ['1h', '2h', '18h'], fontsize=12)  # 增大刻度字体
plt.xlabel('Pre-vacuum Time', fontsize=14)  # 增大x轴标签字体
plt.ylabel('Measured Oxygen Concentration (ppm)', fontsize=14)  # 增大y轴标签字体
plt.title('O2 Concentration at different Pre-vacuum Time', fontsize=16, pad=20)  # 增大标题字体

# Legend settings
plt.legend(loc='upper right', fontsize=12, frameon=True,  # 增大图例字体
          facecolor='white', edgecolor='gray')

# Set axis range
plt.xlim(0.5, 3.5)

# Adjust layout
plt.tight_layout()

# Show plot
plt.show() 
import matplotlib.pyplot as plt

# X 轴标签（温度阶段）与位置
stages = ["18–90°C", "90–290°C", "290–500°C", "500–550°C"]
x = [0, 1, 2, 3]  # 每个阶段的位置（横坐标）

# 数据
measured_o2 = [1.99e+00, 1.93e-18, 3.62e-19, 2.70e-18]
ref_min = [1.99] * 4
ref_max = [3.32] * 4
ref_0026 = [0.0026] * 4  # 添加0.0026ppm的参考线数据

# 创建图形
fig, ax = plt.subplots(figsize=(8, 5))

# 绘制散点图
ax.scatter(x, measured_o2, color='tab:blue', label='Measured O₂', marker='s', s=80, zorder=3)

# 绘制参考线（Min 和 Max 合并成一条线）
ax.plot(x, ref_min, color='orange', linewidth=2.5, label='reference level')
ax.plot(x, ref_max, color='orange', linewidth=2.5)

# 添加0.0026ppm参考线
ax.plot(x, ref_0026, '--', color='black', linewidth=2, label='Measuring range approved by the manufacturer')

# 设置坐标轴
ax.set_yscale('log')  # 对数坐标
ax.set_xticks(x)
ax.set_xticklabels(stages)
ax.set_xlabel("Temperature Range (°C)")
ax.set_ylabel("O₂ Concentration (ppm)")
ax.set_title("O₂ Concentration at Different Heating Stages")

# 设置黑色边框
for spine in ax.spines.values():
    spine.set_color('black')
    spine.set_linewidth(1.0)

# 移除网格线
ax.grid(False)

# 图例与美化
ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
plt.tight_layout()
plt.show()

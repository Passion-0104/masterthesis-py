import matplotlib.pyplot as plt
import numpy as np

# Data definition
flushes = [1, 2, 3, 4, 6]  # Number of Argon flushes
measured_O2 = [1.80, 1.50, 1.65, 1.58, 1.58]  # O2 concentration (ppm)
error = [0.09, 0.075, 0.0825, 0.079, 0.079]  # ±5% error

# Create figure
plt.figure(figsize=(8, 5))

# Create subplot with a box around it
ax = plt.gca()
ax.spines['top'].set_visible(True)
ax.spines['right'].set_visible(True)
ax.spines['left'].set_visible(True)
ax.spines['bottom'].set_visible(True)
for spine in ax.spines.values():
    spine.set_linewidth(1.0)

# Plot main data (points and connecting line)
x_positions = range(1, len(flushes) + 1)  # Positions on x-axis (equidistant)
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
plt.xticks(x_positions, [f'{n} time{"s" if n > 1 else ""}' for n in flushes], fontsize=12)
plt.xlabel('Number of Argon Flushes', fontsize=14)
plt.ylabel('Measured Oxygen Concentration (ppm)', fontsize=14)
plt.title('O2 Concentration at different number of Argon Flushes', fontsize=16, pad=20)

# Set y-axis range to show data clearly
plt.ylim(1.3, 2.0)

# Legend settings
plt.legend(loc='upper right', fontsize=12, frameon=True,
          facecolor='white', edgecolor='gray')

# Set axis range
plt.xlim(0.5, len(flushes) + 0.5)

# Adjust layout
plt.tight_layout()

# Show plot
plt.show() 
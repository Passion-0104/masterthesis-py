import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt

# ✅ Step 1: Your experimental data
pressures = np.array([0.051, 0.103, 0.205, 0.502, 1.005])  # in bar
ppm_measured = np.array([2, 6, 13, 26, 41])                # in ppm

# ✅ Step 2: Set reference pressure (1 bar for normalization)
p_ref = 1.0

# ✅ Step 3: Define your calibration formula
def calibrate(ppm, p, f1, f2):
    ratio = p_ref / p
    exponent = f1 * np.log(ratio) + f2
    return ppm * (ratio ** exponent)

# ✅ Step 4: Target reference value: assume all should calibrate to 41 ppm
ppm_target = np.full_like(ppm_measured, 41.0)

# Fitting wrapper (only p is variable, ppm_measured is constant input)
def fit_func(p, f1, f2):
    return calibrate(ppm_measured, p, f1, f2)

# ✅ Step 5: Fit the model to get f1 and f2 (faktor)
params_opt, _ = curve_fit(fit_func, pressures, ppm_target, p0=(1.0, 0.0))
f1, f2 = params_opt

# ✅ Step 6: Apply calibration
ppm_calibrated = calibrate(ppm_measured, pressures, f1, f2)

# ✅ Step 7: Display results
print("\n✅ Calibration complete.")
print(f"Fitted faktor values:")
print(f"  f1 = {f1:.6f}")
print(f"  f2 = {f2:.6f}")

df = pd.DataFrame({
    'Pressure (bar)': pressures,
    'Measured ppm': ppm_measured,
    'Calibrated ppm (→1 bar)': np.round(ppm_calibrated, 2)
})
print("\n📊 Calibration Table:")
print(df.to_string(index=False))

# ✅ Step 8: Plot
plt.plot(pressures, ppm_measured, 'o-', label='Measured ppm')
plt.plot(pressures, ppm_calibrated, 's--', label='Calibrated ppm (→0.98 bar)')
plt.axhline(41, color='gray', linestyle=':', label='Target: 41 ppm')
plt.xlabel('Measured Pressure (bar)')
plt.ylabel('Humidity (ppm)')
plt.title('Calibration of Humidity to Reference Pressure (0.98 bar)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

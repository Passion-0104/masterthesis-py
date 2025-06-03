"""
Calibration core module for moisture data calibration
"""

import numpy as np
import pandas as pd


class CalibrationCore:
    """Core calibration functionality for moisture data"""
    
    def __init__(self):
        self.f1 = 0.196798
        self.f2 = 0.419073
        self.p_ref = 1.0
        
    def set_parameters(self, f1, f2, p_ref):
        """Set calibration parameters"""
        self.f1 = f1
        self.f2 = f2
        self.p_ref = p_ref
        
    def get_parameters(self):
        """Get current calibration parameters"""
        return {
            'f1': self.f1,
            'f2': self.f2,
            'p_ref': self.p_ref
        }
        
    def calibrate(self, ppm, pressure):
        """
        Calculate calibrated ppm value based on calibration formula
        ppm_calibrated = concentration × (p_ref/p)^(f1·ln(p_ref/p)+f2)
        """
        # Ensure pressure is positive and valid
        pressure = np.maximum(pressure, 1e-10)
        
        ratio = self.p_ref / pressure
        exponent = self.f1 * np.log(ratio) + self.f2
        
        # Limit exponent range to prevent numerical explosion
        exponent = np.clip(exponent, -10, 10)
        
        return ppm * (ratio ** exponent)
        
    def calibrate_single_pair(self, plot_df, moisture_col, pressure_col):
        """Calibrate a single moisture-pressure pair"""
        if moisture_col not in plot_df.columns or pressure_col not in plot_df.columns:
            return None
            
        # Ensure numeric data
        moisture_data = pd.to_numeric(plot_df[moisture_col], errors='coerce')
        pressure_data = pd.to_numeric(plot_df[pressure_col], errors='coerce')
        
        # Remove invalid data (NaN or non-positive pressure)
        valid_mask = ~(pd.isna(moisture_data) | 
                      pd.isna(pressure_data) | 
                      (pressure_data <= 0))
        
        valid_df = plot_df[valid_mask].copy()
        
        if valid_df.empty:
            return None
            
        # Apply calibration
        calibrated_values = self.calibrate(
            valid_df[moisture_col].values, 
            valid_df[pressure_col].values
        )
        
        calib_col = f"{moisture_col}_calib"
        
        return {
            'column': moisture_col,
            'calib_column': calib_col,
            'times': valid_df['relative_time'].values if 'relative_time' in valid_df.columns else valid_df.index.values,
            'values': calibrated_values,
            'valid_df': valid_df,
            'original_values': valid_df[moisture_col].values,
            'pressure_values': valid_df[pressure_col].values
        }
        
    def batch_calibrate(self, plot_df, moisture_pressure_pairs):
        """Calibrate multiple moisture-pressure pairs"""
        calibrated_data = {}
        
        for moisture_col, pressure_col in moisture_pressure_pairs:
            result = self.calibrate_single_pair(plot_df, moisture_col, pressure_col)
            if result:
                calib_col = result['calib_column']
                calibrated_data[calib_col] = result
                
        return calibrated_data
        
    def calculate_calibration_statistics(self, calibrated_data):
        """Calculate statistics for calibrated data"""
        if not calibrated_data:
            return {}
            
        stats = {}
        
        for calib_col, data_dict in calibrated_data.items():
            values = data_dict['values']
            original_values = data_dict['original_values']
            
            # Basic statistics
            stats[calib_col] = {
                'mean': np.mean(values),
                'std': np.std(values),
                'min': np.min(values),
                'max': np.max(values),
                'range': np.max(values) - np.min(values),
                'original_mean': np.mean(original_values),
                'original_std': np.std(original_values),
                'calibration_factor': np.mean(values) / np.mean(original_values) if np.mean(original_values) != 0 else 1.0
            }
            
        return stats
        
    def validate_calibration(self, calibrated_data, tolerance=0.1):
        """Validate calibration results"""
        if len(calibrated_data) < 2:
            return {"status": "insufficient_data", "message": "Need at least 2 calibrated datasets for validation"}
            
        # Get all calibrated values at each time point
        all_values_by_time = {}
        
        for calib_col, data_dict in calibrated_data.items():
            times = data_dict['times']
            values = data_dict['values']
            
            for t, v in zip(times, values):
                time_key = round(t, 3)  # Round to avoid floating point issues
                if time_key not in all_values_by_time:
                    all_values_by_time[time_key] = []
                all_values_by_time[time_key].append(v)
        
        # Calculate variance at each time point
        variances = []
        for time_key, values_list in all_values_by_time.items():
            if len(values_list) > 1:
                variance = np.var(values_list)
                variances.append(variance)
        
        if not variances:
            return {"status": "no_overlap", "message": "No overlapping time points found"}
            
        avg_variance = np.mean(variances)
        max_variance = np.max(variances)
        
        # Determine validation status
        if max_variance < tolerance:
            status = "good"
            message = f"Calibration validation passed. Max variance: {max_variance:.4f}"
        elif avg_variance < tolerance:
            status = "acceptable"
            message = f"Calibration validation acceptable. Average variance: {avg_variance:.4f}, Max variance: {max_variance:.4f}"
        else:
            status = "poor"
            message = f"Calibration validation failed. High variance detected. Average: {avg_variance:.4f}, Max: {max_variance:.4f}"
            
        return {
            "status": status,
            "message": message,
            "avg_variance": avg_variance,
            "max_variance": max_variance,
            "num_time_points": len(all_values_by_time)
        } 
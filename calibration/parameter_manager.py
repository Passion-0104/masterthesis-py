"""
Parameter manager for calibration parameters and pair configurations
"""

import json
import os
from datetime import datetime
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import pandas as pd


class ParameterManager:
    """Manager for calibration parameters and pair configurations"""
    
    def __init__(self):
        self.recent_pairs_file = os.path.join(os.path.expanduser("~"), "h2o_calibration_recent_pairs.json")
        
    def export_parameters(self, parameters, parent_widget):
        """Export calibration parameters to JSON file"""
        try:
            # Add timestamp
            export_data = parameters.copy()
            export_data['export_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # Request save location
            file_path, _ = QFileDialog.getSaveFileName(
                parent_widget, "Export Parameters", "", 
                "JSON files (*.json);;All files (*)"
            )
            
            if not file_path:
                return False
                
            # Save to file
            with open(file_path, 'w') as f:
                json.dump(export_data, f, indent=4)
                
            QMessageBox.information(parent_widget, "Success", f"Parameters exported to {file_path}")
            return True
            
        except Exception as e:
            QMessageBox.critical(parent_widget, "Error", f"Failed to export parameters: {str(e)}")
            return False
            
    def import_parameters(self, parent_widget):
        """Import calibration parameters from JSON file"""
        try:
            # Request file location
            file_path, _ = QFileDialog.getOpenFileName(
                parent_widget, "Import Parameters", "", 
                "JSON files (*.json);;All files (*)"
            )
            
            if not file_path:
                return None
                
            # Load from file
            with open(file_path, 'r') as f:
                data = json.load(f)
                
            # Validate required fields
            required_fields = ['f1', 'f2', 'p_ref']
            if not all(field in data for field in required_fields):
                QMessageBox.critical(parent_widget, "Error", "Invalid parameter file format")
                return None
                
            # Extract parameters
            parameters = {
                'f1': float(data['f1']),
                'f2': float(data['f2']),
                'p_ref': float(data['p_ref'])
            }
            
            # Show import information
            info_msg = f"Imported parameters:\nf1 = {parameters['f1']:.6f}\nf2 = {parameters['f2']:.6f}\np_ref = {parameters['p_ref']:.2f}"
            if 'export_date' in data:
                info_msg += f"\n\nExported on: {data['export_date']}"
                
            QMessageBox.information(parent_widget, "Parameters Imported", info_msg)
            
            return parameters
            
        except Exception as e:
            QMessageBox.critical(parent_widget, "Error", f"Failed to import parameters: {str(e)}")
            return None
            
    def export_pairs_config(self, pairs_dict, parent_widget):
        """Export moisture-pressure pair configuration"""
        try:
            # Add metadata
            export_data = {
                'pairs': pairs_dict,
                'export_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'version': '1.0'
            }
            
            # Request save location
            file_path, _ = QFileDialog.getSaveFileName(
                parent_widget, "Export Pairs Configuration", "", 
                "JSON files (*.json);;All files (*)"
            )
            
            if not file_path:
                return False
                
            # Save to file
            with open(file_path, 'w') as f:
                json.dump(export_data, f, indent=4)
                
            QMessageBox.information(parent_widget, "Success", f"Pairs configuration exported to {file_path}")
            return True
            
        except Exception as e:
            QMessageBox.critical(parent_widget, "Error", f"Failed to export pairs: {str(e)}")
            return False
            
    def import_pairs_config(self, parent_widget):
        """Import moisture-pressure pair configuration"""
        try:
            # Request file location
            file_path, _ = QFileDialog.getOpenFileName(
                parent_widget, "Import Pairs Configuration", "", 
                "JSON files (*.json);;All files (*)"
            )
            
            if not file_path:
                return None
                
            # Load from file
            with open(file_path, 'r') as f:
                data = json.load(f)
                
            # Extract pairs data
            if 'pairs' in data:
                pairs_dict = data['pairs']
            else:
                # Assume the whole file is the pairs data (legacy format)
                pairs_dict = data
                
            # Show import information
            pair_count = len([k for k in pairs_dict.keys() if k.startswith('pair_')])
            info_msg = f"Imported {pair_count} pairs configuration"
            if 'export_date' in data:
                info_msg += f"\n\nExported on: {data['export_date']}"
                
            QMessageBox.information(parent_widget, "Pairs Imported", info_msg)
            
            return pairs_dict
            
        except Exception as e:
            QMessageBox.critical(parent_widget, "Error", f"Failed to import pairs: {str(e)}")
            return None
            
    def save_recent_pairs(self, file_path, pairs_dict):
        """Save recently used pairs for specific file"""
        try:
            if not file_path or not pairs_dict:
                return False
                
            # Load existing recent pairs
            recent_pairs = {}
            if os.path.exists(self.recent_pairs_file):
                try:
                    with open(self.recent_pairs_file, 'r') as f:
                        recent_pairs = json.load(f)
                except:
                    recent_pairs = {}
                    
            # Add current pairs
            file_key = os.path.basename(file_path)
            recent_pairs[file_key] = {
                'pairs': pairs_dict,
                'save_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'file_path': file_path
            }
            
            # Save back to file
            with open(self.recent_pairs_file, 'w') as f:
                json.dump(recent_pairs, f, indent=4)
                
            return True
            
        except Exception as e:
            print(f"Failed to save recent pairs: {str(e)}")
            return False
            
    def load_recent_pairs(self, file_path):
        """Load recently used pairs for specific file"""
        try:
            if not file_path or not os.path.exists(self.recent_pairs_file):
                return None
                
            with open(self.recent_pairs_file, 'r') as f:
                recent_pairs = json.load(f)
                
            file_key = os.path.basename(file_path)
            
            if file_key in recent_pairs:
                return recent_pairs[file_key]['pairs']
                
            return None
            
        except Exception as e:
            print(f"Failed to load recent pairs: {str(e)}")
            return None
            
    def clear_recent_pairs(self, file_path=None):
        """Clear recent pairs (all or for specific file)"""
        try:
            if not os.path.exists(self.recent_pairs_file):
                return True
                
            if file_path is None:
                # Clear all recent pairs
                os.remove(self.recent_pairs_file)
                return True
            else:
                # Clear for specific file
                with open(self.recent_pairs_file, 'r') as f:
                    recent_pairs = json.load(f)
                    
                file_key = os.path.basename(file_path)
                if file_key in recent_pairs:
                    del recent_pairs[file_key]
                    
                    with open(self.recent_pairs_file, 'w') as f:
                        json.dump(recent_pairs, f, indent=4)
                        
                return True
                
        except Exception as e:
            print(f"Failed to clear recent pairs: {str(e)}")
            return False
            
    def get_recent_pairs_list(self):
        """Get list of recent pairs"""
        try:
            if not os.path.exists(self.recent_pairs_file):
                return []
                
            with open(self.recent_pairs_file, 'r') as f:
                recent_pairs = json.load(f)
                
            pairs_list = []
            for file_key, data in recent_pairs.items():
                pairs_info = {
                    'file_name': file_key,
                    'file_path': data.get('file_path', ''),
                    'save_date': data.get('save_date', ''),
                    'pair_count': len([k for k in data.get('pairs', {}).keys() if k.startswith('pair_')])
                }
                pairs_list.append(pairs_info)
                
            # Sort by save date (newest first)
            pairs_list.sort(key=lambda x: x['save_date'], reverse=True)
            
            return pairs_list
            
        except Exception as e:
            print(f"Failed to get recent pairs list: {str(e)}")
            return []
            
    def validate_pairs_config(self, pairs_dict, available_columns):
        """Validate pairs configuration against available columns"""
        if not pairs_dict or not available_columns:
            return {"valid": False, "message": "No pairs or columns available"}
            
        valid_pairs = {}
        invalid_pairs = {}
        
        for pair_key, pair_data in pairs_dict.items():
            if not pair_key.startswith('pair_'):
                continue
                
            if not isinstance(pair_data, dict):
                invalid_pairs[pair_key] = "Invalid pair data format"
                continue
                
            moisture_col = pair_data.get('moisture', '')
            pressure_col = pair_data.get('pressure', '')
            
            if not moisture_col or not pressure_col:
                invalid_pairs[pair_key] = "Missing moisture or pressure column"
                continue
                
            if moisture_col not in available_columns:
                invalid_pairs[pair_key] = f"Moisture column '{moisture_col}' not found"
                continue
                
            if pressure_col not in available_columns:
                invalid_pairs[pair_key] = f"Pressure column '{pressure_col}' not found"
                continue
                
            valid_pairs[pair_key] = pair_data
            
        return {
            "valid": len(valid_pairs) > 0,
            "valid_pairs": valid_pairs,
            "invalid_pairs": invalid_pairs,
            "message": f"Found {len(valid_pairs)} valid pairs, {len(invalid_pairs)} invalid pairs"
        } 
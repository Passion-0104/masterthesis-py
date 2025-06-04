"""
Main window for H2O Data Visualization and Calibration Tool
"""

import sys
import os
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QPushButton, QLabel, QComboBox, QListWidget, 
                           QCheckBox, QSpinBox, QDoubleSpinBox, QGroupBox,
                           QGridLayout, QMessageBox, QFileDialog, QRadioButton,
                           QButtonGroup, QFrame, QScrollArea, QApplication,
                           QAbstractItemView, QSplitter, QTextEdit, QDialog,
                           QDialogButtonBox)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont

from data_processing.data_loader import DataLoader
from calibration.calibration_core import CalibrationCore
from calibration.parameter_manager import ParameterManager
from ui.plot_widget import IndependentPlotWindow
from utils.file_utils import FileUtils


class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.data_loader = DataLoader()
        self.calibration_core = CalibrationCore()
        self.parameter_manager = ParameterManager()
        self.plot_windows = []  # Track opened plot windows
        
        self.setup_ui()
        self.setup_connections()
        
    def setup_ui(self):
        """Setup the main user interface"""
        self.setWindowTitle("H2O Data Visualization and Calibration Tool")
        self.setGeometry(100, 100, 1200, 800)  # Restore reasonable window size
        
        # Create central widget with scroll area
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        self.setCentralWidget(scroll)
        
        # Main layout with proper spacing
        main_layout = QVBoxLayout(scroll_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # File selection section (full width)
        self.create_file_section(main_layout)
        
        # Create two-column layout
        columns_layout = QHBoxLayout()
        columns_layout.setSpacing(15)
        
        # Left column
        left_column = QVBoxLayout()
        left_column.setSpacing(10)
        self.create_column_section(left_column)
        self.create_time_section(left_column)
        
        # Right column  
        right_column = QVBoxLayout()
        right_column.setSpacing(10)
        self.create_calibration_section(right_column)
        self.create_difference_section(right_column)
        
        # Add columns to layout
        columns_layout.addLayout(left_column, 1)  # Equal weight
        columns_layout.addLayout(right_column, 1)  # Equal weight
        
        main_layout.addLayout(columns_layout)
        
        # Action buttons section (full width)
        self.create_action_section(main_layout)
        
        # Initially disable most controls
        self.set_controls_enabled(False)
        
    def create_file_section(self, parent_layout):
        """Create file selection section"""
        file_group = QGroupBox("File Selection")
        file_group.setFont(QFont("Arial", 12))  # Academic standard size
        file_layout = QHBoxLayout(file_group)
        file_layout.setSpacing(10)
        file_layout.setContentsMargins(10, 10, 10, 10)
        
        self.file_button = QPushButton("Select Excel File")
        self.file_button.setFont(QFont("Arial", 11))
        self.file_button.setMinimumHeight(32)
        
        self.file_label = QLabel("No file selected")
        self.file_label.setFont(QFont("Arial", 11))
        
        file_layout.addWidget(self.file_button)
        file_layout.addWidget(self.file_label, 1)
        
        parent_layout.addWidget(file_group)
        
    def create_column_section(self, parent_layout):
        """Create column selection section"""
        column_group = QGroupBox("Column Selection")
        column_group.setFont(QFont("Arial", 12))
        column_layout = QVBoxLayout(column_group)
        column_layout.setSpacing(8)
        column_layout.setContentsMargins(10, 10, 10, 10)
        
        # Instructions
        instructions = QLabel("Select Columns to Plot (Ctrl+Click for multiple selection):")
        instructions.setFont(QFont("Arial", 10))  # Smaller instruction text
        column_layout.addWidget(instructions)
        
        # Buttons row
        button_row = QHBoxLayout()
        button_row.setSpacing(10)
        
        self.clear_selection_btn = QPushButton("Clear Selection")
        self.clear_selection_btn.setFont(QFont("Arial", 10))
        self.clear_selection_btn.setMinimumHeight(28)
        
        self.auto_match_btn = QPushButton("Auto-Match Pairs")
        self.auto_match_btn.setFont(QFont("Arial", 10))
        self.auto_match_btn.setMinimumHeight(28)
        
        button_row.addWidget(self.clear_selection_btn)
        button_row.addWidget(self.auto_match_btn)
        button_row.addStretch()
        column_layout.addLayout(button_row)
        
        # Column list
        self.column_list = QListWidget()
        self.column_list.setFont(QFont("Arial", 10))
        self.column_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.column_list.setMinimumHeight(150)
        column_layout.addWidget(self.column_list)
        
        parent_layout.addWidget(column_group)
        
    def create_time_section(self, parent_layout):
        """Create time settings section"""
        time_group = QGroupBox("Time Settings")
        time_group.setFont(QFont("Arial", 12))
        time_layout = QVBoxLayout(time_group)
        time_layout.setSpacing(8)
        time_layout.setContentsMargins(10, 10, 10, 10)
        
        # Time column selection
        time_col_layout = QHBoxLayout()
        time_col_layout.setSpacing(10)
        
        time_label = QLabel("Time Column:")
        time_label.setFont(QFont("Arial", 11))
        time_col_layout.addWidget(time_label)
        
        self.time_column_combo = QComboBox()
        self.time_column_combo.setFont(QFont("Arial", 10))
        self.time_column_combo.setMinimumHeight(28)
        time_col_layout.addWidget(self.time_column_combo, 1)
        
        time_layout.addLayout(time_col_layout)
        
        # Time range settings
        range_label = QLabel("Time Range:")
        range_label.setFont(QFont("Arial", 11))
        time_layout.addWidget(range_label)
        
        self.time_range_group = QButtonGroup()
        
        self.radio_all = QRadioButton("Show All Data")
        self.radio_all.setFont(QFont("Arial", 10))
        self.radio_2hours = QRadioButton("First 2 Hours Only")
        self.radio_2hours.setFont(QFont("Arial", 10))
        self.radio_custom = QRadioButton("Custom Range")
        self.radio_custom.setFont(QFont("Arial", 10))
        
        self.radio_2hours.setChecked(True)  # Default
        
        self.time_range_group.addButton(self.radio_all, 0)
        self.time_range_group.addButton(self.radio_2hours, 1)
        self.time_range_group.addButton(self.radio_custom, 2)
        
        time_layout.addWidget(self.radio_all)
        time_layout.addWidget(self.radio_2hours)
        time_layout.addWidget(self.radio_custom)
        
        # Custom range inputs
        custom_layout = QHBoxLayout()
        custom_layout.setSpacing(10)
        
        start_label = QLabel("Start Time (h):")
        start_label.setFont(QFont("Arial", 11))
        custom_layout.addWidget(start_label)
        
        self.start_time_spin = QDoubleSpinBox()
        self.start_time_spin.setFont(QFont("Arial", 10))
        self.start_time_spin.setRange(0, 100)
        self.start_time_spin.setValue(0.0)
        self.start_time_spin.setMinimumHeight(28)
        custom_layout.addWidget(self.start_time_spin)
        
        end_label = QLabel("End Time (h):")
        end_label.setFont(QFont("Arial", 11))
        custom_layout.addWidget(end_label)
        
        self.end_time_spin = QDoubleSpinBox()
        self.end_time_spin.setFont(QFont("Arial", 10))
        self.end_time_spin.setRange(0, 100)
        self.end_time_spin.setValue(2.0)
        self.end_time_spin.setMinimumHeight(28)
        custom_layout.addWidget(self.end_time_spin)
        
        time_layout.addLayout(custom_layout)
        parent_layout.addWidget(time_group)
        
    def create_calibration_section(self, parent_layout):
        """Create calibration settings section"""
        calib_group = QGroupBox("Moisture Calibration Settings")
        calib_group.setFont(QFont("Arial", 12))
        calib_layout = QVBoxLayout(calib_group)
        calib_layout.setSpacing(8)
        calib_layout.setContentsMargins(10, 10, 10, 10)
        
        # Enable calibration checkboxes
        calib_options_layout = QVBoxLayout()
        calib_options_layout.setSpacing(5)
        
        self.enable_calibration_cb = QCheckBox("Enable Moisture Calibration")
        self.enable_calibration_cb.setFont(QFont("Arial", 11))
        calib_options_layout.addWidget(self.enable_calibration_cb)
        
        options_row = QHBoxLayout()
        options_row.setSpacing(15)
        
        self.show_original_cb = QCheckBox("Show Original Data")
        self.show_original_cb.setFont(QFont("Arial", 10))
        self.show_original_cb.setChecked(True)
        options_row.addWidget(self.show_original_cb)
        
        self.show_error_cb = QCheckBox("Show Error Range")
        self.show_error_cb.setFont(QFont("Arial", 10))
        options_row.addWidget(self.show_error_cb)
        
        calib_options_layout.addLayout(options_row)
        
        # Error value
        error_layout = QHBoxLayout()
        error_layout.setSpacing(10)
        
        error_label = QLabel("Error Value (ppm):")
        error_label.setFont(QFont("Arial", 11))
        error_layout.addWidget(error_label)
        
        self.error_value_spin = QDoubleSpinBox()
        self.error_value_spin.setFont(QFont("Arial", 10))
        self.error_value_spin.setRange(0, 100)
        self.error_value_spin.setValue(10.0)
        self.error_value_spin.setMinimumHeight(28)
        error_layout.addWidget(self.error_value_spin)
        error_layout.addStretch()
        
        calib_options_layout.addLayout(error_layout)
        calib_layout.addLayout(calib_options_layout)
        
        # Calibration parameters
        param_label = QLabel("Calibration Parameters:")
        param_label.setFont(QFont("Arial", 11))
        calib_layout.addWidget(param_label)
        
        param_grid = QGridLayout()
        param_grid.setSpacing(10)
        
        # f1 parameter
        f1_label = QLabel("f1:")
        f1_label.setFont(QFont("Arial", 11))
        param_grid.addWidget(f1_label, 0, 0)
        
        self.f1_spin = QDoubleSpinBox()
        self.f1_spin.setFont(QFont("Arial", 10))
        self.f1_spin.setDecimals(6)
        self.f1_spin.setRange(-10, 10)
        self.f1_spin.setValue(0.196798)
        self.f1_spin.setMinimumHeight(28)
        param_grid.addWidget(self.f1_spin, 0, 1)
        
        # f2 parameter
        f2_label = QLabel("f2:")
        f2_label.setFont(QFont("Arial", 11))
        param_grid.addWidget(f2_label, 1, 0)
        
        self.f2_spin = QDoubleSpinBox()
        self.f2_spin.setFont(QFont("Arial", 10))
        self.f2_spin.setDecimals(6)
        self.f2_spin.setRange(-10, 10)
        self.f2_spin.setValue(0.419073)
        self.f2_spin.setMinimumHeight(28)
        param_grid.addWidget(self.f2_spin, 1, 1)
        
        # p_ref parameter
        p_ref_label = QLabel("p_ref:")
        p_ref_label.setFont(QFont("Arial", 11))
        param_grid.addWidget(p_ref_label, 2, 0)
        
        self.p_ref_spin = QDoubleSpinBox()
        self.p_ref_spin.setFont(QFont("Arial", 10))
        self.p_ref_spin.setDecimals(2)
        self.p_ref_spin.setRange(0.1, 10.0)
        self.p_ref_spin.setValue(1.0)
        self.p_ref_spin.setMinimumHeight(28)
        param_grid.addWidget(self.p_ref_spin, 2, 1)
        
        calib_layout.addLayout(param_grid)
        
        # Parameter buttons
        param_btn_layout = QHBoxLayout()
        param_btn_layout.setSpacing(10)
        
        self.import_params_btn = QPushButton("Import Parameters")
        self.import_params_btn.setFont(QFont("Arial", 10))
        self.import_params_btn.setMinimumHeight(28)
        param_btn_layout.addWidget(self.import_params_btn)
        
        self.export_params_btn = QPushButton("Export Parameters")
        self.export_params_btn.setFont(QFont("Arial", 10))
        self.export_params_btn.setMinimumHeight(28)
        param_btn_layout.addWidget(self.export_params_btn)
        
        calib_layout.addLayout(param_btn_layout)
        
        # Moisture-Pressure pairs
        pairs_group = QGroupBox("Moisture-Pressure Pairs")
        pairs_group.setFont(QFont("Arial", 12))
        pairs_layout = QVBoxLayout(pairs_group)
        pairs_layout.setSpacing(8)
        pairs_layout.setContentsMargins(10, 10, 10, 10)
        
        # Pairs control buttons
        pairs_btn_layout = QHBoxLayout()
        pairs_btn_layout.setSpacing(10)
        
        self.clear_pairs_btn = QPushButton("Clear All Pairs")
        self.clear_pairs_btn.setFont(QFont("Arial", 10))
        self.clear_pairs_btn.setMinimumHeight(28)
        pairs_btn_layout.addWidget(self.clear_pairs_btn)
        
        self.save_pairs_btn = QPushButton("Save Pairs")
        self.save_pairs_btn.setFont(QFont("Arial", 10))
        self.save_pairs_btn.setMinimumHeight(28)
        pairs_btn_layout.addWidget(self.save_pairs_btn)
        
        self.load_pairs_btn = QPushButton("Load Pairs")
        self.load_pairs_btn.setFont(QFont("Arial", 10))
        self.load_pairs_btn.setMinimumHeight(28)
        pairs_btn_layout.addWidget(self.load_pairs_btn)
        
        pairs_layout.addLayout(pairs_btn_layout)
        
        # Create pairs grid
        self.pairs_widget = QWidget()
        self.pairs_layout = QGridLayout(self.pairs_widget)
        self.pairs_layout.setSpacing(8)
        self.pairs_layout.setContentsMargins(5, 5, 5, 5)
        
        # Headers
        headers = ["No.", "Moisture Column", "Pressure Column", "Actions"]
        for i, header in enumerate(headers):
            label = QLabel(header)
            label.setFont(QFont("Arial", 10, QFont.Bold))
            self.pairs_layout.addWidget(label, 0, i)
        
        # Create 6 pairs
        self.moisture_combos = []
        self.pressure_combos = []
        self.pair_buttons = []
        
        for i in range(6):
            row = i + 1
            
            # Number
            num_label = QLabel(str(i + 1))
            num_label.setFont(QFont("Arial", 10))
            self.pairs_layout.addWidget(num_label, row, 0)
            
            # Moisture combo
            moisture_combo = QComboBox()
            moisture_combo.setFont(QFont("Arial", 9))
            moisture_combo.setMinimumHeight(26)
            self.moisture_combos.append(moisture_combo)
            self.pairs_layout.addWidget(moisture_combo, row, 1)
            
            # Pressure combo
            pressure_combo = QComboBox()
            pressure_combo.setFont(QFont("Arial", 9))
            pressure_combo.setMinimumHeight(26)
            self.pressure_combos.append(pressure_combo)
            self.pairs_layout.addWidget(pressure_combo, row, 2)
            
            # Action buttons
            action_layout = QHBoxLayout()
            action_layout.setSpacing(5)
            
            set_btn = QPushButton("Set")
            set_btn.setFont(QFont("Arial", 9))
            set_btn.setMinimumHeight(24)
            clear_btn = QPushButton("Clear")
            clear_btn.setFont(QFont("Arial", 9))
            clear_btn.setMinimumHeight(24)
            
            # Connect buttons with lambda to capture index
            set_btn.clicked.connect(lambda checked, idx=i: self.set_pair_from_selection(idx))
            clear_btn.clicked.connect(lambda checked, idx=i: self.clear_pair(idx))
            
            action_layout.addWidget(set_btn)
            action_layout.addWidget(clear_btn)
            
            action_widget = QWidget()
            action_widget.setLayout(action_layout)
            self.pairs_layout.addWidget(action_widget, row, 3)
        
        pairs_layout.addWidget(self.pairs_widget)
        calib_layout.addWidget(pairs_group)
        
        parent_layout.addWidget(calib_group)
        
    def create_difference_section(self, parent_layout):
        """Create difference calculation section"""
        diff_group = QGroupBox("Difference Calculation")
        diff_group.setFont(QFont("Arial", 12))
        diff_layout = QVBoxLayout(diff_group)
        diff_layout.setSpacing(8)
        diff_layout.setContentsMargins(10, 10, 10, 10)
        
        # Difference options
        self.enable_30min_diff_cb = QCheckBox("Calculate Difference at 30min")
        self.enable_30min_diff_cb.setFont(QFont("Arial", 10))
        diff_layout.addWidget(self.enable_30min_diff_cb)
        
        self.enable_multi_time_diff_cb = QCheckBox("Calculate at 20/40/60min")
        self.enable_multi_time_diff_cb.setFont(QFont("Arial", 10))
        diff_layout.addWidget(self.enable_multi_time_diff_cb)
        
        # Reference settings
        ref_grid = QGridLayout()
        ref_grid.setSpacing(10)
        
        ref1_label = QLabel("Reference Column:")
        ref1_label.setFont(QFont("Arial", 11))
        ref_grid.addWidget(ref1_label, 0, 0)
        
        self.reference_combo = QComboBox()
        self.reference_combo.setFont(QFont("Arial", 10))
        self.reference_combo.setMinimumHeight(28)
        ref_grid.addWidget(self.reference_combo, 0, 1)
        
        ref2_label = QLabel("Second Reference:")
        ref2_label.setFont(QFont("Arial", 11))
        ref_grid.addWidget(ref2_label, 1, 0)
        
        self.reference2_combo = QComboBox()
        self.reference2_combo.setFont(QFont("Arial", 10))
        self.reference2_combo.setMinimumHeight(28)
        ref_grid.addWidget(self.reference2_combo, 1, 1)
        
        window_label = QLabel("Time Window (min):")
        window_label.setFont(QFont("Arial", 11))
        ref_grid.addWidget(window_label, 2, 0)
        
        self.time_window_spin = QDoubleSpinBox()
        self.time_window_spin.setFont(QFont("Arial", 10))
        self.time_window_spin.setRange(0.1, 30.0)
        self.time_window_spin.setValue(5.0)
        self.time_window_spin.setMinimumHeight(28)
        ref_grid.addWidget(self.time_window_spin, 2, 1)
        
        diff_layout.addLayout(ref_grid)
        
        # Set reference button
        self.set_ref_btn = QPushButton("Set Reference from Selection")
        self.set_ref_btn.setFont(QFont("Arial", 10))
        self.set_ref_btn.setMinimumHeight(28)
        diff_layout.addWidget(self.set_ref_btn)
        
        parent_layout.addWidget(diff_group)
        
    def create_action_section(self, parent_layout):
        """Create action buttons section"""
        action_group = QGroupBox("Actions")
        action_group.setFont(QFont("Arial", 12))
        action_layout = QHBoxLayout(action_group)
        action_layout.setSpacing(20)
        action_layout.setContentsMargins(10, 10, 10, 10)
        
        self.plot_button = QPushButton("Generate Chart")
        self.plot_button.setFont(QFont("Arial", 12, QFont.Bold))  # Academic paper standard
        self.plot_button.setMinimumHeight(36)
        self.plot_button.setMinimumWidth(140)
        
        self.export_button = QPushButton("Export Data")
        self.export_button.setFont(QFont("Arial", 12, QFont.Bold))  # Academic paper standard
        self.export_button.setMinimumHeight(36)
        self.export_button.setMinimumWidth(140)
        
        action_layout.addWidget(self.plot_button)
        action_layout.addWidget(self.export_button)
        action_layout.addStretch()
        
        parent_layout.addWidget(action_group)
        
    def setup_connections(self):
        """Setup signal connections"""
        # File operations
        self.file_button.clicked.connect(self.select_file)
        
        # Column operations
        self.clear_selection_btn.clicked.connect(self.clear_column_selection)
        self.auto_match_btn.clicked.connect(self.auto_match_pairs)
        
        # Calibration operations
        self.enable_calibration_cb.toggled.connect(self.toggle_calibration_controls)
        self.import_params_btn.clicked.connect(self.import_parameters)
        self.export_params_btn.clicked.connect(self.export_parameters)
        
        # Pairs operations
        self.clear_pairs_btn.clicked.connect(self.clear_all_pairs)
        self.save_pairs_btn.clicked.connect(self.save_pairs)
        self.load_pairs_btn.clicked.connect(self.load_pairs)
        
        # Reference operations
        self.set_ref_btn.clicked.connect(self.set_reference_from_selection)
        
        # Main actions
        self.plot_button.clicked.connect(self.generate_chart)
        self.export_button.clicked.connect(self.export_data)
        
        # Initially toggle calibration controls
        self.toggle_calibration_controls()
        
    def set_controls_enabled(self, enabled):
        """Enable or disable controls based on data availability"""
        controls = [
            self.column_list, self.time_column_combo, self.clear_selection_btn,
            self.auto_match_btn, self.plot_button, self.export_button,
            self.reference_combo, self.reference2_combo, self.set_ref_btn
        ]
        
        for control in controls:
            control.setEnabled(enabled)
            
        # Also enable/disable pair combos
        for moisture_combo, pressure_combo in zip(self.moisture_combos, self.pressure_combos):
            moisture_combo.setEnabled(enabled)
            pressure_combo.setEnabled(enabled)
        
    def select_file(self):
        """Select and load Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", 
            "Excel files (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # Load file
                data = self.data_loader.load_file(file_path)
                
                # Update UI
                self.file_label.setText(os.path.basename(file_path))
                self.update_column_lists()
                self.set_controls_enabled(True)
                
                # Try to load recent pairs
                self.load_recent_pairs()
                
                QMessageBox.information(self, "Success", "File loaded successfully!")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")
                
    def update_column_lists(self):
        """Update all column lists and combos"""
        if not self.data_loader.is_data_loaded():
            return
            
        columns = self.data_loader.get_columns()
        time_columns = self.data_loader.get_time_columns()
        
        # Update column list
        self.column_list.clear()
        self.column_list.addItems(columns)
        
        # Update combos
        for combo in [self.time_column_combo, self.reference_combo, 
                     self.reference2_combo] + self.moisture_combos + self.pressure_combos:
            combo.clear()
            combo.addItems(columns)
            
        # Set default time column
        if time_columns:
            self.time_column_combo.setCurrentText(time_columns[0])
        elif columns:
            self.time_column_combo.setCurrentText(columns[0])
            
    def clear_column_selection(self):
        """Clear column selection"""
        self.column_list.clearSelection()
        
    def auto_match_pairs(self):
        """Automatically match moisture-pressure pairs"""
        if not self.data_loader.is_data_loaded():
            return
            
        try:
            pairs = self.data_loader.auto_match_pairs()
            
            # Clear existing pairs
            self.clear_all_pairs()
            
            # Set matched pairs
            for i, (moisture_col, pressure_col) in enumerate(pairs):
                if i < len(self.moisture_combos):
                    self.moisture_combos[i].setCurrentText(moisture_col)
                    self.pressure_combos[i].setCurrentText(pressure_col)
                    
            # Save recent pairs
            self.save_recent_pairs()
            
            if pairs:
                QMessageBox.information(self, "Auto-Match", 
                                      f"Successfully matched {len(pairs)} moisture-pressure pairs")
            else:
                QMessageBox.warning(self, "Auto-Match", 
                                   "No valid moisture-pressure pairs found. Try selecting them manually.")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Auto-match failed: {str(e)}")
            
    def toggle_calibration_controls(self):
        """Enable/disable calibration controls"""
        enabled = self.enable_calibration_cb.isChecked()
        
        controls = [
            self.f1_spin, self.f2_spin, self.p_ref_spin,
            self.show_original_cb, self.show_error_cb, self.error_value_spin
        ] + self.moisture_combos + self.pressure_combos
        
        for control in controls:
            control.setEnabled(enabled)
            
    def set_pair_from_selection(self, pair_index):
        """Set moisture-pressure pair from column selection"""
        selected_items = self.column_list.selectedItems()
        
        if len(selected_items) < 2:
            QMessageBox.warning(self, "Warning", 
                              "Please select at least two columns (moisture and pressure)")
            return
            
        selected_cols = [item.text() for item in selected_items]
        
        # Try to determine which is moisture and which is pressure
        moisture_col, pressure_col = self._guess_moisture_pressure(selected_cols)
        
        if moisture_col and pressure_col:
            self.moisture_combos[pair_index].setCurrentText(moisture_col)
            self.pressure_combos[pair_index].setCurrentText(pressure_col)
            
            # Save recent pairs
            self.save_recent_pairs()
            
            QMessageBox.information(self, "Pair Set", 
                                  f"Pair {pair_index+1} set to:\n"
                                  f"Moisture: {moisture_col}\n"
                                  f"Pressure: {pressure_col}")
        else:
            QMessageBox.warning(self, "Warning", 
                              "Could not determine moisture and pressure columns")
            
    def _guess_moisture_pressure(self, columns):
        """Guess which columns are moisture and pressure"""
        moisture_keywords = ['h2o', 'water', 'humid', 'moisture', 'ppm', 'ppb']
        pressure_keywords = ['pressure', 'press', 'bar', 'pa', 'mpa']
        
        moisture_candidates = []
        pressure_candidates = []
        
        for col in columns:
            col_lower = col.lower()
            
            moisture_score = sum(1 for kw in moisture_keywords if kw in col_lower)
            pressure_score = sum(1 for kw in pressure_keywords if kw in col_lower)
            
            if moisture_score > pressure_score:
                moisture_candidates.append(col)
            elif pressure_score > moisture_score:
                pressure_candidates.append(col)
                
        # Return best candidates
        moisture_col = moisture_candidates[0] if moisture_candidates else columns[0]
        pressure_col = pressure_candidates[0] if pressure_candidates else columns[1] if len(columns) > 1 else columns[0]
        
        return moisture_col, pressure_col
        
    def clear_pair(self, pair_index):
        """Clear specific moisture-pressure pair"""
        self.moisture_combos[pair_index].setCurrentIndex(-1)
        self.pressure_combos[pair_index].setCurrentIndex(-1)
        
    def clear_all_pairs(self):
        """Clear all moisture-pressure pairs"""
        for moisture_combo, pressure_combo in zip(self.moisture_combos, self.pressure_combos):
            moisture_combo.setCurrentIndex(-1)
            pressure_combo.setCurrentIndex(-1)
            
    def set_reference_from_selection(self):
        """Set reference column from selection"""
        selected_items = self.column_list.selectedItems()
        
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a column first")
            return
            
        ref_col = selected_items[0].text()
        self.reference_combo.setCurrentText(ref_col)
        
        QMessageBox.information(self, "Reference Column", 
                              f"Reference column set to: {ref_col}")
        
    def import_parameters(self):
        """Import calibration parameters"""
        params = self.parameter_manager.import_parameters(self)
        
        if params:
            self.f1_spin.setValue(params['f1'])
            self.f2_spin.setValue(params['f2'])
            self.p_ref_spin.setValue(params['p_ref'])
            
            # Update calibration core
            self.calibration_core.set_parameters(params['f1'], params['f2'], params['p_ref'])
            
    def export_parameters(self):
        """Export calibration parameters"""
        params = {
            'f1': self.f1_spin.value(),
            'f2': self.f2_spin.value(),
            'p_ref': self.p_ref_spin.value()
        }
        
        self.parameter_manager.export_parameters(params, self)
        
    def save_pairs(self):
        """Save current pairs configuration"""
        pairs_dict = self._get_current_pairs_dict()
        self.parameter_manager.export_pairs_config(pairs_dict, self)
        
    def load_pairs(self):
        """Load pairs configuration"""
        pairs_dict = self.parameter_manager.import_pairs_config(self)
        
        if pairs_dict:
            self._set_pairs_from_dict(pairs_dict)
            
    def save_recent_pairs(self):
        """Save recent pairs for current file"""
        if self.data_loader.get_file_path():
            pairs_dict = self._get_current_pairs_dict()
            self.parameter_manager.save_recent_pairs(
                self.data_loader.get_file_path(), pairs_dict)
                
    def load_recent_pairs(self):
        """Load recent pairs for current file"""
        if self.data_loader.get_file_path():
            pairs_dict = self.parameter_manager.load_recent_pairs(
                self.data_loader.get_file_path())
                
            if pairs_dict:
                self._set_pairs_from_dict(pairs_dict)
                QMessageBox.information(self, "Recent Pairs", 
                                      "Loaded recent pairs configuration for this file")
                
    def _get_current_pairs_dict(self):
        """Get current pairs as dictionary"""
        pairs_dict = {}
        
        for i, (moisture_combo, pressure_combo) in enumerate(zip(self.moisture_combos, self.pressure_combos)):
            moisture_col = moisture_combo.currentText()
            pressure_col = pressure_combo.currentText()
            
            if moisture_col and pressure_col:
                pairs_dict[f"pair_{i+1}"] = {
                    "moisture": moisture_col,
                    "pressure": pressure_col
                }
                
        return pairs_dict
        
    def _set_pairs_from_dict(self, pairs_dict):
        """Set pairs from dictionary"""
        # Clear existing pairs
        self.clear_all_pairs()
        
        # Set pairs from dictionary
        for i in range(len(self.moisture_combos)):
            pair_key = f"pair_{i+1}"
            if pair_key in pairs_dict:
                pair_data = pairs_dict[pair_key]
                moisture_col = pair_data.get("moisture", "")
                pressure_col = pair_data.get("pressure", "")
                
                if moisture_col and pressure_col:
                    self.moisture_combos[i].setCurrentText(moisture_col)
                    self.pressure_combos[i].setCurrentText(pressure_col)
                    
    def generate_chart(self):
        """Generate and display chart"""
        if not self.data_loader.is_data_loaded():
            QMessageBox.warning(self, "Warning", "Please load a file first")
            return
            
        # Get selected columns
        selected_items = self.column_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select columns to plot")
            return
            
        selected_columns = [item.text() for item in selected_items]
        
        # Get time column
        time_column = self.time_column_combo.currentText()
        if not time_column:
            QMessageBox.warning(self, "Warning", "Please select a time column")
            return
            
        try:
            # Prepare plot settings
            plot_settings = self._get_plot_settings()
            
            # Get data
            data = self.data_loader.data
            
            # Create and show plot window
            plot_window = IndependentPlotWindow(self)
            plot_window.plot_data(data, selected_columns, plot_settings)
            plot_window.show()
            
            # Keep reference to prevent garbage collection
            self.plot_windows.append(plot_window)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate chart: {str(e)}")
            
    def _get_plot_settings(self):
        """Get current plot settings"""
        # Get time range
        time_range = 0  # All data
        if self.radio_2hours.isChecked():
            time_range = 1  # 2 hours
        elif self.radio_custom.isChecked():
            time_range = 2  # Custom
            
        # Get moisture-pressure pairs
        moisture_pressure_pairs = []
        for moisture_combo, pressure_combo in zip(self.moisture_combos, self.pressure_combos):
            moisture_col = moisture_combo.currentText()
            pressure_col = pressure_combo.currentText()
            if moisture_col and pressure_col:
                moisture_pressure_pairs.append((moisture_col, pressure_col))
                
        # Update calibration core parameters
        self.calibration_core.set_parameters(
            self.f1_spin.value(),
            self.f2_spin.value(), 
            self.p_ref_spin.value()
        )
        
        return {
            'time_column': self.time_column_combo.currentText(),
            'time_range': time_range,
            'start_time': self.start_time_spin.value(),
            'end_time': self.end_time_spin.value(),
            'enable_calibration': self.enable_calibration_cb.isChecked(),
            'show_original': self.show_original_cb.isChecked(),
            'show_error': self.show_error_cb.isChecked(),
            'error_value': self.error_value_spin.value(),
            'moisture_pressure_pairs': moisture_pressure_pairs,
            'f1': self.f1_spin.value(),
            'f2': self.f2_spin.value(),
            'p_ref': self.p_ref_spin.value(),
            'enable_30min_diff': self.enable_30min_diff_cb.isChecked(),
            'reference_column': self.reference_combo.currentText(),
            'time_window': self.time_window_spin.value()
        }
        
    def export_data(self):
        """Export data to Excel"""
        if not self.data_loader.is_data_loaded():
            QMessageBox.warning(self, "Warning", "Please load a file first")
            return
            
        try:
            # Get selected columns
            selected_items = self.column_list.selectedItems()
            selected_columns = [item.text() for item in selected_items] if selected_items else []
            
            # Get time column  
            time_column = self.time_column_combo.currentText()
            
            # Get time range settings
            time_range = 0
            if self.radio_2hours.isChecked():
                time_range = 1
            elif self.radio_custom.isChecked():
                time_range = 2
                
            # Prepare data for export
            export_data = self.data_loader.prepare_plot_data(
                time_column, selected_columns, time_range,
                self.start_time_spin.value(), self.end_time_spin.value()
            )
            
            if export_data is not None:
                # Apply calibration if enabled
                if self.enable_calibration_cb.isChecked():
                    export_data = self._apply_calibration_to_export_data(export_data)
                    
                # Export data
                self.data_loader.export_data(export_data, self)
            else:
                QMessageBox.warning(self, "Warning", "No data to export")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export data: {str(e)}")
            
    def _apply_calibration_to_export_data(self, data):
        """Apply calibration to export data"""
        # Get moisture-pressure pairs
        moisture_pressure_pairs = []
        for moisture_combo, pressure_combo in zip(self.moisture_combos, self.pressure_combos):
            moisture_col = moisture_combo.currentText()
            pressure_col = pressure_combo.currentText()
            if moisture_col and pressure_col and moisture_col in data.columns and pressure_col in data.columns:
                moisture_pressure_pairs.append((moisture_col, pressure_col))
                
        if not moisture_pressure_pairs:
            return data
            
        # Update calibration core parameters
        self.calibration_core.set_parameters(
            self.f1_spin.value(),
            self.f2_spin.value(),
            self.p_ref_spin.value()
        )
        
        # Apply calibration to each pair
        export_data = data.copy()
        
        for moisture_col, pressure_col in moisture_pressure_pairs:
            result = self.calibration_core.calibrate_single_pair(
                export_data, moisture_col, pressure_col
            )
            
            if result:
                calib_col = f"{moisture_col}_calibrated"
                
                # Add calibrated data to export
                if 'relative_time' in export_data.columns:
                    calib_df = result['valid_df'][['relative_time']].copy()
                    calib_df[calib_col] = result['values']
                    
                    # Merge with export data
                    export_data = export_data.merge(
                        calib_df, on='relative_time', how='left'
                    )
                    
        return export_data
        
    def closeEvent(self, event):
        """Handle window close event"""
        # Close any open plot windows
        for window in self.plot_windows:
            if window.isVisible():
                window.close()
                
        event.accept() 
#!/usr/bin/env python3
"""
H2O Data Visualization and Calibration Tool
Main entry point for the application
"""

import sys
import os

# Add the current directory to Python path to enable imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def main():
    """Main function to start the application"""
    try:
        from ui.main_window import MainWindow
        from PyQt5.QtWidgets import QApplication
        
        # Create QApplication instance
        app = QApplication(sys.argv)
        app.setApplicationName("H2O Data Visualization and Calibration Tool")
        
        # Create and show main window
        window = MainWindow()
        window.show()
        
        # Start event loop
        sys.exit(app.exec_())
        
    except ImportError as e:
        print(f"Import error: {e}")
        print("Please make sure all required packages are installed:")
        print("pip install PyQt5 matplotlib pandas numpy openpyxl")
        sys.exit(1)
    except Exception as e:
        print(f"Error starting application: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

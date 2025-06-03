#!/usr/bin/env python3
"""
Simple launcher script for H2O Data Visualization and Calibration Tool
"""

if __name__ == "__main__":
    try:
        from main import main
        main()
    except ImportError as e:
        print(f"Import error: {e}")
        print("Please make sure all required packages are installed:")
        print("pip install PyQt5 matplotlib pandas numpy openpyxl")
    except Exception as e:
        print(f"Error: {e}") 
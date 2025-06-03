#!/bin/bash

# H2O Data Visualization and Calibration Tool Startup Script

echo "Starting H2O Data Visualization and Calibration Tool..."
echo "======================================================"

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed or not in PATH"
    echo "Please install Python 3 and try again"
    exit 1
fi

# Check if required packages are installed
echo "Checking required packages..."
python3 -c "import PyQt5, matplotlib, pandas, numpy, openpyxl" 2>/dev/null

if [ $? -ne 0 ]; then
    echo "Some required packages are missing. Installing them now..."
    pip3 install PyQt5 matplotlib pandas numpy openpyxl
    
    if [ $? -ne 0 ]; then
        echo "Failed to install required packages. Please install them manually:"
        echo "pip3 install PyQt5 matplotlib pandas numpy openpyxl"
        exit 1
    fi
fi

echo "All packages are available. Starting application..."
echo ""

# Start the application
python3 main.py

echo ""
echo "Application closed." 
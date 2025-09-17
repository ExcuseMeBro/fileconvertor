#!/bin/bash

# Force File Converter Runner
echo "Starting Force File Converter..."

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    echo "Activating virtual environment..."
    source venv/bin/activate
fi

# Install required packages
echo "Installing required packages..."
pip3 install python-docx openpyxl reportlab Pillow

# Run the force converter
echo "Running force converter..."
python3 force_converter.py

echo "Force conversion completed!"
echo "Check force_conversion.log for detailed results."
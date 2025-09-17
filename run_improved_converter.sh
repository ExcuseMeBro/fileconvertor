#!/bin/bash

echo "Yaxshilangan konverter ishga tushirilmoqda..."
echo "Installing dependencies..."

# Install required packages
pip3 install docx2pdf python-docx openpyxl python-pptx Pillow reportlab xlsxwriter

echo "Starting conversion process..."
python3 improved_converter.py

echo "Konversiya jarayoni tugadi!"
echo "Natijalarni conversion.log faylida ko'ring"
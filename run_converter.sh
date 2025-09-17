#!/bin/bash

echo "Installing Python dependencies..."
pip3 install -r requirements.txt

echo "Starting file conversion process..."
python3 file_converter.py

echo "Conversion process completed!"
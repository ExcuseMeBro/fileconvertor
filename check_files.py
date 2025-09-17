#!/usr/bin/env python3
"""
File checker script to count and verify files
"""

import os
from pathlib import Path

def count_files_by_extension(directory, extensions):
    """Count files by extension in directory"""
    counts = {}
    files_list = []
    
    for ext in extensions:
        counts[ext] = 0
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = Path(root) / file
            ext = file_path.suffix.lower()
            
            if ext in extensions:
                counts[ext] += 1
                files_list.append(str(file_path))
    
    return counts, files_list

def main():
    source_dir = "/Users/bro/PROJECTS/fileconvertor/boltfiles"
    target_dir = "/Users/bro/PROJECTS/fileconvertor/converted"
    
    # Extensions to check
    source_extensions = ['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt', '.png']
    target_extensions = ['.pdf']
    
    print("="*60)
    print("FAYL HISOBOTI / FILE REPORT")
    print("="*60)
    
    # Count source files
    print(f"\nMANBA PAPKA / SOURCE FOLDER: {source_dir}")
    source_counts, source_files = count_files_by_extension(source_dir, source_extensions)
    
    total_source = sum(source_counts.values())
    print(f"Jami fayllar / Total files: {total_source}")
    
    for ext, count in source_counts.items():
        if count > 0:
            print(f"  {ext}: {count} ta")
    
    # Count converted files
    print(f"\nKONVERT QILINGAN PAPKA / CONVERTED FOLDER: {target_dir}")
    if os.path.exists(target_dir):
        target_counts, target_files = count_files_by_extension(target_dir, target_extensions)
        total_converted = sum(target_counts.values())
        print(f"Konvert qilingan fayllar / Converted files: {total_converted}")
        
        print(f"\nKONVERSIYA HOLATI / CONVERSION STATUS:")
        print(f"Muvaffaqiyatli / Successful: {total_converted}/{total_source}")
        print(f"Qolgan / Remaining: {total_source - total_converted}")
        
        if total_converted < total_source:
            print(f"\n⚠️  {total_source - total_converted} ta fayl hali konvert qilinmagan!")
            print("Quyidagi buyruqni ishga tushiring:")
            print("python3 improved_converter.py")
    else:
        print("Converted papka topilmadi!")
    
    print("\n" + "="*60)

if __name__ == "__main__":
    main()
#!/usr/bin/env python3
"""
PDF Ko'chirish Skripti - Mavjud PDF fayllarni converted papkasiga ko'chiradi
"""

import os
import shutil
from pathlib import Path
import logging

# Logging sozlamalari
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_copy.log'),
        logging.StreamHandler()
    ]
)

def find_pdf_files(source_dir):
    """Barcha PDF fayllarni topish"""
    pdf_files = []
    
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                pdf_files.append(file_path)
    
    return pdf_files

def copy_pdf_file(source_file, source_dir, target_dir):
    """PDF faylni maqsadli joyga ko'chirish"""
    try:
        # Nisbiy yo'lni hisoblash
        rel_path = os.path.relpath(os.path.dirname(source_file), source_dir)
        
        # Maqsadli papkani yaratish
        if rel_path == '.':
            target_folder = target_dir
        else:
            target_folder = os.path.join(target_dir, rel_path)
        
        os.makedirs(target_folder, exist_ok=True)
        
        # Fayl nomini olish
        filename = os.path.basename(source_file)
        target_file = os.path.join(target_folder, filename)
        
        # Agar fayl allaqachon mavjud bo'lsa, tekshirish
        if os.path.exists(target_file):
            source_size = os.path.getsize(source_file)
            target_size = os.path.getsize(target_file)
            
            if source_size == target_size:
                logging.info(f"Fayl allaqachon mavjud: {filename}")
                return True
            else:
                logging.info(f"Fayl o'lchamlari farq qiladi, qayta ko'chirilmoqda: {filename}")
        
        # Faylni ko'chirish
        shutil.copy2(source_file, target_file)
        logging.info(f"✓ Ko'chirildi: {filename}")
        return True
        
    except Exception as e:
        logging.error(f"✗ Xatolik {os.path.basename(source_file)}: {e}")
        return False

def main():
    source_dir = "/Users/bro/PROJECTS/fileconvertor/boltfiles"
    target_dir = "/Users/bro/PROJECTS/fileconvertor/converted"
    
    logging.info("PDF fayllarni qidirish boshlandi...")
    
    # PDF fayllarni topish
    pdf_files = find_pdf_files(source_dir)
    
    if not pdf_files:
        logging.info("Hech qanday PDF fayl topilmadi.")
        return
    
    logging.info(f"Topilgan PDF fayllar soni: {len(pdf_files)}")
    
    # Har bir PDF faylni ko'chirish
    copied = 0
    failed = 0
    
    for i, pdf_file in enumerate(pdf_files, 1):
        logging.info(f"Jarayon {i}/{len(pdf_files)}: {os.path.basename(pdf_file)}")
        
        if copy_pdf_file(pdf_file, source_dir, target_dir):
            copied += 1
        else:
            failed += 1
    
    # Natijalarni ko'rsatish
    logging.info(f"\nPDF ko'chirish yakunlandi!")
    logging.info(f"Ko'chirildi: {copied}")
    logging.info(f"Xatolik: {failed}")
    logging.info(f"Jami jarayon: {len(pdf_files)}")
    
    # Qo'shimcha ma'lumot
    if copied > 0:
        logging.info(f"Barcha PDF fayllar {target_dir} papkasiga ko'chirildi.")
    
    return copied, failed

if __name__ == "__main__":
    main()
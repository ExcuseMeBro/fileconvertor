#!/bin/bash

# PDF Ko'chirish Skripti
echo "PDF fayllarni ko'chirish boshlandi..."

# Virtual muhitni faollashtirish
if [ -d "venv" ]; then
    echo "Virtual muhitni faollashtirish..."
    source venv/bin/activate
fi

# PDF ko'chirish skriptini ishga tushirish
echo "PDF fayllarni qidirish va ko'chirish..."
python3 copy_existing_pdfs.py

echo "PDF ko'chirish yakunlandi!"
echo "Batafsil ma'lumot uchun pdf_copy.log faylini ko'ring."
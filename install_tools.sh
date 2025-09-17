#!/bin/bash

# LibreOffice va Pandoc o'rnatish skripti (macOS uchun)
echo "LibreOffice va Pandoc o'rnatish boshlandi..."

# Homebrew mavjudligini tekshirish
if ! command -v brew &> /dev/null; then
    echo "Homebrew topilmadi. Homebrew o'rnatish..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
fi

# LibreOffice o'rnatish
echo "LibreOffice o'rnatish..."
if [ ! -d "/Applications/LibreOffice.app" ]; then
    brew install --cask libreoffice
    echo "✓ LibreOffice o'rnatildi"
else
    echo "✓ LibreOffice allaqachon o'rnatilgan"
fi

# Pandoc o'rnatish
echo "Pandoc o'rnatish..."
if ! command -v pandoc &> /dev/null; then
    brew install pandoc
    echo "✓ Pandoc o'rnatildi"
else
    echo "✓ Pandoc allaqachon o'rnatilgan"
fi

echo "Barcha vositalar muvaffaqiyatli o'rnatildi!"
echo "Endi force_converter.py skriptini ishga tushirishingiz mumkin."
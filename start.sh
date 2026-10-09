#!/bin/bash
# PDF Toolbox start karne ke liye (Mac / Linux)
cd "$(dirname "$0")"
python3 -c "import flask, pymupdf, openpyxl, docx, PIL" 2>/dev/null || {
  echo "Packages install kar rahe hain..."; pip3 install -r requirements.txt; }
echo ""
echo "  ✅ App chalu ho rahi hai → browser me kholein:  http://localhost:8000"
echo "  (band karne ke liye Ctrl + C dabayein)"
echo ""
exec python3 app.py

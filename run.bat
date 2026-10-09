@echo off
REM PDF Toolbox — Windows
cd /d "%~dp0"
python -c "import flask, pymupdf, openpyxl, docx, PIL" 2>nul
if errorlevel 1 (
  echo Packages install kar rahe hain...
  pip install -r requirements.txt
)
echo.
echo   App chalu ho rahi hai -^> browser me kholein:  http://localhost:8000
echo   (band karne ke liye is window me Ctrl + C dabayein)
echo.
start "" http://localhost:8000
python app.py
pause

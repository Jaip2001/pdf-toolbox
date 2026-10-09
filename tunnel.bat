@echo off
REM ============================================================
REM  PDF Toolbox — FREE public URL (no domain needed)
REM  Ye script aapke PC se ek free https link bana dega,
REM  jaise Google AppScript deta hai. Koi account nahi chahiye.
REM
REM  Use:  tunnel.bat  par double-click
REM ============================================================
cd /d "%~dp0"
title PDF Toolbox - Free Public Link
color 0A
echo ==============================================
echo   PDF Toolbox - Free public link ban rahi hai
echo ==============================================
echo.

REM --- Step 1: packages check ---
python -c "import flask, pymupdf, openpyxl, docx, PIL" 2>nul
if errorlevel 1 (
  echo [*] Pehli baar: packages install ho rahe hain (1-2 min)...
  pip install -r requirements.txt
)

REM --- Step 2: cloudflared download (free tunnel tool, no account) ---
if not exist cloudflared.exe (
  echo [*] cloudflared download ho raha hai...
  curl -sL -o cloudflared.exe "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
  if errorlevel 1 (
    echo [!] Download fail. Internet check karein, ya manually download karein:
    echo     https://github.com/cloudflare/cloudflared/releases/latest
    pause
    exit /b 1
  )
)

REM --- Step 3: app chalu karein (background window me) ---
echo [*] App chalu ho rahi hai...
start "PDF Toolbox App" /min python app.py
timeout /t 5 /nobreak >nul

REM --- Step 4: free tunnel ---
echo.
echo ==============================================
echo   Niche jo https://....trycloudflare.com link
necho   dikhega - woh aapki FREE public link hai!
echo   Browser me kholein, ya kisi ko bhi bhej dein.
echo   (Band karne ke liye is window me Ctrl+C)
echo ==============================================
echo.
cloudflared.exe tunnel --url http://localhost:8000 --no-autoupdate

pause

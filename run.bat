@echo off
echo ========================================
echo  VAPT One-Click Report Generator
echo ========================================
echo.

cd /d "%~dp0"

echo [1] Installing dependencies...
pip install -r requirement.txt

echo.
echo [2] Starting VAPT Server...
echo.
echo ========================================
echo  Server started at: http://127.0.0.1:5000
echo  Press Ctrl+C to stop
echo ========================================
echo.

python app.py

pause
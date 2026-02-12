@echo off
title PDFConverter - Professional PDF Conversion Tool
cd /d "%~dp0"

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

REM Create venv if not exists
if not exist venv (
    echo [INFO] Creating virtual environment...
    python -m venv venv
)

REM Activate venv
call venv\Scripts\activate.bat

REM Install dependencies
echo [INFO] Checking dependencies...
pip install -r requirements.txt -q 2>nul

REM Create required directories
if not exist uploads mkdir uploads
if not exist output mkdir output
if not exist logs mkdir logs

echo.
echo ====================================================
echo  PDFConverter - Starting on http://127.0.0.1:5000
echo ====================================================
echo.

python app.py

pause

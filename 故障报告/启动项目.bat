@echo off
chcp 65001 >nul
title Fault Report Generator

echo.
echo ============================================
echo   Fault Report Generation System
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found! Install Python 3.8+ first.
    pause
    exit /b 1
)

cd /d "%~dp0backend"
if errorlevel 1 (
    echo [ERROR] Cannot enter backend directory
    pause
    exit /b 1
)

echo [INFO] Working directory: %CD%

echo.
echo [1/2] Checking dependencies...
pip install flask flask-cors python-docx zhipuai requests openpyxl -q 2>nul
echo [2/2] Starting backend server...
echo.
echo ============================================
echo   Backend:  http://localhost:5000
echo   Frontend: Open frontend/index.html
echo ============================================
echo Press Ctrl+C to stop
echo.

python app.py
pause

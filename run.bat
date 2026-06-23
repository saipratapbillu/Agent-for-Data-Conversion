@echo off
REM Excel to CSV Converter - Windows Batch Launcher
REM This script installs dependencies and runs the GUI agent

echo ========================================
echo Excel to CSV Converter Agent Installer
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher from https://www.python.org/
    pause
    exit /b 1
)

echo Python found!
echo.
echo Installing dependencies...
pip install -r requirements.txt

if errorlevel 1 (
    echo Error: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ========================================
echo Launching Excel to CSV Converter...
echo ========================================
echo.

python main.py
pause

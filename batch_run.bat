@echo off
REM Batch Excel to CSV Converter - Windows Launcher
REM Converts all Excel files from network path to CSV

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo off
REM Batch Excel to CSV Converter - Windows Launcher
REM Opens the GUI to select source and destination folders, then convert Excel files to CSV.

setlocal enabledelayedexpansion
cd /d "%~dp0"

set SCRIPT_DIR=%~dp0
set VENV_PYTHON=%SCRIPT_DIR%\.venv\Scripts\python.exe

if not exist "%VENV_PYTHON%" (
    echo Error: Virtual environment not found.
    echo Run: python -m venv .venv
    pause
    exit /b 1
)

echo.
echo ========================================
echo Excel to CSV Converter - GUI Launcher
echo ========================================
echo.

echo Launching the UI application...

"%VENV_PYTHON%" main.py

if errorlevel 1 (
    echo.
    echo ❌ Application exited with an error.
    pause
    exit /b 1
) else (
    echo.
    echo ✅ Application closed.
    pause
    exit /b 0
)

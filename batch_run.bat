@echo off
REM Batch Excel to CSV Converter - Windows Launcher
REM Converts all Excel files from network path to CSV

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ========================================
echo Batch Excel to CSV Converter
echo ========================================
echo.

REM Define paths
set EXCEL_SOURCE=\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports
set CSV_OUTPUT=C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV
set SCRIPT_DIR=%~dp0
set VENV_PYTHON=%SCRIPT_DIR%\.venv\Scripts\python.exe

REM Check if virtual environment exists
if not exist "%VENV_PYTHON%" (
    echo Error: Virtual environment not found
    echo Run: python -m venv .venv
    pause
    exit /b 1
)

echo Source Directory: %EXCEL_SOURCE%
echo Output Directory: %CSV_OUTPUT%
echo.

REM Run batch converter using virtual environment Python
"%VENV_PYTHON%" batch_convert.py "%EXCEL_SOURCE%" "%CSV_OUTPUT%"

if errorlevel 1 (
    echo.
    echo ❌ Batch conversion encountered errors!
    pause
    exit /b 1
) else (
    echo.
    echo ✅ Batch conversion completed successfully!
    pause
    exit /b 0
)

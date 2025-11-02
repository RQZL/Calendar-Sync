@echo off
echo ============================================
echo Run Calendar Sync
echo ============================================
echo.

REM Change to the directory where this batch file is located
cd /d "%~dp0"

REM Use relative path to Python (works regardless of drive letter)
set PYTHON_PATH=%~dp0system\python\WPy64-31401\python\python.exe

REM Run the Python script
"%PYTHON_PATH%" "%~dp0system\scripts\run-calendar-script.py"

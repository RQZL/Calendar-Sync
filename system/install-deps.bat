@echo off
echo ============================================
echo Installing Python Dependencies
echo ============================================
echo.

REM Change to the directory where this batch file is located
cd /d "%~dp0"

REM Use relative path to Python (works regardless of drive letter)
set PYTHON_PATH=%~dp0python\WPy64-31401\python\python.exe

echo Installing required packages...
echo.

"%PYTHON_PATH%" -m pip install -r "%~dp0scripts\requirements.txt"

echo.
echo ============================================
echo Installation Complete!
echo ============================================
echo.

pause

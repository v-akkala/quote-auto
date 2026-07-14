@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Local Python environment was not found.
    echo Running setup first...
    call "%~dp0setup_portable_quote_generator.bat"
    if errorlevel 1 exit /b 1
)

".venv\Scripts\python.exe" -c "import tkinter; import openpyxl; import PIL" >nul 2>nul
if errorlevel 1 (
    echo Required packages are missing or broken.
    echo Running setup to repair the local environment...
    call "%~dp0setup_portable_quote_generator.bat"
    if errorlevel 1 exit /b 1
)

echo Starting quote generator...
".venv\Scripts\python.exe" quote_app\app.py
if errorlevel 1 (
    echo.
    echo The quote generator closed with an error.
    pause
    exit /b 1
)

endlocal

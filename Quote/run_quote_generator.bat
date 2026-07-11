@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_CMD="
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    )
)

if "%PYTHON_CMD%"=="" (
    echo Python was not found. Please install Python 3.10 or newer, then run this file again.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating local Python environment...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo Could not create the Python environment.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing requirements...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Could not install requirements.
        pause
        exit /b 1
    )
)

echo Starting quote generator...
".venv\Scripts\python.exe" quote_app\app.py
if errorlevel 1 (
    echo The quote generator closed with an error.
    pause
    exit /b 1
)

endlocal

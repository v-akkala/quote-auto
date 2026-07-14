@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON_VERSION=3.12.10"
set "PYTHON_MAJOR_MINOR=312"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe"
set "PYTHON_INSTALLER=%TEMP%\python-%PYTHON_VERSION%-amd64.exe"
set "PYTHON_CMD="

echo.
echo SKC Quote Generator setup
echo =========================
echo.

call :find_python
if not defined PYTHON_CMD (
    echo Python 3 was not found. Setup will download and install Python %PYTHON_VERSION% for this Windows user.
    echo.
    call :install_python
    if errorlevel 1 goto :fail
    call :find_python
)

if not defined PYTHON_CMD (
    echo Python could not be found after installation.
    goto :fail
)

echo Using Python:
%PYTHON_CMD% --version
if errorlevel 1 goto :fail

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import tkinter" >nul 2>nul
    if errorlevel 1 (
        echo Existing local environment is missing Tkinter support. Recreating .venv...
        rmdir /s /q ".venv"
        if errorlevel 1 goto :fail
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo Creating local Python environment in .venv...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 goto :fail
)

echo.
echo Installing/updating required Python packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :fail

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :fail

".venv\Scripts\python.exe" -c "import tkinter; import openpyxl; import PIL" >nul 2>nul
if errorlevel 1 (
    echo Required packages or Tkinter could not be imported after setup.
    goto :fail
)

echo.
echo Setup complete.
echo You can now run run_quote_generator.bat.
echo.
pause
endlocal
exit /b 0

:find_python
set "PYTHON_CMD="
if exist "%LocalAppData%\Programs\Python\Python%PYTHON_MAJOR_MINOR%\python.exe" (
    "%LocalAppData%\Programs\Python\Python%PYTHON_MAJOR_MINOR%\python.exe" -c "import sys; import tkinter; import venv; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD="%LocalAppData%\Programs\Python\Python%PYTHON_MAJOR_MINOR%\python.exe""
        exit /b 0
    )
)
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -c "import sys; import tkinter; import venv; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3"
        exit /b 0
    )
)
where python >nul 2>nul
if not errorlevel 1 (
    python -c "import sys; import tkinter; import venv; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=python"
        exit /b 0
    )
)
exit /b 0

:install_python
echo Downloading Python from:
echo %PYTHON_URL%
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%'"
if errorlevel 1 (
    echo Could not download Python. Check the internet connection and try again.
    exit /b 1
)

echo.
echo Installing Python. This may take a minute...
"%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_pip=1 Include_tcltk=1 Include_test=0 SimpleInstall=1
if errorlevel 1 (
    echo Python installer failed.
    exit /b 1
)
exit /b 0

:fail
echo.
echo Setup failed.
echo If this computer blocks downloads, install Python 3.10 or newer from https://www.python.org/downloads/ and run this setup again.
echo.
pause
endlocal
exit /b 1

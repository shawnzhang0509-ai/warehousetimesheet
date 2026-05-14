@echo off
setlocal

rem Start from the folder where this .bat file lives, so it works after git pull.
cd /d "%~dp0"

set "PYTHON_EXE="
set "PYTHON_ARGS="

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_EXE=python"
    )
)

if not defined PYTHON_EXE (
    echo Python was not found.
    echo Please install Python 3 from https://www.python.org/downloads/ and try again.
    pause
    exit /b 1
)

%PYTHON_EXE% %PYTHON_ARGS% -c "import pandas, openpyxl, xlrd" >nul 2>nul
if errorlevel 1 (
    echo Installing required Python packages...
    %PYTHON_EXE% %PYTHON_ARGS% -m pip install pandas openpyxl xlrd
    if errorlevel 1 (
        echo Failed to install required packages.
        pause
        exit /b 1
    )
)

echo Starting attendance GUI...
%PYTHON_EXE% %PYTHON_ARGS% "%~dp0attendance_gui.py"
if errorlevel 1 (
    echo Attendance GUI exited with an error.
    pause
    exit /b 1
)

@echo off
setlocal

rem Build a Windows executable for the attendance GUI.
rem Run this on a Windows machine from the repository root.
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

echo Installing build dependencies...
%PYTHON_EXE% %PYTHON_ARGS% -m pip install --upgrade pip
if errorlevel 1 goto :fail

%PYTHON_EXE% %PYTHON_ARGS% -m pip install pandas openpyxl xlrd pyinstaller
if errorlevel 1 goto :fail

echo Building attendance_gui.exe...
%PYTHON_EXE% %PYTHON_ARGS% -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name attendance_gui ^
    attendance_gui.py
if errorlevel 1 goto :fail

if exist name_mapping.json (
    copy /Y name_mapping.json dist\name_mapping.json >nul
)

echo.
echo Build complete.
echo EXE: %cd%\dist\attendance_gui.exe
echo Keep name_mapping.json in the same dist folder if you use name mapping.
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1

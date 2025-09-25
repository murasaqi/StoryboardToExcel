@echo off
setlocal

set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%generate_storyboard_workbook.py

if not exist "%SCRIPT_PATH%" (
    echo [ERROR] generate_storyboard_workbook.py not found in %SCRIPT_DIR%
    exit /b 1
)

set PYTHON=python

"%PYTHON%" "%SCRIPT_PATH%" %*
set EXITCODE=%ERRORLEVEL%
endlocal & exit /b %EXITCODE%

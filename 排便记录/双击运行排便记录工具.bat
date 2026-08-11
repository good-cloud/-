@echo off
cd /d "%~dp0"
set "BUNDLED_PY=C:\Users\Eva\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%BUNDLED_PY%" (
    set "PY=%BUNDLED_PY%"
) else (
    set "PY=python"
)

"%PY%" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing required package: openpyxl
    "%PY%" -m pip install openpyxl
    if errorlevel 1 (
        echo.
        echo Failed to install openpyxl.
        echo Please check that Python and pip are installed correctly.
        pause
        exit /b 1
    )
)
"%PY%" "bowel_record_tool.py"
if errorlevel 1 (
    echo.
    echo Failed to start.
    echo Please make sure Python is installed correctly.
    pause
)

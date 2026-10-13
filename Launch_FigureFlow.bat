@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %ERRORLEVEL%==0 (
    py -3 figureflow.py
) else (
    python figureflow.py
)
if errorlevel 1 (
    echo.
    echo If packages are missing, run Install_Once.bat first.
    pause
)
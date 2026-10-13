@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python not found. Install Python 3.10+ from python.org and rerun.
        echo FigureFlow never uploads figures, even when dependency installation needs internet.
        pause
        exit /b 1
    )
    set "PY=python"
)
%PY% -m pip install --upgrade -r requirements.txt
if errorlevel 1 (
    echo Installation failed. For a fully air-gapped PC, install these packages from local wheels.
    pause
    exit /b 1
)
echo.
echo FigureFlow is installed. Double-click Launch_FigureFlow.bat.
pause
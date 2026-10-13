@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python not found. Install Python 3.10+ first.
        pause
        exit /b 1
    )
    set "PY=python"
)
echo This optional step bundles FigureFlow into a standalone local Windows EXE.
echo Python build packages may be downloaded when missing, but PDF files remain local.
%PY% -m pip install pyinstaller
if errorlevel 1 goto fail
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name FigureFlow --collect-all tkinterdnd2 --collect-all pymupdf figureflow.py
if errorlevel 1 goto fail
echo.
echo Built: dist\FigureFlow.exe
pause
exit /b 0
:fail
echo Build failed. You can still run Launch_FigureFlow.bat.
pause
exit /b 1
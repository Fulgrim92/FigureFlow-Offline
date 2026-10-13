@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    py -3 self_test.py
) else (
    python self_test.py
)
echo.
pause
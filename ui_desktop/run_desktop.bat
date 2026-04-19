@echo off
title Cross-Table Check - Desktop

cd /d "%~dp0"

echo.
echo ============================================
echo   Cross-Table Check  (Desktop / PySide6)
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py -3"
    goto :run
)
where python >nul 2>nul
if %errorlevel%==0 (
    set "PY=python"
    goto :run
)
echo [ERROR] Python not found in PATH.
echo Please install Python 3.10+ from https://www.python.org/downloads/
pause
exit /b 1

:run
echo Python: %PY%
echo Starting...
echo.
%PY% "%~dp0app.py"
set EC=%errorlevel%
echo.
if %EC% neq 0 (
    echo [ERROR] App exited with code %EC%
) else (
    echo [OK] App closed normally.
)
echo.
pause

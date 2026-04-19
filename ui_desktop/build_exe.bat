@echo off
title Build Desktop App
cd /d "%~dp0"

echo.
echo ============================================
echo   Building standalone desktop app (no Python needed on target)
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")

echo [1/2] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/2] Running PyInstaller...
%PY% -m PyInstaller app.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build done. Output: dist\跨表核对\
echo   Send the entire folder to users; they run 跨表核对.exe
echo ============================================
echo.
pause

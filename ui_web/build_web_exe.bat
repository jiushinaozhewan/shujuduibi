@echo off
title Build Web App
cd /d "%~dp0"

echo.
echo ============================================
echo   Building standalone Web app (Streamlit)
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")

echo [1/2] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/2] Running PyInstaller...
%PY% -m PyInstaller app_web.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build done. Output: dist\跨表核对Web\
echo   Send the entire folder to users; they run 跨表核对Web.exe
echo ============================================
echo.
pause

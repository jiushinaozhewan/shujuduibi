@echo off
title Cross-Table Check - Web

cd /d "%~dp0"

echo.
echo ============================================
echo   Cross-Table Check  (Web / Streamlit)
echo   Browser opens: http://localhost:8501
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
pause
exit /b 1

:run
%PY% -m streamlit run "%~dp0app.py"
echo.
pause

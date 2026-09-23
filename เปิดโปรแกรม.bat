@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

where pythonw >nul 2>&1
if errorlevel 1 goto usepython
start "" pythonw dq_dashboard_app.py
goto end

:usepython
start "" python dq_dashboard_app.py

:end
endlocal
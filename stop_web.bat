@echo off
REM Stop Estimate AI web UI (kills reload supervisor + port 8000 listeners).
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_web.ps1"
exit /b %ERRORLEVEL%

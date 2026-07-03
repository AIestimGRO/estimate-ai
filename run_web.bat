@echo off
REM Launch the Estimate AI local web UI using the project virtualenv.
REM Stops any previous instance on port 8000, then starts with --reload.
cd /d "%~dp0"

echo Stopping old server on port 8000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8000" ^| findstr "LISTENING"') do (
  taskkill /F /PID %%a >nul 2>&1
)

echo Starting Estimate AI...
".venv\Scripts\python.exe" -m app.web --reload %*
pause

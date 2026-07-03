@echo off
REM Launch the Estimate AI local web UI using the project virtualenv.
REM Stops ALL previous listeners on port 8000, then starts (no --reload: one process, easy to stop).
cd /d "%~dp0"

echo Stopping old servers ...
call "%~dp0stop_web.bat"
if errorlevel 1 exit /b 1

echo.
echo Starting Estimate AI (.venv) ...
echo To stop: Ctrl+C in this window, or run stop_web.bat from another terminal.
".venv\Scripts\python.exe" -m app.web --no-browser %*
pause

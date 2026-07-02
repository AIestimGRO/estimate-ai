@echo off
REM Launch the Estimate AI local web UI using the project virtualenv.
cd /d "%~dp0"
".venv\Scripts\python.exe" -m app.web %*
pause

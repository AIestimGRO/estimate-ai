# Launch the Estimate AI local web UI using the project virtualenv.
Set-Location -LiteralPath $PSScriptRoot

& "$PSScriptRoot\stop_web.ps1"

Write-Host ""
Write-Host "Starting Estimate AI (.venv) ..."
Write-Host "To stop: Ctrl+C in this window, or run stop_web.ps1 from another terminal."
& ".\.venv\Scripts\python.exe" -m app.web --no-browser @args

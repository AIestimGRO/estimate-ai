# Launch the Estimate AI local web UI using the project virtualenv.
Set-Location -LiteralPath $PSScriptRoot
& ".\.venv\Scripts\python.exe" -m app.web @args

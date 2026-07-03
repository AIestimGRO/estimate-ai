# Launch the Estimate AI local web UI using the project virtualenv.
Set-Location -LiteralPath $PSScriptRoot

$listeners = netstat -ano | Select-String ':8000' | Select-String 'LISTENING'
foreach ($line in $listeners) {
    $pid = ($line -split '\s+')[-1]
    if ($pid -match '^\d+$') {
        Write-Host "Stopping old server PID $pid..."
        Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
    }
}

& ".\.venv\Scripts\python.exe" -m app.web --reload @args

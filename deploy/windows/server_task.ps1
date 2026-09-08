Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")

Set-Location -LiteralPath $ProjectRoot
Ensure-ServerDirectories

if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) { throw "Virtualenv Python not found. Run server_install.ps1 first." }
if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) { throw "Database file not found: $DatabasePath" }
if (-not $CertFile -or -not (Test-Path -LiteralPath $CertFile -PathType Leaf)) { throw "TLS certificate is not configured." }
if (-not $KeyFile -or -not (Test-Path -LiteralPath $KeyFile -PathType Leaf)) { throw "TLS private key is not configured." }

$env:ESTIMATE_AI_DB_PATH = $DatabasePath
$env:PYTHONUNBUFFERED = "1"

$logPath = Join-Path $LogDirectory "server.log"
Add-Content -LiteralPath $logPath -Value ("[{0}] Starting Estimate AI on https://0.0.0.0:{1}" -f (Get-Date -Format "s"), $ServerPort)

$arguments = @(
    "-m", "uvicorn", "app.web.app:create_app",
    "--factory",
    "--host", $ServerHost,
    "--port", [string]$ServerPort,
    "--ssl-certfile", $CertFile,
    "--ssl-keyfile", $KeyFile,
    "--log-level", "info"
)
# Uvicorn writes normal INFO logs to stderr. Windows PowerShell converts native
# stderr into ErrorRecord objects when streams are merged. With the script-wide
# ErrorActionPreference=Stop that would terminate this wrapper immediately even
# though Uvicorn started successfully. Keep strict error handling everywhere
# else, but allow the native server process to own its stderr for its lifetime.
$previousErrorActionPreference = $ErrorActionPreference
$exitCode = 1
try {
    $ErrorActionPreference = "Continue"
    & $VenvPython @arguments *>> $logPath
    $exitCode = $LASTEXITCODE
} finally {
    $ErrorActionPreference = $previousErrorActionPreference
}
exit $exitCode

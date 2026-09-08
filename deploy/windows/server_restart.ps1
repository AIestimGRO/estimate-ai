Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

& (Join-Path $PSScriptRoot "server_stop.ps1")
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& (Join-Path $PSScriptRoot "server_start.ps1")
exit $LASTEXITCODE

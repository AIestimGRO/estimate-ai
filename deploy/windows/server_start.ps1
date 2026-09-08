Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")
Assert-Administrator

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) { throw "Scheduled task '$TaskName' is not installed. Run server_install.ps1 first." }

if (Wait-ServerPort -TimeoutSeconds 1) {
    Write-Host "Estimate AI is already listening on port $ServerPort."
    exit 0
}

Start-ScheduledTask -TaskName $TaskName
if (-not (Wait-ServerPort -TimeoutSeconds 45)) { throw "Estimate AI did not open port $ServerPort. Check data\logs\server.log." }

Write-Host "Estimate AI started."
Write-Host "HTTPS port: $ServerPort"

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")
Assert-Administrator

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $task) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
}

Start-Sleep -Seconds 2
Stop-ServerPortListeners -Port $ServerPort
Start-Sleep -Milliseconds 500

if (Wait-ServerPort -TimeoutSeconds 1) { throw "Port $ServerPort is still in use." }

Write-Host "Estimate AI stopped. Port $ServerPort is free."

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")

Set-Location -LiteralPath $ProjectRoot
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
$taskState = if ($null -eq $task) { "NOT_INSTALLED" } else { [string]$task.State }
$portOpen = Wait-ServerPort -TimeoutSeconds 1
$head = (& git rev-parse --short HEAD 2>$null | Select-Object -Last 1)
$branch = (& git branch --show-current 2>$null | Select-Object -Last 1)
if (-not $branch) { $branch = "DETACHED" }

Write-Host "Task:        $TaskName [$taskState]"
Write-Host "Port:        $ServerPort [$(if ($portOpen) { 'LISTENING' } else { 'CLOSED' })]"
Write-Host "Git:         $branch $head"
Write-Host "Database:    $(if (Test-Path -LiteralPath $DatabasePath) { $DatabasePath } else { 'MISSING' })"
Write-Host "Model:       $(if (Test-Path -LiteralPath $ModelDirectory) { $ModelDirectory } else { 'MISSING (TKP semantic matching unavailable)' })"
Write-Host "Certificate: $(if ($CertFile -and (Test-Path -LiteralPath $CertFile)) { $CertFile } else { 'MISSING' })"
Write-Host "Log:         $(Join-Path $LogDirectory 'server.log')"

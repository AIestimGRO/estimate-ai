Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")
Assert-Administrator

$backup = Backup-Database -Label "manual"
if (-not $backup) { throw "Database file not found: $DatabasePath" }

Write-Host "Database backup created: $backup"

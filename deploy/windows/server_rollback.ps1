param([string]$Commit = "")

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")
Assert-Administrator
Set-Location -LiteralPath $ProjectRoot
Ensure-ServerDirectories

$trackedChanges = @(& git status --porcelain --untracked-files=no)
if ($trackedChanges.Count -gt 0) { throw "Tracked files have local changes. Rollback aborted." }

if (-not $Commit) {
    $previousPath = Join-Path $RuntimeDirectory "previous_good_commit.txt"
    if (-not (Test-Path -LiteralPath $previousPath)) { throw "No previous_good_commit.txt found. Pass -Commit explicitly." }
    $Commit = (Get-Content -LiteralPath $previousPath -Raw).Trim()
}

& git cat-file -e "$Commit^{commit}"
if ($LASTEXITCODE -ne 0) { throw "Unknown commit: $Commit" }

$backup = Backup-Database -Label "pre-rollback"
if ($backup) { Write-Host "Database backup: $backup" }

& (Join-Path $PSScriptRoot "server_stop.ps1")
& git switch --detach $Commit
if ($LASTEXITCODE -ne 0) { throw "Could not switch to rollback commit." }

& (Join-Path $PSScriptRoot "server_start.ps1")
if ($LASTEXITCODE -ne 0) { throw "Rollback commit selected, but server did not start." }

Write-Host "Runtime rolled back to $Commit"
Write-Host "Run server_update.ps1 to return to feature/admin-ui."

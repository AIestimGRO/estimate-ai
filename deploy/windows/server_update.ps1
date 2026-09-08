param([switch]$SkipTests)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")
Assert-Administrator
Set-Location -LiteralPath $ProjectRoot
Ensure-ServerDirectories

$trackedChanges = @(& git status --porcelain --untracked-files=no)
if ($trackedChanges.Count -gt 0) { throw "Tracked files have local changes. Update aborted." }

$oldHead = (& git rev-parse HEAD).Trim()
Set-Content -LiteralPath (Join-Path $RuntimeDirectory "previous_good_commit.txt") -Value $oldHead -Encoding ASCII

$backup = Backup-Database -Label "pre-update"
if ($backup) { Write-Host "Database backup: $backup" }

& (Join-Path $PSScriptRoot "server_stop.ps1")

try {
    & git fetch origin
    if ($LASTEXITCODE -ne 0) { throw "git fetch failed." }

    & git switch feature/admin-ui
    if ($LASTEXITCODE -ne 0) { throw "git switch feature/admin-ui failed." }

    & git pull --ff-only origin feature/admin-ui
    if ($LASTEXITCODE -ne 0) { throw "git pull --ff-only failed." }

    & $VenvPython -m pip install -r requirements-semantic.txt
    if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed." }

    if (-not $SkipTests) {
        & $VenvPython -m pytest -q
        if ($LASTEXITCODE -ne 0) { throw "pytest failed." }
    }

    Test-DatabaseIntegrity

    $newHead = (& git rev-parse HEAD).Trim()
    Set-Content -LiteralPath (Join-Path $RuntimeDirectory "last_good_commit.txt") -Value $newHead -Encoding ASCII

    & (Join-Path $PSScriptRoot "server_start.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Server start failed." }

    Write-Host "Update complete: $oldHead -> $newHead"
} catch {
    Write-Warning "Update failed: $($_.Exception.Message)"
    Write-Warning "Returning runtime to previous commit: $oldHead"
    & git switch --detach $oldHead
    & (Join-Path $PSScriptRoot "server_start.ps1")
    throw
}

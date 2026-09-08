param(
    [string]$CertFile = "",
    [string]$KeyFile = "",
    [int]$Port = 7777,
    [switch]$SkipSemantic,
    [switch]$SkipTests
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "server_common.ps1")
Assert-Administrator
Set-Location -LiteralPath $ProjectRoot
Ensure-ServerDirectories

Write-Host "Project: $ProjectRoot"
Write-Host "Port:    $Port"

$pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
if ($null -eq $pythonCommand) { $pythonCommand = Get-Command python -ErrorAction SilentlyContinue }
if ($null -eq $pythonCommand) { throw "Python is not installed or is not in PATH." }

if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
    Write-Host "Creating .venv ..."
    & $pythonCommand.Source -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "python -m venv failed." }
}

Write-Host "Installing Python dependencies ..."
& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed." }

$requirements = if ($SkipSemantic) { "requirements.txt" } else { "requirements-semantic.txt" }
& $VenvPython -m pip install -r $requirements
if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed." }

$pair = Resolve-CertificatePair -RequestedCert $CertFile -RequestedKey $KeyFile
Write-LocalServerConfig -ResolvedCert $pair.Cert -ResolvedKey $pair.Key -Port $Port

. (Join-Path $PSScriptRoot "server_config.ps1")

if (Test-Path -LiteralPath $DatabasePath -PathType Leaf) {
    Write-Host "Checking SQLite database ..."
    Test-DatabaseIntegrity
} else {
    Write-Warning "Database is not present yet: $DatabasePath"
    Write-Warning "Copy the laptop database before starting the service."
}

if (-not $SkipTests) {
    Write-Host "Running full pytest suite ..."
    & $VenvPython -m pytest -q
    if ($LASTEXITCODE -ne 0) { throw "pytest failed. Scheduled task was not installed." }
}

$taskScript = Join-Path $PSScriptRoot "server_task.ps1"
$actionArgs = '-NoProfile -ExecutionPolicy Bypass -File "' + $taskScript + '"'
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $actionArgs
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 5 -RestartInterval (New-TimeSpan -Minutes 1) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Estimate AI HTTPS service on port $ServerPort" -Force | Out-Null

$existingFirewall = Get-NetFirewallRule -DisplayName $FirewallRuleName -ErrorAction SilentlyContinue
if ($null -ne $existingFirewall) { $existingFirewall | Remove-NetFirewallRule }
New-NetFirewallRule -DisplayName $FirewallRuleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $ServerPort | Out-Null

$head = (& git rev-parse HEAD).Trim()
Set-Content -LiteralPath (Join-Path $RuntimeDirectory "last_good_commit.txt") -Value $head -Encoding ASCII

Write-Host ""
Write-Host "Scheduled task installed: $TaskName"
Write-Host "Firewall opened: TCP $ServerPort"
Write-Host "Certificate: $CertFile"
Write-Host "Key:         $KeyFile"

if (Test-Path -LiteralPath $DatabasePath -PathType Leaf) {
    & (Join-Path $PSScriptRoot "server_start.ps1")
} else {
    Write-Host "Service was not started because the database is missing."
}

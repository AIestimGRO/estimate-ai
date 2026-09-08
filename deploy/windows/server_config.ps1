# Estimate AI Windows server defaults.
# Machine-specific values are loaded from server.local.ps1 when present.

$ProjectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))
$ServerHost = "0.0.0.0"
$ServerPort = 7777
$TaskName = "EstimateAI"
$FirewallRuleName = "Estimate AI 7777"
$CertDirectory = "C:\Users\ololl\Desktop\acme\certs"
$DatabasePath = Join-Path $ProjectRoot "data\estimate_ai.db"
$ModelDirectory = Join-Path $ProjectRoot "data\models\qwen3-embedding-0.6b"
$LogDirectory = Join-Path $ProjectRoot "data\logs"
$BackupDirectory = Join-Path $ProjectRoot "data\backups"
$RuntimeDirectory = Join-Path $ProjectRoot "data\runtime"
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$CertFile = ""
$KeyFile = ""

$LocalConfigPath = Join-Path $PSScriptRoot "server.local.ps1"
if (Test-Path -LiteralPath $LocalConfigPath) {
    . $LocalConfigPath
}

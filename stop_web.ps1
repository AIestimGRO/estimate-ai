# Stop Estimate AI web UI (uvicorn reload leaves orphan worker processes on Windows).
Set-Location -LiteralPath $PSScriptRoot

$Port = 8000
$MaxRounds = 4

function Get-ListenerPids([int]$ListenPort) {
    $pids = [System.Collections.Generic.HashSet[int]]::new()
    try {
        foreach ($conn in Get-NetTCPConnection -LocalPort $ListenPort -State Listen -ErrorAction Stop) {
            [void]$pids.Add([int]$conn.OwningProcess)
        }
    } catch {
        # Fallback when Get-NetTCPConnection is unavailable.
    }

    foreach ($line in (netstat -ano | Select-String ":$ListenPort\s" | Select-String 'LISTENING')) {
        $procId = ($line -split '\s+')[-1]
        if ($procId -match '^\d+$') {
            [void]$pids.Add([int]$procId)
        }
    }
    return @($pids)
}

function Stop-ProcessTree([int]$ProcId, [string]$Reason) {
    if ($ProcId -le 0) { return }
    if (-not (Get-Process -Id $ProcId -ErrorAction SilentlyContinue)) { return }
    Write-Host "Stopping PID $ProcId ($Reason) ..."
    & taskkill.exe /F /T /PID $ProcId 2>&1 | ForEach-Object { Write-Host $_ }
}

function Stop-EstimateAiPythonProcesses {
    $procs = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^python(w)?\.exe$' })

    foreach ($proc in $procs) {
        $cmd = [string]$proc.CommandLine
        $procId = [int]$proc.ProcessId
        $reason = $null

        if ($cmd -match 'app\.web') {
            $reason = 'app.web'
        } elseif ($cmd -match 'uvicorn') {
            $reason = 'uvicorn'
        } elseif ($cmd -match 'multiprocessing\.spawn.*spawn_main') {
            $reason = 'uvicorn worker'
            if ($cmd -match 'parent_pid=(\d+)') {
                $parentId = [int]$Matches[1]
                if (-not (Get-Process -Id $parentId -ErrorAction SilentlyContinue)) {
                    $reason = 'orphan uvicorn worker'
                }
            }
        } elseif ($cmd -match [regex]::Escape($PSScriptRoot)) {
            $reason = 'estimate-ai python'
        }

        if ($reason) {
            Stop-ProcessTree -ProcId $procId -Reason $reason
        }
    }
}

for ($round = 1; $round -le $MaxRounds; $round++) {
    Stop-EstimateAiPythonProcesses

    $seen = @{}
    foreach ($procId in (Get-ListenerPids -ListenPort $Port)) {
        if ($seen[$procId]) { continue }
        $seen[$procId] = $true
        Stop-ProcessTree -ProcId $procId -Reason "port $Port listener"
    }

    Start-Sleep -Milliseconds 800
    if (-not (Get-ListenerPids -ListenPort $Port)) { break }
}

$left = Get-ListenerPids -ListenPort $Port
if ($left) {
    Write-Warning "Port $Port still in use by PID(s): $($left -join ', ')"
    Write-Warning "Open Task Manager, end all python.exe, then run stop_web again."
    exit 1
}

Write-Host "OK: port $Port is free."
exit 0

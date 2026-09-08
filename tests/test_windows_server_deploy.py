from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEPLOY = ROOT / "deploy" / "windows"


def _text(name: str) -> str:
    return (DEPLOY / name).read_text(encoding="utf-8")


def test_windows_server_bundle_contains_required_operations() -> None:
    required = {
        "server_config.ps1",
        "server_common.ps1",
        "server_task.ps1",
        "server_install.ps1",
        "server_start.ps1",
        "server_stop.ps1",
        "server_restart.ps1",
        "server_status.ps1",
        "server_update.ps1",
        "server_rollback.ps1",
        "server_backup_db.ps1",
    }
    assert required <= {path.name for path in DEPLOY.iterdir() if path.is_file()}


def test_windows_server_defaults_to_https_port_7777() -> None:
    config = _text("server_config.ps1")
    task = _text("server_task.ps1")

    assert '$ServerPort = 7777' in config
    assert '$ServerHost = "0.0.0.0"' in config
    assert 'C:\\Users\\ololl\\Desktop\\acme\\certs' in config
    assert '"--ssl-certfile"' in task
    assert '"--ssl-keyfile"' in task
    assert '"--port", [string]$ServerPort' in task


def test_windows_install_creates_autostart_task_and_firewall_rule() -> None:
    installer = _text("server_install.ps1")

    assert "New-ScheduledTaskTrigger -AtStartup" in installer
    assert 'New-ScheduledTaskPrincipal -UserId "SYSTEM"' in installer
    assert "Register-ScheduledTask" in installer
    assert "New-NetFirewallRule" in installer
    assert "pytest -q" in installer


def test_windows_update_is_fast_forward_only_and_backs_up_database() -> None:
    update = _text("server_update.ps1")
    rollback = _text("server_rollback.ps1")

    assert "Backup-Database" in update
    assert "git pull --ff-only origin feature/admin-ui" in update
    assert "git switch --detach $oldHead" in update
    assert "git clean" not in update.lower()
    assert "reset --hard" not in update.lower()
    assert "Backup-Database" in rollback


def test_windows_runtime_state_is_gitignored() -> None:
    ignore = (ROOT / ".gitignore").read_text(encoding="utf-8")

    for entry in (
        "deploy/windows/server.local.ps1",
        "data/backups/",
        "data/logs/",
        "data/runtime/",
        "data/workspaces/",
        "data/models/",
    ):
        assert entry in ignore


def test_windows_server_task_tolerates_native_uvicorn_stderr() -> None:
    task = _text("server_task.ps1")

    invoke_marker = "& $VenvPython @arguments *>> $logPath"
    invoke_index = task.index(invoke_marker)
    continue_index = task.rfind('$ErrorActionPreference = "Continue"', 0, invoke_index)
    restore_index = task.index("$ErrorActionPreference = $previousErrorActionPreference", invoke_index)

    assert continue_index >= 0
    assert continue_index < invoke_index < restore_index
    assert "$exitCode = $LASTEXITCODE" in task

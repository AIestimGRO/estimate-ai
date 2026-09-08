# Windows autonomous deployment

This deployment target is intended for the dedicated Windows host used by
Estimate AI.

## Target layout

- Project: `C:\Users\ololl\Desktop\ISLAMS_DIR`
- Git branch: `feature/admin-ui`
- Database: `data\estimate_ai.db`
- Qwen model: `data\models\qwen3-embedding-0.6b`
- TLS certificates: `C:\Users\ololl\Desktop\acme\certs`
- HTTPS port: `7777`
- Scheduled task: `EstimateAI`

The application runs under the Windows `SYSTEM` account through Task
Scheduler. The task starts at Windows startup, runs with highest privileges,
and restarts after failures.

Operational data, logs, backups, workspaces, the database, model weights, and
the generated machine-local server configuration are ignored by Git.

## First deployment

1. Clone `feature/admin-ui` into the target directory.
2. Copy the current laptop database to `data\estimate_ai.db`.
3. Copy the Qwen model directory to
   `data\models\qwen3-embedding-0.6b` when TKP semantic matching is needed.
4. Make sure the ACME certificate and private key are present under the
   certificate directory.
5. Run an elevated PowerShell in the project root:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_install.ps1
```

The installer:

- creates `.venv`;
- installs `requirements-semantic.txt`;
- detects a certificate/key pair;
- runs the full pytest suite;
- checks SQLite integrity when the database exists;
- creates the startup scheduled task;
- opens inbound TCP port 7777 in Windows Firewall;
- writes the local certificate paths to
  `deploy\windows\server.local.ps1`;
- starts the service when the database exists.

If automatic certificate detection is ambiguous, specify the files explicitly:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_install.ps1 -CertFile "C:\path\fullchain.pem" -KeyFile "C:\path\privkey.pem"
```

## Operations

Status:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_status.ps1
```

Start:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_start.ps1
```

Stop:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_stop.ps1
```

Restart:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_restart.ps1
```

Manual SQLite backup:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_backup_db.ps1
```

Server log:

```text
data\logs\server.log
```

## Safe update from GitHub

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_update.ps1
```

The update script refuses tracked local modifications, creates a consistent
SQLite backup, stops the service, performs `git fetch` and
`git pull --ff-only origin feature/admin-ui`, refreshes Python dependencies,
runs the full pytest suite, checks database integrity, and starts the new
version.

If the update or tests fail, runtime is switched to the previous commit in
detached-HEAD mode and the previous service is started again. No force push,
hard reset, database replacement, or Git clean is used.

## Rollback

The update flow stores the previous known-good commit in
`data\runtime\previous_good_commit.txt`.

Rollback to it:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_rollback.ps1
```

Or specify a commit:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\windows\server_rollback.ps1 -Commit <sha>
```

Rollback also creates a SQLite backup before changing the working-tree commit.

Run `server_update.ps1` to return from detached rollback mode to the current
`feature/admin-ui`.

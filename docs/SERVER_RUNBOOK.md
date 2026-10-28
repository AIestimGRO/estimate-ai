# Windows production runbook — Estimate AI

**Updated 2026-10-28.** Operational guide based on the live SSH session reported
on 2026-10-27, not a substitute for checking current server configuration.
**Requires explicit operator approval and a maintenance window** for any
operation that stops the service or changes a production database.

## Confirmed production layout on 2026-10-27

| Item | Value |
| --- | --- |
| Server OS | Windows (SSH shell exposes `cmd.exe`, Python 3.13 available as `py -3`) |
| Project root | `C:\Users\ololl\Desktop\ISLAMS_DIR` |
| Live database | `C:\Users\ololl\Desktop\ISLAMS_DIR\data\estimate_ai.db` |
| Backups | `C:\Users\ololl\Desktop\ISLAMS_DIR\data\backups\` |
| Staging | `C:\Users\ololl\Desktop\ISLAMS_DIR\_db_staging\` |
| Windows Scheduled Task | `EstimateAI` |
| Task wrapper observed | `deploy\windows\server_task.ps1` |
| Web server | Uvicorn on HTTPS port `7777` |
| DB override | `ESTIMATE_AI_DB_PATH` passed by server task wrapper |
| Verified installed data on 2026-10-27 | `tkp_sources = 633`, `tkp_items = 34763` |

**Do not confuse with developer defaults:** `python -m app.web` uses port 8000
without the production SSL/task wrapper. The deployment scripts are **not
present in the provided code snapshot**; verify their current contents on the
real server before editing them or making claims about restart behavior.

Never print or ship TLS private keys, SSH private keys, credentials, tokens,
or real customer data. Do not store the server's public IP/SSH key material
in source documentation; obtain connection details securely from the operator.

## A. Read-only preflight (can be done while app is running)

From `cmd.exe` over SSH:

```bat
powershell -NoProfile -Command "Get-ScheduledTask -TaskName 'EstimateAI' | Select-Object TaskName,State"
netstat -ano | findstr ":7777"
py -3 --version
```

Inspect running server processes if necessary:

```bat
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -match 'uvicorn' } | Select-Object ProcessId,Name,CommandLine | Format-List"
```

Check the **actual** DB path from the wrapper and `ESTIMATE_AI_DB_PATH`, not
merely `data/config/db.json`. The code resolves `ESTIMATE_AI_DB_PATH` before
`data/config/db.json` (`core/storage/connection.py`).

## B. Safe backup (online SQLite backup, no service stop required)

Use SQLite's connection backup API, **not** `copy` of a live WAL-mode SQLite
file. The following `cmd.exe` one-liner reads live SQLite, creates a timestamped
backup and checks integrity. It does not modify the live DB:

```bat
py -3 -c "import sqlite3,datetime,pathlib; root=pathlib.Path(r'C:\Users\ololl\Desktop\ISLAMS_DIR'); live=root/'data'/'estimate_ai.db'; d=root/'data'/'backups'; d.mkdir(parents=True,exist_ok=True); out=d/('estimate_ai_backup_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+'.db'); a=sqlite3.connect(str(live)); b=sqlite3.connect(str(out)); a.backup(b); b.close(); a.close(); v=sqlite3.connect(str(out)); print('BACKUP:',out); print('Integrity:',v.execute('PRAGMA integrity_check').fetchone()[0]); print('KL:',v.execute('SELECT COUNT(*) FROM tkp_sources').fetchone()[0]); print('Items:',v.execute('SELECT COUNT(*) FROM tkp_items').fetchone()[0]); v.close()"
```

Require `Integrity: ok`, record path, timestamps and counts. Store an
**additional off-server copy** according to local data policy, because a backup
in the same project folder will not protect against full disk/project loss.

## C. Changes to catalog DB — staging and reconciliation

Never replace `data/estimate_ai.db` just because a newer-looking `.db` has
more rows. The live DB also contains RNMC, approvals, correction history,
manual rules and import history. To add new TKP to live, prefer creating a
**fresh copy of the stopped live DB** and applying only verified TKP deltas.

Use the procedure below as review criteria (the October 2026 migration used
this shape); it is **not** authorization for blind SQL reuse:

1. Record baseline counts and table structure on both DBs, `PRAGMA
   integrity_check`, `PRAGMA foreign_key_check`, and
   `PRAGMA user_version`/`schema_migrations`.
2. Inspect `tkp_sources` and `tkp_items` schemas. Identify newly added source
   rows by correct **document identity and revision**, not merely by integer ID.
3. Verify old source/item rows are identical (`SELECT ... EXCEPT SELECT ...`)
   when using an append-only snapshot derived from the same original DB.
4. Ensure incoming source IDs and item `source_id` references are consistent,
   old IDs do not conflict, and filename/revision uniqueness rules hold.
5. Make a staging DB snapshot with `sqlite3.Connection.backup`; insert into
   **staging only**, transactionally, with foreign keys enabled. Recheck
   `foreign_key_check`, `integrity_check`, rows, totals, and a sample of original
   source-backed prices. Do not cherry-pick individual `tkp_items` without
   their source record.
6. Verify other tables (especially `catalog_items`, `imported_files`,
   correction/approval/rule tables) remained from the latest live state.
7. **Stop app and repeat staging from the final live snapshot**, since live
   changes may occur during preparation. Only then install.

**Historical, not automatic invariants:** on 2026-10-27 the original DB was
536 sources / 31,925 items; verified incoming DB 633 / 34,763; the two original
TKP tables matched in an `EXCEPT` check; append-only delta was 97 sources and
2,838 items. A future DB may differ; never hardcode these assertions in a
reusable migration tool.

## D. Controlled stop and activation

1. Schedule downtime; tell users not to process files. Disable task to
   prevent auto restart, then stop it:

   ```bat
   powershell -NoProfile -Command "Disable-ScheduledTask -TaskName 'EstimateAI' | Out-Null; Stop-ScheduledTask -TaskName 'EstimateAI'; Start-Sleep -Seconds 3; Get-ScheduledTask -TaskName 'EstimateAI' | Select-Object TaskName,State"
   ```

2. Prove there are **zero matching Uvicorn processes and zero listeners** on
   `7777`; do not kill Jupyter or other Python workers. If a Uvicorn process
   persists, investigate its parent/wrapper before force-termination.
3. Take **another final verified SQLite backup after stop** and record its
   exact filename. The 27.10 historical backup is
   `data\backups\estimate_ai_pre_switch_20261027_100556.db` — do not reuse
   it as the only backup for a later deployment.
4. Verify `estimate_ai_ready.db` from the **final** live snapshot is complete,
   `integrity=ok`, FK valid, and expected audited source records present.
5. Check `estimate_ai.db-wal`, `estimate_ai.db-shm`, and
   `estimate_ai.db-journal`. If present, **STOP** and investigate checkpoint/
   ownership; do **not** delete them reflexively or replace just one member
   of an active SQLite database.
6. Copy staged DB into `data` under an unused temporary name on the **same
   filesystem**; verify full SHA256/file equality and SQLite integrity of
   that temporary file. Replace the live DB **atomically** only after all
   checks pass, retaining the tested pre-switch backup. Do not use raw
   filesystem rename across volumes.
7. Verify installed DB from `data/estimate_ai.db`, then enable and start task:

   ```bat
   powershell -NoProfile -Command "Enable-ScheduledTask -TaskName 'EstimateAI' | Out-Null; Start-ScheduledTask -TaskName 'EstimateAI'; Start-Sleep -Seconds 8; Get-ScheduledTask -TaskName 'EstimateAI' | Select-Object TaskName,State"
   ```

8. Ensure listener is up, then HTTP and **functional** application checks:

   ```bat
   netstat -ano | findstr ":7777"
   curl.exe -k -sS -o NUL -w "HTTP %{http_code}\n" https://127.0.0.1:7777/admin/tkp
   ```

   HTTP **303** is a redirect and confirms only the route responds; it is
   **not** proof the admin UI successfully loaded or that matching works.
   `-k` skips TLS validation and is acceptable for a loopback health probe,
   not as a recommendation for normal remote browsing.
9. Complete an authorized admin session, inspect task `6940345` (historical
   October smoke-test example), and run a known estimate that can pick an
   analog from a **new** TKP task. Confirm task number, name, unit,
   winner price without VAT and output Excel. Record success/failure.

## E. Rollback if service or results are wrong

1. Disable/stop `EstimateAI` and prove no worker owns the database.
2. Preserve the suspect post-deploy database under a **new** diagnostic name,
   outside the live path.
3. Verify the most recent pre-switch SQLite backup (`integrity_check=ok`,
   expected table counts, FK). Do not restore from an unverified raw file.
4. When no live WAL/SHM/journal conflict remains, restore backup to
   `data/estimate_ai.db` using the same verified same-volume staging approach.
5. Enable/start scheduled task and check both HTTP and functional behavior.
6. Record the reason and timestamps; never delete backups automatically.

## F. Prior post-deployment evidence, 2026-10-27

The user reported:

```text
Uvicorn processes: 0
Port 7777 listeners: 0
Task: Disabled
Backup: ...\estimate_ai_pre_switch_20261027_100556.db
DATABASE SWITCH OK
TaskName EstimateAI Running
Integrity: ok
KL: 633
Items: 34763
TCP 0.0.0.0:7777 LISTENING
HTTP 303
```

This is evidence of a successful SQLite switch and server start **at that
moment**, not proof of an end-to-end current matching result or future uptime.

## G. Operational pitfalls

- Port `8000` is the local/dev default, not the observed production port.
- Python process trees may include **other** applications (e.g. Jupyter).
- Do not treat a Scheduled Task's `Running` status as proof of HTTP readiness.
- The `ATTACH DATABASE` parameter in SQLite may reject Windows `file:///...`
  URI with `?mode=ro` when passed as a plain ATTACH filename. During the
  October session this caused `OperationalError: unable to open database`.
  Check Python/SQLite URI usage in **read-only staging** before any write;
  never bypass controls by writing to live while debugging a path.
- `schema_migrations` and real table columns must be checked before importing
  any snapshot; counts alone do not prove compatibility.
- Logs, backups, `data/workspaces`, and `data/models` are runtime/private
  artifacts, not source files. Exclude them from commits and delivery ZIPs.

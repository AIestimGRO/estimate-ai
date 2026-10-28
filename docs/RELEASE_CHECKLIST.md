# Development and release checklist

Reviewed 2026-10-28. This is a safety gate, not authorization to deploy.
Follow `SERVER_RUNBOOK.md` for the full Windows/SQLite procedure.

## Before change

- [ ] Confirm user-provided ZIP is the **code source of truth**; verify GitHub
      `feature/admin-ui` only to detect divergence. Do not silently overwrite
      newer GitHub changes. For autonomous work use a temporary `work/...`
      branch and review/green CI before fast-forward.
- [ ] Read root `AGENTS.md`, `AI_HANDOFF.md`, relevant domain/workflow docs,
      source module, fixtures and tests.
- [ ] Run **full** `python -m pytest -q` and record actual baseline.
- [ ] Identify whether the task affects legacy behavior, persisted schema,
      user data, Excel formatting, server state or a data migration.
- [ ] For database work, obtain a fresh SQLite backup, check
      `PRAGMA integrity_check` and `PRAGMA foreign_key_check`, compare schema
      and current counts, and retain a rollback artifact.

## Implement and review

- [ ] Reproduce failures first (failing test), make minimal root-cause fixes.
- [ ] For KL/TKP, verify winner, unit/scale, row/column origin, arithmetic
      `unit price × quantity ≈ total`, source file/task and duplicate policy.
- [ ] For RNMC, preserve filename-only uniqueness and historical File_Log
      processing semantics.
- [ ] For Excel, compare unrequested cells/formulas/styles/sheets and special
      OOXML parts with the untouched original.
- [ ] Run full `python -m pytest -q` after changes and record result.
- [ ] Verify included files; no production DB, `.db-wal`, backups, client
      originals, logs, SSH/TLS keys, model weights or generated data.

## Publish to GitHub (separate from generating ZIP)

- [ ] Compare working branch HEAD with the supplied ZIP, reconcile any
      differences. Do **not** force push or reset destructively.
- [ ] Review `git status` and `git diff` for **only intended** file changes.
- [ ] Commit on a temporary work branch, push and await green CI, then use
      approved PR or fast-forward into `feature/admin-ui`.
- [ ] Record commit SHA and CI results. **A local ZIP is not a GitHub push.**

## Production activation (operator-approved maintenance window)

- [ ] Disable and stop Scheduled Task `EstimateAI`; confirm Uvicorn workers
      and port `7777` are stopped, do not stop unrelated Python/Jupyter.
- [ ] Snapshot the **final** production DB; test staging changes against it.
- [ ] Check WAL/SHM/journal state; never replace an active SQLite main file.
- [ ] Use same-volume staged/verified replacement, retain pre-switch backup.
- [ ] Enable/start task, check listeners, HTTP status, authenticated UI.
- [ ] Perform end-to-end estimate using at least one known newly added TKP
      task and confirm provenance, unit and price.
- [ ] Record post-deploy state; preserve tested rollback instructions.

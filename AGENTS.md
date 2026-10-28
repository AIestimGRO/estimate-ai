# Estimate AI — agent entry point (updated 2026-10-28)

**Start with `docs/AI_HANDOFF.md`.** It is the dated, source-checked handoff,
including what is known only from the October 2026 production session. Then
read `docs/AGENTS.md`, `docs/DOMAIN_RULES.md`, `docs/RNMC_IMPORT.md`,
`docs/ROADMAP.md`, `docs/RELEASE_CHECKLIST.md`, and `docs/SERVER_RUNBOOK.md` as applicable.

## Authority and freshness

1. The ZIP supplied by the user for **this task** is the source of truth for
   code. The canonical GitHub project is `AIestimGRO/estimate-ai`, working
   branch `feature/admin-ui`, but inspect GitHub only for verification when
   necessary. A ZIP without `.git` cannot prove its branch or commit hash.
2. A live production database and server are **different** from the source ZIP.
   Never infer their contents from GitHub, a local development database, or
   documentation. Verify production state before any change.
3. If documents conflict, prefer verified code/tests for current behavior,
   user-confirmed rules for desired behavior, and real production checks for
   production state. Explicitly flag contradictions rather than assuming.
4. Read-only investigation first. Never place real pricing/client databases,
   backups, workbooks, private keys, models, or runtime artifacts in the
   deliverable ZIP or a commit.

## Mandatory working protocol

- Communicate with the user in Russian. Code, filenames, identifiers and code
  comments should use English/ASCII; documents may explain Russian terms.
- Before **and** after modifications run `python -m pytest -q` and report the
  actual results. Add regression tests for logic changes, starting with a
  reproducing failing test for a bug.
- Keep changes targeted, deterministic, auditable, and minimally invasive.
  Do not alter price or matching policy silently; do not invent analog prices.
- Return a ZIP with the updated files, **not a patch**; give a precise file
  list and test results. Preserve all unrelated files and workbook content.
- For Git changes, prefer a temporary `work/...` branch; do not force-push,
  destructively reset, or silently update `feature/admin-ui`. Integrate only
  after green tests/CI and explicit permission.
- The original VBA under `reference/vba_original/` documents legacy behavior.
  Check `docs/DOMAIN_RULES.md` and tests before changing a ported rule.
- No unreviewed LLM/embedding prices in the live RNMC matching path. Semantic
  TKP prototypes are isolated, optional, read-only shadow comparisons.
- A production SQLite replacement is a deployment operation, **not** a normal
  code change. Follow `docs/SERVER_RUNBOOK.md` and never swap a live DB.

## Quick commands and navigation

```bash
python -m pytest -q
python -m app.web --no-browser  # local development only; port 8000 by default
```

- `core/` — deterministic rules, Excel, importers, SQLite tables.
- `app/services/` — estimate-processing orchestration and RNMC import.
- `app/web/` — FastAPI routes and admin UI.
- `tests/` — behavioral contracts.
- `data/config/` — config. **Do not commit** `data/estimate_ai.db`.
- `docs/AI_HANDOFF.md` — system map, verified status, pitfalls, next tasks.
- `docs/SERVER_RUNBOOK.md` — Windows deployment, backups, verification,
  rollback (including October 27 DB update).
- `docs/DATA_IMPORT_GUARDRAILS.md` — RNMC/TKP/Excel integrity and provenance.

**Warning:** some older docs and `.cursor` instructions used to describe the
project as "Phase 0, no web/db". That is obsolete. The current ZIP contains a
FastAPI app and SQLite persistence. The Windows server wrapper scripts from
production are NOT in this supplied source snapshot; do not fabricate them.

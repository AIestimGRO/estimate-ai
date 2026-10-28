# Estimate AI — binding instructions for AI developers

> Updated 2026-10-28. Read the root `AGENTS.md` and `docs/AI_HANDOFF.md` first.
> The former "Phase 0, no web server/database" rule is obsolete: the current
> source ZIP includes production-shaped FastAPI, SQLite and admin functions.

## Non-negotiable requirements

1. **Source precedence:** work from the user's uploaded ZIP; GitHub
   `AIestimGRO/estimate-ai`, branch `feature/admin-ui`, is for verification
   when necessary. A snapshot without `.git` cannot verify repository HEAD.
2. **Data integrity:** never invent prices or confuse RNMC and TKP sources.
   Retain source file, task number, row, unit, winner identity, and uncertainty.
   Preserve Excel content outside the exact requested cells/areas.
3. **Determinism:** core matching, price arithmetic, risk decisions, section
   resolution, and imports are testable and deterministic. An optional semantic
   TKP shadow does not alter live outputs; no unapproved LLM pricing.
4. **Legacy contract:** read `docs/DOMAIN_RULES.md` and the relevant original
   VBA module before changing a ported business rule. When behavior differs,
   document the explicit product decision and add regression tests.
5. **Code hygiene:** English/ASCII code identifiers, filenames and code
   comments; UI and business data may be Russian. Keep I/O, parsing, business
   rules and persistence separate.
6. **Privacy/secrets:** never put real SQLite databases, client workbooks,
   backups, runtime logs, model weights, or SSH/TLS private keys in Git or an
   ordinary source ZIP. Treat runtime data as sensitive.
7. **Tests:** run `python -m pytest -q` before and after edits; write a failing
   regression case before fixing a bug. Record the actual test totals.
8. **Deliverables:** no patches; return a ZIP with updated files and a clear
   inventory of changed files. Do not modify unrelated files. For spreadsheet
   requests, strictly preserve untouched columns/styles/formulas/sheets.
9. **Git:** use `work/...` for independent changes, avoid force push and
   destructive operations; only merge after green CI and approval.
10. **Server:** never overwrite a running SQLite database or assume a local
    snapshot contains the newest production transactions. Follow
    `docs/SERVER_RUNBOOK.md`; prefer `sqlite3.Connection.backup`, tested
    staging, explicit downtime, and reversible activation.

## Before starting a new task

- Confirm which artifact is authoritative (uploaded source ZIP vs live DB vs
  incoming original KL/RNMC workbooks). Ask for current source when absent.
- Read `docs/AI_HANDOFF.md` and the targeted modules/tests; run full baseline.
- Make the smallest consistent change with a regression test or verification
  script. Validate source provenance and downstream implications.
- Re-run full tests, summarize remaining risks, and return a ZIP.

## Product boundaries

- A live RNMC analog matches the normalized GESN code and base unit using
  deterministic VBA-derived rules; region is metadata, not a filter.
- Optional live TKP analog selection uses deterministic lexical similarity;
  this is **not** a guarantee of engineering equivalence.
- TKP shadow strict unit conversion and semantic diagnostics are a separate
  read-only path. Do not silently apply shadow gating to live matches.
- RNMC imports deduplicate by normalized **filename only**, including
  `File_Log.xlsx` statuses such as `нет данных`/`новая РНМЦ`.
- Original KL winner pricing cannot reconstruct every participant's bids.
  The October 2026 database includes winner prices; reviewing a particular
  subcontractor's historical offers may require original KL files.

See `docs/DATA_IMPORT_GUARDRAILS.md` for the exact handling of price columns,
units, participant identification, and workbook audit trails.

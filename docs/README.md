# Estimate AI — documentation index (2026-10-28)

**Start:** [`../AGENTS.md`](../AGENTS.md) → [`AI_HANDOFF.md`](AI_HANDOFF.md) →
relevant area below → code and tests. Do not assume this snapshot is the
current contents of production or of the remote GitHub working branch.

| Goal | Read | Verify in code / tests |
| --- | --- | --- |
| Continue development as an AI | `../AGENTS.md`, `AGENTS.md`, `AI_HANDOFF.md` | full `python -m pytest -q` baseline |
| Understand the system | `ARCHITECTURE.md`, `MVP.md`, `ROADMAP.md` | `app/web/`, `app/services/`, `core/` |
| Change business calculations | `DOMAIN_RULES.md`, `DECISIONS.md` | `reference/vba_original/`, `core/`, `tests/` |
| Change RNMC ZIP/File_Log import | `RNMC_IMPORT.md`, `DATA_IMPORT_GUARDRAILS.md` | `app/services/rnmc_*.py`, `tests/test_rnmc*` |
| Change KL/TKP winner import | `DATA_IMPORT_GUARDRAILS.md`, `OPEN_ITEMS.md` | `core/tkp_folder_ingest.py`, `core/storage/tkp.py`, `tests/test_tkp*` |
| Publish documentation to GitHub | `GITHUB_SYNC.md`, `RELEASE_CHECKLIST.md` | inspect `origin/feature/admin-ui`, run full tests |
| Deploy or alter SQLite | `SERVER_RUNBOOK.md`, `RELEASE_CHECKLIST.md` | live service/DB, fresh backup, actual table schema |
| Prioritize next work | `ROADMAP.md`, `OPEN_ITEMS.md`, `DECISIONS.md` | reproduce outstanding issue |

**Authority:** user-uploaded source ZIP for code; contemporaneous verified
production SQLite for live data; committed docs for historical rationale.
When these differ, report the conflict. Do not blend them into one “latest”
version. Avoid pushing generated or production data.

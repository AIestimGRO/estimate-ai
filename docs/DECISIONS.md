# Decision register — Estimate AI

Reviewed 2026-10-28. This register distinguishes **confirmed product rules**
from issues awaiting a product/engineering decision. Do not turn "proposed"
into "implemented" without user approval and regression tests.

| ID | Status | Rule / decision | Source of truth |
| --- | --- | --- | --- |
| D-001 | Confirmed | Original RNMC workbooks are deduplicated by **normalized filename only**, independent of parent directory/region. Duplicates across folders are invalid. | User decision; `RNMC_IMPORT.md` |
| D-002 | Confirmed | `File_Log.xlsx` status `нет данных`, `новая РНМЦ` or zero **still counts as processed**. | User decision; `RNMC_IMPORT.md` |
| D-003 | Confirmed | Pricing and primary RNMC matching are deterministic, auditable and testable; AI/semantic tools do not select live prices. | `DOMAIN_RULES.md`; `core/matching.py` |
| D-004 | Confirmed | Historical VBA-derived matching/pricing behavior is a compatibility contract; deviations require approval and tests. | `DOMAIN_RULES.md`; `reference/vba_original/` |
| D-005 | Confirmed | Store selected winner TKP prices with source/task/row/unit and distinguish original KL participant offers from winner-only SQLite entries. | `DATA_IMPORT_GUARDRAILS.md`; `core/storage/tkp.py` |
| D-006 | Confirmed | For an explicitly requested "change only column K" Excel task, leave all other content and formatting alone and validate the result against the original. | User acceptance condition from October 2026 |
| D-007 | Operational | Live production DB is **never** safely replaced solely because a newer archive has more TKP rows; stage a verified append/merge from the latest live snapshot, backup, stop and rollback. | `SERVER_RUNBOOK.md`; operator session 2026-10-27 |
| D-008 | Proposed | Require strict units including scale, action/technology and engineering details for **live** TKP candidates. Existing strict shadow rules are not automatically applied to live selection. | `OPEN_ITEMS.md`; requires user sign-off |
| D-009 | Proposed | Support non-standard KL bidder-group offsets and multi-price ambiguities in the штатный import UI with real regression fixtures. | `OPEN_ITEMS.md`; original files 2026-08..10 |
| D-010 | Pending verification | Confirm post-deploy end-to-end selection of newly imported task `6940345`; HTTP 303 and DB integrity only prove service response and storage consistency. | Operator session 2026-10-27 |

Decision changes: record date, approver/user confirmation, changed behavior,
files/tests, data implications, and rollback plan. Keep historical decisions
rather than silently editing them into requirements.

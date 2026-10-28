# GitHub synchronization — documentation handoff

Updated 2026-10-28. Canonical repo `AIestimGRO/estimate-ai`, target branch
`feature/admin-ui`. **Preparing a ZIP does not change GitHub.** The source
snapshot supplied to the AI has no `.git` provenance, and the GitHub branch may
have moved. Never overwrite other people's changes to make docs appear current.

## Safe publication procedure (Windows PowerShell, for repo maintainers)

1. In the existing, clean local clone (`D:\estimate-ai` for this project),
   inspect `git status --short` and `git branch --show-current`. Back up or
   commit unfinished changes before proceeding; do not blindly discard them.
2. Fetch without overwriting: `git fetch origin`.
3. Inspect the target branch: `git log -5 --oneline origin/feature/admin-ui`.
   Verify the local working branch is a fast-forward ancestor; do not force
   reset or push. If history diverges, reconcile it before publishing.
4. Switch to `feature/admin-ui`, `git pull --ff-only origin feature/admin-ui`,
   then create a temporary docs branch:
   `git switch -c work/docs-handoff-2026-10-28`.
5. Extract **only the documentation-only ZIP** to the repository root. Do not
   extract a full source ZIP over a newer Git checkout. The docs-only ZIP
   includes `README.md`, `AGENTS.md`, `docs/`, `.cursor/rules/`, `.gitignore`.
   Examine the resulting `git diff` and `git status --short`; preserve any
   newer independently edited GitHub documentation by merging thoughtfully.
6. Run `python -m pytest -q` and record the result. Confirm that no production
   SQLite DB, backups, workbooks, logs or private key are staged.
7. Stage only reviewed files (e.g. `git add README.md AGENTS.md docs/ .cursor/rules/ .gitignore`),
   commit, and push `work/docs-handoff-2026-10-28` without force.
8. Open a PR **into** `feature/admin-ui`, wait for green CI and review. If the
   team explicitly uses fast-forward-only, advance the working branch only
   after verification. Record the resulting commit SHA in the handoff notes.

Do not alter the live Windows server/database for a documentation release.
Production source deployment and production data migration are separate,
explicitly authorized operations described in `SERVER_RUNBOOK.md`.

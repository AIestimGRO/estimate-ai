# Estimate AI

Construction pricing assistant. Helps process BOQ (ВОР) Excel files, find price
analogs from historical data, and produce a price corridor (min / median / max /
recommended) with risk flags for human review.

The system does not replace the estimator. It produces a draft for review.

## Current stage

Phase 0: porting existing VBA macro logic into a tested Python core.
No web server, no database yet. Everything runs locally on Excel files.

## How to run (once core/ has code)

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
pytest
```

## Project docs

See `docs/AGENTS.md` for AI agent rules, `docs/DOMAIN_RULES.md` for business
rules extracted from the original macros, `docs/MVP.md` for current scope.

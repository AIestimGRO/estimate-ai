# Agent Instructions

You are working on Estimate AI — a tool that matches construction work items
(BOQ rows) against a historical price database and produces a price corridor.

## Hard rules

1. Never invent prices. Every recommended price must trace back to source
   analog rows that actually exist in the data.
2. Matching and price calculation logic must be deterministic and testable.
   No LLM calls inside the matching/pricing functions themselves.
3. Code, filenames, function names, comments: English/ASCII only.
   Russian business terms (e.g. "демонтаж", unit names, region names) live in
   data/config/dictionaries, never hardcoded inside logic as string literals
   scattered through the code.
4. Keep functions small and single-purpose. No function should both parse
   Excel and make matching decisions.
5. Every new domain rule must come with a test that encodes a real example
   (ideally from an actual BOQ/macro file), not a made-up one.
6. Do not add new dependencies, frameworks, or infrastructure (web server,
   database, Docker) unless explicitly asked. Current phase is a local
   Python library only.
7. docs/DOMAIN_RULES.md is extracted from real, working VBA macros (not a
   speculative design). When porting a piece of logic, cite which
   DOMAIN_RULES.md section and which original module/function it came
   from (e.g. in a code comment or commit message: "ports NormUnit,
   Module3"). If you believe the VBA behavior should change, say so
   explicitly and wait for confirmation — do not silently "improve" it
   during a port.

## Workflow

- Before writing code for a new module, summarize your plan in 3-5 bullet
  points and wait for confirmation if the task is ambiguous.
- After implementation, list exactly which files changed.
- One task = one small module or one bugfix. Do not bundle unrelated changes.
- If a bug is found, first add a failing test reproducing it, then fix.

## What NOT to do without being asked

- Do not scaffold a web framework, API server, or database.
- Do not add authentication, user accounts, or multi-tenancy.
- Do not implement PDF parsing.
- Do not implement semantic/AI-based matching before exact/rule-based
  matching is solid and tested.

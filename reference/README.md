# Reference: original VBA macros

This folder holds the original working VBA modules (Module1–Module7) as a
read-only reference for the Python port. Do not edit these — they are the
ground truth being ported, not code to run or modify.

`docs/DOMAIN_RULES.md` is the human-readable extraction of this logic, with
section references back to specific modules/functions. When in doubt about
exact behavior (edge cases, error handling, exact string normalization),
check the original `.bas` file directly rather than relying only on the
summary in DOMAIN_RULES.md.

| File | Role |
|---|---|
| Module1.bas | Settings (`Instrument` sheet config block, `TSettings` type) |
| Module2_updated.bas | Entry point (`RunAnalogSearch`), file picking, orchestration |
| Module3_updated.bas | Normalization (`NormCode`, `NormUnit`, `HasDemontazh`), section dict, `BuildCatalog` |
| Module4_updated.bas | Core estimate processing: matching, risk checks, output writing, price log |
| Module5.bas | Run log writer (`Log` sheet) |
| Module6.bas | GESN_Exceptions: approved price-range workflow |
| Module7_name_exclusions.bas | Name exclusion rules + task color list (`Name_Exclusions` sheet) |
| Module8_catalog_import.bas | Folder-walk catalog ingestion (renamed from author's "Module6.bas" — name collision with the GESN Exceptions module above) |

A catalog-import module exists but was not included here yet — it is
unfinished per the author and will be reviewed separately before porting
(see DOMAIN_RULES.md §9).

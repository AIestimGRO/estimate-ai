import sqlite3
from pathlib import Path
from datetime import datetime

db = Path(r"D:\estimate-ai\data\estimate_ai.db")
stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
backup = db.with_name(f"estimate_ai.db.before-tkp-rollback-{stamp}.bak")

print("=== CREATE DATABASE BACKUP ===")

source = sqlite3.connect(db)
target = sqlite3.connect(backup)
source.backup(target)
target.close()
source.close()

print("Backup:", backup)

con = sqlite3.connect(db)
con.execute("PRAGMA foreign_keys = ON")

start = "2026-08-31 09:34:00"
finish = "2026-08-31 09:35:00"

source_count = con.execute(
    """
    SELECT COUNT(*)
    FROM tkp_sources
    WHERE imported_at >= ? AND imported_at < ?
    """,
    (start, finish),
).fetchone()[0]

item_count = con.execute(
    """
    SELECT COUNT(*)
    FROM tkp_items
    WHERE source_id IN (
        SELECT id
        FROM tkp_sources
        WHERE imported_at >= ? AND imported_at < ?
    )
    """,
    (start, finish),
).fetchone()[0]

print()
print("=== TARGET BATCH ===")
print("sources:", source_count)
print("items:  ", item_count)

if source_count != 671 or item_count != 43687:
    con.close()
    raise SystemExit(
        "ABORT: expected 671 sources and 43687 items. Nothing deleted."
    )

print()
print("=== DELETE ===")

con.execute("BEGIN IMMEDIATE")

con.execute(
    """
    DELETE FROM tkp_sources
    WHERE imported_at >= ? AND imported_at < ?
    """,
    (start, finish),
)

remaining_sources = con.execute(
    "SELECT COUNT(*) FROM tkp_sources"
).fetchone()[0]

remaining_items = con.execute(
    "SELECT COUNT(*) FROM tkp_items"
).fetchone()[0]

print("remaining sources:", remaining_sources)
print("remaining items:  ", remaining_items)

if remaining_sources != 40 or remaining_items != 557:
    con.rollback()
    con.close()
    raise SystemExit(
        "ABORT: unexpected result. Transaction rolled back."
    )

con.commit()
con.close()

print()
print("=== SUCCESS ===")
print("Removed sources: 671")
print("Removed items:   43687")
print("Remaining sources:", remaining_sources)
print("Remaining items:  ", remaining_items)
print("Backup:", backup)

"""Name exclusion rules and task color persistence."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core.exclusions import NameExclusionRule, TaskColorEntry
from core.macro_workbook import (
    load_all_rules_from_workbook,
    resolve_macro_workbook_path,
)


def list_name_exclusion_rules(connection: sqlite3.Connection) -> list[NameExclusionRule]:
    rows = connection.execute(
        """
        SELECT enabled, scope, match_mode, pattern, rule_group, comment
        FROM name_exclusion_rules
        ORDER BY sort_order, id
        """
    ).fetchall()
    return [_row_to_rule(row) for row in rows]


def list_task_color_entries(connection: sqlite3.Connection) -> list[TaskColorEntry]:
    rows = connection.execute(
        """
        SELECT enabled, task_number, reason, comment
        FROM task_color_entries
        ORDER BY sort_order, id
        """
    ).fetchall()
    return [_row_to_task_color(row) for row in rows]


def replace_name_exclusion_rules(
    connection: sqlite3.Connection,
    rules: list[NameExclusionRule],
) -> int:
    connection.execute("DELETE FROM name_exclusion_rules")
    for index, rule in enumerate(rules):
        connection.execute(
            """
            INSERT INTO name_exclusion_rules (
                enabled, scope, match_mode, pattern, rule_group, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(rule.enabled),
                rule.scope,
                rule.match_mode,
                rule.pattern,
                rule.group,
                rule.comment,
                index,
            ),
        )
    connection.commit()
    return len(rules)


def replace_task_color_entries(
    connection: sqlite3.Connection,
    entries: list[TaskColorEntry],
) -> int:
    connection.execute("DELETE FROM task_color_entries")
    for index, entry in enumerate(entries):
        connection.execute(
            """
            INSERT INTO task_color_entries (
                enabled, task_number, reason, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (
                int(entry.enabled),
                entry.task_number,
                entry.reason,
                entry.comment,
                index,
            ),
        )
    connection.commit()
    return len(entries)


def import_rules_from_workbook(
    connection: sqlite3.Connection,
    workbook_path: str | Path | None = None,
) -> tuple[int, int]:
    path = (
        Path(workbook_path).resolve()
        if workbook_path is not None
        else resolve_macro_workbook_path()
    )
    if path is None or not path.is_file():
        raise FileNotFoundError("macro workbook with Name_Exclusions not found")

    rules, colors = load_all_rules_from_workbook(path)
    replace_name_exclusion_rules(connection, rules)
    replace_task_color_entries(connection, colors)
    return len(rules), len(colors)


def _row_to_rule(row: sqlite3.Row) -> NameExclusionRule:
    return NameExclusionRule(
        enabled=bool(row["enabled"]),
        scope=str(row["scope"]),
        match_mode=str(row["match_mode"]),
        pattern=str(row["pattern"]),
        group=str(row["rule_group"]),
        comment=str(row["comment"]),
    )


def _row_to_task_color(row: sqlite3.Row) -> TaskColorEntry:
    return TaskColorEntry(
        enabled=bool(row["enabled"]),
        task_number=str(row["task_number"]),
        reason=str(row["reason"]),
        comment=str(row["comment"]),
    )

"""Name exclusion rules and task color persistence."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core.exclusions import (
    NameExclusionRule,
    TaskColorEntry,
    TaskHighlightReason,
    VALID_MATCH_MODES,
    VALID_SCOPES,
    normalize_reason_key,
    normalize_task_key,
)

_HEX_DIGITS = set("0123456789ABCDEF")
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


def upsert_name_exclusion_rule(
    connection: sqlite3.Connection,
    scope: str,
    match_mode: str,
    pattern: str,
    rule_group: str = "",
    comment: str = "",
    enabled: bool = True,
) -> None:
    scope_value = _normalize_rule_scope(scope)
    mode_value = _normalize_rule_match_mode(match_mode)
    pattern_value = str(pattern).strip()
    if pattern_value == "":
        raise ValueError("pattern is required")

    rows = connection.execute(
        """
        SELECT id, scope, match_mode, pattern
        FROM name_exclusion_rules
        ORDER BY sort_order, id
        """
    ).fetchall()
    for row in rows:
        if (
            _normalize_rule_scope(row["scope"]) == scope_value
            and _normalize_rule_match_mode(row["match_mode"]) == mode_value
            and str(row["pattern"]).strip().lower() == pattern_value.lower()
        ):
            connection.execute(
                """
                UPDATE name_exclusion_rules
                SET enabled = ?, scope = ?, match_mode = ?, pattern = ?,
                    rule_group = ?, comment = ?
                WHERE id = ?
                """,
                (
                    int(enabled),
                    scope_value,
                    mode_value,
                    pattern_value,
                    rule_group.strip(),
                    comment.strip(),
                    int(row["id"]),
                ),
            )
            connection.commit()
            return

    sort_row = connection.execute(
        "SELECT COALESCE(MAX(sort_order), -1) + 1 AS next_sort_order FROM name_exclusion_rules"
    ).fetchone()
    next_sort_order = int(sort_row["next_sort_order"])
    connection.execute(
        """
        INSERT INTO name_exclusion_rules (
            enabled, scope, match_mode, pattern, rule_group, comment, sort_order
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            int(enabled),
            scope_value,
            mode_value,
            pattern_value,
            rule_group.strip(),
            comment.strip(),
            next_sort_order,
        ),
    )
    connection.commit()


def set_name_exclusion_rule_enabled(
    connection: sqlite3.Connection,
    scope: str,
    match_mode: str,
    pattern: str,
    enabled: bool,
) -> bool:
    scope_value = _normalize_rule_scope(scope)
    mode_value = _normalize_rule_match_mode(match_mode)
    pattern_value = str(pattern).strip()
    if pattern_value == "":
        return False

    rows = connection.execute(
        """
        SELECT id, scope, match_mode, pattern
        FROM name_exclusion_rules
        ORDER BY sort_order, id
        """
    ).fetchall()
    for row in rows:
        if (
            _normalize_rule_scope(row["scope"]) == scope_value
            and _normalize_rule_match_mode(row["match_mode"]) == mode_value
            and str(row["pattern"]).strip().lower() == pattern_value.lower()
        ):
            connection.execute(
                "UPDATE name_exclusion_rules SET enabled = ? WHERE id = ?",
                (int(enabled), int(row["id"])),
            )
            connection.commit()
            return True
    return False


def list_task_color_entries(connection: sqlite3.Connection) -> list[TaskColorEntry]:
    rows = connection.execute(
        """
        SELECT enabled, task_number, reason, comment
        FROM task_color_entries
        ORDER BY sort_order, id
        """
    ).fetchall()
    return [_row_to_task_color(row) for row in rows]


def upsert_task_color_entry(
    connection: sqlite3.Connection,
    task_number: str,
    reason: str = "",
    comment: str = "",
    enabled: bool = True,
) -> None:
    task_value = str(task_number).strip()
    task_key = normalize_task_key(task_value)
    if task_key == "":
        raise ValueError("task_number is required")

    rows = connection.execute(
        "SELECT id, task_number FROM task_color_entries ORDER BY sort_order, id"
    ).fetchall()
    for row in rows:
        if normalize_task_key(row["task_number"]) == task_key:
            connection.execute(
                """
                UPDATE task_color_entries
                SET enabled = ?, task_number = ?, reason = ?, comment = ?
                WHERE id = ?
                """,
                (int(enabled), task_value, reason.strip(), comment.strip(), int(row["id"])),
            )
            connection.commit()
            return

    sort_row = connection.execute(
        "SELECT COALESCE(MAX(sort_order), -1) + 1 AS next_sort_order FROM task_color_entries"
    ).fetchone()
    next_sort_order = int(sort_row["next_sort_order"])
    connection.execute(
        """
        INSERT INTO task_color_entries (
            enabled, task_number, reason, comment, sort_order
        ) VALUES (?, ?, ?, ?, ?)
        """,
        (int(enabled), task_value, reason.strip(), comment.strip(), next_sort_order),
    )
    connection.commit()


def set_task_color_enabled(
    connection: sqlite3.Connection,
    task_number: str,
    enabled: bool,
) -> bool:
    task_key = normalize_task_key(task_number)
    if task_key == "":
        return False

    rows = connection.execute(
        "SELECT id, task_number FROM task_color_entries ORDER BY sort_order, id"
    ).fetchall()
    for row in rows:
        if normalize_task_key(row["task_number"]) == task_key:
            connection.execute(
                "UPDATE task_color_entries SET enabled = ? WHERE id = ?",
                (int(enabled), int(row["id"])),
            )
            connection.commit()
            return True
    return False


def list_task_highlight_reasons(connection: sqlite3.Connection) -> list[TaskHighlightReason]:
    rows = connection.execute(
        """
        SELECT key, label, color_hex, enabled
        FROM task_highlight_reasons
        ORDER BY sort_order, id
        """
    ).fetchall()
    return [_row_to_highlight_reason(row) for row in rows]


def upsert_task_highlight_reason(
    connection: sqlite3.Connection,
    key: str,
    label: str,
    color_hex: str,
    enabled: bool = True,
) -> None:
    key_value = normalize_reason_key(key)
    if key_value == "":
        raise ValueError("key is required")

    label_value = str(label).strip()
    if label_value == "":
        raise ValueError("label is required")

    color_value = str(color_hex).strip().lstrip("#").upper()
    if len(color_value) != 6 or any(char not in _HEX_DIGITS for char in color_value):
        raise ValueError("color_hex must be a 6-digit hex value")

    row = connection.execute(
        "SELECT id FROM task_highlight_reasons WHERE key = ?", (key_value,)
    ).fetchone()
    if row is not None:
        connection.execute(
            """
            UPDATE task_highlight_reasons
            SET label = ?, color_hex = ?, enabled = ?
            WHERE id = ?
            """,
            (label_value, color_value, int(enabled), int(row["id"])),
        )
        connection.commit()
        return

    sort_row = connection.execute(
        "SELECT COALESCE(MAX(sort_order), -1) + 1 AS next_sort_order FROM task_highlight_reasons"
    ).fetchone()
    next_sort_order = int(sort_row["next_sort_order"])
    connection.execute(
        """
        INSERT INTO task_highlight_reasons (key, label, color_hex, enabled, sort_order)
        VALUES (?, ?, ?, ?, ?)
        """,
        (key_value, label_value, color_value, int(enabled), next_sort_order),
    )
    connection.commit()


def set_task_highlight_reason_enabled(
    connection: sqlite3.Connection,
    key: str,
    enabled: bool,
) -> bool:
    key_value = normalize_reason_key(key)
    if key_value == "":
        return False

    cursor = connection.execute(
        "UPDATE task_highlight_reasons SET enabled = ? WHERE key = ?",
        (int(enabled), key_value),
    )
    connection.commit()
    return cursor.rowcount > 0


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


def _row_to_highlight_reason(row: sqlite3.Row) -> TaskHighlightReason:
    return TaskHighlightReason(
        key=str(row["key"]),
        label=str(row["label"]),
        color_hex=str(row["color_hex"]),
        enabled=bool(row["enabled"]),
    )


def _normalize_rule_scope(scope: str) -> str:
    scope_key = str(scope).strip().upper()
    if scope_key == "":
        scope_key = "BOTH"
    if scope_key not in VALID_SCOPES:
        raise ValueError("invalid scope")
    return scope_key


def _normalize_rule_match_mode(match_mode: str) -> str:
    mode_key = str(match_mode).strip().upper()
    if mode_key == "":
        mode_key = "ALL_WORDS"
    if mode_key not in VALID_MATCH_MODES:
        raise ValueError("invalid match_mode")
    return mode_key

"""Manual GESN-to-EKR section mapping persistence."""

from __future__ import annotations

from dataclasses import dataclass
import sqlite3

from core.sections import CanonicalGesnCode, GESnPrefix, NormalizeSectionValue


@dataclass(frozen=True)
class ManualSectionMapping:
    """One manual section mapping row."""

    code: str
    code_norm: str
    section_code: str
    enabled: bool
    comment: str
    created_at: str


def list_manual_section_mappings(
    connection: sqlite3.Connection,
    *,
    enabled_only: bool = False,
) -> list[ManualSectionMapping]:
    where = "WHERE enabled = 1" if enabled_only else ""
    rows = connection.execute(
        f"""
        SELECT code, code_norm, section_code, enabled, comment, created_at
        FROM manual_section_mappings
        {where}
        ORDER BY enabled DESC, code_norm
        """
    ).fetchall()
    return [_row_to_mapping(row) for row in rows]


def load_manual_section_map(connection: sqlite3.Connection) -> dict[str, str]:
    return {
        row.code_norm: row.section_code
        for row in list_manual_section_mappings(connection, enabled_only=True)
    }


def upsert_manual_section_mapping(
    connection: sqlite3.Connection,
    *,
    code: object,
    section_code: object,
    comment: str = "",
    enabled: bool = True,
) -> ManualSectionMapping:
    code_display = str(code).strip()
    code_norm = _normalize_code(code)
    section = _normalize_section(section_code)
    comment_value = str(comment or "").strip()

    connection.execute(
        """
        INSERT INTO manual_section_mappings (
            code, code_norm, section_code, enabled, comment
        ) VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(code_norm) DO UPDATE SET
            code = excluded.code,
            section_code = excluded.section_code,
            enabled = excluded.enabled,
            comment = excluded.comment
        """,
        (code_display, code_norm, section, int(enabled), comment_value),
    )
    connection.commit()
    return get_manual_section_mapping(connection, code_norm)


def get_manual_section_mapping(
    connection: sqlite3.Connection,
    code: object,
) -> ManualSectionMapping:
    code_norm = _normalize_code(code)
    row = connection.execute(
        """
        SELECT code, code_norm, section_code, enabled, comment, created_at
        FROM manual_section_mappings
        WHERE code_norm = ?
        """,
        (code_norm,),
    ).fetchone()
    if row is None:
        raise KeyError(code_norm)
    return _row_to_mapping(row)


def set_manual_section_mapping_enabled(
    connection: sqlite3.Connection,
    code: object,
    enabled: bool,
) -> bool:
    code_norm = _normalize_code(code)
    cursor = connection.execute(
        "UPDATE manual_section_mappings SET enabled = ? WHERE code_norm = ?",
        (int(enabled), code_norm),
    )
    connection.commit()
    return cursor.rowcount > 0


def delete_manual_section_mapping(connection: sqlite3.Connection, code: object) -> bool:
    code_norm = _normalize_code(code)
    cursor = connection.execute(
        "DELETE FROM manual_section_mappings WHERE code_norm = ?",
        (code_norm,),
    )
    connection.commit()
    return cursor.rowcount > 0


def _normalize_code(code: object) -> str:
    code_norm = CanonicalGesnCode(code)
    if code_norm == "":
        raise ValueError("code is required")
    if GESnPrefix(code_norm) == "":
        raise ValueError("code must be GESN, FER, or TER")
    return code_norm


def _normalize_section(section_code: object) -> str:
    section = NormalizeSectionValue(section_code)
    if section == "":
        raise ValueError("section_code must be 01-99")
    return section


def _row_to_mapping(row: sqlite3.Row) -> ManualSectionMapping:
    return ManualSectionMapping(
        code=str(row["code"]),
        code_norm=str(row["code_norm"]),
        section_code=str(row["section_code"]),
        enabled=bool(row["enabled"]),
        comment=str(row["comment"]),
        created_at=str(row["created_at"]),
    )

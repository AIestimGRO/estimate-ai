"""Quarantined GSFX/OS evidence, isolated from RNMC and TKP catalogs."""
from __future__ import annotations

import hashlib
import json
import sqlite3
from pathlib import Path
from collections.abc import Sequence
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from app.services.lsr_evidence import EvidenceReport


def source_set_sha256(sources: Sequence[tuple[str, str]]) -> str:
    """Stable identity for a *set* of GSFX files; legacy singles keep their hash."""
    if len(sources) == 1:
        return sources[0][1]
    payload = json.dumps(sorted(sources), ensure_ascii=False,
                         separators=(',', ':')).encode('utf-8')
    return hashlib.sha256(b'estimate-ai-gsfx-set-v1\0' + payload).hexdigest()


def save_learning_case(connection: sqlite3.Connection, *, gsfx_path: Path | None = None,
                       gsfx_paths: Sequence[Path] | None = None,
                       os_path: Path, gsfx_filename: str = '',
                       gsfx_filenames: Sequence[str] | None = None, os_filename: str,
                       sheet_title: str, author_user_id: int,
                       job_token: str, report: EvidenceReport) -> tuple[int, bool]:
    """Insert once; retries may never undo an administrator's decision."""
    if gsfx_paths is None:
        if gsfx_path is None:
            raise ValueError('At least one GSFX must be supplied')
        gsfx_paths = (gsfx_path,)
        gsfx_filenames = (gsfx_filename,)
    if (not gsfx_paths or gsfx_filenames is None or
            len(gsfx_paths) != len(gsfx_filenames) or len(gsfx_paths) > 20):
        raise ValueError('Invalid set of source GSFX files')
    sources = []
    total_size = 0
    for path, filename in zip(gsfx_paths, gsfx_filenames):
        content = path.read_bytes()
        if not content or len(content) > 12 * 1024 * 1024:
            raise ValueError('GSFX exceeds evidence storage limit (12 MB per file)')
        total_size += len(content)
        if total_size > 64 * 1024 * 1024:
            raise ValueError('GSFX set exceeds evidence storage limit (64 MB)')
        sources.append((filename[:250], hashlib.sha256(content).hexdigest(), content))
    sources.sort(key=lambda item: (item[0].casefold(), item[1]))
    if os_path.stat().st_size > 32 * 1024 * 1024:
        raise ValueError('OS exceeds evidence storage limit (32 MB)')
    os_data = os_path.read_bytes()
    os_sha = hashlib.sha256(os_data).hexdigest()
    gsfx_sha = source_set_sha256([(name, checksum) for name, checksum, _ in sources])
    first_name, _, first_source = sources[0]
    with connection:
        result = connection.execute('''
            INSERT OR IGNORE INTO lsr_learning_cases (
                gsfx_sha256, os_sha256, gsfx_filename, os_filename, gsfx_blob, os_blob,
                os_rows_json, issues_json, sheet_title, author_user_id, job_token,
                quality, os_count, gsfx_count, matched_count, checked_prices
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (gsfx_sha, os_sha, first_name, os_filename[:250],
              sqlite3.Binary(first_source), sqlite3.Binary(os_data),
              json.dumps(report.rows, ensure_ascii=False),
              json.dumps(report.issues, ensure_ascii=False), sheet_title,
              author_user_id, job_token, report.quality,
              report.os_count, report.lsr_count, report.matched, report.checked_prices))
        # Cases collected before schema v17 have no original OS BLOB. The
        # exact same input pair can safely backfill it without changing any
        # administrator decision or reviewer attribution.
        connection.execute('''
            UPDATE lsr_learning_cases SET os_blob=?
            WHERE gsfx_sha256=? AND os_sha256=? AND sheet_title=?
              AND os_blob IS NULL
        ''', (sqlite3.Binary(os_data), gsfx_sha, os_sha, sheet_title))
        row = connection.execute('''
            SELECT id FROM lsr_learning_cases
            WHERE gsfx_sha256=? AND os_sha256=? AND sheet_title=?
        ''', (gsfx_sha, os_sha, sheet_title)).fetchone()
        if result.rowcount == 1:
            connection.executemany('''
                INSERT INTO lsr_learning_sources(case_id, source_index, filename, sha256, gsfx_blob)
                VALUES (?, ?, ?, ?, ?)
            ''', ((int(row['id']), index, name, checksum, sqlite3.Binary(content))
                  for index, (name, checksum, content) in enumerate(sources[1:], 2)))
    return int(row['id']), result.rowcount == 1


def list_learning_cases(connection: sqlite3.Connection, *, limit: int = 100):
    return connection.execute('''
        SELECT c.id, c.gsfx_filename, c.os_filename, c.quality, c.status,
               c.os_count, c.gsfx_count, c.matched_count, c.checked_prices,
               c.issues_json, c.created_at, c.review_comment,
               1 + (SELECT COUNT(*) FROM lsr_learning_sources s WHERE s.case_id=c.id) AS source_count,
               c.os_blob IS NOT NULL AS has_original_os,
               COALESCE(u.full_name, '') AS author_name
        FROM lsr_learning_cases c
        LEFT JOIN app_users u ON u.id=c.author_user_id
        ORDER BY c.id DESC LIMIT ?
    ''', (min(max(limit, 1), 500),)).fetchall()


def list_learning_sources(connection: sqlite3.Connection, case_id: int) -> list[dict]:
    row = connection.execute('''
        SELECT gsfx_filename, gsfx_blob FROM lsr_learning_cases WHERE id=?
    ''', (case_id,)).fetchone()
    if row is None:
        return []
    sources = [{'source_index': 1, 'filename': row['gsfx_filename'],
                'content': bytes(row['gsfx_blob'])}]
    for item in connection.execute('''
        SELECT source_index, filename, sha256, gsfx_blob FROM lsr_learning_sources
        WHERE case_id=? ORDER BY source_index
    ''', (case_id,)):
        sources.append({'source_index': int(item['source_index']),
                        'filename': item['filename'], 'sha256': item['sha256'],
                        'content': bytes(item['gsfx_blob'])})
    return sources


def list_learning_source_names(connection: sqlite3.Connection, case_id: int) -> list[tuple[int, str]]:
    row = connection.execute('SELECT gsfx_filename FROM lsr_learning_cases WHERE id=?',
                             (case_id,)).fetchone()
    if row is None:
        return []
    items = [(1, row['gsfx_filename'])]
    items.extend((int(item['source_index']), item['filename']) for item in connection.execute('''
        SELECT source_index, filename FROM lsr_learning_sources
        WHERE case_id=? ORDER BY source_index
    ''', (case_id,)))
    return items


def get_learning_source(connection: sqlite3.Connection, case_id: int,
                        source_index: int = 1) -> bytes | None:
    if source_index == 1:
        row = connection.execute('SELECT gsfx_blob FROM lsr_learning_cases WHERE id=?',
                                 (case_id,)).fetchone()
    else:
        row = connection.execute('''
            SELECT gsfx_blob FROM lsr_learning_sources WHERE case_id=? AND source_index=?
        ''', (case_id, source_index)).fetchone()
    return bytes(row['gsfx_blob']) if row is not None else None


def get_learning_workbook(connection: sqlite3.Connection, case_id: int) -> bytes | None:
    row = connection.execute('SELECT os_blob FROM lsr_learning_cases WHERE id=?', (case_id,)).fetchone()
    return bytes(row['os_blob']) if row is not None and row['os_blob'] is not None else None


def get_learning_case_for_job(connection: sqlite3.Connection, job_token: str):
    return connection.execute('''
        SELECT id, quality, status, matched_count, os_count, checked_prices,
          1 + (SELECT COUNT(*) FROM lsr_learning_sources s WHERE s.case_id=c.id) AS source_count
        FROM lsr_learning_cases c WHERE job_token=? ORDER BY id DESC LIMIT 1
    ''', (job_token,)).fetchone()


def review_learning_case(connection: sqlite3.Connection, *, case_id: int,
                         action: str, reviewer_user_id: int, comment: str = '') -> bool:
    if action not in ('approved', 'rejected'):
        raise ValueError('Invalid review decision')
    if action == 'approved' and len(comment.strip()) > 500:
        raise ValueError('Review comment is too long')
    with connection:
        result = connection.execute('''
            UPDATE lsr_learning_cases
            SET status=?, reviewed_by=?, reviewed_at=datetime('now'), review_comment=?
            WHERE id=? AND status='pending'
                  AND (?='rejected' OR (quality='candidate' AND os_blob IS NOT NULL))
        ''', (action, reviewer_user_id, comment[:500], case_id, action))
    return result.rowcount == 1

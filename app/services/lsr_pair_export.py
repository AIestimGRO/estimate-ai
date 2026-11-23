"""Export only administrator-approved, complete and hash-verified GSFX/OS pairs."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
from io import BytesIO
import json
import sqlite3
from zipfile import ZIP_DEFLATED, ZipFile


@dataclass(frozen=True)
class VerifiedPairExport:
    content: bytes
    count: int
    last_id: int


def export_verified_pairs(connection: sqlite3.Connection, *, after_id: int = 0,
                          limit: int = 50, max_raw_bytes: int = 128 * 1024 * 1024
                          ) -> VerifiedPairExport:
    """A downloadable batch of original XLSX + GSFX + an audit manifest.

    No pending/review/rejected cases, legacy records without original OS,
    or tampered database BLOBs may enter the archive.
    """
    if not (0 <= after_id and 1 <= limit <= 50):
        raise ValueError('Invalid export range')
    cursor = connection.execute('''
        SELECT id, gsfx_sha256, os_sha256, gsfx_filename, os_filename,
               gsfx_blob, os_blob, os_rows_json, issues_json, sheet_title,
               matched_count, os_count, gsfx_count, checked_prices, created_at
        FROM lsr_learning_cases
        WHERE id>? AND status='approved' AND quality='candidate'
          AND os_blob IS NOT NULL
        ORDER BY id ASC LIMIT ?
    ''', (after_id, limit))

    output = BytesIO()
    manifest = []
    raw_bytes = 0
    with ZipFile(output, 'w', compression=ZIP_DEFLATED, compresslevel=6) as archive:
        for row in cursor:
            gsfx, os_book = bytes(row['gsfx_blob']), bytes(row['os_blob'])
            raw_bytes += len(gsfx) + len(os_book)
            if raw_bytes > max_raw_bytes:
                raise ValueError('Export exceeds 128 MB. Download a smaller batch.')
            if (hashlib.sha256(gsfx).hexdigest() != row['gsfx_sha256'] or
                    hashlib.sha256(os_book).hexdigest() != row['os_sha256']):
                raise ValueError(f'Hash mismatch in approved pair #{row["id"]}')
            prefix = f'pair_{int(row["id"]):06d}/'
            archive.writestr(prefix + 'original_lsr.gsfx', gsfx)
            archive.writestr(prefix + 'original_os.xlsx', os_book)
            metadata = {
                'id': int(row['id']),
                'gsfx_filename': row['gsfx_filename'],
                'os_filename': row['os_filename'],
                'sheet_title': row['sheet_title'],
                'gsfx_sha256': row['gsfx_sha256'],
                'os_sha256': row['os_sha256'],
                'gsfx_count': int(row['gsfx_count']),
                'os_count': int(row['os_count']),
                'matched_count': int(row['matched_count']),
                'checked_prices': int(row['checked_prices']),
                'created_at': row['created_at'],
                'checks': json.loads(row['issues_json']),
                'os_rows': json.loads(row['os_rows_json']),
                'review_status': 'approved_by_admin',
            }
            archive.writestr(prefix + 'verification.json',
                             json.dumps(metadata, ensure_ascii=False, indent=2))
            manifest.append({key: metadata[key] for key in (
                'id', 'gsfx_filename', 'os_filename', 'sheet_title',
                'gsfx_sha256', 'os_sha256', 'matched_count', 'os_count',
                'checked_prices', 'review_status',
            )})
        if not manifest:
            raise ValueError('There are no approved pairs with both original files.')
        archive.writestr('manifest.json', json.dumps({
            'format': 'estimate-ai-verified-gsfx-os-pairs-v1',
            'exported_at_utc': datetime.now(timezone.utc).isoformat(),
            'total_pairs': len(manifest),
            'last_id': manifest[-1]['id'],
            'pairs': manifest,
        }, ensure_ascii=False, indent=2))
    return VerifiedPairExport(output.getvalue(), len(manifest), manifest[-1]['id'])

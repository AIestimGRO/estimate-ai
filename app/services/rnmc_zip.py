"""Dry-run analysis for RNMC zip uploads.

This module deliberately does not import rows into catalog_items. It only scans
an uploaded zip archive, resolves the region from folder names, and checks file
names against imported_files so the UI can show what would happen before a real
import step is enabled.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

from core.storage.catalog import filename_is_processed, normalize_import_filename

EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xls"})
STATUS_WILL_PROCESS = "will_process"
STATUS_SKIPPED_PROCESSED = "skipped_processed"
STATUS_DUPLICATE_NAME = "duplicate_name"


@dataclass(frozen=True)
class RnmcZipDryRunEntry:
    archive_path: str
    filename: str
    region_folder: str
    status: str
    reason: str


@dataclass(frozen=True)
class RnmcZipDryRunResult:
    entries: list[RnmcZipDryRunEntry]
    ignored_files: int

    @property
    def total_excel_files(self) -> int:
        return len(self.entries)

    @property
    def will_process_count(self) -> int:
        return _count_status(self.entries, STATUS_WILL_PROCESS)

    @property
    def skipped_processed_count(self) -> int:
        return _count_status(self.entries, STATUS_SKIPPED_PROCESSED)

    @property
    def duplicate_name_count(self) -> int:
        return _count_status(self.entries, STATUS_DUPLICATE_NAME)


def analyze_rnmc_zip_dry_run(
    connection: sqlite3.Connection,
    zip_path: str,
    *,
    region_override: str = "",
) -> RnmcZipDryRunResult:
    """Return a dry-run plan for an RNMC zip archive.

    File identity is based on the base file name only, matching the accepted
    business rule: duplicated names across folders are rule violations, not new
    versions.
    """
    manual_region = _text(region_override)
    entries: list[RnmcZipDryRunEntry] = []
    ignored_files = 0
    seen_keys: set[str] = set()

    try:
        with ZipFile(zip_path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                path = _normalize_archive_path(info.filename)
                if not _is_supported_excel_path(path):
                    ignored_files += 1
                    continue
                filename = PurePosixPath(path).name
                key = normalize_import_filename(filename)
                region = manual_region or _region_from_archive_path(path)

                if key in seen_keys:
                    status = STATUS_DUPLICATE_NAME
                    reason = "Duplicate filename inside uploaded zip"
                elif filename_is_processed(connection, filename):
                    status = STATUS_SKIPPED_PROCESSED
                    reason = "Filename already exists in imported_files"
                else:
                    status = STATUS_WILL_PROCESS
                    reason = "New filename"
                seen_keys.add(key)

                entries.append(
                    RnmcZipDryRunEntry(
                        archive_path=path,
                        filename=filename,
                        region_folder=region,
                        status=status,
                        reason=reason,
                    )
                )
    except BadZipFile as exc:
        raise ValueError("uploaded file is not a valid zip archive") from exc

    return RnmcZipDryRunResult(entries=entries, ignored_files=ignored_files)


def _normalize_archive_path(value: str) -> str:
    path = str(value).replace("\\", "/").strip("/")
    return "/".join(part for part in path.split("/") if part not in {"", "."})


def _is_supported_excel_path(path: str) -> bool:
    pure = PurePosixPath(path)
    name = pure.name
    if name.startswith("~$"):
        return False
    return pure.suffix.casefold() in EXCEL_SUFFIXES


def _region_from_archive_path(path: str) -> str:
    parent = PurePosixPath(path).parent
    if str(parent) in {"", "."}:
        return ""
    return parent.name.strip()


def _count_status(entries: list[RnmcZipDryRunEntry], status: str) -> int:
    return sum(1 for entry in entries if entry.status == status)


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()

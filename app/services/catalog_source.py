"""Resolve catalog rows for a matching run (database or Excel upload)."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from core.catalog import CatalogRow
from core.excel_io import Settings, read_catalog_rows
from core.storage import connect, count_catalog_rows, init_database, list_catalog_rows

DEFAULT_SOURCE_NAME = "main"


class CatalogNotAvailableError(Exception):
    """Neither a catalog file nor a populated database source is available."""


@dataclass(frozen=True)
class CatalogLoadResult:
    rows: list[CatalogRow]
    source_label: str
    row_count: int


def database_has_catalog(
    *,
    source_name: str = DEFAULT_SOURCE_NAME,
    database_path: str | Path | None = None,
) -> bool:
    if not _database_file_exists(database_path):
        return False

    connection = connect(database_path)
    try:
        init_database(connection)
        return count_catalog_rows(connection, source_name=source_name) > 0
    except OSError:
        return False
    finally:
        connection.close()


def catalog_status_label(
    *,
    source_name: str = DEFAULT_SOURCE_NAME,
    database_path: str | Path | None = None,
) -> str:
    connection = connect(database_path)
    try:
        init_database(connection)
        count = count_catalog_rows(connection, source_name=source_name)
    except OSError:
        return "\u0431\u0430\u0437\u0430 \u043d\u0435 \u0441\u043e\u0437\u0434\u0430\u043d\u0430 (\u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 init-db \u0438 import-catalog)"
    finally:
        connection.close()

    if count <= 0:
        return f"\u043f\u0443\u0441\u0442\u043e ({source_name}) \u2014 \u0437\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u0435 Excel \u0438\u043b\u0438 import-catalog"
    return f"{count} \u0441\u0442\u0440\u043e\u043a ({source_name})"


def load_catalog_for_run(
    catalog_path: str | Path | None = None,
    *,
    source_name: str = DEFAULT_SOURCE_NAME,
    database_path: str | Path | None = None,
    settings: Settings | None = None,
) -> CatalogLoadResult:
    """Load catalog rows from an uploaded file or from the default DB source."""
    if catalog_path is not None:
        path = Path(catalog_path)
        rows = read_catalog_rows(path, settings)
        return CatalogLoadResult(
            rows=rows,
            source_label=f"file:{path.name}",
            row_count=len(rows),
        )

    connection = connect(database_path)
    try:
        init_database(connection)
        row_count = count_catalog_rows(connection, source_name=source_name)
        if row_count <= 0:
            raise CatalogNotAvailableError(
                f"catalog source {source_name!r} is empty; upload Excel or run import-catalog"
            )
        rows = list_catalog_rows(connection, source_name=source_name)
    finally:
        connection.close()

    return CatalogLoadResult(
        rows=rows,
        source_label=f"database:{source_name}",
        row_count=row_count,
    )


def _database_file_exists(database_path: str | Path | None) -> bool:
    from core.storage.connection import default_database_path

    path = default_database_path() if database_path is None else Path(database_path)
    return path.is_file()

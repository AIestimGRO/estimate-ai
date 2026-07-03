"""SQLite persistence for catalog, import log, and admin config."""

from core.storage.connection import connect, default_database_path, init_database
from core.storage.catalog import (
    import_catalog_from_excel,
    list_catalog_rows,
    list_catalog_sources,
    count_catalog_rows,
)
from core.storage.rules import (
    import_rules_from_workbook,
    list_name_exclusion_rules,
    list_task_color_entries,
    replace_name_exclusion_rules,
    replace_task_color_entries,
)

__all__ = [
    "connect",
    "default_database_path",
    "init_database",
    "import_catalog_from_excel",
    "list_catalog_rows",
    "list_catalog_sources",
    "count_catalog_rows",
    "import_rules_from_workbook",
    "list_name_exclusion_rules",
    "list_task_color_entries",
    "replace_name_exclusion_rules",
    "replace_task_color_entries",
]

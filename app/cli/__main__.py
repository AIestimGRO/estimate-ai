"""CLI entry point: database init and one-time imports.

Usage:
    python -m app.cli init-db
    python -m app.cli import-catalog path/to/catalog.xlsx [--source main]
    python -m app.cli import-rules [--xlsm path/to/autopodbor.xlsm]
    python -m app.cli export-catalog path/to/output.xlsx [--region ... --source ...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from core.storage import (
    connect,
    default_database_path,
    import_catalog_from_excel,
    import_rules_from_workbook,
    init_database,
    list_catalog_rows,
    list_catalog_sources,
    list_name_exclusion_rules,
    write_catalog_export_xlsx,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Estimate AI database tools")
    parser.add_argument(
        "--db",
        type=Path,
        default=None,
        help="database file path (default: data/estimate_ai.db or ESTIMATE_AI_DB_PATH)",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("init-db", help="create database schema")

    import_catalog = subparsers.add_parser(
        "import-catalog",
        help="import one RNMC catalog Excel file into the database",
    )
    import_catalog.add_argument("catalog", type=Path, help="path to catalog .xlsx")
    import_catalog.add_argument(
        "--source",
        default="main",
        help="catalog source name (default: main)",
    )
    import_catalog.add_argument(
        "--append",
        action="store_true",
        help="append rows instead of replacing the source catalog",
    )

    import_rules = subparsers.add_parser(
        "import-rules",
        help="import Name_Exclusions and task colors from the macro workbook",
    )
    import_rules.add_argument(
        "--xlsm",
        type=Path,
        default=None,
        help="macro workbook path (default: resolve from macro.json / data/real)",
    )

    status = subparsers.add_parser("status", help="show database summary")

    export_catalog = subparsers.add_parser(
        "export-catalog",
        help="export catalog_items to an .xlsx (fixed column layout)",
    )
    export_catalog.add_argument("output", type=Path, help="path to write the .xlsx to")
    export_catalog.add_argument("--source", default="", help="filter: catalog source name (substring)")
    export_catalog.add_argument("--region", default="", help="filter: region (substring)")
    export_catalog.add_argument("--task-id", default="", help="filter: task/задача number (substring)")
    export_catalog.add_argument("--code", default="", help="filter: ГЭСН/ФЕР/ТЕР code (substring)")
    export_catalog.add_argument("--filename", default="", help="filter: source file name (substring)")
    export_catalog.add_argument("--q", default="", help="filter: free-text search (same as the admin grid)")

    args = parser.parse_args(argv)
    db_path = default_database_path() if args.db is None else args.db.resolve()

    if args.command == "init-db":
        return _cmd_init_db(db_path)
    if args.command == "import-catalog":
        return _cmd_import_catalog(
            db_path,
            args.catalog,
            source_name=args.source,
            replace=not args.append,
        )
    if args.command == "import-rules":
        return _cmd_import_rules(db_path, args.xlsm)
    if args.command == "status":
        return _cmd_status(db_path)
    if args.command == "export-catalog":
        return _cmd_export_catalog(
            db_path,
            args.output,
            filters={
                "source": args.source,
                "region": args.region,
                "task_id": args.task_id,
                "code": args.code,
                "filename": args.filename,
                "q": args.q,
            },
        )

    parser.error(f"unknown command: {args.command}")
    return 2


def _cmd_init_db(db_path: Path) -> int:
    connection = connect(db_path)
    try:
        init_database(connection)
    finally:
        connection.close()
    print(f"Database ready: {db_path}")
    return 0


def _cmd_import_catalog(
    db_path: Path,
    catalog_path: Path,
    *,
    source_name: str,
    replace: bool,
) -> int:
    if not catalog_path.is_file():
        print(f"Catalog file not found: {catalog_path}", file=sys.stderr)
        return 1

    connection = connect(db_path)
    try:
        init_database(connection)
        result = import_catalog_from_excel(
            connection,
            catalog_path,
            source_name=source_name,
            replace=replace,
        )
    finally:
        connection.close()

    print(f"Database: {db_path}")
    print(f"Source:   {result.source_name} (id={result.source_id})")
    print(f"File:     {result.source_filename}")
    print(f"Imported: {result.rows_imported} rows")
    print(f"Skipped:  {result.rows_skipped} rows")
    return 0


def _cmd_import_rules(db_path: Path, xlsm_path: Path | None) -> int:
    connection = connect(db_path)
    try:
        init_database(connection)
        rule_count, color_count = import_rules_from_workbook(connection, xlsm_path)
    except FileNotFoundError as error:
        print(str(error), file=sys.stderr)
        return 1
    finally:
        connection.close()

    print(f"Database: {db_path}")
    print(f"Rules:    {rule_count}")
    print(f"Colors:   {color_count}")
    return 0


def _cmd_status(db_path: Path) -> int:
    if not db_path.is_file():
        print(f"Database not found: {db_path}", file=sys.stderr)
        return 1

    connection = connect(db_path)
    try:
        init_database(connection)
        sources = list_catalog_sources(connection)
        rules = list_name_exclusion_rules(connection)
        enabled_rules = sum(1 for rule in rules if rule.enabled)
        total_rows = 0
        for source in sources:
            rows = list_catalog_rows(connection, source_name=source.name)
            total_rows += len(rows)
            print(f"Source {source.name!r}: {len(rows)} catalog rows ({source.kind})")
        print(f"Name exclusion rules: {enabled_rules} enabled / {len(rules)} total")
        print(f"Total catalog rows:   {total_rows}")
    finally:
        connection.close()

    print(f"Database: {db_path}")
    return 0


def _cmd_export_catalog(
    db_path: Path,
    output_path: Path,
    *,
    filters: dict[str, str],
) -> int:
    if not db_path.is_file():
        print(f"Database not found: {db_path}", file=sys.stderr)
        return 1

    connection = connect(db_path)
    try:
        init_database(connection)
        row_count = write_catalog_export_xlsx(connection, output_path, filters=filters)
    finally:
        connection.close()

    print(f"Database: {db_path}")
    print(f"Written:  {output_path} ({row_count} rows)")
    active_filters = {key: value for key, value in filters.items() if value}
    if active_filters:
        print(f"Filters:  {active_filters}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

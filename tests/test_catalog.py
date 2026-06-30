from datetime import date

from core.catalog import BuildCatalog, CatalogRow
from core.exclusions import NameExclusionRule
from core.normalize import AnalogSearchKey


METER = "\u043c"
DEMOLITION = "\u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436"
INSTALLATION = "\u043c\u043e\u043d\u0442\u0430\u0436"
EXCLUDED = "\u0438\u0441\u043a\u043b\u044e\u0447\u0438\u0442\u044c"


def row(
    *,
    task_id: object = "task-1",
    price: object = 100.0,
    code: object = "gesn01-01-001-01",
    unit: object = METER,
    work_name: object = INSTALLATION,
    region: object = "region-1",
    added_date: object = None,
) -> CatalogRow:
    return CatalogRow(
        task_id=task_id,
        price=price,
        code=code,
        unit=unit,
        work_name=work_name,
        region=region,
        added_date=added_date,
    )


def exclusion_rule() -> NameExclusionRule:
    return NameExclusionRule(
        enabled=True,
        scope="CATALOG",
        match_mode="CONTAINS",
        pattern=EXCLUDED,
        group="test",
        comment="test exclusion",
    )


def entries_for(catalog: dict, catalog_row: CatalogRow) -> list:
    return catalog[AnalogSearchKey(catalog_row.unit, catalog_row.code)][
        str(catalog_row.task_id)
    ]


def test_valid_row_gets_included() -> None:
    catalog_row = row(added_date=date(2024, 1, 1))

    catalog = BuildCatalog([catalog_row])
    entries = entries_for(catalog, catalog_row)

    assert len(entries) == 1
    assert entries[0].price == 100.0
    assert entries[0].region == "region-1"
    assert entries[0].original_row is catalog_row
    assert entries[0].task_id == "task-1"
    assert entries[0].norm_code == "GESN01-01-001-01"
    assert entries[0].norm_unit == METER
    assert entries[0].added_date_serial > 0


def test_missing_task_id_is_skipped() -> None:
    assert BuildCatalog([row(task_id="")]) == {}


def test_non_numeric_price_is_skipped() -> None:
    assert BuildCatalog([row(price="not numeric")]) == {}


def test_zero_price_is_skipped() -> None:
    assert BuildCatalog([row(price=0)]) == {}


def test_negative_price_is_skipped() -> None:
    assert BuildCatalog([row(price=-1)]) == {}


def test_empty_code_is_skipped() -> None:
    assert BuildCatalog([row(code=" \t\r\n")]) == {}


def test_empty_unit_is_skipped() -> None:
    assert BuildCatalog([row(unit="")]) == {}


def test_name_excluded_catalog_row_is_skipped() -> None:
    assert BuildCatalog(
        [row(work_name=f"{EXCLUDED} item")],
        [exclusion_rule()],
    ) == {}


def test_demolition_flag_is_computed_from_work_name() -> None:
    demolition_row = row(task_id="dem", work_name=DEMOLITION)
    regular_row = row(task_id="reg", work_name=INSTALLATION)

    catalog = BuildCatalog([demolition_row, regular_row])

    assert entries_for(catalog, demolition_row)[0].is_demolition
    assert not entries_for(catalog, regular_row)[0].is_demolition


def test_groups_same_key_under_different_task_ids() -> None:
    first = row(task_id="task-1", price=100)
    second = row(task_id="task-2", price=101)

    catalog = BuildCatalog([first, second])
    task_groups = catalog[AnalogSearchKey(METER, "gesn01-01-001-01")]

    assert set(task_groups) == {"task-1", "task-2"}
    assert task_groups["task-1"][0].price == 100
    assert task_groups["task-2"][0].price == 101


def test_dedup_collapses_close_prices_within_same_task_and_demolition_flag() -> None:
    first = row(price=100, work_name=INSTALLATION)
    second = row(price=104, work_name=INSTALLATION)

    catalog = BuildCatalog([first, second])

    assert len(entries_for(catalog, first)) == 1
    assert entries_for(catalog, first)[0].price == 100


def test_dedup_keeps_close_prices_when_demolition_flags_differ() -> None:
    first = row(price=100, work_name=INSTALLATION)
    second = row(price=104, work_name=DEMOLITION)

    catalog = BuildCatalog([first, second])
    entries = entries_for(catalog, first)

    assert len(entries) == 2
    assert [entry.is_demolition for entry in entries] == [False, True]


def test_missing_added_date_results_in_serial_zero() -> None:
    catalog_row = row(added_date=None)

    assert entries_for(BuildCatalog([catalog_row]), catalog_row)[0].added_date_serial == 0


def test_invalid_added_date_results_in_serial_zero() -> None:
    catalog_row = row(added_date="not a date")

    assert entries_for(BuildCatalog([catalog_row]), catalog_row)[0].added_date_serial == 0

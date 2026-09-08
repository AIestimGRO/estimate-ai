from core.storage.connection import connect, init_database


def test_init_database_can_skip_correction_resync(tmp_path, monkeypatch) -> None:
    database = tmp_path / "estimate_ai.db"
    connection = connect(database)
    calls: list[int] = []

    def fake_sync(active_connection):
        calls.append(1)
        return 0

    monkeypatch.setattr(
        "core.storage.corrections.synchronize_catalog_corrections",
        fake_sync,
    )
    try:
        init_database(connection, synchronize_corrections=False)
        assert calls == []

        init_database(connection)
        assert calls == [1]
    finally:
        connection.close()

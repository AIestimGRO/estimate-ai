from fastapi.testclient import TestClient

from app.web.app import create_app


def test_admin_index_shows_navigation(tmp_path) -> None:
    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin")

    assert response.status_code == 200
    assert "\u0410\u0434\u043c\u0438\u043d\u043a\u0430 Estimate AI" in response.text
    assert "/admin/sources" in response.text
    assert "/admin/imports" in response.text
    assert "/admin/risks" in response.text
    assert "/admin/task-colors" in response.text


def test_unknown_admin_section_returns_404(tmp_path) -> None:
    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/unknown")

    assert response.status_code == 404
    assert "\u0420\u0430\u0437\u0434\u0435\u043b \u0430\u0434\u043c\u0438\u043d\u043a\u0438 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d" in response.text


def test_index_links_to_admin(tmp_path) -> None:
    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/")

    assert response.status_code == 200
    assert 'href="/admin"' in response.text

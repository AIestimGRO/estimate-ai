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


def test_admin_placeholder_sections_open(tmp_path) -> None:
    paths = [
        "/admin/approvals",
        "/admin/task-colors",
        "/admin/name-exclusions",
        "/admin/gesn-exceptions",
        "/admin/settings",
    ]

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        for path in paths:
            response = client.get(path)
            assert response.status_code == 200
            assert "\u0420\u0430\u0437\u0434\u0435\u043b \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u044f \u0430\u0432\u0442\u043e\u043f\u043e\u0434\u0431\u043e\u0440\u0449\u0438\u043a\u0430" in response.text
            assert "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441" in response.text


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

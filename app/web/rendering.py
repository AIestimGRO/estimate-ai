"""HTML rendering for the web UI (pure, framework-agnostic).

Kept separate from the FastAPI wiring so the presentation (templates, styles,
result formatting) has no web-framework dependency and stays easy to test.
Russian UI wording lives in the HTML templates and small label maps here, not
in the pipeline logic (AGENTS.md rule 3 is about business logic; these are
presentation strings for the local tool).
"""

import html
from datetime import datetime, timezone
from pathlib import Path
from string import Template
from urllib.parse import quote

from app.services.write_result import RunAndWriteResult
from app.services.catalog_source import catalog_status_label, database_has_catalog
from core.macro_workbook import load_default_macro_settings
from core.risk import DEFAULT_PRICE_SPREAD_LIMIT
from core.storage.catalog import CatalogSource, ImportedFileRecord

_WRITER_MODULE = Path(__file__).resolve().parents[2] / "core" / "excel_writer.py"

TEMPLATES_DIR = Path(__file__).parent / "templates"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

READ_METHOD_LABELS = {
    "template": "\u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443",
    "detected": "\u043f\u043e \u0440\u0430\u0441\u043f\u043e\u0437\u043d\u0430\u043d\u043d\u044b\u043c \u0441\u0442\u043e\u043b\u0431\u0446\u0430\u043c",
}

ADMIN_SECTIONS = [
    {
        "slug": "sources",
        "title": "\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u0431\u0430\u0437\u044b",
        "description": "\u0420\u0430\u0437\u0434\u0435\u043b\u0435\u043d\u0438\u0435 \u0434\u0430\u043d\u043d\u044b\u0445 \u043f\u043e \u043f\u0440\u043e\u0438\u0441\u0445\u043e\u0436\u0434\u0435\u043d\u0438\u044e: \u0420\u041d\u041c\u0426, \u0422\u041a\u041f, \u0440\u0443\u0447\u043d\u044b\u0435 \u0430\u043d\u0430\u043b\u043e\u0433\u0438, \u043f\u0440\u0430\u0439\u0441\u044b \u0438 \u0434\u0440\u0443\u0433\u0438\u0435 \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438.",
        "status": "\u0421\u043f\u0438\u0441\u043e\u043a \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u043e\u0432 \u0438\u0437 catalog_sources \u043f\u043e\u043a\u0430\u0437\u044b\u0432\u0430\u0435\u0442\u0441\u044f \u0432 read-only \u0440\u0435\u0436\u0438\u043c\u0435.",
    },
    {
        "slug": "imports",
        "title": "\u0418\u043c\u043f\u043e\u0440\u0442\u044b \u0444\u0430\u0439\u043b\u043e\u0432",
        "description": "\u0418\u0441\u0442\u043e\u0440\u0438\u044f \u0437\u0430\u0433\u0440\u0443\u0437\u043e\u043a, \u043f\u0440\u0438\u043d\u044f\u0442\u044b\u0435 \u0438 \u043e\u0442\u043a\u043b\u043e\u043d\u0435\u043d\u043d\u044b\u0435 \u0441\u0442\u0440\u043e\u043a\u0438, \u0440\u0435\u0433\u0438\u043e\u043d, \u0444\u0430\u0439\u043b \u0438 \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a.",
        "status": "\u0418\u0441\u0442\u043e\u0440\u0438\u044f \u0438\u043c\u043f\u043e\u0440\u0442\u043e\u0432 \u0438\u0437 imported_files \u043f\u043e\u043a\u0430\u0437\u044b\u0432\u0430\u0435\u0442\u0441\u044f \u0432 read-only \u0440\u0435\u0436\u0438\u043c\u0435.",
    },
    {
        "slug": "risks",
        "title": "\u0420\u0438\u0441\u043a-\u043b\u043e\u0433",
        "description": "\u0421\u0442\u0440\u043e\u043a\u0438, \u0433\u0434\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u0441\u0438\u043b\u044c\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u0440\u043e\u0441 \u0446\u0435\u043d \u0438 \u0442\u0440\u0435\u0431\u0443\u0435\u0442\u0441\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0447\u0435\u043b\u043e\u0432\u0435\u043a\u043e\u043c.",
        "status": "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441. \u041f\u043e\u0437\u0436\u0435 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u043c price_risk_log.",
    },
    {
        "slug": "approvals",
        "title": "\u041e\u0434\u043e\u0431\u0440\u0435\u043d\u0438\u0435 \u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d\u043e\u0432",
        "description": "\u041f\u043e\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u0435 \u0434\u043e\u043f\u0443\u0441\u0442\u0438\u043c\u044b\u0445 \u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d\u043e\u0432 min/max \u0434\u043b\u044f \u0441\u043f\u043e\u0440\u043d\u044b\u0445 \u0413\u042d\u0421\u041d \u0438 \u0435\u0434\u0438\u043d\u0438\u0446 \u0438\u0437\u043c\u0435\u0440\u0435\u043d\u0438\u044f.",
        "status": "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441. \u041f\u043e\u0437\u0436\u0435 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u043c approve workflow.",
    },
    {
        "slug": "task-colors",
        "title": "\u0421\u0438\u043d\u0438\u0435 \u0437\u0430\u0434\u0430\u0447\u0438",
        "description": "\u041d\u043e\u043c\u0435\u0440\u0430 \u0437\u0430\u0434\u0430\u0447, \u0430\u043d\u0430\u043b\u043e\u0433\u0438 \u0438\u0437 \u043a\u043e\u0442\u043e\u0440\u044b\u0445 \u043d\u0435 \u0431\u043b\u043e\u043a\u0438\u0440\u0443\u044e\u0442\u0441\u044f, \u0430 \u043f\u043e\u0434\u0441\u0432\u0435\u0447\u0438\u0432\u0430\u044e\u0442\u0441\u044f \u0441\u0438\u043d\u0438\u043c.",
        "status": "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441. \u041f\u043e\u0437\u0436\u0435 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u043c task_color_entries.",
    },
    {
        "slug": "name-exclusions",
        "title": "\u0418\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u044f \u043f\u043e \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u044f\u043c",
        "description": "\u041f\u0440\u0430\u0432\u0438\u043b\u0430 \u043f\u043e \u0442\u0435\u043a\u0441\u0442\u0443 \u0440\u0430\u0431\u043e\u0442 \u0438 \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0439, \u043e\u0442\u0434\u0435\u043b\u044c\u043d\u043e \u043e\u0442 \u043f\u043e\u0434\u0441\u0432\u0435\u0442\u043a\u0438 \u0437\u0430\u0434\u0430\u0447.",
        "status": "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441. \u041f\u043e\u0437\u0436\u0435 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u043c name_exclusion_rules.",
    },
    {
        "slug": "gesn-exceptions",
        "title": "GESN exceptions",
        "description": "\u041e\u0434\u043e\u0431\u0440\u0435\u043d\u043d\u044b\u0435 \u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d\u044b \u0446\u0435\u043d \u043f\u043e \u0441\u0432\u044f\u0437\u043a\u0435 \u043a\u043e\u0434 + \u0435\u0434\u0438\u043d\u0438\u0446\u0430 + \u043f\u0440\u0438\u0437\u043d\u0430\u043a \u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436\u0430.",
        "status": "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441. \u041f\u043e\u0437\u0436\u0435 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u043c gesn_exceptions.",
    },
    {
        "slug": "settings",
        "title": "\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438",
        "description": "\u0422\u0435\u0445\u043d\u0438\u0447\u0435\u0441\u043a\u0438\u0435 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b \u043f\u0440\u043e\u0435\u043a\u0442\u0430: \u0431\u0430\u0437\u0430, \u043a\u043e\u043d\u0444\u0438\u0433\u0438, \u0441\u0442\u0430\u0442\u0443\u0441 \u0441\u043b\u043e\u0432\u0430\u0440\u0435\u0439 \u0438 \u0432\u0435\u0440\u0441\u0438\u044f writer.",
        "status": "\u041f\u043e\u043a\u0430 \u043a\u0430\u0440\u043a\u0430\u0441. \u041f\u043e\u0437\u0436\u0435 \u0434\u043e\u0431\u0430\u0432\u0438\u043c read-only diagnostics.",
    },
]

ADMIN_SECTION_SLUGS = {section["slug"] for section in ADMIN_SECTIONS}

STYLES = """
* { box-sizing: border-box; }
body {
  margin: 0; min-height: 100vh; display: flex; align-items: center;
  justify-content: center; padding: 24px;
  font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  background: linear-gradient(135deg, #eef2ff, #f8fafc); color: #0f172a;
}
.card {
  width: 100%; max-width: 520px; background: #fff; border-radius: 16px;
  padding: 32px; box-shadow: 0 20px 50px rgba(15, 23, 42, 0.12);
}
.card.wide { max-width: 900px; }
h1 { margin: 0 0 4px; font-size: 24px; }
h2.section { margin: 0 0 10px; font-size: 16px; color: #334155; }
.sub { margin: 0 0 24px; color: #64748b; }
label { display: block; margin-bottom: 16px; font-weight: 600; font-size: 14px; }
input[type=file], input[type=text] {
  display: block; width: 100%; margin-top: 6px; padding: 10px 12px;
  border: 1px solid #cbd5e1; border-radius: 10px; font-size: 14px; font-weight: 400;
}
button {
  width: 100%; margin-top: 8px; padding: 12px 16px; border: 0; border-radius: 10px;
  background: #4f46e5; color: #fff; font-size: 15px; font-weight: 600; cursor: pointer;
}
button:hover { background: #4338ca; }
.stats { margin: 0 0 24px; }
.stats div { display: flex; justify-content: space-between; padding: 8px 0;
  border-bottom: 1px solid #f1f5f9; font-size: 14px; }
.stats dt { color: #64748b; }
.stats dd { margin: 0; font-weight: 600; }
.download {
  display: block; text-align: center; padding: 12px 16px; border-radius: 10px;
  background: #16a34a; color: #fff; text-decoration: none; font-weight: 600;
}
.download:hover { background: #15803d; }
.choices { display: flex; flex-direction: column; gap: 10px; margin-bottom: 20px; }
.choice {
  display: block; padding: 12px 16px; border-radius: 10px; text-decoration: none;
  background: #eef2ff; color: #3730a3; font-weight: 600; text-align: center;
}
.choice:hover { background: #e0e7ff; }
.back { display: inline-block; margin-top: 16px; color: #4f46e5; text-decoration: none; }
.err-title { color: #b91c1c; }
.err { color: #b91c1c; font-weight: 600; }
.report {
  background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px;
  white-space: pre-wrap; word-break: break-word; font-size: 13px; color: #334155;
}
.notice { margin-top: 16px; padding: 12px; border-radius: 10px;
  background: #fef2f2; color: #b91c1c; font-size: 14px; }
.preview-wrap { max-height: 380px; overflow: auto; border: 1px solid #e2e8f0;
  border-radius: 10px; margin-bottom: 20px; }
table.preview { border-collapse: collapse; width: 100%; font-size: 13px; }
table.preview th, table.preview td {
  padding: 7px 10px; border-bottom: 1px solid #f1f5f9; text-align: left;
  white-space: nowrap; }
table.preview th { position: sticky; top: 0; background: #f8fafc; color: #475569;
  font-weight: 600; }
table.preview td.risk { color: #b91c1c; text-align: center; font-weight: 700; }
.muted { color: #94a3b8; font-size: 12px; padding: 8px 10px; margin: 0; }
.build { margin-top: 20px; font-size: 12px; color: #94a3b8; text-align: center; }
.admin-nav {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin: 0 0 24px;
}
.admin-nav-link {
  display: inline-block;
  padding: 8px 10px;
  border-radius: 999px;
  background: #eef2ff;
  color: #3730a3;
  text-decoration: none;
  font-size: 13px;
  font-weight: 600;
}
.admin-nav-link:hover { background: #e0e7ff; }
.admin-nav-link.active { background: #4f46e5; color: #fff; }
.admin-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 12px;
}
.admin-card {
  display: block;
  padding: 16px;
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  background: #f8fafc;
  color: #0f172a;
  text-decoration: none;
}
.admin-card:hover { border-color: #c7d2fe; background: #eef2ff; }
.admin-card strong { display: block; margin-bottom: 8px; color: #1e293b; }
.admin-card span {
  display: block;
  color: #64748b;
  font-size: 13px;
  line-height: 1.35;
}
.admin-panel {
  padding: 16px;
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  background: #f8fafc;
}
.notice-soft {
  margin-top: 14px;
  padding: 12px;
  border-radius: 10px;
  background: #eef2ff;
  color: #3730a3;
  font-size: 14px;
}

"""


def writer_build_stamp() -> str:
    """Human-readable stamp so the UI shows whether the server picked up new code."""
    modified = datetime.fromtimestamp(_WRITER_MODULE.stat().st_mtime, tz=timezone.utc)
    return modified.strftime("%Y-%m-%d %H:%M UTC")


def macro_status_label() -> str:
    """Show whether Name_Exclusions from the autopodbor workbook is loaded."""
    settings = load_default_macro_settings()
    if settings.workbook_path is None:
        return "\u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d (\u043f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 data/config/macro.json)"
    count = len(settings.name_exclusion_rules)
    return f"{count} \u043f\u0440\u0430\u0432\u0438\u043b ({html.escape(settings.workbook_path.name)})"


def _macro_exclusion_label(outcome: RunAndWriteResult) -> str:
    if outcome.name_exclusion_rule_count <= 0 and outcome.macro_workbook is None:
        return ""
    if outcome.macro_workbook is None:
        return str(outcome.name_exclusion_rule_count)
    return f"{outcome.name_exclusion_rule_count} ({outcome.macro_workbook.name})"


def _catalog_source_label(outcome: RunAndWriteResult) -> str:
    if not outcome.catalog_source:
        return ""
    if outcome.catalog_source.startswith("database:"):
        source_name = outcome.catalog_source.split(":", 1)[1]
        return f"{outcome.catalog_row_count} \u0441\u0442\u0440\u043e\u043a \u0438\u0437 \u0411\u0414 ({source_name})"
    if outcome.catalog_source.startswith("file:"):
        file_name = outcome.catalog_source.split(":", 1)[1]
        return f"{outcome.catalog_row_count} \u0441\u0442\u0440\u043e\u043a \u0438\u0437 \u0444\u0430\u0439\u043b\u0430 ({file_name})"
    return f"{outcome.catalog_row_count} ({outcome.catalog_source})"


def render(template_name: str, **context: str) -> str:
    text = (TEMPLATES_DIR / template_name).read_text(encoding="utf-8")
    context.setdefault("styles", STYLES)
    context.setdefault("message", "")
    context.setdefault("detail", "")
    context.setdefault("preview", "")
    context.setdefault("build_stamp", writer_build_stamp())
    context.setdefault("macro_status", macro_status_label())
    context.setdefault("catalog_status", html.escape(catalog_status_label()))
    context.setdefault("catalog_required", "" if database_has_catalog() else "required")
    return Template(text).safe_substitute(context)


def render_index(notice: str | None = None) -> str:
    message = "" if not notice else f'<p class="notice">{html.escape(notice)}</p>'
    return render("index.html", message=message)


def render_error(title: str, message: str, detail: str = "") -> str:
    return render(
        "error.html",
        title=html.escape(title),
        message=html.escape(message),
        detail=html.escape(detail),
    )


def render_choice(token: str, candidates: list[str]) -> str:
    buttons = "\n".join(
        f'<a class="choice" href="/run?token={quote(token)}&sheet={quote(title)}">'
        f"{html.escape(title)}</a>"
        for title in candidates
    )
    return render("choose_sheet.html", sheet_buttons=buttons)


PREVIEW_ROW_LIMIT = 200


def _fmt_number(value: object) -> str:
    if isinstance(value, bool):
        return html.escape(str(value))
    if isinstance(value, (int, float)):
        return f"{value:g}"
    return html.escape("" if value is None else str(value))


def _render_preview(rows: list) -> str:
    header = (
        "<tr><th>#</th><th>\u041a\u043e\u0434</th><th>\u0415\u0434.</th>"
        "<th>\u0411\u0430\u0437. \u0446\u0435\u043d\u0430</th>"
        "<th>\u0420\u0435\u043a\u043e\u043c. \u0446\u0435\u043d\u0430</th>"
        "<th>\u0410\u043d\u0430\u043b\u043e\u0433\u043e\u0432</th>"
        "<th>\u0420\u0438\u0441\u043a</th></tr>"
    )
    body = []
    for row in rows[:PREVIEW_ROW_LIMIT]:
        estimate_row = row.estimate_row
        risk_mark = "\u26a0" if row.risk_result.is_flagged else ""
        body.append(
            "<tr>"
            f"<td>{row.row_index}</td>"
            f"<td>{html.escape('' if estimate_row.code is None else str(estimate_row.code))}</td>"
            f"<td>{html.escape('' if estimate_row.unit is None else str(estimate_row.unit))}</td>"
            f"<td>{_fmt_number(estimate_row.base_price)}</td>"
            f"<td>{_fmt_number(row.recommended_price)}</td>"
            f"<td>{len(row.analogs)}</td>"
            f'<td class="risk">{risk_mark}</td>'
            "</tr>"
        )

    table = f'<table class="preview">{header}{"".join(body)}</table>'
    if len(rows) > PREVIEW_ROW_LIMIT:
        table += (
            f'<p class="muted">\u043f\u043e\u043a\u0430\u0437\u0430\u043d\u044b '
            f"\u043f\u0435\u0440\u0432\u044b\u0435 {PREVIEW_ROW_LIMIT} \u0438\u0437 {len(rows)} "
            "\u0441\u0442\u0440\u043e\u043a</p>"
        )
    return table


def render_result(token: str, output_name: str, outcome: RunAndWriteResult) -> str:
    result = outcome.result
    method_label = READ_METHOD_LABELS.get(outcome.read_method, outcome.read_method)
    exclusion_label = _macro_exclusion_label(outcome)
    catalog_label = _catalog_source_label(outcome)
    rows = [
        ("\u041b\u0438\u0441\u0442", outcome.sheet_title),
        ("\u0421\u043f\u043e\u0441\u043e\u0431 \u0447\u0442\u0435\u043d\u0438\u044f", method_label),
        ("\u0421\u0442\u0440\u043e\u043a \u0441\u043c\u0435\u0442\u044b \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u0430\u043d\u043e", str(len(result.rows))),
        ("\u0421 \u043f\u043e\u0434\u043e\u0431\u0440\u0430\u043d\u043d\u044b\u043c\u0438 \u0430\u043d\u0430\u043b\u043e\u0433\u0430\u043c\u0438", str(result.matched_row_count)),
        ("\u041e\u0442\u043c\u0435\u0447\u0435\u043d\u043e \u043d\u0430 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0443 (\u0440\u0438\u0441\u043a)", str(result.flagged_row_count)),
        ("\u0420\u0435\u0433\u0438\u043e\u043d\u0430\u043b\u044c\u043d\u044b\u0439 \u043a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442", f"{outcome.regional_coefficient:g}"),
        (
            "\u041f\u043e\u0440\u043e\u0433 ratio (max/min)",
            f"{DEFAULT_PRICE_SPREAD_LIMIT:g} (\u043a\u0440\u0430\u0441\u043d\u044b\u0439 \u043f\u0440\u0438 ratio \u2265 \u044d\u0442\u043e\u0433\u043e)",
        ),
        ("\u0417\u0430\u043f\u0438\u0441\u0430\u043d\u043e \u0441\u0442\u0440\u043e\u043a \u0432 \u0444\u0430\u0439\u043b", str(outcome.write_report.written_rows)),
        (
            "\u041a\u043e\u043b\u043e\u043d\u043a\u0438 \u0432\u044b\u0432\u043e\u0434\u0430",
            f"/\u041a\u0420={outcome.write_report.analog_start_column - 2}, "
            f"\u0440\u0430\u0437\u0434\u0435\u043b={outcome.write_report.analog_start_column - 1}, "
            f"\u0430\u043d\u0430\u043b\u043e\u0433\u0438={outcome.write_report.analog_start_column}",
        ),
        (
            "\u0412\u0435\u0440\u0441\u0438\u044f writer",
            writer_build_stamp(),
        ),
    ]
    if catalog_label:
        rows.insert(1, ("\u0411\u0430\u0437\u0430 \u0430\u043d\u0430\u043b\u043e\u0433\u043e\u0432", catalog_label))
    if exclusion_label:
        rows.insert(3, (
            "Name_Exclusions",
            exclusion_label,
        ))
    stats = "\n".join(
        f"<div><dt>{html.escape(label)}</dt><dd>{html.escape(value)}</dd></div>"
        for label, value in rows
    )
    return render(
        "result.html",
        stats=stats,
        preview=_render_preview(result.rows),
        download_url=f"/download?token={quote(token)}",
        filename=html.escape(output_name),
    )


def render_admin_index() -> str:
    cards = "\n".join(_render_admin_card(section) for section in ADMIN_SECTIONS)
    return render(
        "admin.html",
        title="\u0410\u0434\u043c\u0438\u043d\u043a\u0430 Estimate AI",
        subtitle="\u0423\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u0438\u0435 \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0430\u043c\u0438, \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430\u043c\u0438 \u0438 \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0430\u043c\u0438 \u0430\u0432\u0442\u043e\u043f\u043e\u0434\u0431\u043e\u0440\u0449\u0438\u043a\u0430.",
        admin_nav=_render_admin_nav(active_slug=""),
        content=f'<div class="admin-grid">{cards}</div>',
    )



def render_admin_sources(sources: list[CatalogSource]) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u0431\u0430\u0437\u044b</h2>'
        '<p>\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u043d\u0443\u0436\u043d\u044b, \u0447\u0442\u043e\u0431\u044b \u043d\u0435 \u0441\u043c\u0435\u0448\u0438\u0432\u0430\u0442\u044c \u0434\u0430\u043d\u043d\u044b\u0435 \u0420\u041d\u041c\u0426, \u0422\u041a\u041f, \u0440\u0443\u0447\u043d\u044b\u0435 \u0430\u043d\u0430\u043b\u043e\u0433\u0438, \u043f\u0440\u0430\u0439\u0441\u044b \u0438 \u0434\u0440\u0443\u0433\u0438\u0435 \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438.</p>'
        f'{_render_catalog_source_table(sources)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u0431\u0430\u0437\u044b",
        subtitle="\u0420\u0430\u0437\u0434\u0435\u043b \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u044f \u0430\u0432\u0442\u043e\u043f\u043e\u0434\u0431\u043e\u0440\u0449\u0438\u043a\u0430.",
        admin_nav=_render_admin_nav(active_slug="sources"),
        content=content,
    )


def _render_catalog_source_table(sources: list[CatalogSource]) -> str:
    if not sources:
        return '<p class="muted">\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u043f\u043e\u043a\u0430 \u043d\u0435 \u0434\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u044b.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>ID</th>'
        '<th>\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435</th>'
        '<th>\u0422\u0438\u043f</th>'
        '<th>\u0421\u0442\u0440\u043e\u043a</th>'
        '<th>\u0421\u043e\u0437\u0434\u0430\u043d</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for source in sources:
        rows.append(
            '<tr>'
            f'<td>{source.id}</td>'
            f'<td>{html.escape(source.name)}</td>'
            f'<td>{html.escape(source.kind)}</td>'
            f'<td>{source.item_count}</td>'
            f'<td>{html.escape(source.created_at)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'

def render_admin_imports(imports: list[ImportedFileRecord]) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">\u0418\u043c\u043f\u043e\u0440\u0442\u044b \u0444\u0430\u0439\u043b\u043e\u0432</h2>'
        '<p>\u0416\u0443\u0440\u043d\u0430\u043b \u043f\u043e\u043a\u0430\u0437\u044b\u0432\u0430\u0435\u0442, \u043a\u0430\u043a\u0438\u0435 \u0444\u0430\u0439\u043b\u044b \u0431\u044b\u043b\u0438 \u0437\u0430\u0433\u0440\u0443\u0436\u0435\u043d\u044b, \u0438\u0437 \u043a\u0430\u043a\u043e\u0433\u043e \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0430 \u0438 \u0441\u043a\u043e\u043b\u044c\u043a\u043e \u0441\u0442\u0440\u043e\u043a \u043f\u0440\u0438\u043d\u044f\u0442\u043e.</p>'
        f'{_render_imported_file_table(imports)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="\u0418\u043c\u043f\u043e\u0440\u0442\u044b \u0444\u0430\u0439\u043b\u043e\u0432",
        subtitle="\u0420\u0430\u0437\u0434\u0435\u043b \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u044f \u0430\u0432\u0442\u043e\u043f\u043e\u0434\u0431\u043e\u0440\u0449\u0438\u043a\u0430.",
        admin_nav=_render_admin_nav(active_slug="imports"),
        content=content,
    )


def _render_imported_file_table(imports: list[ImportedFileRecord]) -> str:
    if not imports:
        return '<p class="muted">\u0418\u043c\u043f\u043e\u0440\u0442\u044b \u043f\u043e\u043a\u0430 \u043d\u0435 \u0437\u0430\u043f\u0438\u0441\u0430\u043d\u044b.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>ID</th>'
        '<th>\u0424\u0430\u0439\u043b</th>'
        '<th>\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a</th>'
        '<th>\u0422\u0438\u043f</th>'
        '<th>\u0420\u0435\u0433\u0438\u043e\u043d</th>'
        '<th>\u0417\u0430\u0434\u0430\u0447\u0430</th>'
        '<th>\u0421\u0442\u0430\u0442\u0443\u0441</th>'
        '<th>OK</th>'
        '<th>Rejected</th>'
        '<th>\u0418\u043c\u043f\u043e\u0440\u0442</th>'
        '<th>\u041e\u0448\u0438\u0431\u043a\u0430</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for item in imports:
        rows.append(
            '<tr>'
            f'<td>{item.id}</td>'
            f'<td>{html.escape(item.filename)}</td>'
            f'<td>{html.escape(item.source_name)}</td>'
            f'<td>{html.escape(item.source_kind)}</td>'
            f'<td>{html.escape(item.region_folder)}</td>'
            f'<td>{html.escape(item.task_number)}</td>'
            f'<td>{html.escape(item.status)}</td>'
            f'<td>{item.rows_ok}</td>'
            f'<td>{item.rows_rejected}</td>'
            f'<td>{html.escape(item.imported_at)}</td>'
            f'<td>{html.escape(item.failure_reason)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'


def render_admin_section(section_slug: str) -> str:
    section = _get_admin_section(section_slug)
    content = (
        '<section class="admin-panel">'
        f'<h2 class="section">{html.escape(section["title"])}</h2>'
        f'<p>{html.escape(section["description"])}</p>'
        f'<p class="notice-soft">{html.escape(section["status"])}</p>'
        "</section>"
    )
    return render(
        "admin.html",
        title=html.escape(section["title"]),
        subtitle="\u0420\u0430\u0437\u0434\u0435\u043b \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u044f \u0430\u0432\u0442\u043e\u043f\u043e\u0434\u0431\u043e\u0440\u0449\u0438\u043a\u0430.",
        admin_nav=_render_admin_nav(active_slug=section_slug),
        content=content,
    )


def _get_admin_section(section_slug: str) -> dict[str, str]:
    for section in ADMIN_SECTIONS:
        if section["slug"] == section_slug:
            return section
    raise KeyError(section_slug)


def _render_admin_card(section: dict[str, str]) -> str:
    href = f'/admin/{quote(section["slug"])}'
    return (
        f'<a class="admin-card" href="{href}">'
        f'<strong>{html.escape(section["title"])}</strong>'
        f'<span>{html.escape(section["description"])}</span>'
        "</a>"
    )


def _render_admin_nav(active_slug: str) -> str:
    links = ['<a class="admin-nav-link" href="/admin">\u0413\u043b\u0430\u0432\u043d\u0430\u044f</a>']
    for section in ADMIN_SECTIONS:
        classes = "admin-nav-link"
        if section["slug"] == active_slug:
            classes += " active"
        links.append(
            f'<a class="{classes}" href="/admin/{quote(section["slug"])}">'
            f'{html.escape(section["title"])}'
            "</a>"
        )
    return "\n".join(links)

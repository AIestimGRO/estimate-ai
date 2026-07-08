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
from app.services.rnmc_zip import RnmcZipDryRunResult
from app.services.rnmc_excel import RnmcZipCatalogImportResult, RnmcZipRowPreviewResult
from core.macro_workbook import load_default_macro_settings
from core.risk import DEFAULT_PRICE_SPREAD_LIMIT, GesnException
from core.storage.catalog import CatalogItemRecord, CatalogSource, ImportedFileRecord, ImportRowLogRecord
from core.storage.risk_log import PriceRiskLogEntry
from core.exclusions import NameExclusionRule, TaskColorEntry

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
        "status": "Риск-лог из price_risk_log показывается в read-only режиме.",
    },
    {
        "slug": "approvals",
        "title": "\u041e\u0434\u043e\u0431\u0440\u0435\u043d\u0438\u0435 \u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d\u043e\u0432",
        "description": "\u041f\u043e\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u0435 \u0434\u043e\u043f\u0443\u0441\u0442\u0438\u043c\u044b\u0445 \u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d\u043e\u0432 min/max \u0434\u043b\u044f \u0441\u043f\u043e\u0440\u043d\u044b\u0445 \u0413\u042d\u0421\u041d \u0438 \u0435\u0434\u0438\u043d\u0438\u0446 \u0438\u0437\u043c\u0435\u0440\u0435\u043d\u0438\u044f.",
        "status": "Открытые риски из price_risk_log показываются в read-only режиме. Approve workflow подключим отдельным write-этапом.",
    },
    {
        "slug": "task-colors",
        "title": "\u0421\u0438\u043d\u0438\u0435 \u0437\u0430\u0434\u0430\u0447\u0438",
        "description": "\u041d\u043e\u043c\u0435\u0440\u0430 \u0437\u0430\u0434\u0430\u0447, \u0430\u043d\u0430\u043b\u043e\u0433\u0438 \u0438\u0437 \u043a\u043e\u0442\u043e\u0440\u044b\u0445 \u043d\u0435 \u0431\u043b\u043e\u043a\u0438\u0440\u0443\u044e\u0442\u0441\u044f, \u0430 \u043f\u043e\u0434\u0441\u0432\u0435\u0447\u0438\u0432\u0430\u044e\u0442\u0441\u044f \u0441\u0438\u043d\u0438\u043c.",
        "status": "Список задач из task_color_entries показывается в read-only режиме.",
    },
    {
        "slug": "name-exclusions",
        "title": "\u0418\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u044f \u043f\u043e \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u044f\u043c",
        "description": "\u041f\u0440\u0430\u0432\u0438\u043b\u0430 \u043f\u043e \u0442\u0435\u043a\u0441\u0442\u0443 \u0440\u0430\u0431\u043e\u0442 \u0438 \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0439, \u043e\u0442\u0434\u0435\u043b\u044c\u043d\u043e \u043e\u0442 \u043f\u043e\u0434\u0441\u0432\u0435\u0442\u043a\u0438 \u0437\u0430\u0434\u0430\u0447.",
        "status": "Правила из name_exclusion_rules показываются в read-only режиме.",
    },
    {
        "slug": "gesn-exceptions",
        "title": "GESN exceptions",
        "description": "\u041e\u0434\u043e\u0431\u0440\u0435\u043d\u043d\u044b\u0435 \u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d\u044b \u0446\u0435\u043d \u043f\u043e \u0441\u0432\u044f\u0437\u043a\u0435 \u043a\u043e\u0434 + \u0435\u0434\u0438\u043d\u0438\u0446\u0430 + \u043f\u0440\u0438\u0437\u043d\u0430\u043a \u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436\u0430.",
        "status": "Одобренные диапазоны из gesn_exceptions показываются в read-only режиме.",
    },
    {
        "slug": "settings",
        "title": "\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438",
        "description": "\u0422\u0435\u0445\u043d\u0438\u0447\u0435\u0441\u043a\u0438\u0435 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b \u043f\u0440\u043e\u0435\u043a\u0442\u0430: \u0431\u0430\u0437\u0430, \u043a\u043e\u043d\u0444\u0438\u0433\u0438, \u0441\u0442\u0430\u0442\u0443\u0441 \u0441\u043b\u043e\u0432\u0430\u0440\u0435\u0439 \u0438 \u0432\u0435\u0440\u0441\u0438\u044f writer.",
        "status": "Read-only диагностика путей, базы и счетчиков подключена.",
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
input[type=file], input[type=text], select, textarea {
  display: block; width: 100%; margin-top: 6px; padding: 10px 12px;
  border: 1px solid #cbd5e1; border-radius: 10px; font-size: 14px; font-weight: 400;
}
textarea { min-height: 70px; resize: vertical; font-family: inherit; }
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
.admin-form {
  margin: 16px 0 20px; padding: 14px; border: 1px solid #e2e8f0;
  border-radius: 12px; background: #fff;
}
.table-action { margin: 0; }
.table-action button { width: auto; margin: 0; padding: 7px 10px; font-size: 12px; }
.notice-ok {
  margin-top: 14px; padding: 12px; border-radius: 10px;
  background: #ecfdf5; color: #047857; font-size: 14px;
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

def render_admin_imports(
    imports: list[ImportedFileRecord],
    *,
    notice: str = "",
    error: str = "",
    dry_run_result: RnmcZipDryRunResult | None = None,
    row_preview_result: RnmcZipRowPreviewResult | None = None,
    catalog_import_result: RnmcZipCatalogImportResult | None = None,
    status_filter: str = "",
) -> str:
    notice_html = f'<p class="notice-ok">{html.escape(notice)}</p>' if notice else ""
    error_html = f'<p class="notice">{html.escape(error)}</p>' if error else ""
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Импорты файлов</h2>'
        '<p>Журнал показывает, какие файлы были загружены, из какого источника и сколько строк принято.</p>'
        f'{notice_html}'
        f'{error_html}'
        f'{_render_file_log_import_form()}'
        f'{_render_rnmc_zip_dry_run_form()}'
        f'{_render_rnmc_zip_import_log_form()}'
        f'{_render_rnmc_zip_row_preview_form()}'
        f'{_render_rnmc_zip_catalog_import_form()}'
        f'{_render_rnmc_zip_dry_run_result(dry_run_result)}'
        f'{_render_rnmc_zip_row_preview_result(row_preview_result)}'
        f'{_render_rnmc_zip_catalog_import_result(catalog_import_result)}'
        f'{_render_import_status_filters(status_filter)}'
        f'{_render_imported_file_table(imports)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Импорты файлов",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="imports"),
        content=content,
    )


def _render_file_log_import_form() -> str:
    return (
        '<form class="admin-form" action="/admin/imports/file-log" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Импорт старого File_Log.xlsx</h2>'
        '<p class="muted">Файлы из старого лога будут считаться уже обработанными. Повтор имени файла будет отмечен как duplicate_name.</p>'
        '<label>File_Log.xlsx<input type="file" name="file_log" accept=".xlsx,.xlsm" required></label>'
        '<button type="submit">Импортировать FileLog</button>'
        '</form>'
    )



def _render_rnmc_zip_dry_run_form() -> str:
    return (
        '<form class="admin-form" action="/admin/imports/rnmc-dry-run" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Проверка ZIP с новыми РНМЦ</h2>'
        '<p class="muted">Dry-run ничего не пишет в каталог. Он только показывает, какие Excel-файлы будут обработаны, пропущены или отмечены как дубликаты имени.</p>'
        '<label>ZIP-архив РНМЦ<input type="file" name="rnmc_zip" accept=".zip" required></label>'
        '<label>Регион вручную, если нужно<input type="text" name="region_override" placeholder="Оставьте пустым, чтобы взять регион из папки"></label>'
        '<button type="submit">Проверить ZIP без импорта</button>'
        '</form>'
    )


def _render_rnmc_zip_import_log_form() -> str:
    return (
        '<form class="admin-form" action="/admin/imports/rnmc-log" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Зафиксировать ZIP в журнале</h2>'
        '<p class="muted">Эта кнопка записывает только журнал imported_files: новые файлы станут pending, уже обработанные будут skipped, дубликаты имени — duplicate_name. Строки каталога пока не добавляются.</p>'
        '<label>ZIP-архив РНМЦ<input type="file" name="rnmc_zip" accept=".zip" required></label>'
        '<label>Регион вручную, если нужно<input type="text" name="region_override" placeholder="Оставьте пустым, чтобы взять регион из папки"></label>'
        '<button type="submit">Записать ZIP в журнал без каталога</button>'
        '</form>'
    )


def _render_rnmc_zip_row_preview_form() -> str:
    return (
        '<form class="admin-form" action="/admin/imports/rnmc-row-preview" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Предпросмотр строк РНМЦ из ZIP</h2>'
        '<p class="muted">Этот режим открывает Excel-файлы, ищет таблицу работ и считает строки по правилам VBA. В каталог ничего не записывается.</p>'
        '<label>ZIP-архив РНМЦ<input type="file" name="rnmc_zip" accept=".zip" required></label>'
        '<label>Регион вручную, если нужно<input type="text" name="region_override" placeholder="Оставьте пустым, чтобы взять регион из папки"></label>'
        '<button type="submit">Разобрать строки без записи в каталог</button>'
        '</form>'
    )



def _render_rnmc_zip_catalog_import_form() -> str:
    return (
        '<form class="admin-form" action="/admin/imports/rnmc-import" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Импорт строк РНМЦ из ZIP в каталог</h2>'
        '<p class="muted">Этот режим пишет валидные строки в catalog_items и обновляет imported_files. Уже обработанные имена файлов пропускаются, повторы внутри ZIP получают duplicate_name.</p>'
        '<label>ZIP-архив РНМЦ<input type="file" name="rnmc_zip" accept=".zip" required></label>'
        '<label>Регион вручную, если нужно<input type="text" name="region_override" placeholder="Оставьте пустым, чтобы взять регион из папки"></label>'
        '<button type="submit">Импортировать строки в каталог</button>'
        '</form>'
    )


def _render_rnmc_zip_dry_run_result(result: RnmcZipDryRunResult | None) -> str:
    if result is None:
        return ""

    summary = (
        '<div class="admin-form">'
        '<h2 class="section">Результат dry-run ZIP</h2>'
        '<dl class="stats">'
        f'<div><dt>Excel-файлов найдено</dt><dd>{result.total_excel_files}</dd></div>'
        f'<div><dt>Будет обработано</dt><dd>{result.will_process_count}</dd></div>'
        f'<div><dt>Уже было обработано</dt><dd>{result.skipped_processed_count}</dd></div>'
        f'<div><dt>Дубликаты имени в ZIP</dt><dd>{result.duplicate_name_count}</dd></div>'
        f'<div><dt>Прочих файлов проигнорировано</dt><dd>{result.ignored_files}</dd></div>'
        '</dl>'
    )
    if not result.entries:
        return summary + '<p class="muted">В ZIP не найдено Excel-файлов РНМЦ.</p></div>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>Путь в ZIP</th>'
        '<th>Файл</th>'
        '<th>Регион</th>'
        '<th>Статус</th>'
        '<th>Причина</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in result.entries:
        rows.append(
            '<tr>'
            f'<td>{html.escape(entry.archive_path)}</td>'
            f'<td>{html.escape(entry.filename)}</td>'
            f'<td>{html.escape(entry.region_folder)}</td>'
            f'<td>{html.escape(entry.status)}</td>'
            f'<td>{html.escape(entry.reason)}</td>'
            '</tr>'
        )
    return summary + header + ''.join(rows) + '</tbody></table></div>'

def _render_rnmc_zip_row_preview_result(result: RnmcZipRowPreviewResult | None) -> str:
    if result is None:
        return ""

    summary = (
        '<div class="admin-form">'
        '<h2 class="section">Предпросмотр строк РНМЦ</h2>'
        '<dl class="stats">'
        f'<div><dt>Excel-файлов найдено</dt><dd>{result.total_excel_files}</dd></div>'
        f'<div><dt>Файлов с найденными строками</dt><dd>{result.preview_ok_count}</dd></div>'
        f'<div><dt>OK-строк найдено</dt><dd>{result.rows_ok_total}</dd></div>'
        f'<div><dt>Rejected-строк</dt><dd>{result.rows_rejected_total}</dd></div>'
        f'<div><dt>Уже обработано</dt><dd>{result.skipped_processed_count}</dd></div>'
        f'<div><dt>Дубликаты имени</dt><dd>{result.duplicate_name_count}</dd></div>'
        f'<div><dt>Без таблицы</dt><dd>{result.no_table_count}</dd></div>'
        f'<div><dt>Без строк</dt><dd>{result.no_rows_count}</dd></div>'
        f'<div><dt>Неподдерживаемый формат</dt><dd>{result.unsupported_format_count}</dd></div>'
        f'<div><dt>Ошибки чтения</dt><dd>{result.parse_error_count}</dd></div>'
        f'<div><dt>Прочих файлов проигнорировано</dt><dd>{result.ignored_files}</dd></div>'
        '</dl>'
    )
    if not result.entries:
        return summary + '<p class="muted">В ZIP не найдено Excel-файлов РНМЦ.</p></div>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>Путь в ZIP</th>'
        '<th>Файл</th>'
        '<th>Регион</th>'
        '<th>Статус</th>'
        '<th>Причина</th>'
        '<th>Лист</th>'
        '<th>Header row</th>'
        '<th>Задача</th>'
        '<th>ЛСР</th>'
        '<th>Начало</th>'
        '<th>Окончание</th>'
        '<th>OK</th>'
        '<th>Rejected</th>'
        '<th>Примеры строк</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in result.entries:
        rows.append(
            '<tr>'
            f'<td>{html.escape(entry.archive_path)}</td>'
            f'<td>{html.escape(entry.filename)}</td>'
            f'<td>{html.escape(entry.region_folder)}</td>'
            f'<td>{html.escape(entry.status)}</td>'
            f'<td>{html.escape(entry.reason)}</td>'
            f'<td>{html.escape(entry.sheet_name)}</td>'
            f'<td>{entry.header_row}</td>'
            f'<td>{html.escape(entry.task_number)}</td>'
            f'<td>{html.escape(entry.lsr_quarter)}</td>'
            f'<td>{html.escape(entry.planned_start)}</td>'
            f'<td>{html.escape(entry.planned_finish)}</td>'
            f'<td>{entry.rows_ok}</td>'
            f'<td>{entry.rows_rejected}</td>'
            f'<td>{_render_rnmc_samples(entry.sample_rows)}</td>'
            '</tr>'
        )
    return summary + header + ''.join(rows) + '</tbody></table></div>'


def _render_rnmc_samples(samples) -> str:
    if not samples:
        return ''
    parts = []
    for sample in samples:
        parts.append(
            f'{sample.row_number}: '
            f'{html.escape(sample.work_name)} | '
            f'{html.escape(sample.unit)} | '
            f'{html.escape(sample.quantity)}'
        )
    return '<br>'.join(parts)



def _render_rnmc_zip_catalog_import_result(result: RnmcZipCatalogImportResult | None) -> str:
    if result is None:
        return ""

    summary = (
        '<div class="admin-form">'
        '<h2 class="section">Результат импорта ZIP в каталог</h2>'
        '<dl class="stats">'
        f'<div><dt>Excel-файлов найдено</dt><dd>{result.total_excel_files}</dd></div>'
        f'<div><dt>Успешных файлов</dt><dd>{result.success_count}</dd></div>'
        f'<div><dt>Строк добавлено</dt><dd>{result.rows_imported_total}</dd></div>'
        f'<div><dt>Строк отклонено</dt><dd>{result.rows_rejected_total}</dd></div>'
        f'<div><dt>Без данных</dt><dd>{result.no_data_count}</dd></div>'
        f'<div><dt>Пропущено</dt><dd>{result.skipped_count}</dd></div>'
        f'<div><dt>Дубликаты имени</dt><dd>{result.duplicate_name_count}</dd></div>'
        f'<div><dt>Ошибки</dt><dd>{result.failed_count}</dd></div>'
        f'<div><dt>Прочих файлов проигнорировано</dt><dd>{result.ignored_files}</dd></div>'
        '</dl>'
    )
    if not result.entries:
        return summary + '<p class="muted">В ZIP не найдено Excel-файлов РНМЦ.</p></div>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>Путь в ZIP</th>'
        '<th>Файл</th>'
        '<th>Регион</th>'
        '<th>Статус</th>'
        '<th>Причина</th>'
        '<th>Лист</th>'
        '<th>Header row</th>'
        '<th>Задача</th>'
        '<th>ЛСР</th>'
        '<th>Начало</th>'
        '<th>Окончание</th>'
        '<th>Добавлено</th>'
        '<th>Отклонено</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in result.entries:
        rows.append(
            '<tr>'
            f'<td>{html.escape(entry.archive_path)}</td>'
            f'<td>{html.escape(entry.filename)}</td>'
            f'<td>{html.escape(entry.region_folder)}</td>'
            f'<td>{html.escape(entry.status)}</td>'
            f'<td>{html.escape(entry.reason)}</td>'
            f'<td>{html.escape(entry.sheet_name)}</td>'
            f'<td>{entry.header_row}</td>'
            f'<td>{html.escape(entry.task_number)}</td>'
            f'<td>{html.escape(entry.lsr_quarter)}</td>'
            f'<td>{html.escape(entry.planned_start)}</td>'
            f'<td>{html.escape(entry.planned_finish)}</td>'
            f'<td>{entry.rows_imported}</td>'
            f'<td>{entry.rows_rejected}</td>'
            '</tr>'
        )
    return summary + header + ''.join(rows) + '</tbody></table></div>'

def _render_import_status_filters(active_status: str) -> str:
    statuses = [
        ("", "Все"),
        ("success", "success"),
        ("pending", "pending"),
        ("failed", "failed"),
        ("no_data", "no_data"),
        ("duplicate_name", "duplicate_name"),
        ("skipped", "skipped"),
        ("legacy_imported", "legacy_imported"),
    ]
    links = []
    for status, label in statuses:
        href = "/admin/imports" if status == "" else f"/admin/imports?status={quote(status)}"
        active = " active" if status == active_status else ""
        links.append(
            f'<a class="admin-nav-link{active}" href="{href}">{html.escape(label)}</a>'
        )
    return (
        '<div class="admin-form">'
        '<h2 class="section">Фильтры журнала</h2>'
        '<div class="admin-nav">' + ''.join(links) + '</div>'
        '</div>'
    )


def _render_imported_file_table(imports: list[ImportedFileRecord]) -> str:
    if not imports:
        return '<p class="muted">Импорты пока не записаны.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>ID</th>'
        '<th>Файл</th>'
        '<th>Источник</th>'
        '<th>Тип</th>'
        '<th>Регион</th>'
        '<th>Задача</th>'
        '<th>Статус</th>'
        '<th>OK</th>'
        '<th>Rejected</th>'
        '<th>Legacy</th>'
        '<th>ЛСР</th>'
        '<th>Начало</th>'
        '<th>Окончание</th>'
        '<th>Импорт</th>'
        '<th>Ошибка</th>'
        '<th>Действия</th>'
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
            f'<td>{html.escape(item.legacy_note)}</td>'
            f'<td>{html.escape(item.lsr_quarter)}</td>'
            f'<td>{html.escape(item.planned_start)}</td>'
            f'<td>{html.escape(item.planned_finish)}</td>'
            f'<td>{html.escape(item.imported_at)}</td>'
            f'<td>{html.escape(item.failure_reason)}</td>'
            f'<td><a href="/admin/imports/{item.id}">Детали</a></td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'


def render_admin_import_detail(
    record: ImportedFileRecord,
    catalog_rows: list[CatalogItemRecord],
    row_logs: list[ImportRowLogRecord],
    *,
    notice: str = "",
    error: str = "",
) -> str:
    notice_html = f'<p class="notice-ok">{html.escape(notice)}</p>' if notice else ""
    error_html = f'<p class="notice">{html.escape(error)}</p>' if error else ""
    retry_form = ""
    if record.status in {"failed", "no_data"}:
        retry_form = (
            '<form class="admin-form" action="/admin/imports/allow-retry" method="post">'
            '<h2 class="section">Повторная обработка</h2>'
            '<p class="muted">Кнопка переводит файл в pending. После этого загрузите ZIP еще раз, и файл будет обработан повторно.</p>'
            f'<input type="hidden" name="import_id" value="{record.id}">'
            '<button type="submit">Разрешить повторную обработку</button>'
            '</form>'
        )
    content = (
        '<section class="admin-panel">'
        f'<p><a href="/admin/imports">← Назад к импортам</a></p>'
        '<h2 class="section">Детали импорта</h2>'
        f'{notice_html}{error_html}'
        f'{_render_import_detail_summary(record)}'
        f'{_render_import_metadata_form(record)}'
        f'{retry_form}'
        f'{_render_import_catalog_rows(catalog_rows)}'
        f'{_render_import_row_logs(row_logs)}'
        '</section>'
    )
    return render(
        "admin.html",
        title=f"Импорт #{record.id}",
        subtitle="Контроль конкретного файла РНМЦ.",
        admin_nav=_render_admin_nav(active_slug="imports"),
        content=content,
    )


def _render_import_detail_summary(record: ImportedFileRecord) -> str:
    rows = [
        ("ID", str(record.id)),
        ("Файл", record.filename),
        ("Регион", record.region_folder),
        ("Статус", record.status),
        ("Задача", record.task_number),
        ("OK-строки", str(record.rows_ok)),
        ("Rejected-строки", str(record.rows_rejected)),
        ("ЛСР", record.lsr_quarter),
        ("Начало", record.planned_start),
        ("Окончание", record.planned_finish),
        ("Ошибка", record.failure_reason),
    ]
    body = ''.join(
        f'<div><dt>{html.escape(label)}</dt><dd>{html.escape(value)}</dd></div>'
        for label, value in rows
    )
    return f'<dl class="stats">{body}</dl>'


def _render_import_metadata_form(record: ImportedFileRecord) -> str:
    return (
        '<form class="admin-form" action="/admin/imports/update" method="post">'
        '<h2 class="section">Ручная правка данных файла</h2>'
        f'<input type="hidden" name="import_id" value="{record.id}">'
        f'<label>Регион<input type="text" name="region_folder" value="{html.escape(record.region_folder)}"></label>'
        f'<label>Номер задачи<input type="text" name="task_number" value="{html.escape(record.task_number)}"></label>'
        f'<label>Год/квартал ЛСР<input type="text" name="lsr_quarter" value="{html.escape(record.lsr_quarter)}"></label>'
        f'<label>Планируемое начало<input type="text" name="planned_start" value="{html.escape(record.planned_start)}"></label>'
        f'<label>Планируемое окончание<input type="text" name="planned_finish" value="{html.escape(record.planned_finish)}"></label>'
        '<button type="submit">Сохранить данные файла</button>'
        '</form>'
    )


def _render_import_catalog_rows(rows: list[CatalogItemRecord]) -> str:
    if not rows:
        return '<div class="admin-form"><h2 class="section">Импортированные строки</h2><p class="muted">Для этого файла строки catalog_items не найдены.</p></div>'
    header = (
        '<div class="admin-form"><h2 class="section">Импортированные строки</h2>'
        '<table class="preview"><thead><tr>'
        '<th>Excel row</th><th>Код</th><th>Ед.</th><th>Цена</th><th>Работа</th>'
        '</tr></thead><tbody>'
    )
    body = []
    for row in rows[:200]:
        body.append(
            '<tr>'
            f'<td>{row.source_row_number}</td>'
            f'<td>{html.escape(row.code)}</td>'
            f'<td>{html.escape(row.unit)}</td>'
            f'<td>{row.price:g}</td>'
            f'<td>{html.escape(row.work_name)}</td>'
            '</tr>'
        )
    tail = '</tbody></table>'
    if len(rows) > 200:
        tail += f'<p class="muted">Показаны первые 200 строк из {len(rows)}.</p>'
    return header + ''.join(body) + tail + '</div>'


def _render_import_row_logs(rows: list[ImportRowLogRecord]) -> str:
    if not rows:
        return '<div class="admin-form"><h2 class="section">Rejected-лог</h2><p class="muted">Rejected-строки по файлу не записаны.</p></div>'
    header = (
        '<div class="admin-form"><h2 class="section">Rejected-лог</h2>'
        '<table class="preview"><thead><tr>'
        '<th>Excel row</th><th>Статус</th><th>Причина</th>'
        '</tr></thead><tbody>'
    )
    body = []
    for row in rows[:300]:
        body.append(
            '<tr>'
            f'<td>{row.row_number}</td>'
            f'<td>{html.escape(row.status)}</td>'
            f'<td>{html.escape(row.reason)}</td>'
            '</tr>'
        )
    tail = '</tbody></table>'
    if len(rows) > 300:
        tail += f'<p class="muted">Показаны первые 300 записей из {len(rows)}.</p>'
    return header + ''.join(body) + tail + '</div>'


def render_admin_risks(risks: list[PriceRiskLogEntry]) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Риск-лог</h2>'
        '<p>Журнал показывает строки с сильным разбросом цен, которые требуют проверки и последующего одобрения диапазона.</p>'
        f'{_render_price_risk_table(risks)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Риск-лог",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="risks"),
        content=content,
    )


def _render_price_risk_table(risks: list[PriceRiskLogEntry]) -> str:
    if not risks:
        return '<p class="muted">Риск-лог пока пуст.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>ID</th>'
        '<th>Статус</th>'
        '<th>Ключ</th>'
        '<th>Причина</th>'
        '<th>Код</th>'
        '<th>Ед.</th>'
        '<th>Min</th>'
        '<th>Max</th>'
        '<th>Ratio</th>'
        '<th>Реком.</th>'
        '<th>Строка сметы</th>'
        '<th>Первое появление</th>'
        '<th>Последнее появление</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for risk in risks:
        rows.append(
            '<tr>'
            f'<td>{risk.id}</td>'
            f'<td>{html.escape(risk.status)}</td>'
            f'<td>{html.escape(risk.exception_key)}</td>'
            f'<td>{html.escape(risk.reason)}</td>'
            f'<td>{html.escape(risk.code)}</td>'
            f'<td>{html.escape(risk.unit)}</td>'
            f'<td>{_fmt_number(risk.min_price)}</td>'
            f'<td>{_fmt_number(risk.max_price)}</td>'
            f'<td>{_fmt_number(risk.ratio)}</td>'
            f'<td>{_fmt_number(risk.recommended_price)}</td>'
            f'<td>{_fmt_number(risk.estimate_row)}</td>'
            f'<td>{html.escape(risk.first_seen_at)}</td>'
            f'<td>{html.escape(risk.last_seen_at)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'


def render_admin_gesn_exceptions(exceptions: list[GesnException]) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">GESN exceptions</h2>'
        '<p>Одобренные диапазоны показывают, какие min/max цены уже приняты человеком для связки код + единица + признак демонтажа.</p>'
        f'{_render_gesn_exceptions_table(exceptions)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="GESN exceptions",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="gesn-exceptions"),
        content=content,
    )


def _render_gesn_exceptions_table(exceptions: list[GesnException]) -> str:
    if not exceptions:
        return '<p class="muted">Одобренные диапазоны пока не записаны.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>Ключ</th>'
        '<th>Код</th>'
        '<th>Ед.</th>'
        '<th>Демонтаж</th>'
        '<th>Approved min</th>'
        '<th>Approved max</th>'
        '<th>Дата обновления диапазона</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for item in exceptions:
        unit, code, dem_flag = _split_exception_key(item.exception_key)
        rows.append(
            '<tr>'
            f'<td>{html.escape(item.exception_key)}</td>'
            f'<td>{html.escape(code)}</td>'
            f'<td>{html.escape(unit)}</td>'
            f'<td>{html.escape(_dem_flag_label(dem_flag))}</td>'
            f'<td>{_fmt_number(item.approved_min)}</td>'
            f'<td>{_fmt_number(item.approved_max)}</td>'
            f'<td>{_fmt_number(item.last_range_update_date)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'


def _split_exception_key(exception_key: str) -> tuple[str, str, str]:
    parts = exception_key.split("||")
    if len(parts) != 3:
        return "", "", ""
    return parts[0], parts[1], parts[2]


def _dem_flag_label(dem_flag: str) -> str:
    if dem_flag == "1":
        return "да"
    if dem_flag == "0":
        return "нет"
    return dem_flag



def render_admin_approvals(
    risks: list[PriceRiskLogEntry],
    error: str = "",
    notice: str = "",
) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Одобрение диапазонов</h2>'
        '<p>Здесь показываются открытые риски. Кнопка одобрения переносит min/max риска в GESN exceptions и меняет статус риска на approved.</p>'
        f'{_render_admin_message(error, notice)}'
        f'{_render_approval_risk_table(risks)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Одобрение диапазонов",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="approvals"),
        content=content,
    )


def _render_approval_risk_table(risks: list[PriceRiskLogEntry]) -> str:
    if not risks:
        return '<p class="muted">Открытых рисков для одобрения пока нет.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>ID</th>'
        '<th>Ключ</th>'
        '<th>Причина</th>'
        '<th>Код</th>'
        '<th>Ед.</th>'
        '<th>Min</th>'
        '<th>Max</th>'
        '<th>Ratio</th>'
        '<th>Реком.</th>'
        '<th>Строка сметы</th>'
        '<th>Последнее появление</th>'
        '<th>Действие</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for risk in risks:
        rows.append(
            '<tr>'
            f'<td>{risk.id}</td>'
            f'<td>{html.escape(risk.exception_key)}</td>'
            f'<td>{html.escape(risk.reason)}</td>'
            f'<td>{html.escape(risk.code)}</td>'
            f'<td>{html.escape(risk.unit)}</td>'
            f'<td>{_fmt_number(risk.min_price)}</td>'
            f'<td>{_fmt_number(risk.max_price)}</td>'
            f'<td>{_fmt_number(risk.ratio)}</td>'
            f'<td>{_fmt_number(risk.recommended_price)}</td>'
            f'<td>{_fmt_number(risk.estimate_row)}</td>'
            f'<td>{html.escape(risk.last_seen_at)}</td>'
            f'<td>{_render_approval_button(risk)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'



def _render_approval_button(risk: PriceRiskLogEntry) -> str:
    key_value = html.escape(risk.exception_key, quote=True)
    return (
        '<form class="table-action" method="post" action="/admin/approvals/approve">'
        f'<input type="hidden" name="exception_key" value="{key_value}">'
        '<button type="submit">Одобрить</button>'
        '</form>'
    )

def render_admin_name_exclusions(
    rules: list[NameExclusionRule],
    error: str = "",
    notice: str = "",
) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Исключения по наименованиям</h2>'
        '<p>Правила исключений хранятся отдельно от синих задач. Они влияют на фильтрацию по тексту работ и могут включаться или выключаться из админки.</p>'
        f'{_render_admin_message(error, notice)}'
        f'{_render_name_exclusion_form()}'
        f'{_render_name_exclusion_table(rules)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Исключения по наименованиям",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="name-exclusions"),
        content=content,
    )


def _render_name_exclusion_form() -> str:
    return (
        '<form class="admin-form" method="post" action="/admin/name-exclusions/add">'
        '<h2 class="section">Добавить или включить правило</h2>'
        '<label>Scope'
        '<select name="scope">'
        '<option value="BOTH">BOTH</option>'
        '<option value="SMETA">SMETA</option>'
        '<option value="CATALOG">CATALOG</option>'
        '</select>'
        '</label>'
        '<label>Match mode'
        '<select name="match_mode">'
        '<option value="ALL_WORDS">ALL_WORDS</option>'
        '<option value="CONTAINS">CONTAINS</option>'
        '</select>'
        '</label>'
        '<label>Pattern'
        '<input type="text" name="pattern" placeholder="Например, демонтаж|временный" required>'
        '</label>'
        '<label>Группа'
        '<input type="text" name="rule_group" placeholder="Например, noise">'
        '</label>'
        '<label>Комментарий'
        '<textarea name="comment" placeholder="Короткое пояснение для админки"></textarea>'
        '</label>'
        '<button type="submit">Добавить / включить</button>'
        '</form>'
    )


def _render_name_exclusion_table(rules: list[NameExclusionRule]) -> str:
    if not rules:
        return '<p class="muted">Правила исключений пока не записаны.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>Вкл.</th>'
        '<th>Scope</th>'
        '<th>Match mode</th>'
        '<th>Pattern</th>'
        '<th>Группа</th>'
        '<th>Комментарий</th>'
        '<th>Действие</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for rule in rules:
        rows.append(
            '<tr>'
            f'<td>{_enabled_label(rule.enabled)}</td>'
            f'<td>{html.escape(rule.scope)}</td>'
            f'<td>{html.escape(rule.match_mode)}</td>'
            f'<td>{html.escape(rule.pattern)}</td>'
            f'<td>{html.escape(rule.group)}</td>'
            f'<td>{html.escape(rule.comment)}</td>'
            f'<td>{_render_name_exclusion_toggle(rule)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'


def _render_name_exclusion_toggle(rule: NameExclusionRule) -> str:
    next_enabled = "0" if rule.enabled else "1"
    button_label = "Выключить" if rule.enabled else "Включить"
    scope_value = html.escape(rule.scope, quote=True)
    mode_value = html.escape(rule.match_mode, quote=True)
    pattern_value = html.escape(rule.pattern, quote=True)
    return (
        '<form class="table-action" method="post" action="/admin/name-exclusions/toggle">'
        f'<input type="hidden" name="scope" value="{scope_value}">'
        f'<input type="hidden" name="match_mode" value="{mode_value}">'
        f'<input type="hidden" name="pattern" value="{pattern_value}">'
        f'<input type="hidden" name="enabled" value="{next_enabled}">'
        f'<button type="submit">{button_label}</button>'
        '</form>'
    )


def render_admin_settings(settings_rows: list[tuple[str, str]]) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Настройки и диагностика</h2>'
        '<p>Read-only сводка по базе, конфигам и накопленным админ-таблицам. Эта страница ничего не меняет.</p>'
        f'{_render_settings_table(settings_rows)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Настройки",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="settings"),
        content=content,
    )


def _render_settings_table(settings_rows: list[tuple[str, str]]) -> str:
    if not settings_rows:
        return '<p class="muted">Диагностика пока недоступна.</p>'

    rows = []
    for label, value in settings_rows:
        rows.append(
            '<tr>'
            f'<th>{html.escape(label)}</th>'
            f'<td>{html.escape(value)}</td>'
            '</tr>'
        )
    return '<table class="preview"><tbody>' + ''.join(rows) + '</tbody></table>'

def render_admin_task_colors(
    entries: list[TaskColorEntry],
    error: str = "",
    notice: str = "",
) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Синие задачи</h2>'
        '<p>Задачи из этого списка не блокируют подбор аналогов. Их аналоги остаются в результате, но должны подсвечиваться синим.</p>'
        f'{_render_admin_message(error, notice)}'
        f'{_render_task_color_form()}'
        f'{_render_task_color_table(entries)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Синие задачи",
        subtitle="Раздел администрирования автоподборщика.",
        admin_nav=_render_admin_nav(active_slug="task-colors"),
        content=content,
    )


def _render_admin_message(error: str, notice: str) -> str:
    if error:
        return f'<p class="notice">{html.escape(error)}</p>'
    if notice:
        return f'<p class="notice-ok">{html.escape(notice)}</p>'
    return ""


def _render_task_color_form() -> str:
    return (
        '<form class="admin-form" method="post" action="/admin/task-colors/add">'
        '<h2 class="section">Добавить или включить задачу</h2>'
        '<label>Номер задачи'
        '<input type="text" name="task_number" placeholder="Например, TASK-123" required>'
        '</label>'
        '<label>Причина'
        '<input type="text" name="reason" placeholder="Например, manual_review">'
        '</label>'
        '<label>Комментарий'
        '<textarea name="comment" placeholder="Короткое пояснение для админки"></textarea>'
        '</label>'
        '<button type="submit">Добавить / включить</button>'
        '</form>'
    )


def _render_task_color_table(entries: list[TaskColorEntry]) -> str:
    if not entries:
        return '<p class="muted">Синие задачи пока не записаны.</p>'

    header = (
        '<table class="preview"><thead><tr>'
        '<th>Вкл.</th>'
        '<th>Номер задачи</th>'
        '<th>Причина</th>'
        '<th>Комментарий</th>'
        '<th>Действие</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in entries:
        rows.append(
            '<tr>'
            f'<td>{_enabled_label(entry.enabled)}</td>'
            f'<td>{html.escape(entry.task_number)}</td>'
            f'<td>{html.escape(entry.reason)}</td>'
            f'<td>{html.escape(entry.comment)}</td>'
            f'<td>{_render_task_color_toggle(entry)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table>'


def _render_task_color_toggle(entry: TaskColorEntry) -> str:
    next_enabled = "0" if entry.enabled else "1"
    button_label = "Выключить" if entry.enabled else "Включить"
    task_value = html.escape(entry.task_number, quote=True)
    return (
        '<form class="table-action" method="post" action="/admin/task-colors/toggle">'
        f'<input type="hidden" name="task_number" value="{task_value}">'
        f'<input type="hidden" name="enabled" value="{next_enabled}">'
        f'<button type="submit">{button_label}</button>'
        '</form>'
    )


def _enabled_label(enabled: bool) -> str:
    return "да" if enabled else "нет"


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

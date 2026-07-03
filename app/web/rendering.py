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
from core.macro_workbook import load_default_macro_settings
from core.risk import DEFAULT_PRICE_SPREAD_LIMIT

_WRITER_MODULE = Path(__file__).resolve().parents[2] / "core" / "excel_writer.py"

TEMPLATES_DIR = Path(__file__).parent / "templates"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

READ_METHOD_LABELS = {
    "template": "\u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443",
    "detected": "\u043f\u043e \u0440\u0430\u0441\u043f\u043e\u0437\u043d\u0430\u043d\u043d\u044b\u043c \u0441\u0442\u043e\u043b\u0431\u0446\u0430\u043c",
}

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


def render(template_name: str, **context: str) -> str:
    text = (TEMPLATES_DIR / template_name).read_text(encoding="utf-8")
    context.setdefault("styles", STYLES)
    context.setdefault("message", "")
    context.setdefault("detail", "")
    context.setdefault("preview", "")
    context.setdefault("build_stamp", writer_build_stamp())
    context.setdefault("macro_status", macro_status_label())
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

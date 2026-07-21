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
from urllib.parse import quote, urlencode

from app.services.write_result import RunAndWriteResult
from app.services.catalog_source import catalog_status_label, database_has_catalog
from app.services.rnmc_zip import RnmcZipDryRunResult
from app.services.rnmc_excel import (
    DEFAULT_ROW_PREVIEW_LIMIT,
    RnmcZipCatalogImportResult,
    RnmcZipRowPreviewResult,
)
from core.macro_workbook import load_default_macro_settings
from core.risk import DEFAULT_PRICE_SPREAD_LIMIT, GesnException
from core.storage.catalog import CatalogEditorPage, CatalogEditorRow, CatalogItemRecord, CatalogSource, ImportedFileRecord, ImportRowLogRecord, normalize_import_filename
from core.storage.risk_log import PriceRiskLogEntry
from core.exclusions import NameExclusionRule, TaskColorEntry

_WRITER_MODULE = Path(__file__).resolve().parents[2] / "core" / "excel_writer.py"

TEMPLATES_DIR = Path(__file__).parent / "templates"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

READ_METHOD_LABELS = {
    "template": "\u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443",
    "detected": "\u043f\u043e \u0440\u0430\u0441\u043f\u043e\u0437\u043d\u0430\u043d\u043d\u044b\u043c \u0441\u0442\u043e\u043b\u0431\u0446\u0430\u043c",
}

COEFFICIENT_METHOD_LABELS = {
    "explicit": "\u0432\u0432\u0435\u0434\u0451\u043d \u0432\u0440\u0443\u0447\u043d\u0443\u044e",
    "configured_cell": "\u0438\u0437 \u044f\u0447\u0435\u0439\u043a\u0438 \u0448\u0430\u0431\u043b\u043e\u043d\u0430",
    "labeled_region": "\u043d\u0430\u0439\u0434\u0435\u043d \u043f\u043e \u043f\u043e\u0434\u043f\u0438\u0441\u044f\u043c \u00ab\u0420\u0435\u0433\u0438\u043e\u043d:\u00bb/\u00ab\u041a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442:\u00bb",
    "labeled_coefficient": "\u043d\u0430\u0439\u0434\u0435\u043d \u043f\u043e \u043f\u043e\u0434\u043f\u0438\u0441\u0438 \u00ab\u041a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442:\u00bb",
    "default": "\u041d\u0415 \u041d\u0410\u0419\u0414\u0415\u041d \u0432 \u0444\u0430\u0439\u043b\u0435, \u043f\u0440\u0438\u043c\u0435\u043d\u0451\u043d 1.0 \u043f\u043e \u0443\u043c\u043e\u043b\u0447\u0430\u043d\u0438\u044e",
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
        "slug": "catalog",
        "title": "Каталог аналогов",
        "description": "Просмотр, фильтрация и ручная корректировка строк catalog_items после загрузки РНМЦ и других источников.",
        "status": "Подключен редактор строк каталога с фильтрами, удалением и групповыми действиями.",
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
  white-space: nowrap; vertical-align: top; }
table.preview th { position: sticky; top: 0; background: #f8fafc; color: #475569;
  font-weight: 600; z-index: 1; }
table.preview[data-rnmc-table] { table-layout: fixed; width: max-content; min-width: 100%; }
table.preview[data-rnmc-table] th { position: sticky; top: 0; padding-right: 16px; }
.rnmc-preview-shell table.preview[data-rnmc-table] td.path { max-width: none; }
.rnmc-col-resizer {
  position: absolute; top: 0; right: -6px; width: 12px; height: 100%;
  cursor: col-resize; user-select: none; z-index: 6; touch-action: none;
}
.rnmc-col-resizer::after {
  content: ""; position: absolute; top: 5px; bottom: 5px; left: 5px;
  width: 2px; border-radius: 999px; background: #93c5fd;
}
.rnmc-col-resizer:hover::after, .rnmc-col-resizer.active::after { background: #2563eb; width: 3px; }
body.rnmc-resizing, body.rnmc-resizing * { cursor: col-resize !important; user-select: none !important; }
table.preview td.risk { color: #b91c1c; text-align: center; font-weight: 700; }
table.preview td.wrap, table.preview th.wrap {
  white-space: normal; min-width: 260px; max-width: 520px; line-height: 1.35;
}
table.preview td.path {
  max-width: 340px; overflow: hidden; text-overflow: ellipsis;
}
.preview-wide { max-height: 560px; overflow: auto; border: 1px solid #e2e8f0;
  border-radius: 10px; margin-bottom: 16px; background: #fff; }
.rnmc-preview-shell { --rnmc-preview-scale: 1; --rnmc-preview-padding: 7px 10px; }
.rnmc-preview-shell table.preview { font-size: calc(13px * var(--rnmc-preview-scale)); }
.rnmc-preview-shell table.preview th, .rnmc-preview-shell table.preview td {
  padding: var(--rnmc-preview-padding);
}
.rnmc-preview-shell.compact { --rnmc-preview-padding: 4px 7px; }
.rnmc-preview-shell.large { --rnmc-preview-padding: 10px 13px; }
.rnmc-preview-toolbar {
  display: flex; flex-wrap: wrap; gap: 8px; align-items: end; margin: 8px 0 12px;
  padding: 10px; border: 1px solid #e2e8f0; border-radius: 10px; background: #f8fafc;
}
.rnmc-preview-toolbar label { margin: 0; font-size: 12px; color: #475569; font-weight: 600; }
.rnmc-preview-toolbar input[type=search], .rnmc-preview-toolbar select {
  display: block; width: auto; min-width: 170px; margin-top: 4px; padding: 7px 9px;
  border-radius: 8px; font-size: 13px;
}
.rnmc-preview-toolbar button {
  width: auto; margin: 0; padding: 8px 10px; border-radius: 8px; font-size: 13px;
  background: #e0e7ff; color: #3730a3;
}
.rnmc-preview-toolbar button:hover { background: #c7d2fe; }
.rnmc-preview-toolbar .rnmc-toggle { display: flex; gap: 6px; align-items: center; padding: 7px 0; }
.rnmc-preview-toolbar input[type=checkbox] { margin: 0; }
.rnmc-filter-count { color: #64748b; font-size: 12px; padding: 8px 0 0; }
.rnmc-coef-input {
  width: 96px; min-width: 84px; padding: 5px 7px; border: 1px solid #cbd5e1;
  border-radius: 7px; font-size: 13px; background: #fff;
}
.rnmc-coef-input:focus { outline: 2px solid #c7d2fe; border-color: #818cf8; }
.rnmc-hidden-row { display: none; }
.card.wide { max-width: min(1600px, 98vw); }
.rnmc-tabs { margin-top: 14px; }
.rnmc-tab-input { position: absolute; opacity: 0; pointer-events: none; }
.rnmc-tab-labels { display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 12px; }
.rnmc-tab-labels label {
  width: auto; margin: 0; padding: 8px 12px; border-radius: 999px;
  background: #eef2ff; color: #3730a3; font-size: 13px; cursor: pointer;
}
.rnmc-tab-panel { display: none; }
#rnmc-preview-tab-summary:checked ~ .rnmc-tab-labels label[for=rnmc-preview-tab-summary],
#rnmc-preview-tab-files:checked ~ .rnmc-tab-labels label[for=rnmc-preview-tab-files],
#rnmc-preview-tab-metadata:checked ~ .rnmc-tab-labels label[for=rnmc-preview-tab-metadata],
#rnmc-preview-tab-headers:checked ~ .rnmc-tab-labels label[for=rnmc-preview-tab-headers],
#rnmc-preview-tab-rows:checked ~ .rnmc-tab-labels label[for=rnmc-preview-tab-rows] {
  background: #4f46e5; color: #fff;
}
#rnmc-preview-tab-summary:checked ~ .rnmc-tab-panels #rnmc-preview-panel-summary,
#rnmc-preview-tab-files:checked ~ .rnmc-tab-panels #rnmc-preview-panel-files,
#rnmc-preview-tab-metadata:checked ~ .rnmc-tab-panels #rnmc-preview-panel-metadata,
#rnmc-preview-tab-headers:checked ~ .rnmc-tab-panels #rnmc-preview-panel-headers,
#rnmc-preview-tab-rows:checked ~ .rnmc-tab-panels #rnmc-preview-panel-rows {
  display: block;
}
.status-pill { display: inline-block; padding: 3px 7px; border-radius: 999px;
  background: #f1f5f9; color: #334155; font-size: 12px; font-weight: 600; }
.status-pill.ok { background: #ecfdf5; color: #047857; }
.status-pill.warn { background: #fff7ed; color: #c2410c; }
.catalog-editor-form { margin: 0; }
.catalog-filter-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 10px; align-items: end; margin-bottom: 12px;
}
.catalog-filter-grid label { margin: 0; }
.catalog-actions {
  display: flex; flex-wrap: wrap; gap: 8px; align-items: end; margin: 10px 0 12px;
  padding: 10px; border: 1px solid #e2e8f0; border-radius: 10px; background: #f8fafc;
}
.catalog-actions label { width: auto; margin: 0; min-width: 150px; font-size: 12px; color: #475569; }
.catalog-actions select, .catalog-actions input[type=text] { margin-top: 4px; padding: 7px 9px; border-radius: 8px; }
.catalog-actions button { width: auto; margin: 0; padding: 8px 10px; border-radius: 8px; font-size: 13px; }
.catalog-table input[type=text], .catalog-table input[type=number] {
  min-width: 86px; width: 100%; margin: 0; padding: 5px 6px; border-radius: 6px; font-size: 12px;
}
.catalog-table textarea { min-width: 260px; min-height: 42px; margin: 0; padding: 5px 6px; border-radius: 6px; font-size: 12px; }
.catalog-table .short-cell { min-width: 84px; }
.catalog-table .number-cell { min-width: 96px; }
.catalog-table .work-cell { min-width: 320px; }
.catalog-table .file-cell { min-width: 220px; }
.catalog-table button { width: auto; margin: 2px 0; padding: 6px 8px; font-size: 12px; border-radius: 7px; }
.catalog-table button.danger { background: #dc2626; }
.catalog-table button.danger:hover { background: #b91c1c; }
.catalog-pager { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 10px 0; color: #475569; }
.catalog-pager a { padding: 7px 10px; border-radius: 8px; background: #eef2ff; color: #3730a3; text-decoration: none; font-weight: 600; }
.catalog-warning { padding: 10px; margin: 10px 0; background: #fffbeb; color: #92400e; border-radius: 10px; }
.muted { color: #94a3b8; font-size: 12px; padding: 8px 10px; margin: 0; }
.maintenance-tools { margin-top: 18px; border-top: 1px solid #e2e8f0; padding-top: 12px; }
.maintenance-tools summary { cursor: pointer; color: #64748b; font-size: 13px; font-weight: 600; }
.maintenance-tools-body { margin-top: 10px; padding: 10px; border: 1px solid #fee2e2; border-radius: 10px; background: #fffafa; }
button.danger-small { width: auto; margin: 0; padding: 6px 9px; border-radius: 7px; font-size: 12px; background: #b91c1c; }
button.danger-small:hover { background: #991b1b; }
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
  grid-template-columns: repeat(2, minmax(0, 1fr));
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
.admin-section-head { margin-bottom: 14px; }
.admin-section-head p { margin: 4px 0 0; }
.admin-action-grid {
  display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px;
  align-items: start;
}
.admin-action-card { margin: 0; height: 100%; }
.admin-step-list { margin: 0; padding-left: 20px; color: #475569; font-size: 14px; line-height: 1.55; }
.admin-primary-action { margin-top: 12px; }
.admin-confirm-box {
  margin: 12px 0 20px; padding: 14px; border: 1px solid #bbf7d0;
  border-radius: 12px; background: #f0fdf4;
}
.admin-confirm-box form { margin: 0; }
.admin-confirm-box button { width: auto; min-width: 220px; }
@media (max-width: 760px) {
  .admin-grid, .admin-action-grid { grid-template-columns: 1fr; }
}
.admin-section-head { margin-bottom: 14px; }
.admin-section-head p { margin: 4px 0 0; }
.admin-action-grid {
  display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px;
  align-items: start;
}
.admin-action-card { margin: 0; height: 100%; }
.admin-step-list { margin: 0; padding-left: 20px; color: #475569; font-size: 14px; line-height: 1.55; }
.admin-primary-action { margin-top: 12px; }
.admin-confirm-box {
  margin: 12px 0 20px; padding: 14px; border: 1px solid #bbf7d0;
  border-radius: 12px; background: #f0fdf4;
}
.admin-confirm-box form { margin: 0; }
.admin-confirm-box button { width: auto; min-width: 220px; }
@media (max-width: 760px) {
  .admin-grid, .admin-action-grid { grid-template-columns: 1fr; }
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


def render_confirm(
    token: str,
    sheet_title: str,
    *,
    region_value: str,
    region_method: str | None,
    coefficient_value: float | str,
    coefficient_method: str,
    error: str | None = None,
) -> str:
    """Region/coefficient confirmation screen shown before the real run."""
    region_source = (
        "\u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u0432 \u0444\u0430\u0439\u043b\u0435, \u0432\u0432\u0435\u0434\u0438\u0442\u0435 \u0432\u0440\u0443\u0447\u043d\u0443\u044e"
        if not region_method
        else "\u043d\u0430\u0439\u0434\u0435\u043d \u0430\u0432\u0442\u043e\u043c\u0430\u0442\u0438\u0447\u0435\u0441\u043a\u0438"
    )
    coefficient_source = COEFFICIENT_METHOD_LABELS.get(coefficient_method, coefficient_method)

    warning = ""
    if coefficient_method == "default":
        warning = (
            '<p class="notice">\u26a0 \u041a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442 '
            '\u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u0432 \u0444\u0430\u0439\u043b\u0435 '
            '(\u043d\u0435\u0442 \u043f\u043e\u0434\u043f\u0438\u0441\u0438 \u00ab\u0420\u0435\u0433\u0438\u043e\u043d:\u00bb/'
            '\u00ab\u041a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442:\u00bb \u0438\u043b\u0438 '
            '\u043e\u043d\u0430 \u043d\u0435 \u0440\u0430\u0441\u043f\u043e\u0437\u043d\u0430\u043d\u0430). '
            '\u041f\u0440\u0438\u043c\u0435\u043d\u0451\u043d 1.0 \u043f\u043e \u0443\u043c\u043e\u043b\u0447\u0430\u043d\u0438\u044e '
            '\u2014 \u043f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0438 \u0432\u0432\u0435\u0434\u0438\u0442\u0435 '
            '\u043f\u0440\u0430\u0432\u0438\u043b\u044c\u043d\u043e\u0435 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u043d\u0438\u0436\u0435.</p>'
        )
    if error:
        warning += f'<p class="notice">{html.escape(error)}</p>'

    return render(
        "confirm.html",
        token=html.escape(token),
        sheet=html.escape(sheet_title),
        sheet_display=html.escape(sheet_title),
        region_value=html.escape(region_value),
        region_source=region_source,
        coefficient_value=html.escape(
            coefficient_value if isinstance(coefficient_value, str) else f"{coefficient_value:g}"
        ),
        coefficient_source=coefficient_source,
        warning=warning,
    )


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

def render_admin_catalog(
    catalog_page: CatalogEditorPage,
    *,
    notice: str = "",
    error: str = "",
    return_url: str = "/admin/catalog",
) -> str:
    notice_html = f'<p class="notice-ok">{html.escape(notice)}</p>' if notice else ""
    error_html = f'<p class="notice">{html.escape(error)}</p>' if error else ""
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Каталог аналогов</h2>'
        '<p>Здесь можно смотреть актуальные строки catalog_items, фильтровать их, править значения, удалять строки и применять групповые действия к выделенным строкам.</p>'
        '<p class="catalog-warning">Перед массовыми изменениями сделайте backup базы. Matching/pricing использует эти строки как источник аналогов.</p>'
        f'{notice_html}'
        f'{error_html}'
        f'{_render_catalog_editor_filters(catalog_page)}'
        f'{_render_catalog_editor_pager(catalog_page)}'
        f'{_render_catalog_editor_table(catalog_page, return_url)}'
        f'{_render_catalog_editor_pager(catalog_page)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Каталог аналогов",
        subtitle="Редактор строк базы аналогов.",
        admin_nav=_render_admin_nav(active_slug="catalog"),
        content=content,
    )


def _render_catalog_clear_form(total_rows: int) -> str:
    return (
        '<details class="maintenance-tools">'
        '<summary>Служебные действия</summary>'
        '<div class="maintenance-tools-body">'
        '<p class="muted">Полная очистка нужна только перед пересборкой каталога. '
        f'Сейчас строк: {total_rows}.</p>'
        '<form action="/admin/catalog/clear" method="post" '
        'onsubmit="return prompt(\'Для подтверждения введите ОЧИСТИТЬ\') === \'ОЧИСТИТЬ\';">'
        '<input type="hidden" name="confirmation" value="clear_catalog">'
        '<button class="danger-small" type="submit">Очистить каталог</button>'
        '</form>'
        '</div>'
        '</details>'
    )


def _render_catalog_editor_filters(catalog_page: CatalogEditorPage) -> str:
    filters = catalog_page.filters
    def val(name: str) -> str:
        return html.escape(filters.get(name, ""), quote=True)

    return (
        '<form class="admin-form" action="/admin/catalog" method="get">'
        '<h2 class="section">Фильтры</h2>'
        '<div class="catalog-filter-grid">'
        f'<label>Общий поиск<input type="text" name="q" value="{val("q")}" placeholder="работа, код, файл, регион"></label>'
        f'<label>Источник<input type="text" name="source" value="{val("source")}" placeholder="rnmc_zip_upload"></label>'
        f'<label>Регион<input type="text" name="region" value="{val("region")}"></label>'
        f'<label>Задача<input type="text" name="task_id" value="{val("task_id")}"></label>'
        f'<label>Код<input type="text" name="code" value="{val("code")}"></label>'
        f'<label>Ед.<input type="text" name="unit" value="{val("unit")}"></label>'
        f'<label>Файл<input type="text" name="filename" value="{val("filename")}"></label>'
        f'<label>Строк на странице<select name="page_size">{_catalog_page_size_options(catalog_page.page_size)}</select></label>'
        '</div>'
        '<button type="submit">Применить фильтры</button>'
        '<a class="back" href="/admin/catalog">Очистить фильтры</a>'
        f'<a class="back" href="{html.escape(_catalog_export_url(filters))}">⬇ Экспорт в Excel (с текущими фильтрами)</a>'
        '</form>'
    )


def _catalog_export_url(filters: dict[str, str]) -> str:
    params = {key: value for key, value in filters.items() if value}
    return "/admin/catalog/export" + (f"?{urlencode(params)}" if params else "")


def _catalog_page_size_options(current: int) -> str:
    options = []
    for value in [50, 100, 200, 500]:
        selected = " selected" if int(current) == value else ""
        options.append(f'<option value="{value}"{selected}>{value}</option>')
    return "".join(options)


def _render_catalog_editor_pager(catalog_page: CatalogEditorPage) -> str:
    parts = [
        f'<strong>Всего строк: {catalog_page.total_rows}</strong>',
        f'<span>Страница {catalog_page.page} из {catalog_page.total_pages}</span>',
    ]
    if catalog_page.page > 1:
        parts.append(f'<a href="{html.escape(_catalog_editor_url(catalog_page, catalog_page.page - 1))}">← Назад</a>')
    if catalog_page.page < catalog_page.total_pages:
        parts.append(f'<a href="{html.escape(_catalog_editor_url(catalog_page, catalog_page.page + 1))}">Вперед →</a>')
    return '<div class="catalog-pager">' + "".join(parts) + '</div>'


def _catalog_editor_url(catalog_page: CatalogEditorPage, page: int) -> str:
    params = {key: value for key, value in catalog_page.filters.items() if value}
    params["page"] = str(page)
    params["page_size"] = str(catalog_page.page_size)
    return "/admin/catalog?" + urlencode(params)


def _render_catalog_editor_table(catalog_page: CatalogEditorPage, return_url: str) -> str:
    if not catalog_page.rows:
        return '<p class="muted">Строки каталога по заданным фильтрам не найдены.</p>'

    header = (
        '<form class="catalog-editor-form" method="post">'
        f'<input type="hidden" name="return_url" value="{html.escape(return_url, quote=True)}">'
        '<div class="catalog-actions">'
        '<label><input type="checkbox" onclick="document.querySelectorAll(\'.catalog-row-check\').forEach(cb => cb.checked = this.checked)"> Выделить все на странице</label>'
        '<label>Действие<select name="bulk_action"><option value="update">Изменить поле</option><option value="delete">Удалить выделенные</option></select></label>'
        f'<label>Поле<select name="bulk_field">{_catalog_bulk_field_options()}</select></label>'
        '<label>Операция<select name="bulk_operation"><option value="set">Заменить</option><option value="add">Прибавить</option><option value="multiply">Умножить</option></select></label>'
        '<label>Значение<input type="text" name="bulk_value" placeholder="например 1,2"></label>'
        '<button type="submit" formaction="/admin/catalog/bulk" onclick="return prepareCatalogBulkAction(this.form)">Применить к выделенным</button>'
        '</div>'
        f'{_render_catalog_bulk_confirm_script()}'
        '<div class="preview-wide">'
        '<table class="preview catalog-table"><thead><tr>'
        '<th></th><th>ID</th><th>Источник</th><th>Регион</th><th>Задача</th><th>Код</th><th>Ед.</th>'
        '<th>Кол-во</th><th>Цена раб.</th><th>Цена ориг.</th><th>Цена ZLVL</th><th>Итого</th><th>ТЗ ед.</th><th>ТЗ всего</th>'
        '<th>ТЗм ед.</th><th>ТЗм всего</th><th>Коэф.</th><th>ЛСР</th><th>Начало</th><th>Окончание</th><th class="wrap">Работа</th>'
        '<th>Папка</th><th>Файл</th><th>Excel row</th><th>Действия</th>'
        '</tr></thead><tbody>'
    )
    rows = "".join(_render_catalog_editor_row(row) for row in catalog_page.rows)
    return header + rows + '</tbody></table></div></form>'



def _render_catalog_bulk_confirm_script() -> str:
    return """
<script>
function prepareCatalogBulkAction(form) {
  if (!confirmCatalogBulkAction(form)) {
    return false;
  }
  form.querySelectorAll('tbody input:not(.catalog-row-check), tbody textarea').forEach(field => {
    field.disabled = true;
  });
  return true;
}

function confirmCatalogBulkAction(form) {
  const checked = Array.from(form.querySelectorAll('.catalog-row-check:checked'));
  if (checked.length === 0) {
    alert('Выберите хотя бы одну строку каталога.');
    return false;
  }
  if (checked.length < 2) {
    return true;
  }
  const actionSelect = form.querySelector('[name="bulk_action"]');
  const fieldSelect = form.querySelector('[name="bulk_field"]');
  const operationSelect = form.querySelector('[name="bulk_operation"]');
  const valueInput = form.querySelector('[name="bulk_value"]');
  const action = actionSelect ? actionSelect.value : '';
  const actionText = action === 'delete' ? 'удалить' : 'изменить';
  const fieldText = fieldSelect && fieldSelect.selectedOptions.length ? fieldSelect.selectedOptions[0].text : '';
  const operationText = operationSelect && operationSelect.selectedOptions.length ? operationSelect.selectedOptions[0].text : '';
  const valueText = valueInput && valueInput.value ? valueInput.value : '';
  let message = 'Вы уверены, что хотите ' + actionText + ' ' + checked.length + ' строк каталога?';
  if (action !== 'delete') {
    message += '\nПоле: ' + fieldText + '\nОперация: ' + operationText + '\nЗначение: ' + valueText;
  }
  message += '\n\nДействие нельзя отменить автоматически. Перед массовыми изменениями желательно иметь backup базы.';
  return confirm(message);
}
</script>
"""

def _catalog_bulk_field_options() -> str:
    options = [
        ("region", "Регион"),
        ("task_id", "Задача"),
        ("code", "Код"),
        ("unit", "Ед."),
        ("work_name", "Работа"),
        ("quantity", "Кол-во"),
        ("price", "Цена рабочая"),
        ("price_original", "Цена оригинал"),
        ("price_zlvl", "Цена ZLVL"),
        ("total_price", "Итого"),
        ("labor_unit", "ТЗ ед."),
        ("labor_total", "ТЗ всего"),
        ("machine_labor_unit", "ТЗм ед."),
        ("machine_labor_total", "ТЗм всего"),
        ("regional_coefficient", "Коэф."),
        ("lsr_quarter", "ЛСР"),
        ("planned_start", "Начало"),
        ("planned_finish", "Окончание"),
    ]
    return "".join(f'<option value="{html.escape(value)}">{html.escape(label)}</option>' for value, label in options)


def _render_catalog_editor_row(row: CatalogEditorRow) -> str:
    item_id = row.id
    return (
        '<tr>'
        f'<td><input class="catalog-row-check" type="checkbox" name="selected_ids" value="{item_id}"></td>'
        f'<td>{item_id}</td>'
        f'<td title="{html.escape(row.source_kind)}">{html.escape(row.source_name)}</td>'
        f'<td class="short-cell">{_catalog_text_input("region", item_id, row.region)}</td>'
        f'<td class="short-cell">{_catalog_text_input("task_id", item_id, row.task_id)}</td>'
        f'<td class="short-cell">{_catalog_text_input("code", item_id, row.code)}</td>'
        f'<td class="short-cell">{_catalog_text_input("unit", item_id, row.unit)}</td>'
        f'<td class="number-cell">{_catalog_text_input("quantity", item_id, _format_optional_number(row.quantity))}</td>'
        f'<td class="number-cell">{_catalog_text_input("price", item_id, _format_optional_number(row.price))}</td>'
        f'<td class="number-cell">{_catalog_text_input("price_original", item_id, _format_optional_number(row.price_original))}</td>'
        f'<td class="number-cell">{_catalog_text_input("price_zlvl", item_id, _format_optional_number(row.price_zlvl))}</td>'
        f'<td class="number-cell">{_catalog_text_input("total_price", item_id, _format_optional_number(row.total_price))}</td>'
        f'<td class="number-cell">{_catalog_text_input("labor_unit", item_id, _format_optional_number(row.labor_unit))}</td>'
        f'<td class="number-cell">{_catalog_text_input("labor_total", item_id, _format_optional_number(row.labor_total))}</td>'
        f'<td class="number-cell">{_catalog_text_input("machine_labor_unit", item_id, _format_optional_number(row.machine_labor_unit))}</td>'
        f'<td class="number-cell">{_catalog_text_input("machine_labor_total", item_id, _format_optional_number(row.machine_labor_total))}</td>'
        f'<td class="number-cell">{_catalog_text_input("regional_coefficient", item_id, _format_optional_number(row.regional_coefficient))}</td>'
        f'<td class="short-cell">{_catalog_text_input("lsr_quarter", item_id, row.lsr_quarter)}</td>'
        f'<td class="short-cell">{_catalog_text_input("planned_start", item_id, row.planned_start)}</td>'
        f'<td class="short-cell">{_catalog_text_input("planned_finish", item_id, row.planned_finish)}</td>'
        f'<td class="work-cell"><textarea name="work_name_{item_id}">{html.escape(row.work_name)}</textarea></td>'
        f'<td class="file-cell">{_catalog_text_input("source_region_folder", item_id, row.source_region_folder)}</td>'
        f'<td class="file-cell">{_catalog_text_input("source_filename", item_id, row.source_filename)}</td>'
        f'<td class="number-cell">{_catalog_text_input("source_row_number", item_id, str(row.source_row_number))}</td>'
        '<td>'
        f'<button type="submit" name="row_id" value="{item_id}" formaction="/admin/catalog/update-row">Сохранить</button>'
        f'<button class="danger" type="submit" name="row_id" value="{item_id}" formaction="/admin/catalog/delete-row" onclick="return confirm(\'Удалить строку каталога?\')">Удалить</button>'
        '</td>'
        '</tr>'
    )


def _catalog_text_input(name: str, item_id: int, value: object) -> str:
    return f'<input type="text" name="{html.escape(name)}_{item_id}" value="{html.escape(str(value), quote=True)}">'


def _format_optional_number(value: float | None) -> str:
    if value is None:
        return ""
    text = f"{float(value):.10f}".rstrip("0").rstrip(".")
    return text if text else "0"


def render_admin_imports(
    imports: list[ImportedFileRecord],
    *,
    notice: str = "",
    error: str = "",
    dry_run_result: RnmcZipDryRunResult | None = None,
    row_preview_result: RnmcZipRowPreviewResult | None = None,
    catalog_import_result: RnmcZipCatalogImportResult | None = None,
    status_filter: str = "",
    stage_token: str = "",
    stage_filename: str = "",
    stage_mode: str = "",
    legacy_catalog_preview: dict | None = None,
) -> str:
    notice_html = f'<p class="notice-ok">{html.escape(notice)}</p>' if notice else ""
    error_html = f'<p class="notice">{html.escape(error)}</p>' if error else ""
    content = (
        '<section class="admin-panel">'
        '<div class="admin-section-head">'
        '<h2 class="section">Загрузка новых РНМЦ</h2>'
        '<p class="muted">Один сценарий: выберите ZIP, проверьте найденные данные и подтвердите сохранение.</p>'
        '</div>'
        f'{notice_html}{error_html}'
        '<div class="admin-action-grid">'
        f'{_render_rnmc_zip_row_preview_form()}'
        f'{_render_rnmc_single_file_preview_form()}'
        '<div class="admin-form admin-action-card">'
        '<h2 class="section">Как проходит импорт</h2>'
        '<ol class="admin-step-list">'
        '<li>Система проверяет файлы, дубликаты и метаданные.</li>'
        '<li>Показывает сводку, статусы, заголовки и реальные строки.</li>'
        '<li>После подтверждения сохраняет каталог и автоматически обновляет журнал.</li>'
        '</ol>'
        '</div></div>'
        f'{_render_rnmc_zip_dry_run_result(dry_run_result)}'
        f'{_render_rnmc_zip_row_preview_result(row_preview_result)}'
        f'{_render_legacy_catalog_preview(legacy_catalog_preview)}'
        f'{_render_rnmc_stage_commit(stage_token, stage_filename, stage_mode, row_preview_result)}'
        f'{_render_rnmc_zip_catalog_import_result(catalog_import_result)}'
        '<div class="admin-section-head"><h2 class="section">История импортов</h2>'
        '<p class="muted">Журнал обновляется автоматически после сохранения ZIP.</p></div>'
        f'{_render_import_status_filters(status_filter)}'
        f'{_render_imported_file_table(imports)}'
        '</section>'
    )
    return render(
        "admin.html",
        title="Импорты файлов",
        subtitle="Проверка, предпросмотр и сохранение новых РНМЦ.",
        admin_nav=_render_admin_nav(active_slug="imports"),
        content=content,
    )



def _render_legacy_catalog_preview(preview: dict | None) -> str:
    if preview is None:
        return ""
    total_rows = int(preview.get("total_rows", 0))
    positioned_rows = preview.get("rows", [])
    rows_html = []
    for row_number, row in positioned_rows:
        rows_html.append(
            '<tr>'
            f'<td>{row_number}</td>'
            f'<td>{html.escape(str(row.task_id or ""))}</td>'
            f'<td>{html.escape(str(row.region or ""))}</td>'
            f'<td>{html.escape(str(row.code or ""))}</td>'
            f'<td>{html.escape(str(row.work_name or ""))}</td>'
            f'<td>{html.escape(str(row.unit or ""))}</td>'
            f'<td>{html.escape(str(row.quantity or ""))}</td>'
            f'<td>{html.escape(str(row.price_original or ""))}</td>'
            f'<td>{html.escape(str(row.price_zlvl or row.price or ""))}</td>'
            f'<td>{html.escape(str(row.source_filename or ""))}</td>'
            '</tr>'
        )
    table = (
        '<div class="preview-wide"><table class="preview"><thead><tr>'
        '<th>Строка</th><th>Задача</th><th>Регион</th><th>Код</th><th>Наименование</th>'
        '<th>Ед.</th><th>Кол-во</th><th>Цена исходная</th><th>Цена ZLVL</th><th>source_file</th>'
        '</tr></thead><tbody>' + ''.join(rows_html) + '</tbody></table></div>'
    )
    return (
        '<div class="admin-form">'
        '<h2 class="section">Предпросмотр полного каталога</h2>'
        f'<dl class="stats"><div><dt>Строк распознано</dt><dd>{total_rows}</dd></div>'
        f'<div><dt>Показано в предпросмотре</dt><dd>{len(positioned_rows)}</dd></div></dl>'
        '<p class="muted">Показаны первые 30 строк. После подтверждения импортируются все распознанные строки файла.</p>'
        f'{table}</div>'
    )

def _render_rnmc_stage_commit(
    stage_token: str,
    stage_filename: str,
    stage_mode: str = "",
    row_preview_result: RnmcZipRowPreviewResult | None = None,
) -> str:
    if not stage_token:
        return ""
    object_label = 'Файл' if stage_mode else 'Архив'
    mode_note = (
        ' Будет выполнена полная загрузка старого ZLVL-каталога.'
        if stage_mode == 'legacy_catalog'
        else ''
    )
    coefficient_note = ""
    if stage_mode != "legacy_catalog" and row_preview_result is not None:
        editable_count = sum(
            1 for entry in row_preview_result.entries
            if _rnmc_coefficient_is_editable(entry)
        )
        if editable_count:
            coefficient_note = (
                '<p class="muted">Перед сохранением проверьте колонку <strong>Коэф.</strong> '
                'на вкладке <strong>Метаданные</strong>. Значение из этой колонки будет применено '
                'к строкам файла: price_zlvl = price_original / коэффициент.</p>'
            )
    return (
        '<div class="admin-confirm-box">'
        '<h2 class="section">Предпросмотр готов</h2>'
        f'<p>{object_label} <strong>{html.escape(stage_filename)}</strong> проверен.{mode_note} Сохраните данные, если всё верно.</p>'
        f'{coefficient_note}'
        '<form id="rnmc-stage-commit-form" action="/admin/imports/rnmc-stage-commit" method="post">'
        f'<input type="hidden" name="stage_token" value="{html.escape(stage_token, quote=True)}">'
        '<button type="submit">Сохранить в каталог</button>'
        '</form></div>'
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
        '<form class="admin-form admin-action-card" action="/admin/imports/rnmc-row-preview" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Выбрать ZIP</h2>'
        '<p class="muted">На этом шаге данные только проверяются. Запись в каталог начнется после отдельного подтверждения.</p>'
        '<label>ZIP-архив РНМЦ<input type="file" name="rnmc_zip" accept=".zip" required></label>'
        '<label>Регион вручную, если нужно<input type="text" name="region_override" placeholder="Оставьте пустым, чтобы взять регион из папки"></label>'
        '<button class="admin-primary-action" type="submit">Загрузить и проверить</button>'
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


def _render_rnmc_single_file_preview_form() -> str:
    return (
        '<form class="admin-form admin-action-card" action="/admin/imports/rnmc-file-preview" method="post" enctype="multipart/form-data">'
        '<h2 class="section">Выбрать Excel-файл</h2>'
        '<p class="muted">Одиночная новая РНМЦ обрабатывается по обычным правилам. Файл РНМЦ_КА_ЖО_ZLVL_V3.xlsx распознается как полный старый каталог.</p>'
        '<label>Excel-файл РНМЦ<input type="file" name="rnmc_file" accept=".xlsx,.xlsm" required></label>'
        '<label>Регион вручную, если нужно<input type="text" name="region_override" placeholder="Для новой РНМЦ; каталог берет регион из строк"></label>'
        '<button class="admin-primary-action" type="submit">Загрузить и проверить</button>'
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
    summary += (
        f'<p class="muted">Предпросмотр читает не больше {DEFAULT_ROW_PREVIEW_LIMIT} строк на один Excel-файл. '
        'Реальный импорт по-прежнему читает все строки файла.</p>'
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


def _attr(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def _flag(value: bool) -> str:
    return "1" if value else "0"


def _rnmc_filter_toolbar(table_id: str, *, status_filter: bool = True, empty_filter: bool = True) -> str:
    status = (
        f'<label>Статус<select data-rnmc-filter-status="{table_id}">'
        '<option value="">Все статусы</option>'
        '</select></label>'
        if status_filter else ""
    )
    hide_empty = (
        f'<label class="rnmc-toggle"><input type="checkbox" data-rnmc-hide-empty="{table_id}"> Скрыть пустые строки</label>'
        if empty_filter else ""
    )
    return (
        f'<div class="rnmc-preview-toolbar" data-rnmc-toolbar="{table_id}">'
        f'<label>Поиск<input type="search" data-rnmc-search="{table_id}" placeholder="Файл, код, регион, причина..."></label>'
        f'{status}'
        f'<label class="rnmc-toggle"><input type="checkbox" data-rnmc-hide-processed="{table_id}"> Скрыть ранее загруженные</label>'
        f'<label class="rnmc-toggle"><input type="checkbox" data-rnmc-only-problems="{table_id}"> Только проблемные</label>'
        f'{hide_empty}'
        f'<button type="button" data-rnmc-clear="{table_id}">Очистить фильтры</button>'
        f'<span class="rnmc-filter-count" data-rnmc-count="{table_id}"></span>'
        '</div>'
    )


def _rnmc_zoom_toolbar() -> str:
    return (
        '<div class="rnmc-preview-toolbar">'
        '<label>Масштаб таблиц<select data-rnmc-zoom>'
        '<option value="0.85">85%</option>'
        '<option value="1" selected>100%</option>'
        '<option value="1.15">115%</option>'
        '<option value="1.3">130%</option>'
        '</select></label>'
        '<label>Плотность<select data-rnmc-density>'
        '<option value="compact">Компактно</option>'
        '<option value="normal" selected>Обычно</option>'
        '<option value="large">Крупно</option>'
        '</select></label>'
        '<button type="button" data-rnmc-reset-widths>Сбросить ширины</button>'
        '<span class="muted">Масштаб меняет только таблицы предпросмотра. Ширину столбцов можно менять за синюю границу заголовка.</span>'
        '</div>'
    )


def _rnmc_preview_script() -> str:
    return """
<script>
(function () {
  function text(row) { return (row.textContent || '').toLowerCase(); }
  function tableById(id) { return document.querySelector('[data-rnmc-table="' + id + '"]'); }
  function rowsFor(id) {
    var table = tableById(id);
    if (!table) return [];
    return Array.prototype.slice.call(table.querySelectorAll('tbody tr[data-rnmc-row="1"]'));
  }
  function fillStatuses(id) {
    var select = document.querySelector('[data-rnmc-filter-status="' + id + '"]');
    if (!select || select.dataset.ready === '1') return;
    var values = [];
    rowsFor(id).forEach(function (row) {
      var status = row.dataset.status || '';
      if (status && values.indexOf(status) === -1) values.push(status);
    });
    values.sort().forEach(function (status) {
      var option = document.createElement('option');
      option.value = status;
      option.textContent = status;
      select.appendChild(option);
    });
    select.dataset.ready = '1';
  }
  function applyFilter(id) {
    fillStatuses(id);
    var query = (document.querySelector('[data-rnmc-search="' + id + '"]') || {}).value || '';
    query = query.toLowerCase().trim();
    var status = (document.querySelector('[data-rnmc-filter-status="' + id + '"]') || {}).value || '';
    var hideProcessed = !!(document.querySelector('[data-rnmc-hide-processed="' + id + '"]') || {}).checked;
    var onlyProblems = !!(document.querySelector('[data-rnmc-only-problems="' + id + '"]') || {}).checked;
    var hideEmpty = !!(document.querySelector('[data-rnmc-hide-empty="' + id + '"]') || {}).checked;
    var shown = 0;
    var total = 0;
    rowsFor(id).forEach(function (row) {
      total += 1;
      var ok = true;
      if (query && text(row).indexOf(query) === -1) ok = false;
      if (status && row.dataset.status !== status) ok = false;
      if (hideProcessed && row.dataset.processed === '1') ok = false;
      if (onlyProblems && row.dataset.problem !== '1') ok = false;
      if (hideEmpty && row.dataset.empty === '1') ok = false;
      row.classList.toggle('rnmc-hidden-row', !ok);
      if (ok) shown += 1;
    });
    var count = document.querySelector('[data-rnmc-count="' + id + '"]');
    if (count) count.textContent = 'Показано: ' + shown + ' из ' + total;
  }
  function clearFilter(id) {
    var search = document.querySelector('[data-rnmc-search="' + id + '"]');
    var status = document.querySelector('[data-rnmc-filter-status="' + id + '"]');
    var hideProcessed = document.querySelector('[data-rnmc-hide-processed="' + id + '"]');
    var onlyProblems = document.querySelector('[data-rnmc-only-problems="' + id + '"]');
    var hideEmpty = document.querySelector('[data-rnmc-hide-empty="' + id + '"]');
    if (search) search.value = '';
    if (status) status.value = '';
    if (hideProcessed) hideProcessed.checked = false;
    if (onlyProblems) onlyProblems.checked = false;
    if (hideEmpty) hideEmpty.checked = false;
    applyFilter(id);
  }
  function setZoom() {
    var shell = document.querySelector('.rnmc-preview-shell');
    if (!shell) return;
    var zoom = (document.querySelector('[data-rnmc-zoom]') || {}).value || '1';
    var density = (document.querySelector('[data-rnmc-density]') || {}).value || 'normal';
    shell.style.setProperty('--rnmc-preview-scale', zoom);
    shell.classList.toggle('compact', density === 'compact');
    shell.classList.toggle('large', density === 'large');
    window.setTimeout(initResizableTables, 0);
  }
  function widthKey(table) {
    return 'estimate-ai:rnmc-preview-widths:' + (table.getAttribute('data-rnmc-table') || 'table');
  }
  function tableColumns(table) {
    return Array.prototype.slice.call(table.querySelectorAll('colgroup.rnmc-colgroup col'));
  }
  function saveWidths(table) {
    var widths = tableColumns(table).map(function (col) { return parseInt(col.style.width || '0', 10) || 0; });
    try { window.localStorage.setItem(widthKey(table), JSON.stringify(widths)); } catch (error) {}
  }
  function setTableWidth(table) {
    var sum = tableColumns(table).reduce(function (total, col) {
      return total + (parseInt(col.style.width || '0', 10) || 0);
    }, 0);
    if (sum > 0) table.style.width = sum + 'px';
  }
  function initResizableTable(table) {
    if (!table || table.dataset.resizableReady === '1') return;
    var headers = Array.prototype.slice.call(table.querySelectorAll('thead th'));
    if (!headers.length) return;
    var stored = [];
    try { stored = JSON.parse(window.localStorage.getItem(widthKey(table)) || '[]') || []; } catch (error) { stored = []; }
    var group = document.createElement('colgroup');
    group.className = 'rnmc-colgroup';
    table.insertBefore(group, table.firstChild);
    headers.forEach(function (header, index) {
      var width = parseInt(stored[index] || '0', 10);
      if (!width) width = Math.max(70, Math.ceil(header.getBoundingClientRect().width || header.scrollWidth || 120));
      var col = document.createElement('col');
      col.style.width = width + 'px';
      group.appendChild(col);
      var handle = document.createElement('span');
      handle.className = 'rnmc-col-resizer';
      handle.setAttribute('title', 'Потяните, чтобы изменить ширину столбца');
      handle.setAttribute('data-rnmc-col-index', String(index));
      header.appendChild(handle);
    });
    setTableWidth(table);
    table.dataset.resizableReady = '1';
  }
  function initResizableTables() {
    document.querySelectorAll('table[data-rnmc-table]').forEach(initResizableTable);
  }
  function resetResizableTables() {
    document.querySelectorAll('table[data-rnmc-table]').forEach(function (table) {
      try { window.localStorage.removeItem(widthKey(table)); } catch (error) {}
      table.dataset.resizableReady = '0';
      var group = table.querySelector('colgroup.rnmc-colgroup');
      if (group) group.remove();
      table.querySelectorAll('.rnmc-col-resizer').forEach(function (handle) { handle.remove(); });
      table.style.width = '';
      initResizableTable(table);
    });
  }
  var resizeState = null;
  document.addEventListener('pointerdown', function (event) {
    var handle = event.target.closest ? event.target.closest('.rnmc-col-resizer') : null;
    if (!handle) return;
    var table = handle.closest('table[data-rnmc-table]');
    if (!table) return;
    var col = tableColumns(table)[parseInt(handle.getAttribute('data-rnmc-col-index') || '0', 10)];
    if (!col) return;
    resizeState = {
      table: table,
      col: col,
      handle: handle,
      startX: event.clientX,
      startWidth: parseInt(col.style.width || '0', 10) || handle.parentElement.getBoundingClientRect().width
    };
    handle.classList.add('active');
    document.body.classList.add('rnmc-resizing');
    if (handle.setPointerCapture) handle.setPointerCapture(event.pointerId);
    event.preventDefault();
  });
  document.addEventListener('pointermove', function (event) {
    if (!resizeState) return;
    var next = Math.max(60, Math.round(resizeState.startWidth + event.clientX - resizeState.startX));
    resizeState.col.style.width = next + 'px';
    setTableWidth(resizeState.table);
    event.preventDefault();
  });
  document.addEventListener('pointerup', function () {
    if (!resizeState) return;
    saveWidths(resizeState.table);
    resizeState.handle.classList.remove('active');
    document.body.classList.remove('rnmc-resizing');
    resizeState = null;
  });
  document.addEventListener('input', function (event) {
    var id = event.target.getAttribute('data-rnmc-search');
    if (id) applyFilter(id);
  });
  document.addEventListener('change', function (event) {
    var attrs = ['data-rnmc-filter-status', 'data-rnmc-hide-processed', 'data-rnmc-only-problems', 'data-rnmc-hide-empty'];
    for (var i = 0; i < attrs.length; i++) {
      var id = event.target.getAttribute(attrs[i]);
      if (id) { applyFilter(id); return; }
    }
    if (event.target.hasAttribute('data-rnmc-zoom') || event.target.hasAttribute('data-rnmc-density')) setZoom();
  });
  document.addEventListener('click', function (event) {
    var id = event.target.getAttribute('data-rnmc-clear');
    if (id) clearFilter(id);
    if (event.target.hasAttribute('data-rnmc-reset-widths')) resetResizableTables();
  });
  document.querySelectorAll('[data-rnmc-table]').forEach(function (table) {
    var id = table.getAttribute('data-rnmc-table');
    fillStatuses(id);
    applyFilter(id);
  });
  setZoom();
})();
</script>
"""

def _render_rnmc_zip_row_preview_result(result: RnmcZipRowPreviewResult | None) -> str:
    if result is None:
        return ""

    summary_stats = (
        '<dl class="stats">'
        f'<div><dt>Excel-файлов найдено</dt><dd>{result.total_excel_files}</dd></div>'
        f'<div><dt>Файлов с найденными строками</dt><dd>{result.preview_ok_count}</dd></div>'
        f'<div><dt>OK-строк в preview</dt><dd>{result.rows_ok_total}</dd></div>'
        f'<div><dt>Rejected-строк в preview</dt><dd>{result.rows_rejected_total}</dd></div>'
        f'<div><dt>Ограничено лимитом</dt><dd>{result.limited_count}</dd></div>'
        f'<div><dt>Уже обработано</dt><dd>{result.skipped_processed_count}</dd></div>'
        f'<div><dt>Дубликаты имени</dt><dd>{result.duplicate_name_count}</dd></div>'
        f'<div><dt>Без таблицы</dt><dd>{result.no_table_count}</dd></div>'
        f'<div><dt>Без строк</dt><dd>{result.no_rows_count}</dd></div>'
        f'<div><dt>Неподдерживаемый формат</dt><dd>{result.unsupported_format_count}</dd></div>'
        f'<div><dt>Ошибки чтения</dt><dd>{result.parse_error_count}</dd></div>'
        f'<div><dt>Прочих файлов проигнорировано</dt><dd>{result.ignored_files}</dd></div>'
        '</dl>'
    )
    note = (
        f'<p class="muted">Предпросмотр показывает заголовки и первые {DEFAULT_ROW_PREVIEW_LIMIT} '
        'реальных строк тела таблицы на один Excel-файл. Пустые технические строки до начала таблицы '
        'не расходуют лимит. Реальный импорт по-прежнему читает все строки файла.</p>'
    )
    if not result.entries:
        return (
            '<div class="admin-form">'
            '<h2 class="section">Предпросмотр строк РНМЦ</h2>'
            f'{summary_stats}{note}'
            '<p class="muted">В ZIP не найдено Excel-файлов РНМЦ.</p></div>'
        )

    tabs = (
        '<div class="rnmc-preview-shell">'
        f'{_rnmc_zoom_toolbar()}'
        '<div class="rnmc-tabs">'
        '<input class="rnmc-tab-input" type="radio" name="rnmc-preview-tabs" id="rnmc-preview-tab-summary" checked>'
        '<input class="rnmc-tab-input" type="radio" name="rnmc-preview-tabs" id="rnmc-preview-tab-files">'
        '<input class="rnmc-tab-input" type="radio" name="rnmc-preview-tabs" id="rnmc-preview-tab-metadata">'
        '<input class="rnmc-tab-input" type="radio" name="rnmc-preview-tabs" id="rnmc-preview-tab-headers">'
        '<input class="rnmc-tab-input" type="radio" name="rnmc-preview-tabs" id="rnmc-preview-tab-rows">'
        '<div class="rnmc-tab-labels">'
        '<label for="rnmc-preview-tab-summary">Сводка</label>'
        '<label for="rnmc-preview-tab-files">Файлы и статусы</label>'
        '<label for="rnmc-preview-tab-metadata">Метаданные</label>'
        '<label for="rnmc-preview-tab-headers">Заголовки</label>'
        '<label for="rnmc-preview-tab-rows">Строки предпросмотра</label>'
        '</div>'
        '<div class="rnmc-tab-panels">'
        f'<section class="rnmc-tab-panel" id="rnmc-preview-panel-summary">{summary_stats}{note}</section>'
        f'<section class="rnmc-tab-panel" id="rnmc-preview-panel-files">{_render_rnmc_preview_file_table(result.entries)}</section>'
        f'<section class="rnmc-tab-panel" id="rnmc-preview-panel-metadata">{_render_rnmc_preview_metadata_table(result.entries)}</section>'
        f'<section class="rnmc-tab-panel" id="rnmc-preview-panel-headers">{_render_rnmc_preview_header_table(result.entries)}</section>'
        f'<section class="rnmc-tab-panel" id="rnmc-preview-panel-rows">{_render_rnmc_preview_rows_table(result.entries)}</section>'
        '</div></div>'
        '</div>'
        f'{_rnmc_preview_script()}'
    )
    return (
        '<div class="admin-form">'
        '<h2 class="section">Предпросмотр строк РНМЦ</h2>'
        f'{tabs}'
        '</div>'
    )



def _render_rnmc_preview_file_table(entries) -> str:
    table_id = "rnmc-files"
    header = (
        f'{_rnmc_filter_toolbar(table_id)}'
        f'<div class="preview-wide"><table class="preview" data-rnmc-table="{table_id}"><thead><tr>'
        '<th>Файл</th><th>Регион</th><th>Статус</th><th>Причина</th>'
        '<th>Лист</th><th>Header row</th><th>Задача</th><th>OK</th><th>Rejected</th><th>Лимит</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in entries:
        status_class = _preview_status_class(entry.status)
        processed = entry.status == "skipped_processed"
        problem = entry.status != "preview_ok" or bool(entry.reason)
        empty = entry.rows_ok <= 0 and entry.rows_rejected <= 0
        rows.append(
            '<tr data-rnmc-row="1" '
            f'data-status="{_attr(entry.status)}" data-processed="{_flag(processed)}" '
            f'data-problem="{_flag(problem)}" data-empty="{_flag(empty)}">'
            f'<td class="path" title="{_attr(entry.archive_path)}">{html.escape(entry.filename)}</td>'
            f'<td>{html.escape(entry.region_folder)}</td>'
            f'<td><span class="status-pill {status_class}">{html.escape(entry.status)}</span></td>'
            f'<td class="wrap">{html.escape(entry.reason)}</td>'
            f'<td>{html.escape(entry.sheet_name)}</td>'
            f'<td>{entry.header_row or ""}</td>'
            f'<td>{html.escape(entry.task_number)}</td>'
            f'<td>{entry.rows_ok}</td>'
            f'<td>{entry.rows_rejected}</td>'
            f'<td>{"да" if entry.is_limited else ""}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table></div>'


def _render_rnmc_preview_metadata_table(entries) -> str:
    table_id = "rnmc-metadata"
    header = (
        f'{_rnmc_filter_toolbar(table_id)}'
        f'<div class="preview-wide"><table class="preview" data-rnmc-table="{table_id}"><thead><tr>'
        '<th>Файл</th><th>Регион</th><th>Коэф.</th><th>ЛСР</th><th>Начало</th><th>Окончание</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in entries:
        empty = not any([entry.region_folder, entry.regional_coefficient, entry.lsr_quarter, entry.planned_start, entry.planned_finish])
        problem = entry.status != "preview_ok" or not entry.region_folder or not entry.lsr_quarter or not entry.planned_start or not entry.planned_finish
        rows.append(
            '<tr data-rnmc-row="1" '
            f'data-status="{_attr(entry.status)}" data-processed="{_flag(entry.status == "skipped_processed")}" '
            f'data-problem="{_flag(problem)}" data-empty="{_flag(empty)}">'
            f'<td class="path" title="{_attr(entry.archive_path)}">{html.escape(entry.filename)}</td>'
            f'<td>{html.escape(entry.region_folder)}</td>'
            f'<td>{_render_rnmc_coefficient_input(entry)}</td>'
            f'<td>{html.escape(entry.lsr_quarter)}</td>'
            f'<td>{html.escape(entry.planned_start)}</td>'
            f'<td>{html.escape(entry.planned_finish)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table></div>'



def _rnmc_coefficient_is_editable(entry) -> bool:
    return entry.status not in {"skipped_processed", "duplicate_name", "unsupported_format", "parse_error"}


def _render_rnmc_coefficient_input(entry) -> str:
    value = _display_optional_number(entry.regional_coefficient)
    if not _rnmc_coefficient_is_editable(entry):
        return value
    key = normalize_import_filename(entry.filename)
    name = f"coefficient_override__{key}"
    placeholder = "например 1.2"
    return (
        f'<input class="rnmc-coef-input" type="text" inputmode="decimal" '
        f'name="{html.escape(name, quote=True)}" form="rnmc-stage-commit-form" '
        f'value="{html.escape(value, quote=True)}" placeholder="{placeholder}" '
        'title="Региональный коэффициент для пересчета price_original в price_zlvl">'
    )


def _render_rnmc_preview_header_table(entries) -> str:
    table_id = "rnmc-headers"
    header = (
        f'{_rnmc_filter_toolbar(table_id)}'
        f'<div class="preview-wide"><table class="preview" data-rnmc-table="{table_id}"><thead><tr>'
        '<th>Файл</th><th>Header row</th><th>Код</th><th class="wrap">Наименование</th>'
        '<th>Ед.</th><th>Кол-во</th><th class="wrap">Цена ед.</th><th class="wrap">Итого</th>'
        '<th>ТЗ ед.</th><th>ТЗ всего</th><th>ТЗм ед.</th><th>ТЗм всего</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in entries:
        labels = entry.header_preview
        empty = not any([
            labels.code, labels.work_name, labels.unit, labels.quantity, labels.unit_price,
            labels.total_price, labels.labor_unit, labels.labor_total, labels.machine_labor_unit,
            labels.machine_labor_total,
        ])
        problem = entry.status != "preview_ok" or not labels.code or not labels.unit or not labels.quantity or not labels.unit_price
        rows.append(
            '<tr data-rnmc-row="1" '
            f'data-status="{_attr(entry.status)}" data-processed="{_flag(entry.status == "skipped_processed")}" '
            f'data-problem="{_flag(problem)}" data-empty="{_flag(empty)}">'
            f'<td class="path" title="{_attr(entry.archive_path)}">{html.escape(entry.filename)}</td>'
            f'<td>{entry.header_row or ""}</td>'
            f'<td>{html.escape(labels.code)}</td>'
            f'<td class="wrap">{html.escape(labels.work_name)}</td>'
            f'<td>{html.escape(labels.unit)}</td>'
            f'<td>{html.escape(labels.quantity)}</td>'
            f'<td class="wrap">{html.escape(labels.unit_price)}</td>'
            f'<td class="wrap">{html.escape(labels.total_price)}</td>'
            f'<td>{html.escape(labels.labor_unit)}</td>'
            f'<td>{html.escape(labels.labor_total)}</td>'
            f'<td>{html.escape(labels.machine_labor_unit)}</td>'
            f'<td>{html.escape(labels.machine_labor_total)}</td>'
            '</tr>'
        )
    return header + ''.join(rows) + '</tbody></table></div>'


def _render_rnmc_preview_rows_table(entries) -> str:
    table_id = "rnmc-rows"
    header = (
        f'{_rnmc_filter_toolbar(table_id)}'
        f'<div class="preview-wide"><table class="preview" data-rnmc-table="{table_id}"><thead><tr>'
        '<th>Файл</th><th>Excel row</th><th>Код</th><th class="wrap">Работа</th><th>Ед.</th><th>Кол-во</th>'
        '<th>Цена ед. без НДС</th><th>Итого без НДС</th><th>ТЗ ед.</th><th>ТЗ всего</th>'
        '<th>ТЗм ед.</th><th>ТЗм всего</th><th>Проблема</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for entry in entries:
        if not entry.sample_rows:
            continue
        for sample in entry.sample_rows:
            problem = bool(sample.issue)
            empty = not any([sample.code, sample.work_name, sample.unit, sample.quantity, sample.unit_price, sample.total_price])
            rows.append(
                '<tr data-rnmc-row="1" '
                f'data-status="{_attr(entry.status)}" data-processed="{_flag(entry.status == "skipped_processed")}" '
                f'data-problem="{_flag(problem)}" data-empty="{_flag(empty)}">'
                f'<td class="path" title="{_attr(entry.archive_path)}">{html.escape(entry.filename)}</td>'
                f'<td>{sample.row_number}</td>'
                f'<td>{html.escape(sample.code)}</td>'
                f'<td class="wrap">{html.escape(sample.work_name)}</td>'
                f'<td>{html.escape(sample.unit)}</td>'
                f'<td>{html.escape(sample.quantity)}</td>'
                f'<td>{html.escape(sample.unit_price)}</td>'
                f'<td>{html.escape(sample.total_price)}</td>'
                f'<td>{html.escape(sample.labor_unit)}</td>'
                f'<td>{html.escape(sample.labor_total)}</td>'
                f'<td>{html.escape(sample.machine_labor_unit)}</td>'
                f'<td>{html.escape(sample.machine_labor_total)}</td>'
                f'<td>{html.escape(sample.issue)}</td>'
                '</tr>'
            )
        if entry.is_limited:
            rows.append(
                '<tr data-rnmc-row="1" '
                f'data-status="{_attr(entry.status)}" data-processed="{_flag(entry.status == "skipped_processed")}" '
                'data-problem="0" data-empty="1">'
                f'<td class="path" title="{_attr(entry.archive_path)}">{html.escape(entry.filename)}</td>'
                f'<td colspan="12" class="muted">Показаны первые {DEFAULT_ROW_PREVIEW_LIMIT} строк тела таблицы</td>'
                '</tr>'
            )
    if not rows:
        rows.append(
            '<tr><td colspan="13" class="muted">Нет строк для показа. Подробные причины смотрите во вкладке «Файлы и статусы».</td></tr>'
        )
    return header + ''.join(rows) + '</tbody></table></div>'

def _preview_status_class(status: str) -> str:
    if status == "preview_ok":
        return "ok"
    if status in {"skipped_processed", "duplicate_name", "no_table", "no_rows", "unsupported_format", "parse_error"}:
        return "warn"
    return ""


def _render_limited_marker(is_limited: bool) -> str:
    if not is_limited:
        return ''
    return '<br><span class="muted">Остановлено на лимите предпросмотра</span>'





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
    summary += (
        f'<p class="muted">Предпросмотр читает не больше {DEFAULT_ROW_PREVIEW_LIMIT} строк на один Excel-файл. '
        'Реальный импорт по-прежнему читает все строки файла.</p>'
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
        '<th>Коэф.</th>'
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
            f'<td>{_display_optional_number(entry.regional_coefficient)}</td>'
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
        ("Региональный коэффициент", _display_optional_number(record.regional_coefficient)),
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
        f'<label>Региональный коэффициент<input type="text" name="regional_coefficient" value="{_display_optional_number(record.regional_coefficient)}"></label>'
        '<button type="submit">Сохранить данные файла</button>'
        '</form>'
    )


def _render_import_catalog_rows(rows: list[CatalogItemRecord]) -> str:
    if not rows:
        return '<div class="admin-form"><h2 class="section">Импортированные строки</h2><p class="muted">Для этого файла строки catalog_items не найдены.</p></div>'
    header = (
        '<div class="admin-form"><h2 class="section">Импортированные строки</h2>'
        '<table class="preview"><thead><tr>'
        '<th>Excel row</th><th>Код</th><th>Ед.</th><th>Кол-во</th><th>Цена раб.</th><th>Цена ориг.</th><th>Цена ZLVL</th>'
        '<th>Итого</th><th>ТЗ ед.</th><th>ТЗ всего</th>'
        '<th>ТЗм ед.</th><th>ТЗм всего</th><th>Коэф.</th><th>ЛСР</th><th>Начало</th><th>Окончание</th><th>Работа</th>'
        '</tr></thead><tbody>'
    )
    body = []
    for row in rows[:200]:
        body.append(
            '<tr>'
            f'<td>{row.source_row_number}</td>'
            f'<td>{html.escape(row.code)}</td>'
            f'<td>{html.escape(row.unit)}</td>'
            f'<td>{_display_optional_number(row.quantity)}</td>'
            f'<td>{row.price:g}</td>'
            f'<td>{_display_optional_number(row.price_original)}</td>'
            f'<td>{_display_optional_number(row.price_zlvl)}</td>'
            f'<td>{_display_optional_number(row.total_price)}</td>'
            f'<td>{_display_optional_number(row.labor_unit)}</td>'
            f'<td>{_display_optional_number(row.labor_total)}</td>'
            f'<td>{_display_optional_number(row.machine_labor_unit)}</td>'
            f'<td>{_display_optional_number(row.machine_labor_total)}</td>'
            f'<td>{_display_optional_number(row.regional_coefficient)}</td>'
            f'<td>{html.escape(row.lsr_quarter)}</td>'
            f'<td>{html.escape(row.planned_start)}</td>'
            f'<td>{html.escape(row.planned_finish)}</td>'
            f'<td>{html.escape(row.work_name)}</td>'
            '</tr>'
        )
    tail = '</tbody></table>'
    if len(rows) > 200:
        tail += f'<p class="muted">Показаны первые 200 строк из {len(rows)}.</p>'
    return header + ''.join(body) + tail + '</div>'


def _display_optional_number(value: float | None) -> str:
    if value is None:
        return ''
    return f'{value:g}'


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


def render_admin_settings(settings_rows: list[tuple[str, str]], *, catalog_rows: int = 0) -> str:
    content = (
        '<section class="admin-panel">'
        '<h2 class="section">Настройки и диагностика</h2>'
        '<p>Сводка по базе, конфигам и накопленным админ-таблицам.</p>'
        f'{_render_settings_table(settings_rows)}'
        f'{_render_catalog_clear_form(catalog_rows)}'
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

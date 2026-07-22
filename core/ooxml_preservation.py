"""Preserve workbook package features that openpyxl cannot round-trip.

The result writer changes one worksheet only. Some real estimate workbooks
also contain EMF/WMF drawings and binary printer settings. openpyxl can read
the cells in those files, but drops unsupported drawings and related package
parts when it saves the workbook. This module restores those untouched OOXML
parts after the normal cell-level write has completed.
"""

from __future__ import annotations

import copy
import posixpath
import tempfile
from xml.etree import ElementTree as ET
from pathlib import Path, PurePosixPath
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

RELATIONSHIP_ID = f"{{{DOCUMENT_REL_NS}}}id"
PRESERVED_PART_PREFIXES = (
    "xl/drawings/",
    "xl/media/",
    "xl/printerSettings/",
    "xl/vmlDrawings/",
    "xl/embeddings/",
    "xl/ctrlProps/",
    "xl/activeX/",
)


def preserve_workbook_package_features(
    source_path: str | Path,
    output_path: str | Path,
    *,
    modified_sheet_title: str,
) -> None:
    """Restore unsupported package features without reverting written cells.

    Unmodified worksheets are copied byte-for-byte from the source workbook.
    The modified worksheet keeps its new cell data, while relationships to
    preserved drawings and printer settings are merged back into its XML.
    """
    source = Path(source_path)
    output = Path(output_path)
    if not _is_ooxml_workbook(source) or not _is_ooxml_workbook(output):
        return

    with ZipFile(source) as source_zip, ZipFile(output) as output_zip:
        source_names = set(source_zip.namelist())
        output_names = set(output_zip.namelist())
        source_sheets = _sheet_parts(source_zip)
        output_sheets = _sheet_parts(output_zip)

        replacements: dict[str, bytes] = {}
        source_info_by_output_name: dict[str, ZipInfo] = {}
        copied_unmodified_sheet = False

        for title, source_part in source_sheets.items():
            output_part = output_sheets.get(title)
            if output_part is None or title == modified_sheet_title:
                continue
            replacements[output_part] = source_zip.read(source_part)
            source_info_by_output_name[output_part] = source_zip.getinfo(source_part)
            copied_unmodified_sheet = True

            source_rels = _rels_part(source_part)
            output_rels = _rels_part(output_part)
            if source_rels in source_names:
                replacements[output_rels] = source_zip.read(source_rels)
                source_info_by_output_name[output_rels] = source_zip.getinfo(source_rels)
            elif output_rels in output_names:
                replacements[output_rels] = b""

        for name in source_names:
            if name.startswith(PRESERVED_PART_PREFIXES):
                replacements[name] = source_zip.read(name)
                source_info_by_output_name[name] = source_zip.getinfo(name)

        if copied_unmodified_sheet and "xl/sharedStrings.xml" in source_names:
            replacements["xl/sharedStrings.xml"] = source_zip.read("xl/sharedStrings.xml")
            source_info_by_output_name["xl/sharedStrings.xml"] = source_zip.getinfo(
                "xl/sharedStrings.xml"
            )
            replacements["xl/_rels/workbook.xml.rels"] = _merge_workbook_relationships(
                source_zip,
                output_zip,
                replacements,
            )

        source_part = source_sheets.get(modified_sheet_title)
        output_part = output_sheets.get(modified_sheet_title)
        if source_part is not None and output_part is not None:
            _merge_modified_sheet_relationships(
                source_zip,
                output_zip,
                source_part,
                output_part,
                replacements,
            )

        replacements["[Content_Types].xml"] = _merge_content_types(
            source_zip,
            output_zip,
            replacements,
        )

    _rewrite_archive(
        output,
        replacements,
        source_info_by_output_name,
    )


def _is_ooxml_workbook(path: Path) -> bool:
    return path.suffix.casefold() in {".xlsx", ".xlsm", ".xltx", ".xltm"}


def _sheet_parts(workbook_zip: ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
    rels_root = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: _resolve_part("xl/workbook.xml", rel.attrib["Target"])
        for rel in rels_root
        if "Id" in rel.attrib and "Target" in rel.attrib
    }

    parts: dict[str, str] = {}
    sheets = workbook_root.find(f"{{{SPREADSHEET_NS}}}sheets")
    if sheets is None:
        return parts
    for sheet in sheets:
        title = sheet.attrib.get("name")
        relation_id = sheet.attrib.get(RELATIONSHIP_ID)
        if title and relation_id in targets:
            parts[title] = targets[relation_id]
    return parts


def _resolve_part(parent_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(parent_part), target))


def _rels_part(part: str) -> str:
    path = PurePosixPath(part)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _merge_modified_sheet_relationships(
    source_zip: ZipFile,
    output_zip: ZipFile,
    source_part: str,
    output_part: str,
    replacements: dict[str, bytes],
) -> None:
    source_sheet_root = ET.fromstring(source_zip.read(source_part))
    output_sheet_root = ET.fromstring(output_zip.read(output_part))
    _restore_source_row_metadata(source_sheet_root, output_sheet_root)

    source_rels_part = _rels_part(source_part)
    if source_rels_part not in source_zip.namelist():
        replacements[output_part] = ET.tostring(
            output_sheet_root,
            encoding="utf-8",
            xml_declaration=True,
        )
        return

    output_rels_part = _rels_part(output_part)
    source_rels_root = ET.fromstring(source_zip.read(source_rels_part))
    if output_rels_part in output_zip.namelist():
        output_rels_root = ET.fromstring(output_zip.read(output_rels_part))
    else:
        output_rels_root = ET.Element(f"{{{PACKAGE_REL_NS}}}Relationships")

    used_ids = {rel.attrib.get("Id", "") for rel in output_rels_root}
    existing = {
        (rel.attrib.get("Type"), rel.attrib.get("Target")): rel.attrib.get("Id")
        for rel in output_rels_root
    }
    relation_id_map: dict[str, str] = {}

    for source_rel in source_rels_root:
        target = source_rel.attrib.get("Target")
        source_id = source_rel.attrib.get("Id")
        if not target or not source_id:
            continue
        resolved_target = _resolve_part(source_part, target)
        if not resolved_target.startswith(PRESERVED_PART_PREFIXES):
            continue

        key = (source_rel.attrib.get("Type"), target)
        output_id = existing.get(key)
        if not output_id:
            output_id = _next_relation_id(used_ids)
            new_rel = copy.deepcopy(source_rel)
            new_rel.attrib["Id"] = output_id
            output_rels_root.append(new_rel)
            used_ids.add(output_id)
            existing[key] = output_id
        relation_id_map[source_id] = output_id

    if relation_id_map:
        _merge_relationship_elements(source_sheet_root, output_sheet_root, relation_id_map)

    replacements[output_part] = ET.tostring(
        output_sheet_root,
        encoding="utf-8",
        xml_declaration=True,
    )
    replacements[output_rels_part] = ET.tostring(
        output_rels_root,
        encoding="utf-8",
        xml_declaration=True,
    )


def _restore_source_row_metadata(source_root, output_root) -> None:
    """Keep original row sizing flags while retaining rewritten cell data."""
    source_sheet_data = source_root.find(f"{{{SPREADSHEET_NS}}}sheetData")
    output_sheet_data = output_root.find(f"{{{SPREADSHEET_NS}}}sheetData")
    if source_sheet_data is None or output_sheet_data is None:
        return

    source_rows = {row.attrib.get("r"): row for row in source_sheet_data}
    for output_row in output_sheet_data:
        source_row = source_rows.get(output_row.attrib.get("r"))
        if source_row is not None:
            output_row.attrib.clear()
            output_row.attrib.update(source_row.attrib)


def _next_relation_id(used_ids: set[str]) -> str:
    index = 1
    while f"rId{index}" in used_ids:
        index += 1
    return f"rId{index}"


def _merge_relationship_elements(source_root, output_root, relation_id_map: dict[str, str]) -> None:
    for source_element in source_root.iter():
        source_id = source_element.attrib.get(RELATIONSHIP_ID)
        if source_id not in relation_id_map:
            continue

        local_name = source_element.tag.rsplit("}", 1)[-1]
        output_element = next(
            (
                candidate
                for candidate in output_root.iter()
                if candidate.tag.rsplit("}", 1)[-1] == local_name
            ),
            None,
        )
        if output_element is None:
            output_element = copy.deepcopy(source_element)
            output_root.append(output_element)
        output_element.attrib[RELATIONSHIP_ID] = relation_id_map[source_id]


def _merge_content_types(
    source_zip: ZipFile,
    output_zip: ZipFile,
    replacements: dict[str, bytes],
) -> bytes:
    source_root = ET.fromstring(source_zip.read("[Content_Types].xml"))
    output_root = ET.fromstring(output_zip.read("[Content_Types].xml"))
    preserved_names = set(replacements)

    existing_defaults = {
        element.attrib.get("Extension")
        for element in output_root
        if element.tag.rsplit("}", 1)[-1] == "Default"
    }
    existing_overrides = {
        element.attrib.get("PartName")
        for element in output_root
        if element.tag.rsplit("}", 1)[-1] == "Override"
    }

    preserved_extensions = {
        PurePosixPath(name).suffix.lstrip(".")
        for name in preserved_names
        if PurePosixPath(name).suffix
    }
    for element in source_root:
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name == "Default":
            extension = element.attrib.get("Extension")
            if extension in preserved_extensions and extension not in existing_defaults:
                output_root.append(copy.deepcopy(element))
                existing_defaults.add(extension)
        elif local_name == "Override":
            part_name = element.attrib.get("PartName")
            if (
                part_name
                and part_name.lstrip("/") in preserved_names
                and part_name not in existing_overrides
            ):
                output_root.append(copy.deepcopy(element))
                existing_overrides.add(part_name)

    return ET.tostring(output_root, encoding="utf-8", xml_declaration=True)


def _merge_workbook_relationships(
    source_zip: ZipFile,
    output_zip: ZipFile,
    replacements: dict[str, bytes],
) -> bytes:
    part = "xl/_rels/workbook.xml.rels"
    source_root = ET.fromstring(source_zip.read(part))
    output_root = ET.fromstring(output_zip.read(part))
    used_ids = {rel.attrib.get("Id", "") for rel in output_root}
    existing = {
        (rel.attrib.get("Type"), rel.attrib.get("Target"))
        for rel in output_root
    }

    for source_rel in source_root:
        target = source_rel.attrib.get("Target")
        if not target:
            continue
        resolved_target = _resolve_part("xl/workbook.xml", target)
        key = (source_rel.attrib.get("Type"), target)
        if resolved_target not in replacements or key in existing:
            continue
        new_rel = copy.deepcopy(source_rel)
        new_rel.attrib["Id"] = _next_relation_id(used_ids)
        output_root.append(new_rel)
        used_ids.add(new_rel.attrib["Id"])
        existing.add(key)

    return ET.tostring(output_root, encoding="utf-8", xml_declaration=True)


def _rewrite_archive(
    output_path: Path,
    replacements: dict[str, bytes],
    source_info_by_output_name: dict[str, ZipInfo],
) -> None:
    with tempfile.NamedTemporaryFile(
        prefix=f".{output_path.stem}-",
        suffix=output_path.suffix,
        dir=output_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)

    try:
        written: set[str] = set()
        with (
            ZipFile(output_path) as output_zip,
            ZipFile(temporary_path, "w", compression=ZIP_DEFLATED) as target_zip,
        ):
            for info in output_zip.infolist():
                name = info.filename
                if name in replacements:
                    data = replacements[name]
                    if data == b"":
                        written.add(name)
                        continue
                else:
                    data = output_zip.read(name)
                target_zip.writestr(info, data)
                written.add(name)

            for name, data in replacements.items():
                if name in written or data == b"":
                    continue
                info = source_info_by_output_name.get(name)
                if info is None:
                    target_zip.writestr(name, data)
                else:
                    target_zip.writestr(info, data)
                written.add(name)

        temporary_path.replace(output_path)
    finally:
        temporary_path.unlink(missing_ok=True)

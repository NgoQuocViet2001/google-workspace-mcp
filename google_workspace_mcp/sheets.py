from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .common import (
    IMAGE_FORMULA_RE,
    NS,
    a1_from_zero_based,
    compact_dict,
    rel_join,
    safe_filename,
    scalar_value,
    text_style_summary,
)
from .docs import ensure_output_dir


def simplify_text_runs(text_runs: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    output = []
    for run in text_runs or []:
        output.append(
            compact_dict(
                {
                    "start_index": run.get("startIndex"),
                    "format": text_style_summary(run.get("format")),
                }
            )
        )
    return output


def _style_markers(style: dict[str, Any] | None) -> tuple[list[str], list[str]]:
    if not style:
        return [], []

    open_markers: list[str] = []
    close_markers: list[str] = []
    if style.get("strikethrough"):
        open_markers.append("[[STRIKE]]")
        close_markers.insert(0, "[[/STRIKE]]")
    if style.get("underline"):
        open_markers.append("[[UNDERLINE]]")
        close_markers.insert(0, "[[/UNDERLINE]]")
    return open_markers, close_markers


def annotate_formatted_text(
    formatted_value: str | None,
    text_runs: list[dict[str, Any]] | None,
) -> str | None:
    if not formatted_value:
        return None
    if not text_runs:
        return None

    segments = []
    for index, run in enumerate(text_runs):
        start_index = run.get("start_index", run.get("startIndex"))
        if start_index is None and index == 0:
            start_index = 0
        if start_index is None:
            continue
        end_index = (
            text_runs[index + 1].get("start_index", text_runs[index + 1].get("startIndex"))
            if index + 1 < len(text_runs)
            else len(formatted_value)
        )
        if end_index is None or end_index <= start_index:
            continue
        chunk = formatted_value[start_index:end_index]
        if not chunk:
            continue
        open_markers, close_markers = _style_markers(run.get("format"))
        segments.append("".join(open_markers) + chunk + "".join(close_markers))

    annotated = "".join(segments)
    return annotated if annotated != formatted_value else None


def simplify_chip_runs(chip_runs: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    result = []
    for run in chip_runs or []:
        chip = run.get("chip", {})
        person = chip.get("personProperties", {})
        rich_link = chip.get("richLinkProperties", {})
        result.append(
            compact_dict(
                {
                    "start_index": run.get("startIndex"),
                    "person": compact_dict(
                        {
                            "display_format": person.get("displayFormat"),
                            "email": person.get("email"),
                            "name": person.get("name"),
                        }
                    ),
                    "rich_link": compact_dict(
                        {
                            "mime_type": rich_link.get("mimeType"),
                            "title": rich_link.get("title"),
                            "uri": rich_link.get("uri"),
                        }
                    ),
                }
            )
        )
    return result


def simplify_grid_data(sheet_payload: dict[str, Any]) -> list[dict[str, Any]]:
    simplified_sheets = []
    for sheet in sheet_payload.get("sheets", []):
        properties = sheet.get("properties", {})
        data_sections = []
        for section in sheet.get("data", []):
            start_row = section.get("startRow", 0)
            start_col = section.get("startColumn", 0)
            rows = []
            for row_offset, row in enumerate(section.get("rowData", [])):
                row_cells = []
                for col_offset, cell in enumerate(row.get("values", [])):
                    user_value = scalar_value(cell.get("userEnteredValue"))
                    effective_value = scalar_value(cell.get("effectiveValue"))
                    formula = cell.get("userEnteredValue", {}).get("formulaValue")
                    image_formula_match = IMAGE_FORMULA_RE.match(formula or "")
                    text_runs = simplify_text_runs(cell.get("textFormatRuns"))
                    formatted_value = cell.get("formattedValue")
                    row_cells.append(
                        compact_dict(
                            {
                                "a1": a1_from_zero_based(start_row + row_offset, start_col + col_offset),
                                "row_index": start_row + row_offset + 1,
                                "column_index": start_col + col_offset + 1,
                                "formatted_value": formatted_value,
                                "user_entered_value": user_value,
                                "effective_value": effective_value,
                                "formula": formula,
                                "note": cell.get("note"),
                                "hyperlink": cell.get("hyperlink")
                                or cell.get("userEnteredFormat", {})
                                .get("textFormat", {})
                                .get("link", {})
                                .get("uri"),
                                "text_runs": text_runs,
                                "annotated_text": annotate_formatted_text(formatted_value, text_runs),
                                "chip_runs": simplify_chip_runs(cell.get("chipRuns")),
                                "detected_image_formula_url": image_formula_match.group(1)
                                if image_formula_match
                                else None,
                            }
                        )
                    )
                rows.append(
                    {
                        "row_index": start_row + row_offset + 1,
                        "cells": row_cells,
                    }
                )
            data_sections.append(
                {
                    "start_row": start_row + 1,
                    "start_column": start_col + 1,
                    "rows": rows,
                }
            )
        simplified_sheets.append(
            {
                "sheet_id": properties.get("sheetId"),
                "sheet_name": properties.get("title"),
                "index": properties.get("index"),
                "row_count": properties.get("gridProperties", {}).get("rowCount"),
                "column_count": properties.get("gridProperties", {}).get("columnCount"),
                "data": data_sections,
            }
        )
    return simplified_sheets


def normalize_headers(header_values: list[Any]) -> list[str]:
    seen: dict[str, int] = {}
    headers = []
    for index, value in enumerate(header_values, start=1):
        base = str(value).strip() or f"column_{index}"
        count = seen.get(base, 0) + 1
        seen[base] = count
        headers.append(base if count == 1 else f"{base}_{count}")
    return headers


def flatten_values_rows(values_payload: dict[str, Any]) -> list[list[Any]]:
    return values_payload.get("values", [])


def parse_relationships(xml_bytes: bytes) -> dict[str, str]:
    root = ET.fromstring(xml_bytes)
    mapping: dict[str, str] = {}
    for relationship in root.findall("rel:Relationship", NS):
        rel_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if rel_id and target:
            mapping[rel_id] = target
    return mapping


def parse_workbook_sheets(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    workbook_rels = parse_relationships(zf.read("xl/_rels/workbook.xml.rels"))
    sheets = []
    for sheet in workbook_root.findall("main:sheets/main:sheet", NS):
        rel_id = sheet.attrib.get(f"{{{NS['r']}}}id")
        target = workbook_rels.get(rel_id or "", "")
        sheets.append(
            {
                "name": sheet.attrib.get("name"),
                "sheet_id": sheet.attrib.get("sheetId"),
                "path": rel_join("xl/workbook.xml", target),
            }
        )
    return sheets


def anchor_cell(anchor_root: ET.Element | None) -> dict[str, Any] | None:
    if anchor_root is None:
        return None
    col = anchor_root.findtext("xdr:col", default="0", namespaces=NS)
    row = anchor_root.findtext("xdr:row", default="0", namespaces=NS)
    col_off = anchor_root.findtext("xdr:colOff", default="0", namespaces=NS)
    row_off = anchor_root.findtext("xdr:rowOff", default="0", namespaces=NS)
    col_index = int(col)
    row_index = int(row)
    return {
        "a1": a1_from_zero_based(row_index, col_index),
        "row_index": row_index + 1,
        "column_index": col_index + 1,
        "col_offset_emu": int(col_off),
        "row_offset_emu": int(row_off),
    }


def extract_sheet_images_from_xlsx(
    xlsx_bytes: bytes,
    *,
    output_dir: str | None = None,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    target_dir = ensure_output_dir(output_dir, "google-sheet-images-") if output_dir else None
    workbook_summary = []
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
        temp_file.write(xlsx_bytes)
        temp_path = Path(temp_file.name)
    try:
        with zipfile.ZipFile(temp_path) as zf:
            sheets = parse_workbook_sheets(zf)
            for sheet in sheets:
                if sheet_name and sheet["name"] != sheet_name:
                    continue
                sheet_path = sheet["path"]
                if sheet_path not in zf.namelist():
                    continue
                sheet_root = ET.fromstring(zf.read(sheet_path))
                drawing = sheet_root.find("main:drawing", NS)
                if drawing is None:
                    workbook_summary.append({"sheet_name": sheet["name"], "images": []})
                    continue
                rel_id = drawing.attrib.get(f"{{{NS['r']}}}id")
                sheet_rels_path = rel_join(sheet_path, f"_rels/{Path(sheet_path).name}.rels")
                if sheet_rels_path not in zf.namelist():
                    workbook_summary.append({"sheet_name": sheet["name"], "images": []})
                    continue
                sheet_rels = parse_relationships(zf.read(sheet_rels_path))
                drawing_path = rel_join(sheet_path, sheet_rels.get(rel_id or "", ""))
                drawing_rels_path = rel_join(drawing_path, f"_rels/{Path(drawing_path).name}.rels")
                drawing_rels = (
                    parse_relationships(zf.read(drawing_rels_path))
                    if drawing_rels_path in zf.namelist()
                    else {}
                )
                drawing_root = ET.fromstring(zf.read(drawing_path))
                images = []
                for anchor_type in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                    for anchor in drawing_root.findall(f"xdr:{anchor_type}", NS):
                        pic = anchor.find("xdr:pic", NS)
                        if pic is None:
                            continue
                        non_visual = pic.find("xdr:nvPicPr/xdr:cNvPr", NS)
                        blip = pic.find("xdr:blipFill/a:blip", NS)
                        embed_id = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
                        media_target = drawing_rels.get(embed_id or "", "")
                        media_path = rel_join(drawing_path, media_target) if media_target else None
                        saved_path = None
                        if target_dir and media_path and media_path in zf.namelist():
                            extension = Path(media_path).suffix or ".bin"
                            filename = (
                                f"{safe_filename(sheet['name'])}_"
                                f"{safe_filename(non_visual.attrib.get('name', 'image'))}{extension}"
                            )
                            output_path = target_dir / filename
                            output_path.write_bytes(zf.read(media_path))
                            saved_path = str(output_path)
                        images.append(
                            compact_dict(
                                {
                                    "name": non_visual.attrib.get("name") if non_visual is not None else None,
                                    "description": non_visual.attrib.get("descr")
                                    if non_visual is not None
                                    else None,
                                    "anchor_type": anchor_type,
                                    "from": anchor_cell(anchor.find("xdr:from", NS)),
                                    "to": anchor_cell(anchor.find("xdr:to", NS)),
                                    "media_path_in_xlsx": media_path,
                                    "saved_path": saved_path,
                                }
                            )
                        )
                workbook_summary.append({"sheet_name": sheet["name"], "images": images})
    finally:
        temp_path.unlink(missing_ok=True)

    return compact_dict(
        {
            "output_dir": str(target_dir) if target_dir else None,
            "sheets": workbook_summary,
        }
    )


def collect_formula_images(sheet_payload: dict[str, Any]) -> list[dict[str, Any]]:
    results = []
    for sheet in simplify_grid_data(sheet_payload):
        for section in sheet.get("data", []):
            for row in section.get("rows", []):
                for cell in row.get("cells", []):
                    formula = cell.get("formula") or ""
                    match = IMAGE_FORMULA_RE.match(formula)
                    if match:
                        results.append(
                            {
                                "sheet_name": sheet.get("sheet_name"),
                                "a1": cell.get("a1"),
                                "formula": formula,
                                "image_url": match.group(1),
                            }
                        )
    return results

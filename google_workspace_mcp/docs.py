from __future__ import annotations

import tempfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

import requests

from .common import compact_dict, safe_filename, text_style_summary

if TYPE_CHECKING:
    from .client import GoogleWorkspaceClient


def doc_tabs(document: dict[str, Any]) -> list[dict[str, Any]]:
    flat_tabs: list[dict[str, Any]] = []

    def _walk(items: list[dict[str, Any]]) -> None:
        for tab in items or []:
            flat_tabs.append(tab)
            _walk(tab.get("childTabs", []))

    _walk(document.get("tabs", []))
    return flat_tabs


def extract_embedded_object(
    object_id: str,
    embedded_object: dict[str, Any],
    *,
    object_type: str,
    positioning: dict[str, Any] | None = None,
) -> dict[str, Any]:
    image_properties = embedded_object.get("imageProperties", {})
    linked_chart = embedded_object.get("linkedContentReference", {}).get("sheetsChartReference", {})
    return compact_dict(
        {
            "object_id": object_id,
            "object_type": object_type,
            "title": embedded_object.get("title"),
            "description": embedded_object.get("description"),
            "alt_text": " ".join(
                item.strip()
                for item in [embedded_object.get("title", ""), embedded_object.get("description", "")]
                if item and item.strip()
            )
            or None,
            "size": embedded_object.get("size"),
            "margin_top": embedded_object.get("marginTop"),
            "margin_bottom": embedded_object.get("marginBottom"),
            "margin_left": embedded_object.get("marginLeft"),
            "margin_right": embedded_object.get("marginRight"),
            "content_uri": image_properties.get("contentUri"),
            "source_uri": image_properties.get("sourceUri"),
            "brightness": image_properties.get("brightness"),
            "contrast": image_properties.get("contrast"),
            "transparency": image_properties.get("transparency"),
            "crop_properties": image_properties.get("cropProperties"),
            "rotation_radians": image_properties.get("angle"),
            "linked_chart": linked_chart or None,
            "positioning": positioning or None,
            "is_image": bool(image_properties),
            "is_drawing": "embeddedDrawingProperties" in embedded_object,
        }
    )


def simplify_paragraph_element(element: dict[str, Any], tab_doc: dict[str, Any]) -> dict[str, Any]:
    if "textRun" in element:
        text_run = element["textRun"]
        return compact_dict(
            {
                "type": "text",
                "text": text_run.get("content"),
                "text_style": text_style_summary(text_run.get("textStyle")),
            }
        )

    if "inlineObjectElement" in element:
        inline_object_id = element["inlineObjectElement"].get("inlineObjectId")
        inline_object = tab_doc.get("inlineObjects", {}).get(inline_object_id, {})
        embedded_object = inline_object.get("inlineObjectProperties", {}).get("embeddedObject", {})
        return {
            "type": "inline_object",
            "object_id": inline_object_id,
            "object": extract_embedded_object(
                inline_object_id or "",
                embedded_object,
                object_type="inline_object",
            ),
        }

    if "footnoteReference" in element:
        return {"type": "footnote_reference", "footnote_id": element["footnoteReference"].get("footnoteId")}
    if "pageBreak" in element:
        return {"type": "page_break"}
    if "columnBreak" in element:
        return {"type": "column_break"}
    if "horizontalRule" in element:
        return {"type": "horizontal_rule"}
    if "equation" in element:
        return {"type": "equation"}
    if "person" in element:
        return compact_dict(
            {
                "type": "person",
                "person_id": element["person"].get("personId"),
                "person_properties": element["person"].get("personProperties"),
            }
        )
    if "richLink" in element:
        return compact_dict({"type": "rich_link", "rich_link": element["richLink"]})
    if "autoText" in element:
        return compact_dict({"type": "auto_text", "auto_text": element["autoText"]})
    return {"type": "unknown", "raw": element}


def simplify_structural_elements(
    content: list[dict[str, Any]],
    tab_doc: dict[str, Any],
) -> list[dict[str, Any]]:
    simplified = []
    for item in content or []:
        if "paragraph" in item:
            paragraph = item["paragraph"]
            elements = [simplify_paragraph_element(el, tab_doc) for el in paragraph.get("elements", [])]
            text_preview = "".join(
                el.get("text", "")
                if el.get("type") == "text"
                else f"[{el.get('type')}:{el.get('object_id', '')}]"
                for el in elements
            )
            simplified.append(
                compact_dict(
                    {
                        "type": "paragraph",
                        "style": paragraph.get("paragraphStyle"),
                        "bullet": paragraph.get("bullet"),
                        "elements": elements,
                        "text": text_preview.rstrip("\n") or None,
                    }
                )
            )
            continue

        if "table" in item:
            table = item["table"]
            rows = []
            for row in table.get("tableRows", []):
                cells = []
                for cell in row.get("tableCells", []):
                    cells.append(
                        {
                            "col_span": cell.get("columnSpan"),
                            "row_span": cell.get("rowSpan"),
                            "content": simplify_structural_elements(cell.get("content", []), tab_doc),
                        }
                    )
                rows.append({"cells": cells})
            simplified.append({"type": "table", "rows": rows})
            continue

        if "tableOfContents" in item:
            toc = item["tableOfContents"]
            simplified.append(
                {
                    "type": "table_of_contents",
                    "content": simplify_structural_elements(toc.get("content", []), tab_doc),
                }
            )
            continue

        if "sectionBreak" in item:
            simplified.append({"type": "section_break", "style": item["sectionBreak"].get("sectionStyle")})
            continue

        simplified.append({"type": "unknown_structural_element", "raw": item})
    return simplified


def simplify_document(document: dict[str, Any], tab_id: str | None = None) -> dict[str, Any]:
    simplified_tabs = []
    for tab in doc_tabs(document):
        tab_properties = tab.get("tabProperties", {})
        current_tab_id = tab_properties.get("tabId")
        if tab_id and current_tab_id != tab_id:
            continue
        document_tab = tab.get("documentTab", {})
        inline_objects = []
        for object_id, inline_object in document_tab.get("inlineObjects", {}).items():
            embedded = inline_object.get("inlineObjectProperties", {}).get("embeddedObject", {})
            inline_objects.append(
                extract_embedded_object(object_id, embedded, object_type="inline_object")
            )
        positioned_objects = []
        for object_id, positioned_object in document_tab.get("positionedObjects", {}).items():
            properties = positioned_object.get("positionedObjectProperties", {})
            positioned_objects.append(
                extract_embedded_object(
                    object_id,
                    properties.get("embeddedObject", {}),
                    object_type="positioned_object",
                    positioning=properties.get("positioning"),
                )
            )
        simplified_tabs.append(
            {
                "tab_id": current_tab_id,
                "title": tab_properties.get("title"),
                "index": tab_properties.get("index"),
                "nesting_level": tab_properties.get("nestingLevel"),
                "content": simplify_structural_elements(
                    document_tab.get("body", {}).get("content", []),
                    document_tab,
                ),
                "inline_objects": inline_objects,
                "positioned_objects": positioned_objects,
            }
        )
    return {
        "document_id": document.get("documentId"),
        "title": document.get("title"),
        "revision_id": document.get("revisionId"),
        "tabs": simplified_tabs,
    }


def ensure_output_dir(output_dir: str | None, default_prefix: str) -> Path:
    if output_dir:
        target = Path(output_dir)
    else:
        target = Path(tempfile.mkdtemp(prefix=default_prefix))
    target.mkdir(parents=True, exist_ok=True)
    return target


def download_url(session: requests.Session, url: str, path: Path, timeout: int) -> None:
    response = session.get(url, timeout=timeout)
    response.raise_for_status()
    path.write_bytes(response.content)


def download_doc_images_payload(
    client: GoogleWorkspaceClient,
    document: dict[str, Any],
    *,
    output_dir: str | None,
    tab_id: str | None,
) -> dict[str, Any]:
    simplified = simplify_document(document, tab_id=tab_id)
    image_objects = []
    for tab in simplified["tabs"]:
        for image in tab.get("inline_objects", []) + tab.get("positioned_objects", []):
            if image.get("content_uri"):
                image_objects.append(
                    {
                        "tab_id": tab.get("tab_id"),
                        "tab_title": tab.get("title"),
                        **image,
                    }
                )

    folder = ensure_output_dir(output_dir, "google-doc-images-")
    downloads = []
    for index, image in enumerate(image_objects, start=1):
        extension = ".bin"
        source_uri = image.get("source_uri") or ""
        content_uri = image.get("content_uri") or ""
        for candidate in (source_uri, content_uri):
            lower = candidate.lower()
            if ".png" in lower:
                extension = ".png"
                break
            if ".jpg" in lower or ".jpeg" in lower:
                extension = ".jpg"
                break
            if ".gif" in lower:
                extension = ".gif"
                break
            if ".webp" in lower:
                extension = ".webp"
                break
        filename = f"{index:03d}_{safe_filename(image.get('object_id', 'image'))}{extension}"
        file_path = folder / filename
        download_url(client.session, image["content_uri"], file_path, client.timeout)
        downloads.append(
            {
                "object_id": image.get("object_id"),
                "tab_id": image.get("tab_id"),
                "tab_title": image.get("tab_title"),
                "path": str(file_path),
                "source_uri": image.get("source_uri"),
                "alt_text": image.get("alt_text"),
            }
        )
    return {
        "output_dir": str(folder),
        "count": len(downloads),
        "images": downloads,
    }

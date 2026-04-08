from __future__ import annotations

from pathlib import Path
from typing import Any

from .chat import (
    simplify_chat_membership,
    simplify_chat_message,
    simplify_chat_space,
)
from .client import get_client
from .common import (
    SHEET_FORMULA_FIELDS,
    SHEET_GRID_FIELDS,
    SHEETS_READ_SCOPE,
    SHEETS_SCOPE,
    XLSX_MIME,
    compact_dict,
    detect_google_file_kind,
    extract_file_id,
    parse_chat_url_context,
    parse_sheet_url_context,
    quote_sheet_title,
    safe_filename,
    split_sheet_range,
    unquote_sheet_title,
)
from .docs import download_doc_images_payload, simplify_document
from .server import mcp
from .sheets import (
    collect_formula_images,
    extract_sheet_images_from_xlsx,
    flatten_values_rows,
    grid_from_csv_rows,
    normalize_headers,
    parse_csv_rows,
    simplify_grid_data,
    values_from_csv_rows,
)


def _is_missing_sheets_scope_error(exc: RuntimeError) -> bool:
    return "spreadsheets" in str(exc).lower()


def _range_sheet_name(range_a1: str | None) -> str | None:
    if not range_a1:
        return None
    sheet_name, _ = split_sheet_range(range_a1)
    if not sheet_name:
        return None
    return unquote_sheet_title(sheet_name)


def _drive_export_csv_fallback(
    client: Any,
    spreadsheet_id_or_url: str,
    range_a1: str | None,
    *,
    major_dimension: str = "ROWS",
) -> dict[str, Any]:
    context = parse_sheet_url_context(spreadsheet_id_or_url)
    gid = context.get("gid")
    effective_range = range_a1.strip() if range_a1 else context.get("range_a1")
    if gid is None:
        raise RuntimeError("Drive export fallback requires a Google Sheets URL that includes `gid`.")

    csv_text = client.export_sheet_via_drive(
        spreadsheet_id_or_url,
        export_format="csv",
        gid=gid,
        range_a1=effective_range,
    ).decode("utf-8-sig")
    rows = parse_csv_rows(csv_text)
    return {
        "spreadsheet_id": extract_file_id(spreadsheet_id_or_url, kind="sheet"),
        "gid": gid,
        "range_a1": effective_range,
        "sheet_name": _range_sheet_name(effective_range),
        "rows": rows,
        "values": values_from_csv_rows(rows, major_dimension=major_dimension),
        "auth_warning": (
            "Cached OAuth token is missing Google Sheets API read scope "
            f"(`{SHEETS_READ_SCOPE}` or `{SHEETS_SCOPE}`), so this response came from "
            "Drive export fallback. Formulas, notes, hyperlinks, and rich text metadata may be omitted."
        ),
        "source": "drive_export_csv_fallback",
    }


@mcp.tool()
def diagnose_google_auth() -> dict[str, Any]:
    """Return a quick summary of the active Google authentication setup."""
    return get_client().auth_summary()


@mcp.tool()
def resolve_google_file(file_id_or_url: str) -> dict[str, Any]:
    """Resolve basic metadata for a Google Docs, Sheets, or Drive file."""
    client = get_client()
    try:
        metadata = client.get_drive_file(file_id_or_url)
        return {
            "id": metadata.get("id"),
            "name": metadata.get("name"),
            "mime_type": metadata.get("mimeType"),
            "web_view_link": metadata.get("webViewLink"),
            "owners": metadata.get("owners", []),
            "has_export_links": bool(metadata.get("exportLinks")),
        }
    except RuntimeError as exc:
        error_text = str(exc).lower()
        if "insufficient authentication scopes" not in error_text and "drive.readonly" not in error_text:
            raise

        file_kind = detect_google_file_kind(file_id_or_url)
        if file_kind == "sheet":
            metadata = client.get_sheet_metadata(file_id_or_url)
            spreadsheet_id = metadata.get("spreadsheetId") or extract_file_id(file_id_or_url, kind="sheet")
            return {
                "id": spreadsheet_id,
                "name": metadata.get("properties", {}).get("title"),
                "mime_type": "application/vnd.google-apps.spreadsheet",
                "web_view_link": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
                "owners": [],
                "has_export_links": False,
                "auth_warning": "Cached OAuth token is missing drive.readonly; rerun auth if you need Drive metadata.",
                "source": "sheets_metadata_fallback",
            }
        if file_kind == "doc":
            document = client.get_doc(file_id_or_url)
            document_id = extract_file_id(file_id_or_url, kind="doc")
            return {
                "id": document_id,
                "name": document.get("title"),
                "mime_type": "application/vnd.google-apps.document",
                "web_view_link": f"https://docs.google.com/document/d/{document_id}/edit",
                "owners": [],
                "has_export_links": False,
                "auth_warning": "Cached OAuth token is missing drive.readonly; rerun auth if you need Drive metadata.",
                "source": "docs_metadata_fallback",
            }
        raise RuntimeError(
            "Cached OAuth token is missing drive.readonly. Re-run `google-workspace-mcp auth` to refresh the token."
        ) from exc


@mcp.tool()
def list_google_chat_spaces(
    page_size: int = 100,
    page_token: str | None = None,
    filter_text: str | None = None,
) -> dict[str, Any]:
    """List Google Chat spaces the authenticated user can access."""
    client = get_client()
    payload = client.list_chat_spaces(
        page_size=page_size,
        page_token=page_token,
        filter_text=filter_text,
    )
    spaces = [simplify_chat_space(space) for space in payload.get("spaces", [])]
    return {
        "space_count": len(spaces),
        "spaces": spaces,
        "next_page_token": payload.get("nextPageToken"),
    }


@mcp.tool()
def get_google_chat_space(space_name_or_url: str) -> dict[str, Any]:
    """Read metadata for one Google Chat space from a resource name or Chat URL."""
    client = get_client()
    return simplify_chat_space(client.get_chat_space(space_name_or_url))


@mcp.tool()
def get_google_chat_message(message_name_or_url: str) -> dict[str, Any]:
    """Read one Google Chat message from a resource name or Chat thread URL."""
    client = get_client()
    return simplify_chat_message(client.get_chat_message(message_name_or_url))


@mcp.tool()
def read_google_chat_messages(
    space_name_or_url: str,
    page_size: int = 100,
    page_token: str | None = None,
    filter_text: str | None = None,
    order_by: str | None = "DESC",
    show_deleted: bool = False,
) -> dict[str, Any]:
    """Read messages in a Google Chat space from a resource name or Chat URL."""
    client = get_client()
    payload = client.list_chat_messages(
        space_name_or_url,
        page_size=page_size,
        page_token=page_token,
        filter_text=filter_text,
        order_by=order_by,
        show_deleted=show_deleted,
    )
    messages = [simplify_chat_message(message) for message in payload.get("messages", [])]
    return {
        "space": space_name_or_url,
        "message_count": len(messages),
        "messages": messages,
        "next_page_token": payload.get("nextPageToken"),
    }


@mcp.tool()
def read_google_chat_thread(
    thread_name_or_url: str,
    page_size: int = 100,
    page_token: str | None = None,
    order_by: str | None = "ASC",
    show_deleted: bool = False,
) -> dict[str, Any]:
    """Read one Google Chat thread from a thread URL or resource name, including the linked message and root message."""
    client = get_client()
    context = parse_chat_url_context(thread_name_or_url)
    payload = client.list_chat_thread_messages(
        thread_name_or_url,
        page_size=page_size,
        page_token=page_token,
        order_by=order_by,
        show_deleted=show_deleted,
    )
    messages = [simplify_chat_message(message) for message in payload.get("messages", [])]
    root_message = next((message for message in messages if not message.get("thread_reply")), None)
    if root_message is None and messages:
        root_message = messages[0]
    linked_message = None
    linked_message_lookup_warning = None
    if context.get("message_name"):
        linked_message = simplify_chat_message(client.get_chat_message(context["message_name"]))
    elif context.get("message_lookup_hint"):
        hint = context["message_lookup_hint"]
        linked_message = next(
            (
                message
                for message in messages
                if str(message.get("name", "")).endswith(f"/{hint}")
                or message.get("client_assigned_message_id") == hint
            ),
            None,
        )
        if linked_message is None:
            linked_message_lookup_warning = (
                "Google Chat room URLs don't expose a message resource ID that can always be resolved through the "
                "Chat API. The thread was loaded successfully, but the exact linked reply couldn't be mapped from "
                "the URL token alone."
            )
    return {
        "space": context.get("space_name"),
        "thread": context.get("thread_name"),
        "linked_message": linked_message,
        "linked_message_lookup_warning": linked_message_lookup_warning,
        "root_message": root_message,
        "message_count": len(messages),
        "messages": messages,
        "next_page_token": payload.get("nextPageToken"),
    }


@mcp.tool()
def list_google_chat_memberships(
    space_name_or_url: str,
    page_size: int = 100,
    page_token: str | None = None,
    filter_text: str | None = None,
    show_groups: bool = False,
    show_invited: bool = False,
) -> dict[str, Any]:
    """List users, bots, and optional groups invited to a Google Chat space."""
    client = get_client()
    payload = client.list_chat_memberships(
        space_name_or_url,
        page_size=page_size,
        page_token=page_token,
        filter_text=filter_text,
        show_groups=show_groups,
        show_invited=show_invited,
    )
    memberships = [simplify_chat_membership(item) for item in payload.get("memberships", [])]
    return {
        "space": space_name_or_url,
        "membership_count": len(memberships),
        "memberships": memberships,
        "next_page_token": payload.get("nextPageToken"),
    }


@mcp.tool()
def read_sheet_values(
    spreadsheet_id_or_url: str,
    range_a1: str | None = None,
    major_dimension: str = "ROWS",
    value_render_option: str = "FORMATTED_VALUE",
    date_time_render_option: str = "SERIAL_NUMBER",
) -> dict[str, Any]:
    """Read raw Google Sheets values for an A1 range."""
    client = get_client()
    try:
        payload = client.get_sheet_values(
            spreadsheet_id_or_url,
            range_a1,
            major_dimension=major_dimension,
            value_render_option=value_render_option,
            date_time_render_option=date_time_render_option,
        )
    except RuntimeError as exc:
        if not _is_missing_sheets_scope_error(exc):
            raise
        fallback = _drive_export_csv_fallback(
            client,
            spreadsheet_id_or_url,
            range_a1,
            major_dimension=major_dimension,
        )
        return {
            "spreadsheet_id": fallback["spreadsheet_id"],
            "range": fallback["range_a1"],
            "major_dimension": major_dimension,
            "row_count": len(fallback["values"]),
            "values": fallback["values"],
            "auth_warning": fallback["auth_warning"],
            "source": fallback["source"],
        }
    return {
        "spreadsheet_id": payload.get("spreadsheetId"),
        "range": payload.get("range"),
        "major_dimension": payload.get("majorDimension"),
        "row_count": len(payload.get("values", [])),
        "values": payload.get("values", []),
    }


@mcp.tool()
def read_sheet_grid(spreadsheet_id_or_url: str, range_a1: str | None = None) -> dict[str, Any]:
    """Read Google Sheets grid data including formatted values, formulas, notes, and links."""
    client = get_client()
    try:
        payload = client.get_sheet_grid(
            spreadsheet_id_or_url,
            range_a1,
            fields=SHEET_GRID_FIELDS,
        )
    except RuntimeError as exc:
        if not _is_missing_sheets_scope_error(exc):
            raise
        fallback = _drive_export_csv_fallback(client, spreadsheet_id_or_url, range_a1)
        title = None
        try:
            title = client.get_drive_file(spreadsheet_id_or_url).get("name")
        except RuntimeError:
            title = None
        return {
            "spreadsheet_id": fallback["spreadsheet_id"],
            "title": title,
            "sheets": grid_from_csv_rows(
                fallback["rows"],
                range_a1=fallback["range_a1"],
                sheet_name=fallback["sheet_name"],
                sheet_id=fallback["gid"],
            ),
            "auth_warning": fallback["auth_warning"],
            "source": fallback["source"],
        }
    return {
        "spreadsheet_id": payload.get("spreadsheetId"),
        "title": payload.get("properties", {}).get("title"),
        "sheets": simplify_grid_data(payload),
    }


@mcp.tool()
def get_sheet_row(
    spreadsheet_id_or_url: str,
    sheet_name: str | None,
    row_index: int,
    header_row: int = 1,
) -> dict[str, Any]:
    """Fetch one Google Sheets row and map it to the header row."""
    client = get_client()
    if not sheet_name:
        context = client.resolve_sheet_range_context(spreadsheet_id_or_url)
        sheet_name = context["resolved_sheet_name"]
    if not sheet_name:
        raise ValueError("Pass `sheet_name` explicitly or use a Google Sheets URL with `gid`.")
    header_payload = client.get_sheet_values(
        spreadsheet_id_or_url,
        f"{quote_sheet_title(sheet_name)}!{header_row}:{header_row}",
        major_dimension="ROWS",
        value_render_option="FORMATTED_VALUE",
        date_time_render_option="SERIAL_NUMBER",
    )
    row_payload = client.get_sheet_values(
        spreadsheet_id_or_url,
        f"{quote_sheet_title(sheet_name)}!{row_index}:{row_index}",
        major_dimension="ROWS",
        value_render_option="FORMATTED_VALUE",
        date_time_render_option="SERIAL_NUMBER",
    )
    headers = normalize_headers((header_payload.get("values") or [[]])[0])
    row_values = (row_payload.get("values") or [[]])[0]
    mapped = {header: row_values[index] if index < len(row_values) else None for index, header in enumerate(headers)}
    return {
        "sheet_name": sheet_name,
        "header_row": header_row,
        "row_index": row_index,
        "headers": headers,
        "values": row_values,
        "mapped": mapped,
    }


@mcp.tool()
def search_sheet(
    spreadsheet_id_or_url: str,
    needle: str,
    sheet_name: str | None = None,
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Search text across one sheet or all sheets and return matching cells."""
    client = get_client()
    if sheet_name:
        sheet_names = [sheet_name]
    else:
        context = client.resolve_sheet_range_context(spreadsheet_id_or_url)
        if context["resolved_sheet_name"]:
            sheet_names = [context["resolved_sheet_name"]]
        else:
            metadata = context["metadata"]
            sheet_names = [sheet["properties"]["title"] for sheet in metadata.get("sheets", [])]
    needle_cmp = needle if case_sensitive else needle.lower()
    matches = []
    for current_sheet_name in sheet_names:
        payload = client.get_sheet_grid(
            spreadsheet_id_or_url,
            quote_sheet_title(current_sheet_name),
            fields=SHEET_GRID_FIELDS,
        )
        haystack = simplify_grid_data(payload)
        for sheet in haystack:
            for section in sheet.get("data", []):
                for row in section.get("rows", []):
                    for cell in row.get("cells", []):
                        candidates = [
                            cell.get("formatted_value"),
                            cell.get("formula"),
                            cell.get("note"),
                            cell.get("hyperlink"),
                        ]
                        joined = " | ".join(str(value) for value in candidates if value)
                        if not joined:
                            continue
                        compare_target = joined if case_sensitive else joined.lower()
                        if needle_cmp in compare_target:
                            matches.append(
                                compact_dict(
                                    {
                                        "sheet_name": sheet.get("sheet_name"),
                                        "a1": cell.get("a1"),
                                        "formatted_value": cell.get("formatted_value"),
                                        "formula": cell.get("formula"),
                                        "note": cell.get("note"),
                                        "hyperlink": cell.get("hyperlink"),
                                    }
                                )
                            )
    return {"needle": needle, "match_count": len(matches), "matches": matches}


@mcp.tool()
def sheet_to_json(
    spreadsheet_id_or_url: str,
    sheet_name: str | None,
    header_row: int = 1,
    start_row: int | None = None,
    end_row: int | None = None,
) -> dict[str, Any]:
    """Convert a Google Sheets tab into JSON records using the header row."""
    client = get_client()
    if not sheet_name:
        context = client.resolve_sheet_range_context(spreadsheet_id_or_url)
        sheet_name = context["resolved_sheet_name"]
    if not sheet_name:
        raise ValueError("Pass `sheet_name` explicitly or use a Google Sheets URL with `gid`.")
    start = start_row or header_row
    if end_row:
        range_a1 = f"{quote_sheet_title(sheet_name)}!{start}:{end_row}"
    else:
        range_a1 = f"{quote_sheet_title(sheet_name)}!{start}:999999"
    payload = client.get_sheet_values(
        spreadsheet_id_or_url,
        range_a1,
        major_dimension="ROWS",
        value_render_option="FORMATTED_VALUE",
        date_time_render_option="SERIAL_NUMBER",
    )
    rows = flatten_values_rows(payload)
    if not rows:
        return {"sheet_name": sheet_name, "header_row": header_row, "rows": []}

    header_offset = max(header_row - start, 0)
    if header_offset >= len(rows):
        return {"sheet_name": sheet_name, "header_row": header_row, "rows": []}

    headers = normalize_headers(rows[header_offset])
    records = []
    for absolute_offset, row_values in enumerate(rows[header_offset + 1 :], start=header_row + 1):
        if not any(str(value).strip() for value in row_values):
            continue
        records.append(
            {
                "_row": absolute_offset,
                **{header: row_values[index] if index < len(row_values) else None for index, header in enumerate(headers)},
            }
        )
    return {
        "sheet_name": sheet_name,
        "header_row": header_row,
        "headers": headers,
        "row_count": len(records),
        "rows": records,
    }


@mcp.tool()
def inspect_sheet_images(
    spreadsheet_id_or_url: str,
    sheet_name: str | None = None,
    output_dir: str | None = None,
) -> dict[str, Any]:
    """Inspect Google Sheets images via XLSX export and detect IMAGE() formulas."""
    client = get_client()
    metadata, xlsx_bytes = client.export_drive_file(spreadsheet_id_or_url, XLSX_MIME)
    exported_images = extract_sheet_images_from_xlsx(
        xlsx_bytes,
        output_dir=output_dir,
        sheet_name=sheet_name,
    )
    formula_images = []
    if sheet_name:
        ranges = [quote_sheet_title(sheet_name)]
    else:
        context = client.resolve_sheet_range_context(spreadsheet_id_or_url)
        if context["resolved_sheet_name"]:
            ranges = [quote_sheet_title(context["resolved_sheet_name"])]
        else:
            sheet_metadata = client.get_sheet_metadata(spreadsheet_id_or_url)
            ranges = [quote_sheet_title(sheet["properties"]["title"]) for sheet in sheet_metadata.get("sheets", [])]
    for current_range in ranges:
        formula_payload = client.get_sheet_grid(
            spreadsheet_id_or_url,
            current_range,
            fields=SHEET_FORMULA_FIELDS,
        )
        formula_images.extend(collect_formula_images(formula_payload))
    return {
        "spreadsheet_id": metadata.get("id"),
        "name": metadata.get("name"),
        "exported_drawing_images": exported_images,
        "image_formulas": formula_images,
        "notes": [
            "Over-grid images are extracted from the XLSX export.",
            "In-cell IMAGE() formulas are detected from Sheets grid data.",
            "Private files require Drive read access via a service account or OAuth token.",
        ],
    }


@mcp.tool()
def read_google_doc(
    document_id_or_url: str,
    tab_id: str | None = None,
    download_images: bool = False,
    output_dir: str | None = None,
) -> dict[str, Any]:
    """Read a Google Doc as structured JSON with text, tables, and image metadata."""
    client = get_client()
    document = client.get_doc(document_id_or_url)
    payload = simplify_document(document, tab_id=tab_id)
    if download_images:
        payload["downloaded_images"] = download_doc_images_payload(
            client,
            document,
            output_dir=output_dir,
            tab_id=tab_id,
        )
    return payload


@mcp.tool()
def download_google_doc_images(
    document_id_or_url: str,
    output_dir: str | None = None,
    tab_id: str | None = None,
) -> dict[str, Any]:
    """Download image objects from a Google Doc to a local folder."""
    client = get_client()
    document = client.get_doc(document_id_or_url)
    return download_doc_images_payload(client, document, output_dir=output_dir, tab_id=tab_id)


@mcp.tool()
def export_google_file(
    file_id_or_url: str,
    mime_type: str,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Export a Google Workspace file to PDF, XLSX, HTML zip, Markdown, or plain text."""
    client = get_client()
    metadata, content = client.export_drive_file(file_id_or_url, mime_type)
    if output_path:
        destination = Path(output_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
    else:
        folder = client.export_root
        folder.mkdir(parents=True, exist_ok=True)
        extension = {
            "application/pdf": ".pdf",
            "application/zip": ".zip",
            "text/markdown": ".md",
            "text/plain": ".txt",
            XLSX_MIME: ".xlsx",
        }.get(mime_type, ".bin")
        destination = folder / f"{safe_filename(metadata.get('name', metadata.get('id', 'export')))}{extension}"
    destination.write_bytes(content)
    return {
        "file_id": metadata.get("id"),
        "name": metadata.get("name"),
        "mime_type": mime_type,
        "output_path": str(destination),
        "bytes": len(content),
    }

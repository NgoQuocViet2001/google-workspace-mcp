from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any
from urllib.parse import parse_qsl, quote, urlparse


DOCS_SCOPE = "https://www.googleapis.com/auth/documents.readonly"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly"
SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
CHAT_SPACES_SCOPE = "https://www.googleapis.com/auth/chat.spaces.readonly"
CHAT_MESSAGES_SCOPE = "https://www.googleapis.com/auth/chat.messages.readonly"
CHAT_MEMBERSHIPS_SCOPE = "https://www.googleapis.com/auth/chat.memberships.readonly"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_READONLY_SCOPES = [
    DOCS_SCOPE,
    DRIVE_SCOPE,
    SHEETS_SCOPE,
    CHAT_SPACES_SCOPE,
    CHAT_MESSAGES_SCOPE,
    CHAT_MEMBERSHIPS_SCOPE,
]

DOC_URL_RE = re.compile(r"https?://docs\.google\.com/document/d/([a-zA-Z0-9_-]+)")
SHEET_URL_RE = re.compile(r"https?://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)")
DRIVE_FILE_RE = re.compile(r"https?://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)")
DRIVE_OPEN_RE = re.compile(r"[?&]id=([a-zA-Z0-9_-]+)")
CHAT_SPACE_RESOURCE_RE = re.compile(r"^spaces/[A-Za-z0-9._=-]+$")
CHAT_THREAD_RESOURCE_RE = re.compile(r"^spaces/(?P<space>[A-Za-z0-9._=-]+)/threads/(?P<thread>[A-Za-z0-9._=-]+)$")
CHAT_MESSAGE_RESOURCE_RE = re.compile(
    r"^spaces/(?P<space>[A-Za-z0-9._=-]+)/messages/(?P<message>[A-Za-z0-9._=-]+)$"
)
CHAT_SPACE_EMBEDDED_RE = re.compile(r"spaces/[A-Za-z0-9._=-]+")
CHAT_THREAD_EMBEDDED_RE = re.compile(r"spaces/[A-Za-z0-9._=-]+/threads/[A-Za-z0-9._=-]+")
CHAT_MESSAGE_EMBEDDED_RE = re.compile(r"spaces/[A-Za-z0-9._=-]+/messages/[A-Za-z0-9._=-]+")
CHAT_SPACE_UI_RE = re.compile(r"#chat/(?:space|dm)/([A-Za-z0-9._=-]+)")
CHAT_ROOM_RE = re.compile(r"/room/([A-Za-z0-9._=-]+)")
CHAT_ID_RE = re.compile(r"^[A-Za-z0-9._=-]{3,}$")
IMAGE_FORMULA_RE = re.compile(r'(?is)^=IMAGE\(\s*"([^"]+)"')
ID_RE = re.compile(r"^[a-zA-Z0-9_-]{15,}$")
ROW_ONLY_RANGE_RE = re.compile(r"^\d+(?::\d+)?$")
MAX_SHEET_COLUMN_A1 = "ZZZ"

SHEET_GRID_FIELDS = ",".join(
    [
        "spreadsheetId",
        "properties.title",
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.properties.index",
        "sheets.properties.gridProperties.rowCount",
        "sheets.properties.gridProperties.columnCount",
        "sheets.data.startRow",
        "sheets.data.startColumn",
        "sheets.data.rowData.values.formattedValue",
        "sheets.data.rowData.values.hyperlink",
        "sheets.data.rowData.values.note",
        "sheets.data.rowData.values.userEnteredValue",
        "sheets.data.rowData.values.effectiveValue",
        "sheets.data.rowData.values.textFormatRuns",
        "sheets.data.rowData.values.chipRuns",
        "sheets.data.rowData.values.userEnteredFormat.textFormat.link",
    ]
)

SHEET_FORMULA_FIELDS = ",".join(
    [
        "spreadsheetId",
        "sheets.properties.title",
        "sheets.data.startRow",
        "sheets.data.startColumn",
        "sheets.data.rowData.values.userEnteredValue",
    ]
)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}


def compact_dict(values: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in values.items() if value not in (None, "", [], {})}


def scalar_value(value: dict[str, Any] | None) -> Any:
    if not value:
        return None
    for key in ("stringValue", "numberValue", "boolValue", "formulaValue", "errorValue"):
        if key in value:
            return value[key]
    return value


def path_from_env(value: str | None) -> Path | None:
    if not value:
        return None
    return Path(os.path.expandvars(os.path.expanduser(value)))


def default_oauth_token_file() -> Path:
    return Path.home() / ".google-workspace-mcp" / "oauth-token.json"


def default_oauth_client_secrets_file() -> Path:
    return Path.home() / ".google-workspace-mcp" / "oauth-client-secret.json"


def normalize_scopes(raw_scopes: Any) -> list[str]:
    if raw_scopes is None:
        return []
    if isinstance(raw_scopes, str):
        return sorted({scope for scope in re.split(r"[\s,]+", raw_scopes) if scope})
    if isinstance(raw_scopes, (list, tuple, set)):
        return sorted({str(scope).strip() for scope in raw_scopes if str(scope).strip()})
    return []


def extract_file_id(value: str, kind: str | None = None) -> str:
    value = value.strip()
    if kind == "doc":
        match = DOC_URL_RE.search(value)
        if match:
            return match.group(1)
    elif kind == "sheet":
        match = SHEET_URL_RE.search(value)
        if match:
            return match.group(1)
    else:
        for pattern in (DOC_URL_RE, SHEET_URL_RE, DRIVE_FILE_RE):
            match = pattern.search(value)
            if match:
                return match.group(1)
        match = DRIVE_OPEN_RE.search(value)
        if match:
            return match.group(1)

    if ID_RE.match(value):
        return value
    raise ValueError(f"Could not extract a Google file ID from: {value}")


def detect_google_file_kind(value: str) -> str | None:
    trimmed = value.strip()
    if DOC_URL_RE.search(trimmed):
        return "doc"
    if SHEET_URL_RE.search(trimmed):
        return "sheet"
    if DRIVE_FILE_RE.search(trimmed) or DRIVE_OPEN_RE.search(trimmed):
        return "drive"
    return None


def parse_chat_url_context(value: str) -> dict[str, Any]:
    trimmed = value.strip()
    if not trimmed:
        raise ValueError("Could not extract a Google Chat space name from an empty value.")

    message_match = CHAT_MESSAGE_RESOURCE_RE.fullmatch(trimmed)
    if message_match:
        space_id = message_match.group("space")
        message_id = message_match.group("message")
        return {
            "space_id": space_id,
            "space_name": f"spaces/{space_id}",
            "thread_id": None,
            "thread_name": None,
            "message_id": message_id,
            "message_name": trimmed,
        }

    thread_match = CHAT_THREAD_RESOURCE_RE.fullmatch(trimmed)
    if thread_match:
        space_id = thread_match.group("space")
        thread_id = thread_match.group("thread")
        return {
            "space_id": space_id,
            "space_name": f"spaces/{space_id}",
            "thread_id": thread_id,
            "thread_name": trimmed,
            "message_id": None,
            "message_name": None,
        }

    if CHAT_SPACE_RESOURCE_RE.fullmatch(trimmed):
        space_id = trimmed.split("/", 1)[1]
        return {
            "space_id": space_id,
            "space_name": trimmed,
            "thread_id": None,
            "thread_name": None,
            "message_id": None,
            "message_name": None,
        }

    embedded_message_match = CHAT_MESSAGE_EMBEDDED_RE.search(trimmed)
    if embedded_message_match:
        message_name = embedded_message_match.group(0)
        context = parse_chat_url_context(message_name)
        return context

    embedded_thread_match = CHAT_THREAD_EMBEDDED_RE.search(trimmed)
    if embedded_thread_match:
        thread_name = embedded_thread_match.group(0)
        context = parse_chat_url_context(thread_name)
        return context

    embedded_match = CHAT_SPACE_EMBEDDED_RE.search(trimmed)
    if embedded_match:
        return parse_chat_url_context(embedded_match.group(0))

    chat_ui_match = CHAT_SPACE_UI_RE.search(trimmed)
    if chat_ui_match:
        return parse_chat_url_context(chat_ui_match.group(1))

    parsed = urlparse(trimmed)
    path_segments = [segment for segment in parsed.path.split("/") if segment]
    if "room" in path_segments:
        room_index = path_segments.index("room")
        space_id = path_segments[room_index + 1] if len(path_segments) > room_index + 1 else None
        thread_id = path_segments[room_index + 2] if len(path_segments) > room_index + 2 else None
        message_id = path_segments[room_index + 3] if len(path_segments) > room_index + 3 else None
        if space_id:
            return {
                "space_id": space_id,
                "space_name": f"spaces/{space_id}",
                "thread_id": thread_id,
                "thread_name": f"spaces/{space_id}/threads/{thread_id}" if thread_id else None,
                "message_id": message_id,
                "message_name": f"spaces/{space_id}/messages/{message_id}" if message_id else None,
            }

    for pattern in (CHAT_ROOM_RE,):
        match = pattern.search(trimmed)
        if match:
            return parse_chat_url_context(match.group(1))

    if CHAT_ID_RE.fullmatch(trimmed):
        return {
            "space_id": trimmed,
            "space_name": f"spaces/{trimmed}",
            "thread_id": None,
            "thread_name": None,
            "message_id": None,
            "message_name": None,
        }

    raise ValueError(f"Could not extract a Google Chat space name from: {value}")


def extract_chat_space_name(value: str) -> str:
    return parse_chat_url_context(value)["space_name"]


def extract_chat_thread_name(value: str) -> str:
    thread_name = parse_chat_url_context(value).get("thread_name")
    if thread_name:
        return thread_name
    raise ValueError(f"Could not extract a Google Chat thread name from: {value}")


def extract_chat_message_name(value: str) -> str:
    message_name = parse_chat_url_context(value).get("message_name")
    if message_name:
        return message_name
    raise ValueError(f"Could not extract a Google Chat message name from: {value}")


def parse_sheet_url_context(value: str) -> dict[str, Any]:
    spreadsheet_id = extract_file_id(value, kind="sheet")
    if not SHEET_URL_RE.search(value):
        return {
            "spreadsheet_id": spreadsheet_id,
            "gid": None,
            "range_a1": None,
        }

    parsed = urlparse(value)
    params: dict[str, str] = {}
    for segment in (parsed.query, parsed.fragment):
        params.update(dict(parse_qsl(segment, keep_blank_values=True)))
    gid_text = params.get("gid", "").strip()
    range_a1 = params.get("range", "").strip() or None
    gid = int(gid_text) if gid_text.isdigit() else None
    return {
        "spreadsheet_id": spreadsheet_id,
        "gid": gid,
        "range_a1": range_a1,
    }


def quote_range(range_a1: str) -> str:
    return quote(range_a1, safe="!():,$")


def quote_sheet_title(sheet_name: str) -> str:
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'"


def unquote_sheet_title(sheet_name: str) -> str:
    if len(sheet_name) >= 2 and sheet_name[0] == "'" and sheet_name[-1] == "'":
        return sheet_name[1:-1].replace("''", "'")
    return sheet_name


def split_sheet_range(range_a1: str) -> tuple[str | None, str]:
    trimmed = range_a1.strip()
    if "!" not in trimmed:
        return None, trimmed
    sheet_name, body = trimmed.split("!", 1)
    return sheet_name, body.strip()


def normalize_values_range(range_a1: str, *, max_column_a1: str = MAX_SHEET_COLUMN_A1) -> str:
    trimmed = range_a1.strip()
    sheet_name, body = split_sheet_range(trimmed)
    sheet_prefix = f"{sheet_name}!" if sheet_name else ""
    if not ROW_ONLY_RANGE_RE.fullmatch(body):
        return trimmed
    if ":" in body:
        start_row, end_row = body.split(":", 1)
    else:
        start_row = end_row = body
    return f"{sheet_prefix}A{start_row}:{max_column_a1}{end_row}"


def column_to_a1(column_index_zero_based: int) -> str:
    column_number = column_index_zero_based + 1
    letters = []
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def a1_from_zero_based(row_index_zero_based: int, column_index_zero_based: int) -> str:
    return f"{column_to_a1(column_index_zero_based)}{row_index_zero_based + 1}"


def safe_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._") or "file"


def rel_join(base: str, target: str) -> str:
    parts = base.split("/")
    if parts and "." in parts[-1]:
        parts = parts[:-1]
    for piece in target.split("/"):
        if piece in ("", "."):
            continue
        if piece == "..":
            if parts:
                parts.pop()
        else:
            parts.append(piece)
    return "/".join(parts)


def text_style_summary(text_style: dict[str, Any] | None) -> dict[str, Any]:
    if not text_style:
        return {}
    link_url = (
        text_style.get("link", {}).get("url")
        or text_style.get("link", {}).get("bookmarkId")
        or text_style.get("link", {}).get("headingId")
    )
    return compact_dict(
        {
            "bold": text_style.get("bold"),
            "italic": text_style.get("italic"),
            "underline": text_style.get("underline"),
            "strikethrough": text_style.get("strikethrough"),
            "small_caps": text_style.get("smallCaps"),
            "baseline_offset": text_style.get("baselineOffset"),
            "font_size": text_style.get("fontSize"),
            "link": link_url,
        }
    )

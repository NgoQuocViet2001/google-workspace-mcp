from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import time
import zipfile
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import parse_qsl, quote, urlparse
from xml.etree import ElementTree as ET

import requests
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials as UserOAuthCredentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from mcp.server.fastmcp import FastMCP


DOCS_SCOPE = "https://www.googleapis.com/auth/documents.readonly"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly"
SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_READONLY_SCOPES = [DOCS_SCOPE, DRIVE_SCOPE, SHEETS_SCOPE]

DOC_URL_RE = re.compile(r"https?://docs\.google\.com/document/d/([a-zA-Z0-9_-]+)")
SHEET_URL_RE = re.compile(r"https?://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)")
DRIVE_FILE_RE = re.compile(r"https?://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)")
DRIVE_OPEN_RE = re.compile(r"[?&]id=([a-zA-Z0-9_-]+)")
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

mcp = FastMCP("Google Workspace MCP", json_response=True)


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
    return re.sub(r'[^A-Za-z0-9._-]+', "_", name).strip("._") or "file"


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


class GoogleWorkspaceClient:
    def __init__(self) -> None:
        self.api_key = os.getenv("GOOGLE_API_KEY")
        self.oauth_access_token = os.getenv("GOOGLE_OAUTH_ACCESS_TOKEN")
        self.oauth_client_secrets_file = path_from_env(os.getenv("GOOGLE_OAUTH_CLIENT_SECRETS_FILE"))
        self.oauth_client_config_json = os.getenv("GOOGLE_OAUTH_CLIENT_CONFIG_JSON")
        self.oauth_token_file = (
            path_from_env(os.getenv("GOOGLE_OAUTH_TOKEN_FILE")) or default_oauth_token_file()
        )
        self.oauth_local_server_port = int(os.getenv("GOOGLE_OAUTH_LOCAL_SERVER_PORT", "0"))
        self.oauth_open_browser = os.getenv("GOOGLE_OAUTH_OPEN_BROWSER", "true").lower() not in {
            "0",
            "false",
            "no",
        }
        self.service_account_file = path_from_env(os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE"))
        self.service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
        self.timeout = int(os.getenv("GOOGLE_HTTP_TIMEOUT_SECONDS", "60"))
        export_dir = path_from_env(os.getenv("GOOGLE_WORKSPACE_EXPORT_DIR"))
        self.export_root = export_dir or Path.cwd() / "exports"
        self._base_service_account: ServiceAccountCredentials | None = None
        self._scoped_credentials: dict[tuple[str, ...], ServiceAccountCredentials] = {}
        self._user_credentials: dict[tuple[str, ...], UserOAuthCredentials] = {}
        self._sheet_metadata_cache: dict[str, dict[str, Any]] = {}
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "google-workspace-mcp/1.0"})

    def _cached_oauth_token_payload(self) -> dict[str, Any] | None:
        if not self.oauth_token_file.exists():
            return None
        try:
            data = json.loads(self.oauth_token_file.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
        return data if isinstance(data, dict) else None

    def _cached_oauth_token_format(self) -> str | None:
        payload = self._cached_oauth_token_payload()
        if not payload:
            return None
        if "installed" in payload or "web" in payload:
            return "oauth_client_secret"
        required = {"client_id", "client_secret", "refresh_token"}
        if required.issubset(payload):
            return "authorized_user"
        return "unknown_json"

    def _cached_oauth_token_scopes(self) -> list[str]:
        payload = self._cached_oauth_token_payload()
        if not payload:
            return []
        return normalize_scopes(payload.get("scopes"))

    def _missing_cached_oauth_scopes(self, required_scopes: Iterable[str]) -> list[str]:
        cached_scopes = set(self._cached_oauth_token_scopes())
        return [scope for scope in sorted(set(required_scopes)) if scope not in cached_scopes]

    def auth_summary(self) -> dict[str, Any]:
        oauth_client_ready = bool(self.oauth_client_secrets_file or self.oauth_client_config_json)
        oauth_client_source = None
        if self.oauth_client_secrets_file:
            oauth_client_source = str(self.oauth_client_secrets_file)
        elif self.oauth_client_config_json:
            oauth_client_source = "GOOGLE_OAUTH_CLIENT_CONFIG_JSON"
        service_account_ready = bool(self.service_account_file or self.service_account_json)
        service_account_source = None
        if self.service_account_file:
            service_account_source = str(self.service_account_file)
        elif self.service_account_json:
            service_account_source = "GOOGLE_SERVICE_ACCOUNT_JSON"
        cached_oauth_scopes = self._cached_oauth_token_scopes()
        return {
            "api_key_configured": bool(self.api_key),
            "oauth_access_token_configured": bool(self.oauth_access_token),
            "oauth_client_configured": oauth_client_ready,
            "oauth_client_source": oauth_client_source,
            "oauth_token_file": str(self.oauth_token_file),
            "oauth_token_cached": self.oauth_token_file.exists(),
            "oauth_token_format": self._cached_oauth_token_format(),
            "oauth_token_scopes": cached_oauth_scopes,
            "oauth_token_missing_scopes": self._missing_cached_oauth_scopes(DEFAULT_READONLY_SCOPES),
            "service_account_configured": service_account_ready,
            "service_account_source": service_account_source,
            "recommended_mode": self._recommended_mode(),
            "active_auth_mode": self._active_auth_mode(),
            "notes": [
                "Public Sheets can be read with GOOGLE_API_KEY.",
                "OAuth desktop client credentials can read private files shared to your Google account.",
                "Docs API and Drive export are most reliable with a service account or OAuth access token.",
                "A service account must be granted access to private files or shared drives.",
            ],
        }

    def _recommended_mode(self) -> str:
        if self.oauth_access_token:
            return "oauth_access_token"
        if self.oauth_client_secrets_file or self.oauth_client_config_json:
            return "oauth_client"
        if self.service_account_file or self.service_account_json:
            return "service_account"
        if self.api_key:
            return "api_key_public_only"
        return "missing_credentials"

    def _active_auth_mode(self) -> str:
        if self.oauth_access_token:
            return "oauth_access_token"
        if self._oauth_client_is_configured():
            if self.oauth_token_file.exists():
                return "oauth_client_cached_token"
            return "oauth_client_not_authorized"
        if self.service_account_file or self.service_account_json:
            return "service_account"
        if self.api_key:
            return "api_key_public_only"
        return "missing_credentials"

    def _service_account_base(self) -> ServiceAccountCredentials:
        if self._base_service_account is not None:
            return self._base_service_account
        if self.service_account_file:
            self._base_service_account = ServiceAccountCredentials.from_service_account_file(
                str(self.service_account_file)
            )
            return self._base_service_account
        if self.service_account_json:
            info = json.loads(self.service_account_json)
            self._base_service_account = ServiceAccountCredentials.from_service_account_info(info)
            return self._base_service_account
        raise RuntimeError(
            "No service account is configured. Set GOOGLE_SERVICE_ACCOUNT_FILE or "
            "GOOGLE_SERVICE_ACCOUNT_JSON."
        )

    def _oauth_client_is_configured(self) -> bool:
        return bool(self.oauth_client_secrets_file or self.oauth_client_config_json)

    def _oauth_flow(self, scopes: Iterable[str]) -> InstalledAppFlow:
        if self.oauth_client_secrets_file:
            return InstalledAppFlow.from_client_secrets_file(
                str(self.oauth_client_secrets_file),
                list(scopes),
            )
        if self.oauth_client_config_json:
            return InstalledAppFlow.from_client_config(
                json.loads(self.oauth_client_config_json),
                list(scopes),
            )
        raise RuntimeError(
            "No OAuth client credentials are configured. Set GOOGLE_OAUTH_CLIENT_SECRETS_FILE "
            "or GOOGLE_OAUTH_CLIENT_CONFIG_JSON."
        )

    def _save_user_credentials(self, credentials: UserOAuthCredentials) -> None:
        self.oauth_token_file.parent.mkdir(parents=True, exist_ok=True)
        self.oauth_token_file.write_text(credentials.to_json(), encoding="utf-8")

    def run_oauth_login(
        self,
        scopes: Iterable[str] | None = None,
        *,
        open_browser: bool | None = None,
        port: int | None = None,
    ) -> dict[str, Any]:
        requested_scopes = list(scopes or DEFAULT_READONLY_SCOPES)
        flow = self._oauth_flow(requested_scopes)
        credentials = flow.run_local_server(
            port=self.oauth_local_server_port if port is None else port,
            open_browser=self.oauth_open_browser if open_browser is None else open_browser,
            authorization_prompt_message="Open this URL in your browser to authorize access: {url}",
            success_message="Google Workspace MCP authorization completed. You can close this window.",
        )
        self._save_user_credentials(credentials)
        self._user_credentials.clear()
        return {
            "oauth_token_file": str(self.oauth_token_file),
            "scopes": requested_scopes,
            "account": getattr(credentials, "account", None),
            "has_refresh_token": bool(credentials.refresh_token),
        }

    def _sheet_properties_by_title(
        self,
        metadata: dict[str, Any],
        sheet_name: str | None,
    ) -> dict[str, Any] | None:
        if not sheet_name:
            return None
        requested_name = unquote_sheet_title(sheet_name)
        for sheet in metadata.get("sheets", []):
            properties = sheet.get("properties", {})
            if properties.get("title") == requested_name:
                return properties
        return None

    def _sheet_properties_by_gid(
        self,
        metadata: dict[str, Any],
        gid: int | None,
    ) -> dict[str, Any] | None:
        if gid is None:
            return None
        for sheet in metadata.get("sheets", []):
            properties = sheet.get("properties", {})
            if properties.get("sheetId") == gid:
                return properties
        return None

    def resolve_sheet_range_context(
        self,
        spreadsheet_id_or_url: str,
        *,
        range_a1: str | None = None,
        sheet_name: str | None = None,
    ) -> dict[str, Any]:
        spreadsheet_id = extract_file_id(spreadsheet_id_or_url, kind="sheet")
        metadata = self.get_sheet_metadata(spreadsheet_id)
        url_context = parse_sheet_url_context(spreadsheet_id_or_url)
        explicit_range = range_a1.strip() if range_a1 else None
        range_sheet_name, range_body = split_sheet_range(explicit_range) if explicit_range else (None, None)

        resolved_properties = (
            self._sheet_properties_by_title(metadata, sheet_name)
            or self._sheet_properties_by_title(metadata, range_sheet_name)
            or self._sheet_properties_by_gid(metadata, url_context.get("gid"))
        )
        default_properties = resolved_properties
        if default_properties is None:
            sheets = metadata.get("sheets", [])
            if sheets:
                default_properties = sheets[0].get("properties", {})

        effective_range = explicit_range or url_context.get("range_a1")
        if effective_range:
            prefix, body = split_sheet_range(effective_range)
            is_whole_sheet_reference = (
                resolved_properties is not None
                and prefix is None
                and unquote_sheet_title(body) == resolved_properties.get("title")
            )
            if is_whole_sheet_reference:
                effective_range = quote_sheet_title(resolved_properties["title"])
            elif resolved_properties and (sheet_name or prefix is None):
                effective_range = f"{quote_sheet_title(resolved_properties['title'])}!{body}"
        elif resolved_properties:
            effective_range = quote_sheet_title(resolved_properties["title"])

        return {
            "spreadsheet_id": spreadsheet_id,
            "url_context": url_context,
            "metadata": metadata,
            "resolved_sheet_name": resolved_properties.get("title") if resolved_properties else None,
            "resolved_gid": resolved_properties.get("sheetId") if resolved_properties else url_context.get("gid"),
            "resolved_range_a1": effective_range,
            "resolved_sheet_properties": resolved_properties,
            "default_sheet_properties": default_properties,
            "range_body": range_body,
        }

    def _user_oauth_credentials(self, scopes: Iterable[str]) -> UserOAuthCredentials:
        scope_list = tuple(sorted(set(scopes)))
        credentials = self._user_credentials.get(scope_list)
        if credentials is not None and credentials.valid and credentials.token:
            return credentials

        if not self._oauth_client_is_configured():
            raise RuntimeError(
                "No OAuth client configuration found. Set GOOGLE_OAUTH_CLIENT_SECRETS_FILE or "
                "GOOGLE_OAUTH_CLIENT_CONFIG_JSON, then run `google-workspace-mcp auth`."
            )

        if not self.oauth_token_file.exists():
            raise RuntimeError(
                "OAuth client credentials are configured, but no cached user token was found. "
                "Run `google-workspace-mcp auth` to complete the browser login flow."
            )

        missing_scopes = self._missing_cached_oauth_scopes(scope_list)
        if missing_scopes:
            missing_display = ", ".join(missing_scopes)
            raise RuntimeError(
                "Cached OAuth token is missing required scopes: "
                f"{missing_display}. Re-run `google-workspace-mcp auth` to refresh the token."
            )

        try:
            credentials = UserOAuthCredentials.from_authorized_user_file(
                str(self.oauth_token_file),
                list(scope_list),
            )
        except ValueError as exc:
            raise RuntimeError(
                "Cached OAuth token file is not in the authorized-user format. "
                "Run `google-workspace-mcp auth` again to regenerate it."
            ) from exc

        if credentials.expired and credentials.refresh_token:
            credentials.refresh(GoogleAuthRequest())
            self._save_user_credentials(credentials)

        if not credentials.valid or not credentials.token:
            raise RuntimeError(
                "The cached OAuth token is invalid or missing. Run `google-workspace-mcp auth` again."
            )

        self._user_credentials[scope_list] = credentials
        return credentials

    def _auth_headers(self, scopes: Iterable[str], allow_api_key: bool) -> tuple[dict[str, str], dict[str, str]]:
        scope_list = tuple(sorted(set(scopes)))
        headers: dict[str, str] = {}
        params: dict[str, str] = {}

        if self.oauth_access_token:
            headers["Authorization"] = f"Bearer {self.oauth_access_token}"
            return headers, params

        if scope_list and self._oauth_client_is_configured():
            credentials = self._user_oauth_credentials(scope_list)
            headers["Authorization"] = f"Bearer {credentials.token}"
            return headers, params

        if scope_list and (self.service_account_file or self.service_account_json):
            creds = self._scoped_credentials.get(scope_list)
            if creds is None:
                creds = self._service_account_base().with_scopes(scope_list)
                self._scoped_credentials[scope_list] = creds
            if not creds.valid or not creds.token:
                creds.refresh(GoogleAuthRequest())
            headers["Authorization"] = f"Bearer {creds.token}"
            return headers, params

        if allow_api_key and self.api_key:
            params["key"] = self.api_key
            return headers, params

        mode = "missing credentials"
        if allow_api_key:
            mode = "requires GOOGLE_API_KEY or credentials with read access"
        raise RuntimeError(
            "No valid Google credentials are configured. "
            f"This operation {mode}. Use an OAuth desktop client, service account, or OAuth token "
            "for private Docs/Drive, or GOOGLE_API_KEY for public Sheets."
        )

    def _retry_delay_seconds(self, response: requests.Response, attempt_index_zero_based: int) -> float:
        retry_after = response.headers.get("Retry-After")
        if retry_after:
            retry_after = retry_after.strip()
            if retry_after.isdigit():
                return max(float(retry_after), 0.0)
            try:
                retry_after_dt = parsedate_to_datetime(retry_after)
                if retry_after_dt.tzinfo is None:
                    retry_after_dt = retry_after_dt.replace(tzinfo=timezone.utc)
                return max((retry_after_dt - datetime.now(timezone.utc)).total_seconds(), 0.0)
            except (TypeError, ValueError, OverflowError):
                pass
        return min(2**attempt_index_zero_based, 8.0)

    def _request(
        self,
        method: str,
        url: str,
        *,
        scopes: Iterable[str] = (),
        params: dict[str, Any] | None = None,
        allow_api_key: bool = False,
        expect_json: bool = True,
    ) -> Any:
        headers, auth_params = self._auth_headers(scopes, allow_api_key)
        final_params = {**auth_params, **(params or {})}
        response: requests.Response | None = None
        for attempt_index in range(4):
            response = self.session.request(
                method,
                url,
                params=final_params,
                headers=headers,
                timeout=self.timeout,
            )
            if response.ok or response.status_code not in {429, 500, 502, 503, 504} or attempt_index == 3:
                break
            time.sleep(self._retry_delay_seconds(response, attempt_index))
        assert response is not None
        if not response.ok:
            error_message = response.text
            try:
                payload = response.json()
                error = payload.get("error", {})
                if isinstance(error, dict):
                    error_message = error.get("message", error_message)
            except ValueError:
                pass
            raise RuntimeError(
                f"Google API returned HTTP {response.status_code} for {url}: {error_message}"
            )
        if expect_json:
            return response.json()
        return response.content

    def get_sheet_metadata(self, spreadsheet_id_or_url: str) -> dict[str, Any]:
        spreadsheet_id = extract_file_id(spreadsheet_id_or_url, kind="sheet")
        cached = self._sheet_metadata_cache.get(spreadsheet_id)
        if cached is not None:
            return cached
        metadata = self._request(
            "GET",
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}",
            scopes=[SHEETS_SCOPE],
            allow_api_key=True,
            params={
                "fields": "spreadsheetId,properties.title,sheets.properties.sheetId,"
                "sheets.properties.title,sheets.properties.index,"
                "sheets.properties.gridProperties.rowCount,"
                "sheets.properties.gridProperties.columnCount"
            },
        )
        self._sheet_metadata_cache[spreadsheet_id] = metadata
        return metadata

    def get_drive_file(self, file_id_or_url: str) -> dict[str, Any]:
        file_id = extract_file_id(file_id_or_url)
        return self._request(
            "GET",
            f"https://www.googleapis.com/drive/v3/files/{file_id}",
            scopes=[DRIVE_SCOPE],
            allow_api_key=True,
            params={
                "fields": "id,name,mimeType,webViewLink,iconLink,owners(displayName,emailAddress),exportLinks"
            },
        )

    def export_drive_file(self, file_id_or_url: str, mime_type: str) -> tuple[dict[str, Any], bytes]:
        file_id = extract_file_id(file_id_or_url)
        metadata = self.get_drive_file(file_id)
        content = self._request(
            "GET",
            f"https://www.googleapis.com/drive/v3/files/{file_id}/export",
            scopes=[DRIVE_SCOPE],
            params={"mimeType": mime_type},
            expect_json=False,
        )
        return metadata, content

    def get_doc(self, document_id_or_url: str) -> dict[str, Any]:
        document_id = extract_file_id(document_id_or_url, kind="doc")
        return self._request(
            "GET",
            f"https://docs.googleapis.com/v1/documents/{document_id}",
            scopes=[DOCS_SCOPE],
            allow_api_key=False,
            params={"includeTabsContent": "true"},
        )

    def get_sheet_values(
        self,
        spreadsheet_id_or_url: str,
        range_a1: str | None,
        major_dimension: str,
        value_render_option: str,
        date_time_render_option: str,
    ) -> dict[str, Any]:
        context = self.resolve_sheet_range_context(spreadsheet_id_or_url, range_a1=range_a1)
        if not context["resolved_range_a1"]:
            raise ValueError(
                "No A1 range could be resolved. Pass `range_a1` explicitly or include `gid`/`range` in the sheet URL."
            )
        spreadsheet_id = context["spreadsheet_id"]
        active_properties = context["resolved_sheet_properties"] or context["default_sheet_properties"] or {}
        column_count = active_properties.get("gridProperties", {}).get("columnCount")
        max_column_a1 = (
            column_to_a1(column_count - 1) if column_count and column_count > 0 else MAX_SHEET_COLUMN_A1
        )
        encoded_range = quote_range(
            normalize_values_range(context["resolved_range_a1"], max_column_a1=max_column_a1)
        )
        return self._request(
            "GET",
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{encoded_range}",
            scopes=[SHEETS_SCOPE],
            allow_api_key=True,
            params={
                "majorDimension": major_dimension,
                "valueRenderOption": value_render_option,
                "dateTimeRenderOption": date_time_render_option,
            },
        )

    def get_sheet_grid(
        self,
        spreadsheet_id_or_url: str,
        range_a1: str | None,
        *,
        fields: str = SHEET_GRID_FIELDS,
    ) -> dict[str, Any]:
        context = self.resolve_sheet_range_context(spreadsheet_id_or_url, range_a1=range_a1)
        spreadsheet_id = context["spreadsheet_id"]
        active_properties = context["resolved_sheet_properties"] or context["default_sheet_properties"] or {}
        column_count = active_properties.get("gridProperties", {}).get("columnCount")
        max_column_a1 = (
            column_to_a1(column_count - 1) if column_count and column_count > 0 else MAX_SHEET_COLUMN_A1
        )
        params = {
            "includeGridData": "true",
            "fields": fields,
        }
        if context["resolved_range_a1"]:
            params["ranges"] = normalize_values_range(
                context["resolved_range_a1"],
                max_column_a1=max_column_a1,
            )
        return self._request(
            "GET",
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}",
            scopes=[SHEETS_SCOPE],
            allow_api_key=True,
            params=params,
        )


def get_client() -> GoogleWorkspaceClient:
    return GoogleWorkspaceClient()


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
                    row_cells.append(
                        compact_dict(
                            {
                                "a1": a1_from_zero_based(start_row + row_offset, start_col + col_offset),
                                "row_index": start_row + row_offset + 1,
                                "column_index": start_col + col_offset + 1,
                                "formatted_value": cell.get("formattedValue"),
                                "user_entered_value": user_value,
                                "effective_value": effective_value,
                                "formula": formula,
                                "note": cell.get("note"),
                                "hyperlink": cell.get("hyperlink")
                                or cell.get("userEnteredFormat", {})
                                .get("textFormat", {})
                                .get("link", {})
                                .get("uri"),
                                "text_runs": simplify_text_runs(cell.get("textFormatRuns")),
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
def read_sheet_values(
    spreadsheet_id_or_url: str,
    range_a1: str | None = None,
    major_dimension: str = "ROWS",
    value_render_option: str = "FORMATTED_VALUE",
    date_time_render_option: str = "SERIAL_NUMBER",
) -> dict[str, Any]:
    """Read raw Google Sheets values for an A1 range."""
    client = get_client()
    payload = client.get_sheet_values(
        spreadsheet_id_or_url,
        range_a1,
        major_dimension=major_dimension,
        value_render_option=value_render_option,
        date_time_render_option=date_time_render_option,
    )
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
    payload = client.get_sheet_grid(spreadsheet_id_or_url, range_a1)
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
        payload = client.get_sheet_grid(spreadsheet_id_or_url, quote_sheet_title(current_sheet_name))
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


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Google Workspace MCP server")
    subparsers = parser.add_subparsers(dest="command")

    auth_parser = subparsers.add_parser(
        "auth",
        help="Run the OAuth desktop login flow or inspect cached OAuth status.",
    )
    auth_parser.add_argument(
        "action",
        nargs="?",
        choices=("login", "status"),
        default="login",
        help="Choose `login` to launch the browser flow or `status` to inspect the current setup.",
    )
    auth_parser.add_argument(
        "--client-secrets",
        dest="client_secrets",
        help="Path to the OAuth client secrets JSON file.",
    )
    auth_parser.add_argument(
        "--token-file",
        dest="token_file",
        help="Path where the OAuth token cache should be stored.",
    )
    auth_parser.add_argument(
        "--scope",
        dest="scopes",
        action="append",
        help="Additional scope to request. Repeat for multiple scopes.",
    )
    auth_parser.add_argument(
        "--port",
        dest="port",
        type=int,
        help="Local callback port for the OAuth browser flow. Defaults to the Google-assigned port.",
    )
    auth_parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Do not auto-open the browser. The authorization URL will still be printed.",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    if args.command == "auth":
        client = get_client()
        if args.client_secrets:
            client.oauth_client_secrets_file = Path(args.client_secrets)
            client.oauth_client_config_json = None
        if args.token_file:
            client.oauth_token_file = Path(args.token_file)

        if args.action == "status":
            print(json.dumps(client.auth_summary(), indent=2))
            return

        try:
            result = client.run_oauth_login(
                scopes=args.scopes or DEFAULT_READONLY_SCOPES,
                open_browser=not args.no_browser,
                port=args.port,
            )
            print(json.dumps(result, indent=2))
            return
        except RuntimeError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            raise SystemExit(1) from exc

    mcp.run()


if __name__ == "__main__":
    main(sys.argv[1:])

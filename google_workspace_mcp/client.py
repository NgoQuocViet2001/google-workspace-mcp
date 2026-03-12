from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Iterable

import requests
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials as UserOAuthCredentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow

from .common import (
    DEFAULT_READONLY_SCOPES,
    DOCS_SCOPE,
    DRIVE_SCOPE,
    MAX_SHEET_COLUMN_A1,
    SHEETS_SCOPE,
    column_to_a1,
    default_oauth_token_file,
    extract_file_id,
    normalize_scopes,
    normalize_values_range,
    parse_sheet_url_context,
    path_from_env,
    quote_range,
    quote_sheet_title,
    split_sheet_range,
    unquote_sheet_title,
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
        fields: str,
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

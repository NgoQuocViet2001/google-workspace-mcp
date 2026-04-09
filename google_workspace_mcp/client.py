from __future__ import annotations

import importlib.resources
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
    CHAT_MEMBERSHIPS_WRITE_SCOPE,
    CHAT_MEMBERSHIPS_SCOPE,
    CHAT_MESSAGES_WRITE_SCOPE,
    CHAT_MESSAGES_SCOPE,
    CHAT_SPACES_WRITE_SCOPE,
    CHAT_SPACES_SCOPE,
    DEFAULT_ALL_WRITE_SCOPES,
    DEFAULT_READONLY_SCOPES,
    DEFAULT_READWRITE_SCOPES,
    DOCS_SCOPE,
    DOCS_WRITE_SCOPE,
    DRIVE_SCOPE,
    DRIVE_WRITE_SCOPE,
    MAX_SHEET_COLUMN_A1,
    SHEETS_SCOPE,
    SHEETS_WRITE_SCOPE,
    column_to_a1,
    default_oauth_client_secrets_file,
    default_oauth_token_file,
    extract_chat_message_name,
    extract_chat_space_name,
    extract_chat_thread_name,
    extract_file_id,
    normalize_scopes,
    normalize_values_range,
    parse_chat_url_context,
    parse_sheet_url_context,
    path_from_env,
    quote_range,
    quote_sheet_title,
    scope_is_satisfied,
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

    def _configured_oauth_client_secrets_file(self) -> Path | None:
        if not self.oauth_client_secrets_file:
            return None
        return self.oauth_client_secrets_file.expanduser()

    def _oauth_client_secrets_search_patterns(self) -> tuple[str, ...]:
        return (
            "oauth-client-secret*.json",
            "oauth-client-secrets*.json",
            "client-secret*.json",
            "client-secrets*.json",
            "client_secret*.json",
            "*apps.googleusercontent.com*.json",
        )

    def _oauth_client_secrets_search_roots(self) -> list[Path]:
        roots = [default_oauth_client_secrets_file().parent]
        deduped: list[Path] = []
        seen: set[Path] = set()
        for root in roots:
            normalized = root.expanduser()
            if normalized in seen:
                continue
            seen.add(normalized)
            deduped.append(normalized)
        return deduped

    def _looks_like_oauth_client_secrets_file(self, path: Path) -> bool:
        if not path.is_file():
            return False
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return False
        if not isinstance(payload, dict):
            return False
        for key in ("installed", "web"):
            section = payload.get(key)
            if isinstance(section, dict) and section.get("client_id") and section.get("client_secret"):
                return True
        return False

    def _oauth_client_secrets_candidates_in_dir(self, directory: Path) -> list[Path]:
        if not directory.is_dir():
            return []
        candidates: list[Path] = []
        seen: set[Path] = set()
        for pattern in self._oauth_client_secrets_search_patterns():
            for match in sorted(directory.glob(pattern)):
                normalized = match.expanduser()
                if normalized in seen or not self._looks_like_oauth_client_secrets_file(normalized):
                    continue
                seen.add(normalized)
                candidates.append(normalized)
        return candidates

    def _auto_discovered_oauth_client_secrets_candidates(self) -> list[Path]:
        candidates: list[Path] = []
        seen: set[Path] = set()
        for root in self._oauth_client_secrets_search_roots():
            for match in self._oauth_client_secrets_candidates_in_dir(root):
                if match in seen:
                    continue
                seen.add(match)
                candidates.append(match)
        return candidates

    def _oauth_client_secrets_candidates_preview(self, candidates: list[Path]) -> str:
        preview = ", ".join(str(path) for path in candidates[:3])
        if len(candidates) > 3:
            preview += ", ..."
        return preview

    def _bundled_oauth_client_config_json(self) -> dict[str, Any] | None:
        try:
            resource = importlib.resources.files("google_workspace_mcp").joinpath("oauth-default-client.json")
        except (ModuleNotFoundError, FileNotFoundError):
            return None
        try:
            if not resource.is_file():
                return None
            payload = json.loads(resource.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
        if not isinstance(payload, dict):
            return None
        for key in ("installed", "web"):
            section = payload.get(key)
            if isinstance(section, dict) and section.get("client_id") and section.get("client_secret"):
                return payload
        return None

    def _oauth_client_secrets_fallback_candidates(self, configured_path: Path) -> list[Path]:
        candidates: list[Path] = []
        seen: set[Path] = set()

        def add(path: Path) -> None:
            normalized = path.expanduser()
            if normalized == configured_path or normalized in seen:
                return
            seen.add(normalized)
            candidates.append(normalized)

        file_name = configured_path.name
        if "oauth-client-secret" in file_name:
            add(configured_path.with_name(file_name.replace("oauth-client-secret", "oauth-client-secrets", 1)))
        if "oauth-client-secrets" in file_name:
            add(configured_path.with_name(file_name.replace("oauth-client-secrets", "oauth-client-secret", 1)))
        if "client-secret" in file_name:
            add(configured_path.with_name(file_name.replace("client-secret", "client-secrets", 1)))
        if "client-secrets" in file_name:
            add(configured_path.with_name(file_name.replace("client-secrets", "client-secret", 1)))

        for candidate_name in (
            "oauth-client-secret.json",
            "oauth-client-secrets.json",
            "client-secret.json",
            "client-secrets.json",
        ):
            add(configured_path.with_name(candidate_name))

        for match in self._oauth_client_secrets_candidates_in_dir(configured_path.parent):
            add(match)

        candidates = [path for path in candidates if self._looks_like_oauth_client_secrets_file(path)]
        return candidates

    def _oauth_client_secrets_path_issue(self) -> str | None:
        configured_path = self._configured_oauth_client_secrets_file()
        if configured_path is None:
            auto_discovered = self._auto_discovered_oauth_client_secrets_candidates()
            if len(auto_discovered) > 1:
                return (
                    "Multiple OAuth client secrets files were auto-discovered: "
                    f"{self._oauth_client_secrets_candidates_preview(auto_discovered)}. "
                    "Pass --client-secrets with the exact file you want to use, or keep only one desktop-app client "
                    "JSON in ~/.google-workspace-mcp."
                )
            return None
        if configured_path is None or configured_path.is_file():
            return None

        fallback_matches = [
            path for path in self._oauth_client_secrets_fallback_candidates(configured_path) if path.is_file()
        ]
        if len(fallback_matches) == 1:
            return None
        if len(fallback_matches) > 1:
            return (
                f"OAuth client secrets file was not found at '{configured_path}', and multiple nearby JSON files "
                "look like possible client secrets: "
                f"{self._oauth_client_secrets_candidates_preview(fallback_matches)}. "
                "Pass --client-secrets with the exact file path."
            )
        return (
            f"OAuth client secrets file was not found at '{configured_path}'. Pass the exact path to the desktop-app "
            "client JSON downloaded from Google Cloud (often named "
            "'client_secret_<id>.apps.googleusercontent.com.json')."
        )

    def _resolved_oauth_client_secrets_file(self, *, raise_on_error: bool = False) -> Path | None:
        configured_path = self._configured_oauth_client_secrets_file()
        if configured_path is None:
            auto_discovered = self._auto_discovered_oauth_client_secrets_candidates()
            if len(auto_discovered) == 1:
                return auto_discovered[0]
            if raise_on_error:
                issue = self._oauth_client_secrets_path_issue()
                if issue:
                    raise RuntimeError(issue)
            return None
        if configured_path is None:
            return None
        if configured_path.is_file():
            return configured_path

        fallback_matches = [
            path for path in self._oauth_client_secrets_fallback_candidates(configured_path) if path.is_file()
        ]
        if len(fallback_matches) == 1:
            return fallback_matches[0]
        if raise_on_error:
            issue = self._oauth_client_secrets_path_issue()
            if issue:
                raise RuntimeError(issue)
        return None

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

    def _scopes_for_cached_user_credentials(self, required_scopes: Iterable[str]) -> list[str]:
        cached_scopes = self._cached_oauth_token_scopes()
        if cached_scopes:
            return cached_scopes
        return normalize_scopes(required_scopes)

    def _missing_cached_oauth_scopes(self, required_scopes: Iterable[str]) -> list[str]:
        cached_scopes = self._cached_oauth_token_scopes()
        return [
            scope for scope in sorted(set(required_scopes)) if not scope_is_satisfied(cached_scopes, scope)
        ]

    def _cached_oauth_capabilities(self) -> dict[str, bool]:
        cached_scopes = self._cached_oauth_token_scopes()
        return {
            "drive_readonly": scope_is_satisfied(cached_scopes, DRIVE_SCOPE),
            "drive_write": scope_is_satisfied(cached_scopes, DRIVE_WRITE_SCOPE),
            "docs_readonly": scope_is_satisfied(cached_scopes, DOCS_SCOPE),
            "docs_write": scope_is_satisfied(cached_scopes, DOCS_WRITE_SCOPE),
            "sheets_readonly": scope_is_satisfied(cached_scopes, SHEETS_SCOPE),
            "sheets_write": scope_is_satisfied(cached_scopes, SHEETS_WRITE_SCOPE),
            "chat_spaces_readonly": scope_is_satisfied(cached_scopes, CHAT_SPACES_SCOPE),
            "chat_spaces_write": scope_is_satisfied(cached_scopes, CHAT_SPACES_WRITE_SCOPE),
            "chat_messages_readonly": scope_is_satisfied(cached_scopes, CHAT_MESSAGES_SCOPE),
            "chat_messages_write": scope_is_satisfied(cached_scopes, CHAT_MESSAGES_WRITE_SCOPE),
            "chat_memberships_readonly": scope_is_satisfied(cached_scopes, CHAT_MEMBERSHIPS_SCOPE),
            "chat_memberships_write": scope_is_satisfied(cached_scopes, CHAT_MEMBERSHIPS_WRITE_SCOPE),
            "sheets_url_drive_export_fallback": scope_is_satisfied(cached_scopes, DRIVE_SCOPE)
            and not scope_is_satisfied(cached_scopes, SHEETS_SCOPE),
        }

    def _auth_summary_notes(self, cached_oauth_capabilities: dict[str, bool]) -> list[str]:
        notes = [
            "Public Sheets can be read with GOOGLE_API_KEY.",
            "OAuth desktop client credentials can read private files shared to your Google account.",
            "Docs, Drive, and Google Chat reads are most reliable with OAuth user credentials or an OAuth access token.",
            "A service account must be granted access to private files or shared drives.",
        ]
        if cached_oauth_capabilities["sheets_url_drive_export_fallback"]:
            notes.append(
                "This cached token can still read Google Sheets URLs that include gid/range via Drive export fallback. "
                "Expect values-only output until you re-run `google-workspace-mcp auth login` with "
                "spreadsheets.readonly."
            )
        if cached_oauth_capabilities["drive_readonly"] and not cached_oauth_capabilities["docs_readonly"]:
            notes.append(
                "drive.readonly does not unlock Google Docs content reads. Google Docs still require "
                "documents.readonly."
            )
        if not cached_oauth_capabilities["sheets_write"]:
            notes.append(
                "Google Sheets edits require spreadsheets write scope. Re-run "
                "`google-workspace-mcp auth login --scope-preset sheets-write` to refresh the cached token."
            )
        if not (
            cached_oauth_capabilities["docs_write"]
            and cached_oauth_capabilities["drive_write"]
            and cached_oauth_capabilities["sheets_write"]
            and cached_oauth_capabilities["chat_spaces_write"]
            and cached_oauth_capabilities["chat_messages_write"]
            and cached_oauth_capabilities["chat_memberships_write"]
        ):
            notes.append(
                "If you want broad read/write scopes across Docs, Drive, Sheets, and Google Chat, re-run "
                "`google-workspace-mcp auth login --scope-preset all-write`."
            )
        return notes

    def auth_summary(self) -> dict[str, Any]:
        resolved_oauth_client_file = self._resolved_oauth_client_secrets_file()
        bundled_oauth_client_config = self._bundled_oauth_client_config_json()
        oauth_client_issue = self._oauth_client_secrets_path_issue()
        oauth_client_ready = bool(
            resolved_oauth_client_file
            or bundled_oauth_client_config
            or self.oauth_client_config_json
            or self.oauth_token_file.exists()
        )
        oauth_client_source = None
        if resolved_oauth_client_file:
            oauth_client_source = str(resolved_oauth_client_file)
        elif bundled_oauth_client_config:
            oauth_client_source = "google_workspace_mcp/oauth-default-client.json"
        elif self.oauth_client_secrets_file:
            oauth_client_source = str(self._configured_oauth_client_secrets_file())
        elif self.oauth_client_config_json:
            oauth_client_source = "GOOGLE_OAUTH_CLIENT_CONFIG_JSON"
        service_account_ready = bool(self.service_account_file or self.service_account_json)
        service_account_source = None
        if self.service_account_file:
            service_account_source = str(self.service_account_file)
        elif self.service_account_json:
            service_account_source = "GOOGLE_SERVICE_ACCOUNT_JSON"
        cached_oauth_scopes = self._cached_oauth_token_scopes()
        cached_oauth_capabilities = self._cached_oauth_capabilities()
        return {
            "api_key_configured": bool(self.api_key),
            "oauth_access_token_configured": bool(self.oauth_access_token),
            "oauth_client_configured": oauth_client_ready,
            "oauth_client_source": oauth_client_source,
            "oauth_client_path_issue": oauth_client_issue,
            "oauth_token_file": str(self.oauth_token_file),
            "oauth_token_cached": self.oauth_token_file.exists(),
            "oauth_token_format": self._cached_oauth_token_format(),
            "oauth_token_scopes": cached_oauth_scopes,
            "oauth_token_missing_scopes": self._missing_cached_oauth_scopes(DEFAULT_READONLY_SCOPES),
            "oauth_token_missing_readwrite_scopes": self._missing_cached_oauth_scopes(DEFAULT_READWRITE_SCOPES),
            "oauth_token_missing_all_write_scopes": self._missing_cached_oauth_scopes(DEFAULT_ALL_WRITE_SCOPES),
            "oauth_token_capabilities": cached_oauth_capabilities,
            "service_account_configured": service_account_ready,
            "service_account_source": service_account_source,
            "recommended_mode": self._recommended_mode(),
            "active_auth_mode": self._active_auth_mode(),
            "notes": self._auth_summary_notes(cached_oauth_capabilities),
        }

    def _recommended_mode(self) -> str:
        if self.oauth_access_token:
            return "oauth_access_token"
        if self.oauth_token_file.exists():
            return "oauth_client"
        if self._oauth_client_is_configured():
            return "oauth_client"
        if self.service_account_file or self.service_account_json:
            return "service_account"
        if self.api_key:
            return "api_key_public_only"
        return "missing_credentials"

    def _active_auth_mode(self) -> str:
        if self.oauth_access_token:
            return "oauth_access_token"
        if self.oauth_token_file.exists():
            return "oauth_client_cached_token"
        if self._oauth_client_is_configured():
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
        return bool(
            self._resolved_oauth_client_secrets_file()
            or self._bundled_oauth_client_config_json()
            or self.oauth_client_config_json
        )

    def _oauth_flow(self, scopes: Iterable[str]) -> InstalledAppFlow:
        oauth_client_secrets_file = self._resolved_oauth_client_secrets_file(raise_on_error=True)
        if oauth_client_secrets_file:
            try:
                return InstalledAppFlow.from_client_secrets_file(
                    str(oauth_client_secrets_file),
                    list(scopes),
                )
            except OSError as exc:
                raise RuntimeError(
                    f"Unable to read OAuth client secrets file at '{oauth_client_secrets_file}': {exc}"
                ) from exc
            except ValueError as exc:
                raise RuntimeError(
                    "OAuth client secrets file is not a valid Google OAuth client configuration JSON."
                ) from exc
        bundled_oauth_client_config = self._bundled_oauth_client_config_json()
        if bundled_oauth_client_config:
            try:
                return InstalledAppFlow.from_client_config(
                    bundled_oauth_client_config,
                    list(scopes),
                )
            except ValueError as exc:
                raise RuntimeError(
                    "The bundled OAuth client configuration is not a valid Google OAuth client configuration."
                ) from exc
        if self.oauth_client_config_json:
            try:
                client_config = json.loads(self.oauth_client_config_json)
            except ValueError as exc:
                raise RuntimeError("GOOGLE_OAUTH_CLIENT_CONFIG_JSON is not valid JSON.") from exc
            try:
                return InstalledAppFlow.from_client_config(
                    client_config,
                    list(scopes),
                )
            except ValueError as exc:
                raise RuntimeError(
                    "GOOGLE_OAUTH_CLIENT_CONFIG_JSON is not a valid Google OAuth client configuration."
                ) from exc
        raise RuntimeError(
            "No OAuth client credentials are configured. Pass --client-secrets, set "
            "GOOGLE_OAUTH_CLIENT_SECRETS_FILE or GOOGLE_OAUTH_CLIENT_CONFIG_JSON, or place the desktop-app client "
            "JSON in ~/.google-workspace-mcp."
        )

    def _save_user_credentials(self, credentials: UserOAuthCredentials) -> None:
        self.oauth_token_file.parent.mkdir(parents=True, exist_ok=True)
        self.oauth_token_file.write_text(credentials.to_json(), encoding="utf-8")

    def _persist_oauth_client_secrets_file(self, source_path: Path) -> tuple[Path | None, str | None]:
        target_path = default_oauth_client_secrets_file()
        source_path = source_path.expanduser()
        target_path = target_path.expanduser()
        try:
            if source_path.resolve() == target_path.resolve():
                return target_path, None
        except OSError:
            pass

        try:
            payload = source_path.read_text(encoding="utf-8")
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_text(payload, encoding="utf-8")
        except OSError as exc:
            return None, (
                f"Login succeeded, but the OAuth client secrets file could not be copied to '{target_path}': {exc}"
            )
        return target_path, None

    def persist_oauth_client_config(self, client_id: str, client_secret: str) -> Path:
        normalized_client_id = client_id.strip()
        normalized_client_secret = client_secret.strip()
        if not normalized_client_id or not normalized_client_secret:
            raise RuntimeError("Both OAuth client ID and client secret are required.")

        target_path = default_oauth_client_secrets_file().expanduser()
        payload = {
            "installed": {
                "client_id": normalized_client_id,
                "client_secret": normalized_client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "redirect_uris": ["http://localhost"],
            }
        }
        try:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except OSError as exc:
            raise RuntimeError(
                f"Failed to save OAuth client configuration to '{target_path}': {exc}"
            ) from exc

        self.oauth_client_secrets_file = target_path
        self.oauth_client_config_json = None
        return target_path

    def _persist_oauth_client_config_payload(self, payload: dict[str, Any]) -> tuple[Path | None, str | None]:
        target_path = default_oauth_client_secrets_file().expanduser()
        try:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except OSError as exc:
            return None, (
                f"Login succeeded, but the OAuth client configuration could not be saved to '{target_path}': {exc}"
            )
        self.oauth_client_secrets_file = target_path
        self.oauth_client_config_json = None
        return target_path, None

    def _revoke_oauth_token(self, token: str) -> tuple[bool, str | None]:
        try:
            response = self.session.post(
                "https://oauth2.googleapis.com/revoke",
                params={"token": token},
                headers={"content-type": "application/x-www-form-urlencoded"},
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            return False, f"Failed to contact Google's token revocation endpoint: {exc}"

        if response.status_code == 200:
            return True, None

        detail = response.text.strip()
        if detail:
            detail = f": {detail}"
        return False, f"Google token revocation returned HTTP {response.status_code}{detail}"

    def run_oauth_login(
        self,
        scopes: Iterable[str] | None = None,
        *,
        open_browser: bool | None = None,
        port: int | None = None,
    ) -> dict[str, Any]:
        requested_scopes = list(scopes or DEFAULT_READONLY_SCOPES)
        resolved_oauth_client_secrets_file = self._resolved_oauth_client_secrets_file()
        bundled_oauth_client_config = self._bundled_oauth_client_config_json()
        flow = self._oauth_flow(requested_scopes)
        credentials = flow.run_local_server(
            port=self.oauth_local_server_port if port is None else port,
            open_browser=self.oauth_open_browser if open_browser is None else open_browser,
            authorization_prompt_message="Open this URL in your browser to authorize access: {url}",
            success_message="Google Workspace MCP authorization completed. You can close this window.",
        )
        self._save_user_credentials(credentials)
        self._user_credentials.clear()
        oauth_client_secrets_file = None
        notes: list[str] = []
        if resolved_oauth_client_secrets_file:
            persisted_path, persist_error = self._persist_oauth_client_secrets_file(
                resolved_oauth_client_secrets_file
            )
            oauth_client_secrets_file = str(persisted_path or resolved_oauth_client_secrets_file)
            if persist_error:
                notes.append(persist_error)
        elif bundled_oauth_client_config:
            persisted_path, persist_error = self._persist_oauth_client_config_payload(bundled_oauth_client_config)
            oauth_client_secrets_file = str(persisted_path) if persisted_path else None
            if persist_error:
                notes.append(persist_error)
        return {
            "oauth_client_secrets_file": oauth_client_secrets_file,
            "oauth_token_file": str(self.oauth_token_file),
            "scopes": requested_scopes,
            "account": getattr(credentials, "account", None),
            "has_refresh_token": bool(credentials.refresh_token),
            "notes": notes,
        }

    def run_oauth_logout(self) -> dict[str, Any]:
        cached_payload = self._cached_oauth_token_payload() or {}
        oauth_access_token_configured = bool(self.oauth_access_token)
        token_file_existed = self.oauth_token_file.exists()

        token_to_revoke = None
        revoked_token_type = None
        for candidate_key, candidate_label in (
            ("refresh_token", "refresh_token"),
            ("token", "access_token"),
        ):
            candidate_value = str(cached_payload.get(candidate_key) or "").strip()
            if candidate_value:
                token_to_revoke = candidate_value
                revoked_token_type = candidate_label
                break

        revoked = False
        revoke_error = None
        if token_to_revoke:
            revoked, revoke_error = self._revoke_oauth_token(token_to_revoke)

        token_file_deleted = False
        if token_file_existed:
            try:
                self.oauth_token_file.unlink()
            except OSError as exc:
                raise RuntimeError(
                    f"Failed to delete cached OAuth token file '{self.oauth_token_file}': {exc}"
                ) from exc
            token_file_deleted = True

        self._user_credentials.clear()

        notes: list[str] = []
        if not token_file_existed:
            notes.append("No cached OAuth token file was found.")
        if token_to_revoke and not revoked and revoke_error:
            notes.append(
                "Cached token file was deleted, but token revocation may not have completed on Google's side: "
                + revoke_error
            )
        if oauth_access_token_configured:
            notes.append(
                "GOOGLE_OAUTH_ACCESS_TOKEN is configured separately. Remove it from your shell or MCP env config "
                "if you want to fully disable bearer-token auth."
            )

        return {
            "oauth_token_file": str(self.oauth_token_file),
            "oauth_token_file_existed": token_file_existed,
            "oauth_token_file_deleted": token_file_deleted,
            "revoked": revoked,
            "revoked_token_type": revoked_token_type,
            "oauth_access_token_configured": oauth_access_token_configured,
            "notes": notes,
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

        if not self.oauth_token_file.exists():
            raise RuntimeError(
                "No cached OAuth user token was found. Set GOOGLE_OAUTH_CLIENT_SECRETS_FILE or "
                "GOOGLE_OAUTH_CLIENT_CONFIG_JSON, then run `google-workspace-mcp auth` to complete the browser login flow."
            )

        missing_scopes = self._missing_cached_oauth_scopes(scope_list)
        if missing_scopes:
            missing_display = ", ".join(missing_scopes)
            raise RuntimeError(
                "Cached OAuth token is missing required scopes: "
                f"{missing_display}. Re-run `google-workspace-mcp auth` to refresh the token."
            )

        # Refresh tokens must be reloaded with the scopes originally granted by Google.
        # Replacing a cached write scope with its readonly alias can trigger `invalid_scope`
        # once the access token expires and refresh is required.
        credential_scopes = self._scopes_for_cached_user_credentials(scope_list)
        try:
            credentials = UserOAuthCredentials.from_authorized_user_file(
                str(self.oauth_token_file),
                credential_scopes,
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
        auth_errors: list[str] = []

        if self.oauth_access_token:
            headers["Authorization"] = f"Bearer {self.oauth_access_token}"
            return headers, params

        if scope_list and (self.oauth_token_file.exists() or self._oauth_client_is_configured()):
            try:
                credentials = self._user_oauth_credentials(scope_list)
            except RuntimeError as exc:
                auth_errors.append(str(exc))
            else:
                headers["Authorization"] = f"Bearer {credentials.token}"
                return headers, params

        if scope_list and (self.service_account_file or self.service_account_json):
            try:
                creds = self._scoped_credentials.get(scope_list)
                if creds is None:
                    creds = self._service_account_base().with_scopes(scope_list)
                    self._scoped_credentials[scope_list] = creds
                if not creds.valid or not creds.token:
                    creds.refresh(GoogleAuthRequest())
                headers["Authorization"] = f"Bearer {creds.token}"
                return headers, params
            except RuntimeError as exc:
                auth_errors.append(str(exc))

        if allow_api_key and self.api_key:
            params["key"] = self.api_key
            return headers, params

        if auth_errors:
            if len(auth_errors) == 1:
                raise RuntimeError(auth_errors[0])
            raise RuntimeError(
                "No usable Google credentials matched the requested scopes. "
                + " ".join(auth_errors)
            )

        mode = "missing credentials"
        if allow_api_key:
            mode = "requires GOOGLE_API_KEY or credentials with read access"
        raise RuntimeError(
            "No valid Google credentials are configured. "
            f"This operation {mode}. Use an OAuth desktop client, service account, or OAuth token "
            "for private Google Workspace data such as Docs, Drive, Sheets, and Chat, or GOOGLE_API_KEY "
            "for public Sheets."
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
        json_body: Any | None = None,
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
                json=json_body,
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
            if "chat.googleapis.com" in url:
                error_message = self._augment_google_chat_error_message(response.status_code, error_message)
            raise RuntimeError(
                f"Google API returned HTTP {response.status_code} for {url}: {error_message}"
            )
        if expect_json:
            if response.status_code == 204 or not response.content:
                return {}
            return response.json()
        return response.content

    def _augment_google_chat_error_message(self, status_code: int, error_message: str) -> str:
        lowered = error_message.lower()
        if status_code == 404 and "google chat app not found" in lowered:
            return (
                "Google Chat API is enabled, but this Google Cloud project does not have a Chat app "
                "configuration yet. In Google Cloud Console, open Google Chat API > Configuration and "
                "set at least App name, Avatar URL, and Description, then save. Google requires a "
                "configured Chat app before Chat API reads succeed, even when you authenticate as a user."
            )
        return error_message

    def _normalize_chat_message_order_by(self, order_by: str | None) -> str | None:
        if order_by is None:
            return None
        trimmed = order_by.strip()
        if not trimmed:
            return None
        upper = trimmed.upper()
        if upper in {"ASC", "DESC"}:
            return f"createTime {upper}"
        return trimmed

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

    def get_drive_file(self, file_id_or_url: str, *, include_capabilities: bool = False) -> dict[str, Any]:
        file_id = extract_file_id(file_id_or_url)
        fields = "id,name,mimeType,webViewLink,iconLink,owners(displayName,emailAddress),exportLinks"
        if include_capabilities:
            fields += ",capabilities(canEdit,canModifyContent)"
        return self._request(
            "GET",
            f"https://www.googleapis.com/drive/v3/files/{file_id}",
            scopes=[DRIVE_SCOPE],
            allow_api_key=True,
            params={"fields": fields},
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

    def export_sheet_via_drive(
        self,
        spreadsheet_id_or_url: str,
        *,
        export_format: str,
        gid: int,
        range_a1: str | None = None,
    ) -> bytes:
        spreadsheet_id = extract_file_id(spreadsheet_id_or_url, kind="sheet")
        params: dict[str, Any] = {
            "format": export_format,
            "gid": str(gid),
        }
        if range_a1:
            params["range"] = range_a1
        return self._request(
            "GET",
            f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export",
            scopes=[DRIVE_SCOPE],
            allow_api_key=False,
            params=params,
            expect_json=False,
        )

    def get_doc(self, document_id_or_url: str) -> dict[str, Any]:
        document_id = extract_file_id(document_id_or_url, kind="doc")
        return self._request(
            "GET",
            f"https://docs.googleapis.com/v1/documents/{document_id}",
            scopes=[DOCS_SCOPE],
            allow_api_key=False,
            params={"includeTabsContent": "true"},
        )

    def list_chat_spaces(
        self,
        *,
        page_size: int = 100,
        page_token: str | None = None,
        filter_text: str | None = None,
    ) -> dict[str, Any]:
        params = {"pageSize": page_size}
        if page_token:
            params["pageToken"] = page_token
        if filter_text:
            params["filter"] = filter_text
        return self._request(
            "GET",
            "https://chat.googleapis.com/v1/spaces",
            scopes=[CHAT_SPACES_SCOPE],
            allow_api_key=False,
            params=params,
        )

    def get_chat_space(self, space_name_or_url: str) -> dict[str, Any]:
        space_name = extract_chat_space_name(space_name_or_url)
        return self._request(
            "GET",
            f"https://chat.googleapis.com/v1/{space_name}",
            scopes=[CHAT_SPACES_SCOPE],
            allow_api_key=False,
        )

    def get_chat_message(self, message_name_or_url: str) -> dict[str, Any]:
        message_name = extract_chat_message_name(message_name_or_url)
        return self._request(
            "GET",
            f"https://chat.googleapis.com/v1/{message_name}",
            scopes=[CHAT_MESSAGES_SCOPE],
            allow_api_key=False,
        )

    def list_chat_messages(
        self,
        space_name_or_url: str,
        *,
        page_size: int = 100,
        page_token: str | None = None,
        filter_text: str | None = None,
        order_by: str | None = None,
        show_deleted: bool = False,
    ) -> dict[str, Any]:
        space_name = extract_chat_space_name(space_name_or_url)
        params = {"pageSize": page_size}
        if page_token:
            params["pageToken"] = page_token
        if filter_text:
            params["filter"] = filter_text
        normalized_order_by = self._normalize_chat_message_order_by(order_by)
        if normalized_order_by:
            params["orderBy"] = normalized_order_by
        if show_deleted:
            params["showDeleted"] = "true"
        return self._request(
            "GET",
            f"https://chat.googleapis.com/v1/{space_name}/messages",
            scopes=[CHAT_MESSAGES_SCOPE],
            allow_api_key=False,
            params=params,
        )

    def list_chat_thread_messages(
        self,
        thread_name_or_url: str,
        *,
        page_size: int = 100,
        page_token: str | None = None,
        order_by: str | None = "ASC",
        show_deleted: bool = False,
    ) -> dict[str, Any]:
        context = parse_chat_url_context(thread_name_or_url)
        space_name = context.get("space_name") or extract_chat_space_name(thread_name_or_url)
        thread_name = context.get("thread_name") or extract_chat_thread_name(thread_name_or_url)
        return self.list_chat_messages(
            space_name,
            page_size=page_size,
            page_token=page_token,
            filter_text=f"thread.name = {thread_name}",
            order_by=order_by,
            show_deleted=show_deleted,
        )

    def list_chat_memberships(
        self,
        space_name_or_url: str,
        *,
        page_size: int = 100,
        page_token: str | None = None,
        filter_text: str | None = None,
        show_groups: bool = False,
        show_invited: bool = False,
    ) -> dict[str, Any]:
        space_name = extract_chat_space_name(space_name_or_url)
        params = {"pageSize": page_size}
        if page_token:
            params["pageToken"] = page_token
        if filter_text:
            params["filter"] = filter_text
        if show_groups:
            params["showGroups"] = "true"
        if show_invited:
            params["showInvited"] = "true"
        return self._request(
            "GET",
            f"https://chat.googleapis.com/v1/{space_name}/members",
            scopes=[CHAT_MEMBERSHIPS_SCOPE],
            allow_api_key=False,
            params=params,
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

    def update_sheet_values(
        self,
        spreadsheet_id_or_url: str,
        range_a1: str,
        values: list[list[Any]],
        *,
        value_input_option: str,
        major_dimension: str,
        include_values_in_response: bool,
    ) -> dict[str, Any]:
        context = self.resolve_sheet_range_context(spreadsheet_id_or_url, range_a1=range_a1)
        if not context["resolved_range_a1"]:
            raise ValueError("Pass `range_a1` explicitly when writing Google Sheets values.")
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
            "PUT",
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{encoded_range}",
            scopes=[SHEETS_WRITE_SCOPE],
            allow_api_key=False,
            params={
                "valueInputOption": value_input_option,
                "includeValuesInResponse": str(include_values_in_response).lower(),
                "responseValueRenderOption": "FORMATTED_VALUE",
            },
            json_body={
                "majorDimension": major_dimension,
                "values": values,
            },
        )


def get_client() -> GoogleWorkspaceClient:
    return GoogleWorkspaceClient()

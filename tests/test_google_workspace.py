import json
import os
import runpy
import tempfile
import unittest
from argparse import Namespace
from pathlib import Path
from unittest.mock import Mock, patch

import mcp_google_workspace as workspace


TEST_SPREADSHEET_ID = "1AbCdEfGhIjKlMnOp"


class GoogleWorkspaceClientTests(unittest.TestCase):
    def make_client(self, extra_env: dict[str, str] | None = None) -> workspace.GoogleWorkspaceClient:
        env = {"USERPROFILE": "C:/Users/TestUser"}
        env.update(extra_env or {})
        with patch.dict(os.environ, env, clear=True):
            return workspace.GoogleWorkspaceClient()

    def test_parse_sheet_url_context_reads_gid_and_range_from_fragment(self) -> None:
        context = workspace.parse_sheet_url_context(
            "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=111#gid=1544244212&range=38:38"
        )
        self.assertEqual(context["spreadsheet_id"], TEST_SPREADSHEET_ID)
        self.assertEqual(context["gid"], 1544244212)
        self.assertEqual(context["range_a1"], "38:38")

    def test_extract_chat_space_name_accepts_chat_urls_and_ids(self) -> None:
        self.assertEqual(
            workspace.extract_chat_space_name("https://mail.google.com/chat/u/0/#chat/space/AAAAB3NzaC1yc2E"),
            "spaces/AAAAB3NzaC1yc2E",
        )
        self.assertEqual(
            workspace.extract_chat_space_name("AAAAB3NzaC1yc2E"),
            "spaces/AAAAB3NzaC1yc2E",
        )
        self.assertEqual(
            workspace.extract_chat_space_name("spaces/AAAAB3NzaC1yc2E/messages/123"),
            "spaces/AAAAB3NzaC1yc2E",
        )

    def test_parse_chat_url_context_reads_space_thread_and_message_from_room_url(self) -> None:
        context = workspace.parse_chat_url_context(
            "https://chat.google.com/room/AAQAyxdRoZo/jVIpmenXnO0/WNSdv6IyQf0?cls=10"
        )

        self.assertEqual(context["space_name"], "spaces/AAQAyxdRoZo")
        self.assertEqual(context["thread_name"], "spaces/AAQAyxdRoZo/threads/jVIpmenXnO0")
        self.assertEqual(context["message_id"], "WNSdv6IyQf0")
        self.assertIsNone(context["message_name"])
        self.assertEqual(context["message_lookup_hint"], "WNSdv6IyQf0")

    def test_extract_chat_thread_and_message_names_accept_room_url(self) -> None:
        url = "https://chat.google.com/room/AAQAyxdRoZo/jVIpmenXnO0/WNSdv6IyQf0?cls=10"
        self.assertEqual(
            workspace.extract_chat_thread_name(url),
            "spaces/AAQAyxdRoZo/threads/jVIpmenXnO0",
        )
        with self.assertRaises(ValueError) as caught:
            workspace.extract_chat_message_name(url)
        self.assertIn("don't reliably expose the API message ID", str(caught.exception))

    def test_get_sheet_values_uses_gid_for_row_only_range(self) -> None:
        client = self.make_client()
        client.get_sheet_metadata = Mock(
            return_value={
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 1544244212,
                            "title": "Spec",
                            "gridProperties": {"rowCount": 200, "columnCount": 23},
                        }
                    }
                ]
            }
        )
        client._request = Mock(return_value={"range": "'Spec'!A38:W38", "values": [[]]})

        client.get_sheet_values(
            "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1544244212#gid=1544244212&range=38:38",
            "38:38",
            major_dimension="ROWS",
            value_render_option="FORMATTED_VALUE",
            date_time_render_option="SERIAL_NUMBER",
        )

        called_url = client._request.call_args.args[1]
        self.assertIn("%27Spec%27!A38:W38", called_url)

    def test_get_sheet_values_uses_gid_when_explicit_range_has_no_prefix(self) -> None:
        client = self.make_client()
        client.get_sheet_metadata = Mock(
            return_value={
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 1436003411,
                            "title": "Feedback",
                            "gridProperties": {"rowCount": 300, "columnCount": 17},
                        }
                    }
                ]
            }
        )
        client._request = Mock(return_value={"range": "'Feedback'!C119:F119", "values": [[]]})

        client.get_sheet_values(
            "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411#gid=1436003411&range=C119:F119",
            "C119:F119",
            major_dimension="ROWS",
            value_render_option="FORMATTED_VALUE",
            date_time_render_option="SERIAL_NUMBER",
        )

        called_url = client._request.call_args.args[1]
        self.assertIn("%27Feedback%27!C119:F119", called_url)

    def test_get_sheet_values_uses_fragment_range_when_range_argument_is_missing(self) -> None:
        client = self.make_client()
        client.get_sheet_metadata = Mock(
            return_value={
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 1436003411,
                            "title": "Feedback",
                            "gridProperties": {"rowCount": 300, "columnCount": 17},
                        }
                    }
                ]
            }
        )
        client._request = Mock(return_value={"range": "'Feedback'!C119:F119", "values": [[]]})

        client.get_sheet_values(
            "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411#gid=1436003411&range=C119:F119",
            None,
            major_dimension="ROWS",
            value_render_option="FORMATTED_VALUE",
            date_time_render_option="SERIAL_NUMBER",
        )

        called_url = client._request.call_args.args[1]
        self.assertIn("%27Feedback%27!C119:F119", called_url)

    def test_read_sheet_values_falls_back_to_drive_export_for_gid_url_when_sheets_scope_is_missing(self) -> None:
        client = self.make_client()
        client.get_sheet_values = Mock(
            side_effect=RuntimeError(
                "Cached OAuth token is missing required scopes: "
                f"{workspace.SHEETS_SCOPE}. Re-run `google-workspace-mcp auth` to refresh the token."
            )
        )
        client.export_sheet_via_drive = Mock(return_value=b"left,right\r\nvalue-1,value-2\r\n")

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.read_sheet_values(
                "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411#gid=1436003411&range=E187:F187"
            )

        client.export_sheet_via_drive.assert_called_once_with(
            "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411#gid=1436003411&range=E187:F187",
            export_format="csv",
            gid=1436003411,
            range_a1="E187:F187",
        )
        self.assertEqual(result["spreadsheet_id"], TEST_SPREADSHEET_ID)
        self.assertEqual(result["range"], "E187:F187")
        self.assertEqual(result["values"], [["left", "right"], ["value-1", "value-2"]])
        self.assertEqual(result["source"], "drive_export_csv_fallback")
        self.assertIn("spreadsheets.readonly", result["auth_warning"])

    def test_read_sheet_grid_falls_back_to_drive_export_for_gid_url_when_sheets_scope_is_missing(self) -> None:
        client = self.make_client()
        client.get_sheet_grid = Mock(
            side_effect=RuntimeError(
                "Cached OAuth token is missing required scopes: "
                f"{workspace.SHEETS_SCOPE}. Re-run `google-workspace-mcp auth` to refresh the token."
            )
        )
        client.export_sheet_via_drive = Mock(return_value=b"alpha,beta\r\ngamma,delta\r\n")
        client.get_drive_file = Mock(return_value={"name": "Feedback tracker"})

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.read_sheet_grid(
                "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411#gid=1436003411&range=E187:F187"
            )

        self.assertEqual(result["spreadsheet_id"], TEST_SPREADSHEET_ID)
        self.assertEqual(result["title"], "Feedback tracker")
        self.assertEqual(result["source"], "drive_export_csv_fallback")
        self.assertEqual(result["sheets"][0]["sheet_id"], 1436003411)
        self.assertEqual(result["sheets"][0]["data"][0]["start_row"], 187)
        self.assertEqual(result["sheets"][0]["data"][0]["start_column"], 5)
        first_cell = result["sheets"][0]["data"][0]["rows"][0]["cells"][0]
        self.assertEqual(first_cell["a1"], "E187")
        self.assertEqual(first_cell["formatted_value"], "alpha")
        self.assertIn("spreadsheets.readonly", result["auth_warning"])

    def test_search_sheet_limits_search_to_gid_sheet(self) -> None:
        client = self.make_client()
        metadata = {
            "sheets": [
                {"properties": {"sheetId": 1, "title": "Cover", "gridProperties": {"rowCount": 10, "columnCount": 5}}},
                {
                    "properties": {
                        "sheetId": 1436003411,
                        "title": "Feedback",
                        "gridProperties": {"rowCount": 300, "columnCount": 17},
                    }
                },
            ]
        }
        client.resolve_sheet_range_context = Mock(
            return_value={
                "resolved_sheet_name": "Feedback",
                "metadata": metadata,
            }
        )
        client.get_sheet_grid = Mock(return_value={"sheets": []})

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            with patch("google_workspace_mcp.tools.simplify_grid_data", return_value=[]):
                result = workspace.search_sheet(
                    "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411",
                    "needle",
                )

        self.assertEqual(result["match_count"], 0)
        self.assertEqual(client.get_sheet_grid.call_count, 1)
        self.assertEqual(client.get_sheet_grid.call_args.args[1], workspace.quote_sheet_title("Feedback"))

    def test_list_chat_messages_uses_space_name_and_chat_scope(self) -> None:
        client = self.make_client()
        client._request = Mock(return_value={"messages": []})

        client.list_chat_messages(
            "https://mail.google.com/chat/u/0/#chat/space/AAAAB3NzaC1yc2E",
            page_size=25,
            page_token="next-token",
            filter_text='thread.name = "spaces/AAAAB3NzaC1yc2E/threads/123"',
            order_by="ASC",
        )

        client._request.assert_called_once_with(
            "GET",
            "https://chat.googleapis.com/v1/spaces/AAAAB3NzaC1yc2E/messages",
            scopes=[workspace.CHAT_MESSAGES_SCOPE],
            allow_api_key=False,
            params={
                "pageSize": 25,
                "pageToken": "next-token",
                "filter": 'thread.name = "spaces/AAAAB3NzaC1yc2E/threads/123"',
                "orderBy": "createTime ASC",
            },
        )

    def test_normalize_chat_message_order_by_expands_short_directions(self) -> None:
        client = self.make_client()

        self.assertEqual(client._normalize_chat_message_order_by("ASC"), "createTime ASC")
        self.assertEqual(client._normalize_chat_message_order_by("desc"), "createTime DESC")
        self.assertEqual(
            client._normalize_chat_message_order_by("createTime DESC"),
            "createTime DESC",
        )

    def test_get_chat_message_uses_message_name_and_chat_scope(self) -> None:
        client = self.make_client()
        client._request = Mock(return_value={"name": "spaces/AAQAyxdRoZo/messages/WNSdv6IyQf0"})

        client.get_chat_message("spaces/AAQAyxdRoZo/messages/WNSdv6IyQf0")

        client._request.assert_called_once_with(
            "GET",
            "https://chat.googleapis.com/v1/spaces/AAQAyxdRoZo/messages/WNSdv6IyQf0",
            scopes=[workspace.CHAT_MESSAGES_SCOPE],
            allow_api_key=False,
        )

    def test_list_chat_thread_messages_filters_by_thread_name(self) -> None:
        client = self.make_client()
        client.list_chat_messages = Mock(return_value={"messages": []})

        client.list_chat_thread_messages(
            "https://chat.google.com/room/AAQAyxdRoZo/jVIpmenXnO0/WNSdv6IyQf0?cls=10",
            page_size=50,
            page_token="next-token",
            order_by="ASC",
            show_deleted=True,
        )

        client.list_chat_messages.assert_called_once_with(
            "spaces/AAQAyxdRoZo",
            page_size=50,
            page_token="next-token",
            filter_text="thread.name = spaces/AAQAyxdRoZo/threads/jVIpmenXnO0",
            order_by="ASC",
            show_deleted=True,
        )

    def test_list_google_chat_spaces_tool_returns_simplified_spaces(self) -> None:
        client = self.make_client()
        client.list_chat_spaces = Mock(
            return_value={
                "spaces": [
                    {
                        "name": "spaces/AAAAB3NzaC1yc2E",
                        "displayName": "Product Launch",
                        "spaceType": "SPACE",
                        "spaceThreadingState": "THREADED_MESSAGES",
                    }
                ],
                "nextPageToken": "page-2",
            }
        )

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.list_google_chat_spaces(page_size=10)

        client.list_chat_spaces.assert_called_once_with(page_size=10, page_token=None, filter_text=None)
        self.assertEqual(result["space_count"], 1)
        self.assertEqual(result["spaces"][0]["display_name"], "Product Launch")
        self.assertEqual(result["spaces"][0]["threading_state"], "THREADED_MESSAGES")
        self.assertEqual(result["next_page_token"], "page-2")

    def test_read_google_chat_messages_tool_returns_simplified_messages(self) -> None:
        client = self.make_client()
        client.list_chat_messages = Mock(
            return_value={
                "messages": [
                    {
                        "name": "spaces/AAAAB3NzaC1yc2E/messages/msg-1",
                        "text": "Hello team",
                        "formattedText": "*Hello* team",
                        "createTime": "2026-03-25T10:00:00Z",
                        "threadReply": False,
                        "sender": {
                            "name": "users/123",
                            "displayName": "Viet Ngo",
                            "type": "HUMAN",
                        },
                    }
                ]
            }
        )

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.read_google_chat_messages("spaces/AAAAB3NzaC1yc2E")

        self.assertEqual(result["message_count"], 1)
        self.assertEqual(result["messages"][0]["text"], "Hello team")
        self.assertEqual(result["messages"][0]["formatted_text"], "*Hello* team")
        self.assertEqual(result["messages"][0]["sender"]["display_name"], "Viet Ngo")

    def test_read_google_chat_thread_tool_returns_root_and_lookup_warning_when_room_token_cannot_map(self) -> None:
        client = self.make_client()
        client.list_chat_thread_messages = Mock(
            return_value={
                "messages": [
                    {
                        "name": "spaces/AAQAyxdRoZo/messages/root-msg",
                        "text": "Root message",
                        "threadReply": False,
                    },
                    {
                        "name": "spaces/AAQAyxdRoZo/messages/reply-msg",
                        "text": "Reply message",
                        "threadReply": True,
                    },
                ]
            }
        )

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.read_google_chat_thread(
                "https://chat.google.com/room/AAQAyxdRoZo/jVIpmenXnO0/WNSdv6IyQf0?cls=10"
            )

        self.assertEqual(result["space"], "spaces/AAQAyxdRoZo")
        self.assertEqual(result["thread"], "spaces/AAQAyxdRoZo/threads/jVIpmenXnO0")
        self.assertIsNone(result["linked_message"])
        self.assertIn("couldn't be mapped", result["linked_message_lookup_warning"])
        self.assertEqual(result["root_message"]["name"], "spaces/AAQAyxdRoZo/messages/root-msg")
        self.assertEqual(result["message_count"], 2)

    def test_list_google_chat_memberships_tool_returns_simplified_memberships(self) -> None:
        client = self.make_client()
        client.list_chat_memberships = Mock(
            return_value={
                "memberships": [
                    {
                        "name": "spaces/AAAAB3NzaC1yc2E/members/123",
                        "state": "JOINED",
                        "role": "ROLE_MEMBER",
                        "member": {
                            "name": "users/123",
                            "displayName": "Viet Ngo",
                            "type": "HUMAN",
                        },
                    }
                ]
            }
        )

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.list_google_chat_memberships("spaces/AAAAB3NzaC1yc2E", show_invited=True)

        client.list_chat_memberships.assert_called_once_with(
            "spaces/AAAAB3NzaC1yc2E",
            page_size=100,
            page_token=None,
            filter_text=None,
            show_groups=False,
            show_invited=True,
        )
        self.assertEqual(result["membership_count"], 1)
        self.assertEqual(result["memberships"][0]["member"]["display_name"], "Viet Ngo")

    def test_user_oauth_credentials_fails_fast_when_token_scopes_are_missing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "scopes": [workspace.SHEETS_SCOPE],
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(Path(temp_dir) / "client-secret.json"),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(token_file),
                }
            )

            with self.assertRaises(RuntimeError) as caught:
                client._user_oauth_credentials([workspace.DRIVE_SCOPE])

        self.assertIn("missing required scopes", str(caught.exception))
        self.assertIn(workspace.DRIVE_SCOPE, str(caught.exception))

    def test_user_oauth_credentials_uses_cached_write_scope_when_readonly_alias_is_requested(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "expired-access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "scopes": [workspace.SHEETS_WRITE_SCOPE],
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(Path(temp_dir) / "client-secret.json"),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(token_file),
                }
            )

            class FakeCredentials:
                def __init__(self) -> None:
                    self.valid = False
                    self.token = None
                    self.expired = True
                    self.refresh_token = "refresh-token"

                def refresh(self, _request: object) -> None:
                    self.valid = True
                    self.expired = False
                    self.token = "fresh-access-token"

            fake_credentials = FakeCredentials()

            with patch(
                "google_workspace_mcp.client.UserOAuthCredentials.from_authorized_user_file",
                return_value=fake_credentials,
            ) as from_file, patch.object(client, "_save_user_credentials") as save_mock:
                credentials = client._user_oauth_credentials([workspace.SHEETS_SCOPE])

        self.assertIs(credentials, fake_credentials)
        self.assertEqual(from_file.call_args.args[1], [workspace.SHEETS_WRITE_SCOPE])
        save_mock.assert_called_once_with(fake_credentials)
        self.assertEqual(credentials.token, "fresh-access-token")

    def test_auth_summary_reports_cached_and_missing_scopes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "scopes": [workspace.SHEETS_SCOPE],
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(Path(temp_dir) / "client-secret.json"),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(token_file),
                }
            )
            summary = client.auth_summary()

        self.assertEqual(summary["oauth_token_scopes"], [workspace.SHEETS_SCOPE])
        self.assertIn(workspace.DRIVE_SCOPE, summary["oauth_token_missing_scopes"])
        self.assertFalse(summary["oauth_token_capabilities"]["sheets_url_drive_export_fallback"])
        self.assertEqual(summary["active_auth_mode"], "oauth_client_cached_token")

    def test_auth_summary_treats_sheets_write_scope_as_covering_reads(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "scopes": [workspace.SHEETS_WRITE_SCOPE],
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(Path(temp_dir) / "client-secret.json"),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(token_file),
                }
            )
            summary = client.auth_summary()

        self.assertEqual(
            summary["oauth_token_missing_scopes"],
            [
                workspace.CHAT_MEMBERSHIPS_SCOPE,
                workspace.CHAT_MESSAGES_SCOPE,
                workspace.CHAT_SPACES_SCOPE,
                workspace.DOCS_SCOPE,
                workspace.DRIVE_SCOPE,
            ],
        )
        self.assertTrue(summary["oauth_token_capabilities"]["sheets_readonly"])
        self.assertTrue(summary["oauth_token_capabilities"]["sheets_write"])
        self.assertIn(workspace.DOCS_WRITE_SCOPE, summary["oauth_token_missing_readwrite_scopes"])
        self.assertIn(workspace.DRIVE_WRITE_SCOPE, summary["oauth_token_missing_all_write_scopes"])

    def test_auth_summary_reports_drive_export_fallback_when_only_drive_scope_is_cached(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "scopes": [workspace.DRIVE_SCOPE],
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(Path(temp_dir) / "client-secret.json"),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(token_file),
                }
            )
            summary = client.auth_summary()

        self.assertTrue(summary["oauth_token_capabilities"]["drive_readonly"])
        self.assertFalse(summary["oauth_token_capabilities"]["docs_readonly"])
        self.assertFalse(summary["oauth_token_capabilities"]["sheets_readonly"])
        self.assertTrue(summary["oauth_token_capabilities"]["sheets_url_drive_export_fallback"])
        self.assertIn("gid/range", " ".join(summary["notes"]))
        self.assertIn("documents.readonly", " ".join(summary["notes"]))

    def test_auth_summary_treats_broader_write_scopes_as_read_capable(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "scopes": [
                            workspace.DOCS_WRITE_SCOPE,
                            workspace.DRIVE_WRITE_SCOPE,
                            workspace.SHEETS_WRITE_SCOPE,
                            workspace.CHAT_SPACES_WRITE_SCOPE,
                            workspace.CHAT_MESSAGES_WRITE_SCOPE,
                            workspace.CHAT_MEMBERSHIPS_WRITE_SCOPE,
                        ],
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(Path(temp_dir) / "client-secret.json"),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(token_file),
                }
            )
            summary = client.auth_summary()

        self.assertEqual(summary["oauth_token_missing_scopes"], [])
        self.assertEqual(summary["oauth_token_missing_readwrite_scopes"], [])
        self.assertEqual(summary["oauth_token_missing_all_write_scopes"], [])
        self.assertTrue(summary["oauth_token_capabilities"]["drive_readonly"])
        self.assertTrue(summary["oauth_token_capabilities"]["drive_write"])
        self.assertTrue(summary["oauth_token_capabilities"]["chat_spaces_readonly"])
        self.assertTrue(summary["oauth_token_capabilities"]["chat_spaces_write"])
        self.assertTrue(summary["oauth_token_capabilities"]["chat_messages_readonly"])
        self.assertTrue(summary["oauth_token_capabilities"]["chat_messages_write"])
        self.assertTrue(summary["oauth_token_capabilities"]["chat_memberships_readonly"])
        self.assertTrue(summary["oauth_token_capabilities"]["chat_memberships_write"])

    def test_auth_summary_marks_missing_oauth_client_secret_file_as_not_configured(self) -> None:
        client = self.make_client(
            {
                "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": "C:/Users/TestUser/.google-workspace-mcp/oauth-client-secret.json",
            }
        )
        summary = client.auth_summary()

        self.assertFalse(summary["oauth_client_configured"])
        self.assertEqual(summary["recommended_mode"], "missing_credentials")
        self.assertIn("not found", summary["oauth_client_path_issue"])

    def test_auth_headers_falls_back_to_api_key_when_cached_oauth_token_lacks_scope(self) -> None:
        client = self.make_client({"GOOGLE_API_KEY": "public-key"})

        with patch.object(
            client,
            "_user_oauth_credentials",
            side_effect=RuntimeError("Cached OAuth token is missing required scopes."),
        ), patch.object(client, "_oauth_client_is_configured", return_value=True):
            headers, params = client._auth_headers([workspace.SHEETS_SCOPE], allow_api_key=True)

        self.assertEqual(headers, {})
        self.assertEqual(params, {"key": "public-key"})

    def test_auth_headers_falls_back_to_service_account_when_cached_oauth_token_lacks_scope(self) -> None:
        client = self.make_client()
        client.service_account_json = "{}"

        class FakeScopedCredentials:
            def __init__(self) -> None:
                self.valid = False
                self.token = None

            def refresh(self, _request: object) -> None:
                self.valid = True
                self.token = "service-account-token"

        class FakeBaseCredentials:
            def with_scopes(self, scopes: list[str] | tuple[str, ...]) -> FakeScopedCredentials:
                self.requested_scopes = tuple(scopes)
                return FakeScopedCredentials()

        fake_base = FakeBaseCredentials()

        with patch.object(
            client,
            "_user_oauth_credentials",
            side_effect=RuntimeError("Cached OAuth token is missing required scopes."),
        ), patch.object(client, "_oauth_client_is_configured", return_value=True), patch.object(
            client,
            "_service_account_base",
            return_value=fake_base,
        ):
            headers, params = client._auth_headers([workspace.SHEETS_SCOPE], allow_api_key=False)

        self.assertEqual(params, {})
        self.assertEqual(headers["Authorization"], "Bearer service-account-token")
        self.assertEqual(fake_base.requested_scopes, (workspace.SHEETS_SCOPE,))

    def test_chat_request_error_explains_missing_chat_app_configuration(self) -> None:
        client = self.make_client()
        failed_response = Mock()
        failed_response.ok = False
        failed_response.status_code = 404
        failed_response.text = '{"error":{"message":"Google Chat app not found. To create a Chat app, you must turn on the Chat API and configure the app in the Google Cloud console."}}'
        failed_response.headers = {}
        failed_response.json.return_value = {
            "error": {
                "message": "Google Chat app not found. To create a Chat app, you must turn on the Chat API and configure the app in the Google Cloud console."
            }
        }
        client.session.request = Mock(return_value=failed_response)

        with patch.object(client, "_auth_headers", return_value=({"Authorization": "Bearer test"}, {})):
            with self.assertRaises(RuntimeError) as caught:
                client.list_chat_spaces()

        self.assertIn("does not have a Chat app configuration yet", str(caught.exception))
        self.assertIn("Google Chat API > Configuration", str(caught.exception))

    def test_oauth_flow_uses_detected_google_download_filename(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            downloaded_file = temp_path / "client_secret_123.apps.googleusercontent.com.json"
            downloaded_file.write_text(
                json.dumps(
                    {
                        "installed": {
                            "client_id": "client-id",
                            "client_secret": "client-secret",
                            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                            "token_uri": "https://oauth2.googleapis.com/token",
                        }
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client(
                {
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(temp_path / "oauth-client-secret.json"),
                }
            )

            with patch(
                "google_workspace_mcp.client.InstalledAppFlow.from_client_secrets_file",
                return_value=Mock(name="flow"),
            ) as flow_factory:
                flow = client._oauth_flow([workspace.SHEETS_SCOPE])

        self.assertEqual(flow, flow_factory.return_value)
        flow_factory.assert_called_once_with(
            str(downloaded_file),
            [workspace.SHEETS_SCOPE],
        )

    def test_oauth_flow_raises_helpful_error_when_client_secret_file_is_missing(self) -> None:
        client = self.make_client(
            {
                "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": "C:/Users/TestUser/.google-workspace-mcp/oauth-client-secret.json",
            }
        )

        with self.assertRaises(RuntimeError) as caught:
            client._oauth_flow([workspace.SHEETS_SCOPE])

        self.assertIn("not found", str(caught.exception))
        self.assertIn("apps.googleusercontent.com", str(caught.exception))

    def test_oauth_flow_auto_discovers_client_secret_in_downloads_without_explicit_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            home_dir = Path(temp_dir)
            config_dir = home_dir / ".google-workspace-mcp"
            config_dir.mkdir()
            downloaded_file = config_dir / "client_secret_123.apps.googleusercontent.com.json"
            downloaded_file.write_text(
                json.dumps(
                    {
                        "installed": {
                            "client_id": "client-id",
                            "client_secret": "client-secret",
                            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                            "token_uri": "https://oauth2.googleapis.com/token",
                        }
                    }
                ),
                encoding="utf-8",
            )

            with patch.dict(os.environ, {"USERPROFILE": str(home_dir)}, clear=True):
                client = workspace.GoogleWorkspaceClient()
                with patch(
                    "google_workspace_mcp.client.InstalledAppFlow.from_client_secrets_file",
                    return_value=Mock(name="flow"),
                ) as flow_factory:
                    flow = client._oauth_flow([workspace.SHEETS_SCOPE])
                    summary = client.auth_summary()

        self.assertEqual(flow, flow_factory.return_value)
        flow_factory.assert_called_once_with(
            str(downloaded_file),
            [workspace.SHEETS_SCOPE],
        )
        self.assertTrue(summary["oauth_client_configured"])
        self.assertEqual(summary["oauth_client_source"], str(downloaded_file))

    def test_oauth_flow_raises_helpful_error_when_multiple_client_secret_files_are_auto_discovered(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            home_dir = Path(temp_dir)
            config_dir = home_dir / ".google-workspace-mcp"
            config_dir.mkdir()
            payload = json.dumps(
                {
                    "installed": {
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                    }
                }
            )
            (config_dir / "client_secret_one.apps.googleusercontent.com.json").write_text(
                payload,
                encoding="utf-8",
            )
            (config_dir / "client_secret_two.apps.googleusercontent.com.json").write_text(
                payload,
                encoding="utf-8",
            )

            with patch.dict(os.environ, {"USERPROFILE": str(home_dir)}, clear=True):
                client = workspace.GoogleWorkspaceClient()
                with self.assertRaises(RuntimeError) as caught:
                    client._oauth_flow([workspace.SHEETS_SCOPE])

        self.assertIn("Multiple OAuth client secrets files were auto-discovered", str(caught.exception))
        self.assertIn("Pass --client-secrets", str(caught.exception))

    def test_oauth_flow_uses_bundled_client_config_when_available(self) -> None:
        client = self.make_client()
        bundled_config = {
            "installed": {
                "client_id": "client-id",
                "client_secret": "client-secret",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        }

        with patch.object(
            client,
            "_resolved_oauth_client_secrets_file",
            return_value=None,
        ), patch.object(
            client,
            "_bundled_oauth_client_config_json",
            return_value=bundled_config,
        ), patch(
            "google_workspace_mcp.client.InstalledAppFlow.from_client_config",
            return_value=Mock(name="flow"),
        ) as flow_factory:
            flow = client._oauth_flow([workspace.SHEETS_SCOPE])
            summary = client.auth_summary()

        self.assertEqual(flow, flow_factory.return_value)
        flow_factory.assert_called_once()
        call_args, call_kwargs = flow_factory.call_args
        self.assertEqual(call_args[0], bundled_config)
        requested_scopes = call_args[1] if len(call_args) > 1 else call_kwargs.get("scopes")
        self.assertEqual(requested_scopes, [workspace.SHEETS_SCOPE])
        self.assertTrue(summary["oauth_client_configured"])
        self.assertEqual(summary["oauth_client_source"], "google_workspace_mcp/oauth-default-client.json")

    def test_run_oauth_login_persists_client_secret_file_to_default_config_location(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            home_dir = Path(temp_dir)
            source_dir = home_dir / "source"
            source_dir.mkdir()
            source_file = source_dir / "client_secret_123.apps.googleusercontent.com.json"
            source_file.write_text(
                json.dumps(
                    {
                        "installed": {
                            "client_id": "client-id",
                            "client_secret": "client-secret",
                            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                            "token_uri": "https://oauth2.googleapis.com/token",
                        }
                    }
                ),
                encoding="utf-8",
            )

            with patch.dict(
                os.environ,
                {
                    "USERPROFILE": str(home_dir),
                    "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": str(source_file),
                    "GOOGLE_OAUTH_TOKEN_FILE": str(home_dir / ".google-workspace-mcp" / "oauth-token.json"),
                },
                clear=True,
            ):
                client = workspace.GoogleWorkspaceClient()
                fake_credentials = Mock()
                fake_credentials.to_json.return_value = '{"token":"access-token"}'
                fake_credentials.refresh_token = "refresh-token"
                fake_flow = Mock()
                fake_flow.run_local_server.return_value = fake_credentials

                with patch.object(client, "_oauth_flow", return_value=fake_flow):
                    result = client.run_oauth_login([workspace.SHEETS_SCOPE], open_browser=False, port=0)

            persisted_file = home_dir / ".google-workspace-mcp" / "oauth-client-secret.json"
            self.assertTrue(persisted_file.exists())
            self.assertEqual(result["oauth_client_secrets_file"], str(persisted_file))
            self.assertEqual(result["notes"], [])

    def test_persist_oauth_client_config_writes_default_config_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            home_dir = Path(temp_dir)
            default_client_file = home_dir / ".google-workspace-mcp" / "oauth-client-secret.json"
            with patch.dict(os.environ, {"USERPROFILE": str(home_dir)}, clear=True), patch(
                "google_workspace_mcp.client.default_oauth_client_secrets_file",
                return_value=default_client_file,
            ):
                client = workspace.GoogleWorkspaceClient()
                saved_path = client.persist_oauth_client_config("client-id", "client-secret")
                summary = client.auth_summary()
                self.assertTrue(saved_path.exists())
                payload = json.loads(saved_path.read_text(encoding="utf-8"))
                self.assertEqual(payload["installed"]["client_id"], "client-id")
                self.assertEqual(payload["installed"]["client_secret"], "client-secret")
                self.assertEqual(summary["oauth_client_source"], str(saved_path))

    def test_cli_login_prompts_for_client_credentials_when_missing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            home_dir = Path(temp_dir)
            default_client_file = home_dir / ".google-workspace-mcp" / "oauth-client-secret.json"
            with patch.dict(os.environ, {"USERPROFILE": str(home_dir)}, clear=True), patch(
                "google_workspace_mcp.client.default_oauth_client_secrets_file",
                return_value=default_client_file,
            ), patch(
                "google_workspace_mcp.cli.default_oauth_client_secrets_file",
                return_value=default_client_file,
            ):
                client = workspace.GoogleWorkspaceClient()
                fake_result = {"oauth_token_file": "token.json", "notes": []}

                with patch("google_workspace_mcp.cli.parse_args", return_value=Namespace(
                    command="auth",
                    action="login",
                    client_secrets=None,
                    client_id=None,
                    client_secret=None,
                    token_file=None,
                    scopes=None,
                    scope_preset="readonly",
                    port=None,
                    no_browser=True,
                )), patch("google_workspace_mcp.cli.get_client", return_value=client), patch(
                    "google_workspace_mcp.cli.input", return_value="client-id"
                ), patch("google_workspace_mcp.cli.getpass.getpass", return_value="client-secret"), patch(
                    "google_workspace_mcp.cli.sys.stdin.isatty", return_value=True
                ), patch.object(client, "run_oauth_login", return_value=fake_result) as login_mock, patch(
                    "google_workspace_mcp.cli.print"
                ) as print_mock:
                    workspace.main([])
                persisted_file = default_client_file
                self.assertTrue(persisted_file.exists())
                login_mock.assert_called_once()
                printed_payload = print_mock.call_args.args[0]
                self.assertIn("oauth_client_secrets_file", printed_payload)

    def test_cli_login_skips_prompt_when_bundled_client_is_available(self) -> None:
        client = self.make_client()
        fake_result = {"oauth_token_file": "token.json", "notes": []}

        with patch("google_workspace_mcp.cli.parse_args", return_value=Namespace(
            command="auth",
            action="login",
            client_secrets=None,
            client_id=None,
            client_secret=None,
            token_file=None,
            scopes=None,
            scope_preset="readonly",
            port=None,
            no_browser=True,
        )), patch("google_workspace_mcp.cli.get_client", return_value=client), patch.object(
            client,
            "_oauth_client_is_configured",
            return_value=True,
        ), patch.object(
            client,
            "_resolved_oauth_client_secrets_file",
            return_value=None,
        ), patch.object(
            client,
            "run_oauth_login",
            return_value=fake_result,
        ) as login_mock, patch("google_workspace_mcp.cli.input") as input_mock, patch(
            "google_workspace_mcp.cli.getpass.getpass"
        ) as getpass_mock, patch("google_workspace_mcp.cli.print") as print_mock:
            workspace.main([])

        login_mock.assert_called_once()
        input_mock.assert_not_called()
        getpass_mock.assert_not_called()
        self.assertTrue(print_mock.called)

    def test_cli_login_accepts_client_id_and_client_secret_flags(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            home_dir = Path(temp_dir)
            default_client_file = home_dir / ".google-workspace-mcp" / "oauth-client-secret.json"
            with patch.dict(os.environ, {"USERPROFILE": str(home_dir)}, clear=True), patch(
                "google_workspace_mcp.client.default_oauth_client_secrets_file",
                return_value=default_client_file,
            ), patch(
                "google_workspace_mcp.cli.default_oauth_client_secrets_file",
                return_value=default_client_file,
            ):
                client = workspace.GoogleWorkspaceClient()
                fake_result = {"oauth_token_file": "token.json", "notes": []}

                with patch("google_workspace_mcp.cli.parse_args", return_value=Namespace(
                    command="auth",
                    action="login",
                    client_secrets=None,
                    client_id="client-id",
                    client_secret="client-secret",
                    token_file=None,
                    scopes=None,
                    scope_preset="readonly",
                    port=None,
                    no_browser=True,
                )), patch("google_workspace_mcp.cli.get_client", return_value=client), patch.object(
                    client, "run_oauth_login", return_value=fake_result
                ) as login_mock, patch("google_workspace_mcp.cli.print") as print_mock:
                    workspace.main([])
                persisted_file = default_client_file
                self.assertTrue(persisted_file.exists())
                login_mock.assert_called_once()
                printed_payload = print_mock.call_args.args[0]
                self.assertIn("oauth_client_secrets_file", printed_payload)

    def test_cli_login_merges_scope_preset_with_extra_scopes(self) -> None:
        client = self.make_client()
        fake_result = {"oauth_token_file": "token.json", "notes": []}

        with patch("google_workspace_mcp.cli.parse_args", return_value=Namespace(
            command="auth",
            action="login",
            client_secrets=None,
            client_id=None,
            client_secret=None,
            token_file=None,
            scopes=["https://www.googleapis.com/auth/drive.metadata.readonly"],
            scope_preset="sheets-write",
            port=None,
            no_browser=True,
        )), patch("google_workspace_mcp.cli.get_client", return_value=client), patch.object(
            client,
            "_oauth_client_is_configured",
            return_value=True,
        ), patch.object(
            client,
            "_resolved_oauth_client_secrets_file",
            return_value=None,
        ), patch.object(
            client,
            "run_oauth_login",
            return_value=fake_result,
        ) as login_mock, patch("google_workspace_mcp.cli.print"):
            workspace.main([])

        requested_scopes = login_mock.call_args.kwargs["scopes"]
        self.assertIn(workspace.SHEETS_WRITE_SCOPE, requested_scopes)
        self.assertIn(workspace.DOCS_SCOPE, requested_scopes)
        self.assertIn("https://www.googleapis.com/auth/drive.metadata.readonly", requested_scopes)

    def test_cli_login_all_write_preset_requests_chat_drive_and_workspace_write_scopes(self) -> None:
        client = self.make_client()
        fake_result = {"oauth_token_file": "token.json", "notes": []}

        with patch("google_workspace_mcp.cli.parse_args", return_value=Namespace(
            command="auth",
            action="login",
            client_secrets=None,
            client_id=None,
            client_secret=None,
            token_file=None,
            scopes=None,
            scope_preset="all-write",
            port=None,
            no_browser=True,
        )), patch("google_workspace_mcp.cli.get_client", return_value=client), patch.object(
            client,
            "_oauth_client_is_configured",
            return_value=True,
        ), patch.object(
            client,
            "_resolved_oauth_client_secrets_file",
            return_value=None,
        ), patch.object(
            client,
            "run_oauth_login",
            return_value=fake_result,
        ) as login_mock, patch("google_workspace_mcp.cli.print"):
            workspace.main([])

        requested_scopes = login_mock.call_args.kwargs["scopes"]
        self.assertIn(workspace.DOCS_WRITE_SCOPE, requested_scopes)
        self.assertIn(workspace.DRIVE_WRITE_SCOPE, requested_scopes)
        self.assertIn(workspace.SHEETS_WRITE_SCOPE, requested_scopes)
        self.assertIn(workspace.CHAT_SPACES_WRITE_SCOPE, requested_scopes)
        self.assertIn(workspace.CHAT_MESSAGES_WRITE_SCOPE, requested_scopes)
        self.assertIn(workspace.CHAT_MEMBERSHIPS_WRITE_SCOPE, requested_scopes)

    def test_run_oauth_logout_deletes_cached_token_file_and_revokes_refresh_token(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "refresh_token": "refresh-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client({"GOOGLE_OAUTH_TOKEN_FILE": str(token_file)})
            client._user_credentials[("scope",)] = Mock()
            client.session.post = Mock(return_value=Mock(status_code=200, text=""))

            result = client.run_oauth_logout()

        self.assertFalse(token_file.exists())
        self.assertEqual(client._user_credentials, {})
        self.assertTrue(result["oauth_token_file_existed"])
        self.assertTrue(result["oauth_token_file_deleted"])
        self.assertTrue(result["revoked"])
        self.assertEqual(result["revoked_token_type"], "refresh_token")
        client.session.post.assert_called_once()

    def test_run_oauth_logout_deletes_cached_token_even_when_revocation_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "oauth-user-token.json"
            token_file.write_text(
                json.dumps(
                    {
                        "token": "access-token",
                        "client_id": "client-id",
                        "client_secret": "client-secret",
                    }
                ),
                encoding="utf-8",
            )
            client = self.make_client({"GOOGLE_OAUTH_TOKEN_FILE": str(token_file)})
            client.session.post = Mock(return_value=Mock(status_code=400, text='{"error":"invalid_token"}'))

            result = client.run_oauth_logout()

        self.assertFalse(token_file.exists())
        self.assertFalse(result["revoked"])
        self.assertEqual(result["revoked_token_type"], "access_token")
        self.assertIn("token revocation may not have completed", result["notes"][0])

    def test_run_oauth_logout_reports_missing_token_file(self) -> None:
        client = self.make_client()

        result = client.run_oauth_logout()

        self.assertFalse(result["oauth_token_file_existed"])
        self.assertFalse(result["oauth_token_file_deleted"])
        self.assertFalse(result["revoked"])
        self.assertIn("No cached OAuth token file was found.", result["notes"])

    def test_run_oauth_logout_warns_when_env_access_token_is_configured(self) -> None:
        client = self.make_client({"GOOGLE_OAUTH_ACCESS_TOKEN": "ya29.test-token"})

        result = client.run_oauth_logout()

        self.assertTrue(result["oauth_access_token_configured"])
        self.assertTrue(any("GOOGLE_OAUTH_ACCESS_TOKEN" in note for note in result["notes"]))

    def test_python_module_entrypoint_invokes_cli_main(self) -> None:
        with patch("google_workspace_mcp.cli.main") as mocked_main:
            runpy.run_module("google_workspace_mcp", run_name="__main__")

        mocked_main.assert_called_once_with()

    def test_resolve_google_file_falls_back_to_sheet_metadata_when_drive_scope_is_missing(self) -> None:
        client = self.make_client()
        client.get_drive_file = Mock(
            side_effect=RuntimeError(
                "Google API returned HTTP 403 for https://www.googleapis.com/drive/v3/files/test: Request had insufficient authentication scopes."
            )
        )
        client.get_sheet_metadata = Mock(
            return_value={
                "spreadsheetId": TEST_SPREADSHEET_ID,
                "properties": {"title": "Spec Workbook"},
                "sheets": [],
            }
        )

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            resolved = workspace.resolve_google_file(
                "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1544244212"
            )

        self.assertEqual(resolved["source"], "sheets_metadata_fallback")
        self.assertEqual(resolved["name"], "Spec Workbook")
        self.assertIn("drive.readonly", resolved["auth_warning"])

    def test_check_sheet_edit_access_reports_missing_write_scope_when_drive_says_editable(self) -> None:
        client = self.make_client()
        client.auth_summary = Mock(
            return_value={
                "oauth_token_scopes": [workspace.DRIVE_SCOPE],
                "oauth_token_capabilities": {"drive_readonly": True, "sheets_write": False},
            }
        )
        client._auth_headers = Mock(side_effect=RuntimeError("Cached OAuth token is missing required scopes."))
        client.get_drive_file = Mock(
            return_value={
                "id": TEST_SPREADSHEET_ID,
                "name": "BD_Promotion",
                "webViewLink": "https://docs.google.com/spreadsheets/d/test/edit",
                "owners": [{"displayName": "Chi Vu", "emailAddress": "chi.vu@sotatek.com"}],
                "capabilities": {"canEdit": True, "canModifyContent": True},
            }
        )

        with patch("google_workspace_mcp.tools.get_client", return_value=client):
            result = workspace.check_sheet_edit_access(
                "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1544244212"
            )

        self.assertFalse(result["api_write_ready"])
        self.assertTrue(result["drive_capabilities"]["can_edit"])
        self.assertFalse(result["can_write_via_api"])
        self.assertIn("scope-preset sheets-write", " ".join(result["notes"]))

    def test_update_sheet_values_uses_write_scope_and_put_request(self) -> None:
        client = self.make_client()
        client.get_sheet_metadata = Mock(
            return_value={
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 1436003411,
                            "title": "Feedback",
                            "gridProperties": {"rowCount": 300, "columnCount": 17},
                        }
                    }
                ]
            }
        )
        client._request = Mock(return_value={"updatedRange": "'Feedback'!C119:D119", "updatedCells": 2})

        payload = client.update_sheet_values(
            "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit?gid=1436003411",
            "C119:D119",
            [["ok", "done"]],
            value_input_option="RAW",
            major_dimension="ROWS",
            include_values_in_response=False,
        )

        self.assertEqual(payload["updatedCells"], 2)
        client._request.assert_called_once_with(
            "PUT",
            "https://sheets.googleapis.com/v4/spreadsheets/1AbCdEfGhIjKlMnOp/values/%27Feedback%27!C119:D119",
            scopes=[workspace.SHEETS_WRITE_SCOPE],
            allow_api_key=False,
            params={
                "valueInputOption": "RAW",
                "includeValuesInResponse": "false",
                "responseValueRenderOption": "FORMATTED_VALUE",
            },
            json_body={
                "majorDimension": "ROWS",
                "values": [["ok", "done"]],
            },
        )

    def test_annotate_formatted_text_marks_strikethrough_and_underline_segments(self) -> None:
        annotated = workspace.annotate_formatted_text(
            "Thời gian đăng kí chốt",
            [
                {},
                {"start_index": 10, "format": {"strikethrough": True}},
                {"start_index": 17},
                {"start_index": 18, "format": {"underline": True}},
            ],
        )

        self.assertEqual(
            annotated,
            "Thời gian [[STRIKE]]đăng kí[[/STRIKE]] [[UNDERLINE]]chốt[[/UNDERLINE]]",
        )

    def test_simplify_grid_data_includes_annotated_text_for_styled_cells(self) -> None:
        payload = {
            "sheets": [
                {
                    "properties": {
                        "sheetId": 1544244212,
                        "title": "Spec",
                        "index": 0,
                        "gridProperties": {"rowCount": 100, "columnCount": 20},
                    },
                    "data": [
                        {
                            "startRow": 37,
                            "startColumn": 9,
                            "rowData": [
                                {
                                    "values": [
                                        {
                                            "formattedValue": "登録日時\n確定日時\n",
                                            "userEnteredValue": {"stringValue": "登録日時\n確定日時\n"},
                                            "effectiveValue": {"stringValue": "登録日時\n確定日時\n"},
                                            "textFormatRuns": [
                                                {"startIndex": 0, "format": {"strikethrough": True}},
                                                {"startIndex": 5, "format": {}},
                                            ],
                                        }
                                    ]
                                }
                            ],
                        }
                    ],
                }
            ]
        }

        sheets = workspace.simplify_grid_data(payload)
        cell = sheets[0]["data"][0]["rows"][0]["cells"][0]

        self.assertEqual(cell["a1"], "J38")
        self.assertEqual(cell["annotated_text"], "[[STRIKE]]登録日時\n[[/STRIKE]]確定日時\n")
        self.assertEqual(cell["text_runs"][0]["format"]["strikethrough"], True)


if __name__ == "__main__":
    unittest.main()

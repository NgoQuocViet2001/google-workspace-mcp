import json
import os
import tempfile
import unittest
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
        self.assertEqual(summary["active_auth_mode"], "oauth_client_cached_token")

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

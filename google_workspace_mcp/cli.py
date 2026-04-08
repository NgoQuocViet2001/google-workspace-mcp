from __future__ import annotations

import argparse
import getpass
import json
import sys
from pathlib import Path

from .client import GoogleWorkspaceClient, get_client
from .common import DEFAULT_OAUTH_SCOPES, default_oauth_client_secrets_file
from .server import mcp
from . import tools as _tools  # noqa: F401


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
        choices=("login", "status", "logout"),
        default="login",
        help="Choose `login` to launch the browser flow, `status` to inspect the current setup, or `logout` to delete the cached OAuth token.",
    )
    auth_parser.add_argument(
        "--client-secrets",
        dest="client_secrets",
        help="Optional path to the OAuth client secrets JSON file. If omitted, the CLI looks for a desktop-app client JSON in ~/.google-workspace-mcp.",
    )
    auth_parser.add_argument(
        "--client-id",
        dest="client_id",
        help="Optional Google OAuth desktop-app client ID. If paired with --client-secret, the CLI saves a local OAuth client config automatically.",
    )
    auth_parser.add_argument(
        "--client-secret",
        dest="client_secret",
        help="Optional Google OAuth desktop-app client secret. If paired with --client-id, the CLI saves a local OAuth client config automatically.",
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


def _prompt_non_empty(prompt: str, *, secret: bool = False) -> str:
    reader = getpass.getpass if secret else input
    while True:
        value = reader(prompt).strip()
        if value:
            return value
        print("Value is required.", file=sys.stderr)


def _prepare_oauth_client_config(client: GoogleWorkspaceClient, args: argparse.Namespace) -> Path | None:
    if args.client_id or args.client_secret:
        if not args.client_id or not args.client_secret:
            raise RuntimeError("Pass both --client-id and --client-secret together.")
        return client.persist_oauth_client_config(args.client_id, args.client_secret)

    if client._oauth_client_is_configured():
        return client._resolved_oauth_client_secrets_file()

    if not sys.stdin.isatty():
        return None

    print(
        "No OAuth client configuration found. Enter your Google OAuth Desktop App client ID and client secret once, "
        f"and they will be saved to {default_oauth_client_secrets_file()}.",
        file=sys.stderr,
    )
    client_id = _prompt_non_empty("Client ID: ")
    client_secret = _prompt_non_empty("Client Secret: ", secret=True)
    return client.persist_oauth_client_config(client_id, client_secret)


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
            if args.action == "logout":
                print(json.dumps(client.run_oauth_logout(), indent=2))
                return
            prepared_client_config = _prepare_oauth_client_config(client, args)
            result = client.run_oauth_login(
                scopes=args.scopes or DEFAULT_OAUTH_SCOPES,
                open_browser=not args.no_browser,
                port=args.port,
            )
            if prepared_client_config and not result.get("oauth_client_secrets_file"):
                result["oauth_client_secrets_file"] = str(prepared_client_config)
            print(json.dumps(result, indent=2))
            return
        except RuntimeError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            raise SystemExit(1) from exc

    mcp.run()

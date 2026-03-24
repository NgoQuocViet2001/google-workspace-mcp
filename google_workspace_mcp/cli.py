from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .client import get_client
from .common import DEFAULT_READONLY_SCOPES
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
            if args.action == "logout":
                print(json.dumps(client.run_oauth_logout(), indent=2))
                return
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

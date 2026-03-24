# Google Workspace MCP

Python MCP server for reading Google Docs, Google Sheets, and Google Chat with structured output and better image handling.

## What It Does

- Reads Google Docs as structured JSON with paragraphs, tables, inline objects, positioned objects, and image metadata.
- Reads Google Sheets values, grid data, formulas, notes, hyperlinks, and chip runs.
- Lists Google Chat spaces, space members, and messages as structured JSON.
- Preserves partial text styling in Sheets cells via `text_runs` and an `annotated_text` helper field for segments such as strikethrough and underline.
- Extracts over-grid sheet images from `Drive export -> XLSX`.
- Detects in-cell `IMAGE("...")` formulas separately from drawing exports.

## Project Layout

```text
google_workspace_mcp/
  chat.py       # Google Chat normalization helpers
  cli.py        # command-line entrypoint
  client.py     # Google API auth + HTTP client
  common.py     # shared constants and parsing helpers
  docs.py       # Google Docs normalization helpers
  server.py     # FastMCP server instance
  sheets.py     # Google Sheets normalization helpers
  tools.py      # MCP tool definitions
mcp_google_workspace.py  # compatibility wrapper for local scripts/config
tests/
```

## Quick Start

### Install from GitHub

```powershell
pip install "git+https://github.com/NgoQuocViet2001/google-workspace-mcp.git"
```

The most reliable way to run it across platforms is:

```powershell
python -m google_workspace_mcp
```

If your Python Scripts directory is already on `PATH`, this shorter launcher works too:

```powershell
google-workspace-mcp
```

### Login OAuth quickly

1. In Google Cloud, create an OAuth client ID with application type `Desktop app`.
2. If the published package already includes a bundled OAuth desktop client, just run:

```powershell
python -m google_workspace_mcp auth login
```

That opens the browser OAuth flow directly.

3. If the package does not include a bundled OAuth client and no local OAuth client config exists yet, the CLI prompts once for:
   - `Client ID`
   - `Client Secret`

It then saves a reusable desktop-app client config at:

```powershell
$HOME/.google-workspace-mcp/oauth-client-secret.json
```

and opens the browser OAuth flow automatically. After the first successful login, the cached token is stored at:

```powershell
$HOME/.google-workspace-mcp/oauth-token.json
```

### Other common commands

```powershell
python -m google_workspace_mcp auth status
python -m google_workspace_mcp auth logout
```

## CLI Commands

Use `python -m google_workspace_mcp ...` everywhere below. If `google-workspace-mcp` is already on `PATH`, the same commands also work with that shorter launcher.

- `python -m google_workspace_mcp`
- `python -m google_workspace_mcp auth`
- `python -m google_workspace_mcp auth login`
- `python -m google_workspace_mcp auth login --client-secrets C:\path\to\oauth-client-secret.json`
- `python -m google_workspace_mcp auth login --client-id <client-id> --client-secret <client-secret>`
- `python -m google_workspace_mcp auth login --token-file C:\path\to\oauth-token.json`
- `python -m google_workspace_mcp auth status`
- `python -m google_workspace_mcp auth logout`

## Authentication Options

### Recommended for private files shared to your Google account: OAuth desktop client

Use a Google OAuth client ID for Desktop App if the files are private but shared to your personal Google account.

If you want end users to be able to run `python -m google_workspace_mcp auth login` and jump straight into the browser OAuth flow with no extra setup, publish the package with a bundled desktop-app client at:

```text
google_workspace_mcp/oauth-default-client.json
```

If no bundled client is shipped, the CLI falls back to prompting once for `Client ID` and `Client Secret`, or it can read them from a local JSON file.

1. Enable:
   - Google Sheets API
   - Google Docs API
   - Google Drive API
   - Google Chat API
2. Create an OAuth client ID with application type `Desktop app`.
3. Choose one setup method:
   - Easiest: run `python -m google_workspace_mcp auth login`, paste the `Client ID` and `Client Secret` once, and let the CLI save them for future logins.
   - If you prefer files: download the client secret JSON. The downloaded filename is often something like `client_secret_<id>.apps.googleusercontent.com.json`.
4. Optional file-based setup:

```powershell
$HOME/.google-workspace-mcp/oauth-client-secret.json
```

Or set:

```powershell
$env:GOOGLE_OAUTH_CLIENT_SECRETS_FILE="C:\path\to\oauth-client-secret.json"
```

5. Run the one-time browser login flow:

```powershell
python -m google_workspace_mcp auth
```

After the first successful login, the server automatically uses the cached OAuth token for private Docs, Sheets, Drive, and Google Chat calls. You do not need to provide a separate API key for that flow.

This stores a refreshable token by default at:

```powershell
$HOME\.google-workspace-mcp\oauth-token.json
```

Use this to inspect the cached token scopes and see which scopes are still missing:

```powershell
python -m google_workspace_mcp auth status
```

If you need to overwrite the cached token with a specific client secret file and token path, you can also run:

```powershell
python -m google_workspace_mcp auth login --client-secrets C:\path\to\oauth-client-secret.json --token-file C:\path\to\oauth-token.json
```

If the desktop-app client JSON is already in `$HOME/.google-workspace-mcp/`, this shorter command also works:

```powershell
python -m google_workspace_mcp auth login
```

When you log in with `--client-secrets`, or with `--client-id` plus `--client-secret`, the CLI also saves a reusable desktop-app client JSON into `$HOME/.google-workspace-mcp/oauth-client-secret.json` so future logins can omit the extra flags.

To delete the cached OAuth token later, run:

```powershell
python -m google_workspace_mcp auth logout
```

If you separately configured `GOOGLE_OAUTH_ACCESS_TOKEN`, remove that environment variable from your shell or MCP config as well.

### Recommended: service account

Use a Google Cloud service account for the most reliable setup.

1. Enable:
   - Google Sheets API
   - Google Docs API
   - Google Drive API
   - Google Chat API if you plan to call the Chat tools with a user-scoped bearer token
2. Create a service account key.
3. Share the target Docs/Sheets files with the service account email.
4. Set:

```powershell
$env:GOOGLE_SERVICE_ACCOUNT_FILE="C:\path\to\service-account.json"
```

### Public Sheets only: API key

Suitable for public Google Sheets reads. Not recommended for Docs or Drive export.

```powershell
$env:GOOGLE_API_KEY="your_api_key"
```

### Existing bearer token: OAuth access token

```powershell
$env:GOOGLE_OAUTH_ACCESS_TOKEN="ya29...."
```

## Installation

```powershell
git clone https://github.com/NgoQuocViet2001/google-workspace-mcp.git
cd google-workspace-mcp
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Install from GitHub

### Option 1: clone the repository

```powershell
git clone https://github.com/NgoQuocViet2001/google-workspace-mcp.git
cd google-workspace-mcp
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Option 2: install directly from GitHub

```powershell
pip install "git+https://github.com/NgoQuocViet2001/google-workspace-mcp.git"
```

After installing from GitHub, the most reliable way to run it is:

```powershell
python -m google_workspace_mcp
```

This avoids `PATH` issues when `pip` installs console scripts into a user-site Scripts directory.

If your Python Scripts directory is already on `PATH`, the standalone command also works:

```powershell
google-workspace-mcp
```

## Running the Server

```powershell
cd <path-to-repo>
.venv\Scripts\python.exe mcp_google_workspace.py
```

Or, if you installed it directly from GitHub:

```powershell
python -m google_workspace_mcp
```

To bootstrap OAuth for a private user account:

```powershell
python -m google_workspace_mcp auth
```

If your `google-workspace-mcp` launcher is already available on `PATH`, the equivalent shorter command is:

```powershell
google-workspace-mcp auth login
```

To inspect the current auth setup:

```powershell
python -m google_workspace_mcp auth status
```

To remove the cached OAuth login:

```powershell
python -m google_workspace_mcp auth logout
```

## Codex MCP Configuration

```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "<path-to-repo>/.venv/Scripts/python.exe",
      "args": ["<path-to-repo>/mcp_google_workspace.py"],
      "env": {
        "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": "C:/path/to/oauth-client-secret.json",
        "GOOGLE_OAUTH_TOKEN_FILE": "C:/path/to/oauth-token.json"
      }
    }
  }
}
```

For public Sheets only, replace the `env` block with:

```json
{
  "GOOGLE_API_KEY": "your_api_key"
}
```

If you installed the package directly from GitHub into an environment on your PATH, you can also use:

```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "google-workspace-mcp",
      "env": {
        "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": "C:/path/to/oauth-client-secret.json",
        "GOOGLE_OAUTH_TOKEN_FILE": "C:/path/to/oauth-token.json"
      }
    }
  }
}
```

## Available Tools

- `diagnose_google_auth`
- `resolve_google_file`
- `list_google_chat_spaces`
- `get_google_chat_space`
- `get_google_chat_message`
- `read_google_chat_messages`
- `read_google_chat_thread`
- `list_google_chat_memberships`
- `read_sheet_values`
- `read_sheet_grid`
- `get_sheet_row`
- `search_sheet`
- `sheet_to_json`
- `inspect_sheet_images`
- `read_google_doc`
- `download_google_doc_images`
- `export_google_file`

## Example Prompts

Replace placeholders such as `<spreadsheet-id>`, `<sheet-name>`, and `<output-dir>` with your own values.

Google Sheets URLs with `gid` and `range` are resolved automatically. If the caller omits the sheet prefix, the server uses the tab identified by `gid`.

### Read one row from a sheet

```text
get_sheet_row(
  "<spreadsheet-id>",
  "<sheet-name>",
  42,
  1
)
```

`read_sheet_values` also accepts row-style input such as `<sheet-name>!42:42` and normalizes it to a valid full-row A1 range automatically.

### Read directly from a Sheets URL with `gid` and `range`

```text
read_sheet_values(
  "https://docs.google.com/spreadsheets/d/<spreadsheet-id>/edit?gid=<gid>#gid=<gid>&range=38:38"
)
```

### Read grid data with formulas, notes, and links

```text
read_sheet_grid(
  "<spreadsheet-id>",
  "<sheet-name>!A1:Z200"
)
```

For cells with partial formatting, `read_sheet_grid()` now includes:

- `text_runs`: structured offsets plus style flags from Google Sheets
- `annotated_text`: a plain-text helper string such as `[[STRIKE]]old[[/STRIKE]] new`

### Search across a sheet

```text
search_sheet(
  "<spreadsheet-id>",
  "login"
)
```

If you pass a Sheets URL with `gid`, `search_sheet()` searches only that tab by default instead of scanning the full workbook.

### List Google Chat spaces

```text
list_google_chat_spaces()
```

### Read messages from a Google Chat space

`read_google_chat_messages()` accepts either a resource name like `spaces/AAAA...` or a Chat UI URL that contains the space id.

```text
read_google_chat_messages(
  "spaces/AAAA1234567",
  50,
  null,
  null,
  "DESC",
  false
)
```

### Read one Google Chat thread from a thread URL

`read_google_chat_thread()` accepts either a thread resource like `spaces/<space>/threads/<thread>` or a Chat UI URL like `https://chat.google.com/room/<space>/<thread>/<message>`.

```text
read_google_chat_thread(
  "https://chat.google.com/room/AAQAyxdRoZo/jVIpmenXnO0/WNSdv6IyQf0?cls=10"
)
```

When the URL includes a message ID, the response includes both:

- `linked_message`: the exact message referenced by the link
- `root_message`: the first message in the thread

If Google Chat doesn't expose an API message resource that matches the URL token, the response sets `linked_message` to `null` and explains the limitation in `linked_message_lookup_warning`.

### List members in a Google Chat space

```text
list_google_chat_memberships(
  "spaces/AAAA1234567"
)
```

### Convert a sheet to JSON

```text
sheet_to_json(
  "<spreadsheet-id>",
  "<sheet-name>",
  1
)
```

### Extract images from a sheet

```text
inspect_sheet_images(
  "<spreadsheet-id>",
  "<sheet-name>",
  "C:/path/to/output/sheet-images"
)
```

### Read a Google Doc with text and image metadata

```text
read_google_doc(
  "https://docs.google.com/document/d/<doc-id>/edit",
  null,
  false,
  null
)
```

### Download images from a Google Doc

```text
download_google_doc_images(
  "https://docs.google.com/document/d/<doc-id>/edit",
  "C:/path/to/output/doc-images",
  null
)
```

## Practical Limitations

- Google Docs image metadata is available directly through the Docs API, so document extraction is strong.
- Google Sheets does not expose over-grid images as cleanly as cell data, so this server uses XLSX export to recover them.
- Google Chat reads require OAuth scopes such as `chat.spaces.readonly`, `chat.messages.readonly`, and `chat.memberships.readonly`. If your cached token is older, rerun `python -m google_workspace_mcp auth login`.
- Google Chat API requests also require a configured Chat app in the same Google Cloud project. In `Google Chat API > Configuration`, fill in at least `App name`, `Avatar URL`, and `Description`, then save.
- Google Chat room URLs always expose the space ID, and often work for thread reads, but the final URL token isn't guaranteed to be a `spaces/{space}/messages/{message}` API resource name. In those cases the server can still return the thread and root message, but not reliably resolve the exact linked reply through the Chat API alone.
- Google Chat private user conversations are most reliable with OAuth user credentials. A plain service account usually needs a properly configured Chat app flow to access Chat resources.
- In-cell `IMAGE("...")` formulas are detected separately from exported drawing images.
- Private files shared to your user account should use the OAuth desktop client flow.
- Private files shared to a robot identity should use a service account.
- An API key is only suitable for public Sheets and can't read Google Chat.

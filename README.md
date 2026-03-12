# Google Workspace MCP

Python MCP server for reading Google Docs and Google Sheets with structured output and better image handling.

## What It Does

- Reads Google Docs as structured JSON with paragraphs, tables, inline objects, positioned objects, and image metadata.
- Reads Google Sheets values, grid data, formulas, notes, hyperlinks, and chip runs.
- Extracts over-grid sheet images from `Drive export -> XLSX`.
- Detects in-cell `IMAGE("...")` formulas separately from drawing exports.

## Authentication Options

### Recommended for private files shared to your Google account: OAuth desktop client

Use a Google OAuth client ID for Desktop App if the files are private but shared to your personal Google account.

1. Enable:
   - Google Sheets API
   - Google Docs API
   - Google Drive API
2. Create an OAuth client ID with application type `Desktop app`.
3. Download the client secret JSON.
4. Set:

```powershell
$env:GOOGLE_OAUTH_CLIENT_SECRETS_FILE="C:\path\to\oauth-client-secret.json"
```

5. Run the one-time browser login flow:

```powershell
google-workspace-mcp auth
```

This stores a refreshable token by default at:

```powershell
$HOME\.google-workspace-mcp\oauth-token.json
```

Use this to inspect the cached token scopes and see which scopes are still missing:

```powershell
google-workspace-mcp auth status
```

If you need to overwrite the cached token with a specific client secret file and token path, you can also run:

```powershell
google-workspace-mcp auth login --client-secrets C:\path\to\oauth-client-secret.json --token-file C:\path\to\oauth-token.json
```

### Recommended: service account

Use a Google Cloud service account for the most reliable setup.

1. Enable:
   - Google Sheets API
   - Google Docs API
   - Google Drive API
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

If you install it this way, the console entrypoint is:

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
google-workspace-mcp
```

To bootstrap OAuth for a private user account:

```powershell
google-workspace-mcp auth
```

To inspect the current auth setup:

```powershell
google-workspace-mcp auth status
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

### Search across a sheet

```text
search_sheet(
  "<spreadsheet-id>",
  "login"
)
```

If you pass a Sheets URL with `gid`, `search_sheet()` searches only that tab by default instead of scanning the full workbook.

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
- In-cell `IMAGE("...")` formulas are detected separately from exported drawing images.
- Private files shared to your user account should use the OAuth desktop client flow.
- Private files shared to a robot identity should use a service account.
- An API key is only suitable for public Sheets.

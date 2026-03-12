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
cd C:\Users\Admin\google-workspace-mcp
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
cd C:\Users\Admin\google-workspace-mcp
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
      "command": "C:/Users/Admin/google-workspace-mcp/.venv/Scripts/python.exe",
      "args": ["C:/Users/Admin/google-workspace-mcp/mcp_google_workspace.py"],
      "env": {
        "GOOGLE_OAUTH_CLIENT_SECRETS_FILE": "C:/path/to/oauth-client-secret.json",
        "GOOGLE_OAUTH_TOKEN_FILE": "C:/Users/Admin/.google-workspace-mcp/oauth-token.json"
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
        "GOOGLE_OAUTH_TOKEN_FILE": "C:/Users/Admin/.google-workspace-mcp/oauth-token.json"
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

### Read one row from a sheet

```text
get_sheet_row(
  "1_6tB3R932HqKHYJoJRZByEFueFdOnJhR4v6IyDkwdnU",
  "Sheet1",
  129,
  1
)
```

`read_sheet_values` also accepts row-style input such as `Sheet1!129:129` and normalizes it to a valid full-row A1 range automatically.

### Read grid data with formulas, notes, and links

```text
read_sheet_grid(
  "1_6tB3R932HqKHYJoJRZByEFueFdOnJhR4v6IyDkwdnU",
  "Sheet1!A1:Z200"
)
```

### Search across a sheet

```text
search_sheet(
  "1_6tB3R932HqKHYJoJRZByEFueFdOnJhR4v6IyDkwdnU",
  "login"
)
```

### Convert a sheet to JSON

```text
sheet_to_json(
  "1_6tB3R932HqKHYJoJRZByEFueFdOnJhR4v6IyDkwdnU",
  "Sheet1",
  1
)
```

### Extract images from a sheet

```text
inspect_sheet_images(
  "1_6tB3R932HqKHYJoJRZByEFueFdOnJhR4v6IyDkwdnU",
  "Sheet1",
  "C:/Users/Admin/google-workspace-mcp/out/sheet-images"
)
```

### Read a Google Doc with text and image metadata

```text
read_google_doc(
  "https://docs.google.com/document/d/FILE_ID/edit",
  null,
  false,
  null
)
```

### Download images from a Google Doc

```text
download_google_doc_images(
  "https://docs.google.com/document/d/FILE_ID/edit",
  "C:/Users/Admin/google-workspace-mcp/out/doc-images",
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

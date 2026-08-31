# @lnwebworks/outlook-mcp-server

An MCP (Model Context Protocol) server that connects Claude (or any MCP-compatible client) to Outlook via the Microsoft Graph API, using client-credentials (app-only) auth.

## Installation

```bash
npm install -g @lnwebworks/outlook-mcp-server
```

Or run directly with `npx`:

```bash
npx @lnwebworks/outlook-mcp-server
```

## Configuration

Copy `.env.example` to `.env` and fill in your Azure AD app registration details:

```bash
cp .env.example .env
```

Required:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `OUTLOOK_USER_EMAIL` — the mailbox this server acts on by default

Optional:

- `OUTLOOK_EXCLUDED_FOLDER_IDS`
- `OUTLOOK_SKIP_EXCLUDED_FOLDERS`
- `OUTLOOK_SIGNOFF_NAME`

## Usage with an MCP client

Add this server to your MCP client's config, e.g.:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@lnwebworks/outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "OUTLOOK_USER_EMAIL": "..."
      }
    }
  }
}
```

## License

MIT

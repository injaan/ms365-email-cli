# @injaan.dev/ms365-email-cli

MS365 mailbox CLI for AI-agent and automation workflows using Microsoft Graph API.

## Install

```bash
npm i -g @injaan.dev/ms365-email-cli
```

Verify:

```bash
ms365-email-cli --help
```

## What It Does

- Initialize credentials with an interactive wizard
- List unread or recent emails
- Search by text, sender, subject, date, or folder
- Read full email body and thread history
- List/download attachments
- Mark messages as read
- Send, reply, and reply-all (with optional attachments)

## Requirements

- Node.js 20+
- Azure AD app with Graph permissions:
  - `Mail.ReadWrite`
  - `Mail.Send`
- Admin consent granted

## Setup

Run the wizard:

```bash
ms365-email-cli init
```

`init` replaces the existing CLI `.env` values with a fresh configuration.

It creates or updates `.env` with:

- `AUTH_MODE` (`client_credentials` or `delegated`)
- `MS365_CLIENT_ID`

Auth mode guidance:

- `AUTH_MODE=delegated` is for personal Microsoft accounts (Outlook/Hotmail/Live)
- `AUTH_MODE=client_credentials` is for company/work accounts

When `AUTH_MODE=client_credentials`:

- `MS365_TENANT_ID`
- `MS365_CLIENT_SECRET`
- `MS365_EMAIL_ADDRESS`

When `AUTH_MODE=delegated`:

- Sign-in happens via device-code prompt in terminal
- API calls use the signed-in mailbox (`/me`)
- Wizard defaults `MS365_CLIENT_ID=90819426-b785-4919-a65e-818d7a8e9952` (you can enter your own client ID)
- Wizard auto-sets `MS365_TENANT_ID=consumers`
- If your tenant is set to `common` but app is Microsoft-account-only, the CLI auto-falls back to `/consumers`
- Access/refresh tokens are cached locally, so login is reused across runs until refresh expires or is revoked

## Quick Commands

```bash
ms365-email-cli unread -n 5
ms365-email-cli list -n 20
ms365-email-cli search -q "invoice"
ms365-email-cli read <MESSAGE_ID>
ms365-email-cli thread <MESSAGE_ID>
ms365-email-cli attachment <MESSAGE_ID> -o ./downloads
ms365-email-cli mark-read <MESSAGE_ID>
ms365-email-cli send -t user@example.com -s "Hello" -b "Hi there"
ms365-email-cli send -t user@example.com -s "Hello" --body-file email.html --html
ms365-email-cli send -t user@example.com -s "Hello" --body-stdin --html < email.html
ms365-email-cli send -t user@example.com -c manager@example.com -s "Hello" -b "Hi there"
ms365-email-cli reply <MESSAGE_ID> -b "Thanks"
ms365-email-cli reply <MESSAGE_ID> --body-file reply.html --html
ms365-email-cli reply-all <MESSAGE_ID> -b "Thanks everyone"
```

For HTML email, prefer `--body-file` or `--body-stdin` when the body contains
quotes, newlines, `<`, or `>` characters. This avoids shell-specific escaping
differences across macOS, Linux, Windows PowerShell, and Windows cmd.exe.

## AI-Agent Friendly

- Predictable command surface for tooling
- Script-friendly terminal output
- Full inbox workflows via CLI (read/search/reply/send/attachments)

## MCP Server

This package also includes an MCP stdio server. After installing this package,
you get both binaries:

```bash
ms365-email-cli
ms365-email-cli-mcp
```

Before using the MCP server, initialize the mailbox config once:

```bash
ms365-email-cli init
```

Run the server directly:

```bash
ms365-email-cli-mcp
```

MCP clients launch it as a stdio server. Example configs:

Claude Code:

```bash
claude mcp add --transport stdio ms365-email-cli -- ms365-email-cli-mcp
```

OpenAI Codex:

```bash
codex mcp add ms365-email-cli -- ms365-email-cli-mcp
```

Codex config TOML:

```toml
[mcp_servers."ms365-email-cli"]
command = "ms365-email-cli-mcp"
```

GitHub Copilot in VS Code `.vscode/mcp.json`:

```json
{
  "servers": {
    "ms365-email-cli": {
      "type": "stdio",
      "command": "ms365-email-cli-mcp"
    }
  }
}
```

Available MCP tools:

- `list_emails`
- `list_unread_emails`
- `read_email`
- `thread`
- `mark_read`
- `search_emails`
- `send_email`
- `reply`
- `reply_all`
- `attachment`

The `send_email`, `reply`, and `reply_all` tools accept either `body` or
`body_file`; use `body_file` for large or multiline HTML.

## Troubleshooting

- `command not found: ms365-email-cli`:
  reinstall globally with `npm i -g @injaan.dev/ms365-email-cli`
- `command not found: ms365-email-cli-mcp`:
  reinstall or upgrade this package with `npm i -g @injaan.dev/ms365-email-cli`
- `Missing MS365 credentials`:
  run `ms365-email-cli init`
- Personal mailbox returns invalid user:
  set `AUTH_MODE=delegated` and re-run (app-only mode cannot access personal Outlook users by `/users/{email}`)
- Delegated sign-in fails with `AADSTS7000218`:
  your app registration is requiring client auth; either add `MS365_CLIENT_SECRET` in `.env` or enable public client flows in Azure App Registration
- Delegated sign-in fails with `AADSTS70002` (client must be marked as mobile):
  in Azure App Registration -> Authentication, enable public client flows (mobile and desktop)
- Need to force delegated re-login:
  delete `~/.@injaan.dev/ms365-email-cli/delegated-token.json` and run a command again
- Graph auth/permission errors:
  confirm app permissions and admin consent in Azure

## License

MIT

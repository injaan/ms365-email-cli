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

It creates or updates `.env` with:

- `AUTH_MODE` (`client_credentials` or `delegated`)
- `MS365_EMAIL_CLIENT_ID`

Auth mode guidance:

- `AUTH_MODE=delegated` is for personal Microsoft accounts (Outlook/Hotmail/Live)
- `AUTH_MODE=client_credentials` is for company/work accounts

When `AUTH_MODE=client_credentials`:

- `MS365_EMAIL_TENANT_ID`
- `MS365_EMAIL_CLIENT_SECRET`
- `MS365_EMAIL_ADDRESS`

When `AUTH_MODE=delegated`:

- Sign-in happens via device-code prompt in terminal
- API calls use the signed-in mailbox (`/me`)
- Wizard auto-sets `MS365_EMAIL_TENANT_ID=consumers`
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
ms365-email-cli reply <MESSAGE_ID> -b "Thanks"
ms365-email-cli reply-all <MESSAGE_ID> -b "Thanks everyone"
```

## AI-Agent Friendly

- Predictable command surface for tooling
- Script-friendly terminal output
- Full inbox workflows via CLI (read/search/reply/send/attachments)

## Troubleshooting

- `command not found: ms365-email-cli`:
  reinstall globally with `npm i -g @injaan.dev/ms365-email-cli`
- `Missing MS365 credentials`:
  run `ms365-email-cli init`
- Personal mailbox returns invalid user:
  set `AUTH_MODE=delegated` and re-run (app-only mode cannot access personal Outlook users by `/users/{email}`)
- Delegated sign-in fails with `AADSTS7000218`:
  your app registration is requiring client auth; either add `MS365_EMAIL_CLIENT_SECRET` in `.env` or enable public client flows in Azure App Registration
- Delegated sign-in fails with `AADSTS70002` (client must be marked as mobile):
  in Azure App Registration -> Authentication, enable public client flows (mobile and desktop)
- Need to force delegated re-login:
  delete `~/.@injaan.dev/ms365-email-cli/delegated-token.json` and run a command again
- Graph auth/permission errors:
  confirm app permissions and admin consent in Azure

## License

MIT

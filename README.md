# ms365-email-cli

CLI tool to manage an MS365 mailbox via Microsoft Graph API.

This project is specially built for AI agents to read, search, triage, and act on mailbox data through simple terminal commands.

## Built For AI Agents

- Predictable CLI command patterns for automation
- Structured, script-friendly output for parsing
- End-to-end email workflows (read, search, reply, send, attachments)
- Easy local setup with environment-based credentials

## Features

- Configure credentials with a guided wizard (`init`)
- List unread or recent emails
- Search emails by query, sender, subject, date, and folder
- Read email body and thread messages
- List or download attachments
- Mark messages as read
- Send, reply, and reply-all with optional attachments

## Requirements

- Node.js (compatible with your installed dependencies)
- An Azure AD app with Microsoft Graph permissions:
  - `Mail.ReadWrite`
  - `Mail.Send`
- Admin consent granted for the app

## Installation

```bash
npm install
```

Global install from npm:

```bash
npm i -g @injaan.dev/ms365-email-cli
```

After install, run:

```bash
ms365-email-cli --help
```

For global CLI usage while developing:

```bash
npm link
```

## Configuration

Run:

```bash
ms365-email-cli init
```

This creates/updates a local `.env` file with:

- `MS365_EMAIL_CLIENT_ID`
- `MS365_EMAIL_TENANT_ID`
- `MS365_EMAIL_CLIENT_SECRET`
- `MS365_FROM_EMAIL`

## Usage

```bash
ms365-email-cli --help
```

Examples:

```bash
ms365-email-cli init
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

## Development

Run the CLI directly:

```bash
npm start
```

## License

MIT

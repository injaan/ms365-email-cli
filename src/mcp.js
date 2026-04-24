#!/usr/bin/env node

const path = require("path");
const { spawnSync } = require("child_process");
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const {
  StdioServerTransport,
} = require("@modelcontextprotocol/sdk/server/stdio.js");
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require("@modelcontextprotocol/sdk/types.js");
const { version } = require("../package.json");

const CLI_ENTRY = path.resolve(__dirname, "..", "index.js");
const MAX_OUTPUT_BYTES = 10 * 1024 * 1024;

const server = new Server(
  { name: "ms365-email-cli", version },
  {
    capabilities: { tools: {} },
  },
);

function bodySourceProperties(label) {
  return {
    body: {
      type: "string",
      description: `${label} body content (plain text or HTML)`,
    },
    body_file: {
      type: "string",
      description: `Path to a UTF-8 file containing the ${label.toLowerCase()} body`,
    },
    html: {
      type: "boolean",
      description: "Send body as HTML (default: false)",
    },
  };
}

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: "list_emails",
      description:
        "List recent emails (all statuses, newest first). Output includes ID, From, Subject, Received, Status for each email.",
      inputSchema: {
        type: "object",
        properties: {
          count: {
            type: "number",
            description: "Number of emails to list (default 10)",
          },
        },
      },
    },
    {
      name: "list_unread_emails",
      description:
        "List unread emails (newest first). Output includes ID, From, Subject, Received. [+attachments] shown for file attachments.",
      inputSchema: {
        type: "object",
        properties: {
          count: {
            type: "number",
            description: "Number of unread emails to list (default 10)",
          },
        },
      },
    },
    {
      name: "read_email",
      description:
        "Read full email content by message ID. Shows: From, To, CC, Subject, Date, Body, Attachments.",
      inputSchema: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description:
              "Message ID (get from list_emails or list_unread_emails)",
          },
        },
        required: ["id"],
      },
    },
    {
      name: "thread",
      description:
        "Show full conversation thread for an email. Fetches all messages in the same conversation, sorted oldest first.",
      inputSchema: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description: "Message ID of any email in the thread",
          },
        },
        required: ["id"],
      },
    },
    {
      name: "mark_read",
      description: "Mark an email as read by message ID.",
      inputSchema: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description:
              "Message ID to mark as read (get from list_emails or list_unread_emails)",
          },
        },
        required: ["id"],
      },
    },
    {
      name: "search_emails",
      description:
        "Search emails by various criteria. Use query for full-text search or combine from/subject/since filters.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description: "Full-text search across subject, body, sender",
          },
          from: {
            type: "string",
            description: "Filter by sender email address",
          },
          subject: { type: "string", description: "Filter by subject text" },
          since: {
            type: "string",
            description: "Filter emails since date (YYYY-MM-DD)",
          },
          folder: {
            type: "string",
            enum: ["inbox", "sent"],
            description: "Folder to search: inbox or sent (default: inbox)",
          },
          count: {
            type: "number",
            description: "Max results to return (default: 20)",
          },
        },
      },
    },
    {
      name: "send_email",
      description:
        "Send an email via MS365 mailbox. Provide body or body_file. Use body_file for large or multiline HTML.",
      inputSchema: {
        type: "object",
        properties: {
          to: { type: "string", description: "Recipient email address" },
          cc: {
            anyOf: [
              { type: "string" },
              {
                type: "array",
                items: { type: "string" },
              },
            ],
            description:
              "CC recipient email address(es); supports arrays and comma-separated entries",
          },
          subject: { type: "string", description: "Email subject" },
          ...bodySourceProperties("Email"),
          attachments: {
            type: "array",
            items: { type: "string" },
            description: "List of file paths to attach",
          },
        },
        required: ["to", "subject"],
      },
    },
    {
      name: "reply",
      description:
        "Reply to the sender of an email. Provide body or body_file.",
      inputSchema: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description: "Message ID of the email to reply to",
          },
          ...bodySourceProperties("Reply"),
          attachments: {
            type: "array",
            items: { type: "string" },
            description: "List of file paths to attach",
          },
        },
        required: ["id"],
      },
    },
    {
      name: "reply_all",
      description:
        "Reply to all recipients of an email. Provide body or body_file.",
      inputSchema: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description: "Message ID of the email to reply to",
          },
          ...bodySourceProperties("Reply"),
          attachments: {
            type: "array",
            items: { type: "string" },
            description: "List of file paths to attach",
          },
        },
        required: ["id"],
      },
    },
    {
      name: "attachment",
      description:
        "List or download attachments from an email. Lists attachments if no output directory is set; downloads to directory if output is set.",
      inputSchema: {
        type: "object",
        properties: {
          id: { type: "string", description: "Message ID of the email" },
          output_dir: {
            type: "string",
            description:
              "Directory path to download attachments to (omit to just list attachments)",
          },
        },
        required: ["id"],
      },
    },
  ],
}));

function normalizeEmailList(value) {
  if (!value) return [];

  const items = Array.isArray(value) ? value : [value];

  return items
    .flatMap((item) => String(item).split(","))
    .map((email) => email.trim())
    .filter(Boolean);
}

function addCount(cmdArgs, count) {
  if (count === undefined || count === null || count === "") return;
  cmdArgs.push("-n", String(count));
}

function addAttachments(cmdArgs, attachments) {
  if (!Array.isArray(attachments)) return;
  for (const file of attachments) {
    cmdArgs.push("-a", String(file));
  }
}

function addBodySource(cmdArgs, args, label) {
  const hasBody = typeof args.body === "string";
  const hasBodyFile = typeof args.body_file === "string";

  if (!hasBody && !hasBodyFile) {
    throw new Error(`${label} requires either body or body_file.`);
  }
  if (hasBody && hasBodyFile) {
    throw new Error(`${label} accepts only one of body or body_file.`);
  }

  if (hasBodyFile) {
    cmdArgs.push("--body-file", args.body_file);
  } else {
    cmdArgs.push("-b", args.body);
  }

  if (args.html) {
    cmdArgs.push("--html");
  }
}

function runCli(cmdArgs) {
  const result = spawnSync(process.execPath, [CLI_ENTRY, ...cmdArgs], {
    cwd: process.cwd(),
    encoding: "utf-8",
    env: process.env,
    maxBuffer: MAX_OUTPUT_BYTES,
  });

  if (result.error) {
    throw result.error;
  }

  const output = `${result.stdout || ""}${result.stderr || ""}`;

  if (result.status !== 0) {
    throw new Error(output.trim() || `CLI exited with code ${result.status}`);
  }

  return output || "Done.";
}

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;
  const cmdArgs = [];

  try {
    if (name === "list_emails") {
      cmdArgs.push("list");
      addCount(cmdArgs, args.count);
    } else if (name === "list_unread_emails") {
      cmdArgs.push("unread");
      addCount(cmdArgs, args.count);
    } else if (name === "read_email") {
      cmdArgs.push("read", args.id);
    } else if (name === "thread") {
      cmdArgs.push("thread", args.id);
    } else if (name === "mark_read") {
      cmdArgs.push("mark-read", args.id);
    } else if (name === "search_emails") {
      cmdArgs.push("search");
      if (args.query) cmdArgs.push("-q", args.query);
      if (args.from) cmdArgs.push("--from", args.from);
      if (args.subject) cmdArgs.push("--subject", args.subject);
      if (args.since) cmdArgs.push("--since", args.since);
      if (args.folder) cmdArgs.push("--folder", args.folder);
      addCount(cmdArgs, args.count);
    } else if (name === "send_email") {
      cmdArgs.push("send", "-t", args.to, "-s", args.subject);
      addBodySource(cmdArgs, args, "send_email");
      for (const cc of normalizeEmailList(args.cc)) {
        cmdArgs.push("-c", cc);
      }
      addAttachments(cmdArgs, args.attachments);
    } else if (name === "reply") {
      cmdArgs.push("reply", args.id);
      addBodySource(cmdArgs, args, "reply");
      addAttachments(cmdArgs, args.attachments);
    } else if (name === "reply_all") {
      cmdArgs.push("reply-all", args.id);
      addBodySource(cmdArgs, args, "reply_all");
      addAttachments(cmdArgs, args.attachments);
    } else if (name === "attachment") {
      cmdArgs.push("attachment", args.id);
      if (args.output_dir) cmdArgs.push("-o", args.output_dir);
    } else {
      throw new Error(`Unknown tool: ${name}`);
    }

    return { content: [{ type: "text", text: runCli(cmdArgs) }] };
  } catch (err) {
    return {
      content: [{ type: "text", text: `Error: ${err.message}` }],
      isError: true,
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

#!/usr/bin/env node

const { Command } = require("commander");
const fs = require("fs");
const path = require("path");
const { version: packageVersion } = require("../package.json");
const { getAccessToken } = require("./auth");
const {
  getUnreadEmails,
  getAllEmails,
  searchEmails,
  getEmail,
  getThreadMessages,
  getEmailAttachments,
  getAttachmentContent,
  markAsRead,
  sendEmail,
  replyEmail,
} = require("./graph");
const { runWizard, checkConfig } = require("./config");

async function ensureConfig() {
  const missing = checkConfig();
  if (missing.length > 0) {
    console.error(`Missing config: ${missing.map((v) => v.label).join(", ")}`);
    console.error("Run: ms365-email-cli init");
    process.exit(1);
  }
}

function buildConfigHelpIndicator() {
  const missing = checkConfig();
  const initRequired = missing.length > 0;

  if (!initRequired) {
    return [
      "Configuration status:",
      "  config required: no (init already done)",
      "",
    ].join("\n");
  }

  return [
    "Configuration status:",
    "  config required: yes (run: ms365-email-cli init)",
    `  missing: ${missing.map((v) => v.label).join(", ")}`,
    "",
  ].join("\n");
}

function sanitizeTerminalOutput(value) {
  return String(value ?? "").replace(
    /[\u0000-\u0008\u000b-\u001f\u007f-\u009f\u001b]/g,
    "",
  );
}

function decodeHtmlEntities(value) {
  const namedEntities = {
    amp: "&",
    lt: "<",
    gt: ">",
    quot: '"',
    apos: "'",
    nbsp: " ",
  };

  return value.replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (match, entity) => {
    const normalized = entity.toLowerCase();
    if (namedEntities[normalized]) {
      return namedEntities[normalized];
    }
    if (normalized.startsWith("#x")) {
      const codePoint = Number.parseInt(normalized.slice(2), 16);
      return Number.isNaN(codePoint) ? match : String.fromCodePoint(codePoint);
    }
    if (normalized.startsWith("#")) {
      const codePoint = Number.parseInt(normalized.slice(1), 10);
      return Number.isNaN(codePoint) ? match : String.fromCodePoint(codePoint);
    }
    return match;
  });
}

function htmlToReadableText(html) {
  return decodeHtmlEntities(String(html ?? ""))
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|section|article|header|footer|li|tr|h[1-6])>/gi, "\n")
    .replace(/<li\b[^>]*>/gi, "- ")
    .replace(/<\/td>/gi, "\t")
    .replace(/<\/th>/gi, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/\r/g, "")
    .replace(/\t+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trimEnd())
    .join("\n")
    .trim();
}

function safeDisplay(value, fallback = "Unknown") {
  const sanitized = sanitizeTerminalOutput(value).trim();
  return sanitized || fallback;
}

function formatMessageBody(body) {
  if (!body?.content) {
    return "(empty body)";
  }

  const contentType = String(body.contentType || "").toUpperCase();
  const rawContent = String(body.content);
  const formatted =
    contentType === "HTML" ? htmlToReadableText(rawContent) : rawContent;
  const sanitized = sanitizeTerminalOutput(formatted).trim();
  return sanitized || "(empty body)";
}

function sanitizeAttachmentName(name) {
  const baseName = path
    .basename(String(name || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_"))
    .trim();
  return baseName || "attachment";
}

function resolveAttachmentOutputPath(outDir, attachmentName) {
  const safeName = sanitizeAttachmentName(attachmentName);
  const destination = path.resolve(outDir, safeName);
  const relativePath = path.relative(outDir, destination);
  if (relativePath.startsWith("..") || path.isAbsolute(relativePath)) {
    throw new Error(`Unsafe attachment name: ${attachmentName}`);
  }
  return destination;
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function collectEmailList(value, previous = []) {
  const parsed = String(value ?? "")
    .split(",")
    .map((entry) => sanitizeTerminalOutput(entry).trim())
    .filter(Boolean);
  return previous.concat(parsed);
}

function resolveInputFilePath(filePath) {
  const rawPath = String(filePath || "").trim();
  if (!rawPath) {
    throw new Error("File path is required.");
  }
  if (rawPath === "~") {
    return process.env.HOME || process.env.USERPROFILE || rawPath;
  }
  if (rawPath.startsWith(`~${path.sep}`) || rawPath.startsWith("~/")) {
    const home = process.env.HOME || process.env.USERPROFILE;
    if (home) {
      return path.resolve(home, rawPath.slice(2));
    }
  }
  return path.resolve(rawPath);
}

function readUtf8File(filePath, label) {
  const resolvedPath = resolveInputFilePath(filePath);
  const stat = fs.existsSync(resolvedPath) ? fs.statSync(resolvedPath) : null;
  if (!stat || !stat.isFile()) {
    throw new Error(`${label} not found: ${filePath}`);
  }
  return fs.readFileSync(resolvedPath, "utf-8");
}

function readStdin() {
  if (process.stdin.isTTY) {
    throw new Error("No stdin content available for --body-stdin.");
  }
  return new Promise((resolve, reject) => {
    let content = "";
    process.stdin.setEncoding("utf-8");
    process.stdin.on("data", (chunk) => {
      content += chunk;
    });
    process.stdin.on("end", () => resolve(content));
    process.stdin.on("error", reject);
  });
}

async function resolveMessageBody(opts) {
  const sources = [opts.body, opts.bodyFile, opts.bodyStdin].filter(
    (value) => value !== undefined && value !== false,
  );
  if (sources.length === 0) {
    throw new Error(
      "Email body is required. Use --body, --body-file, or --body-stdin.",
    );
  }
  if (sources.length > 1) {
    throw new Error(
      "Use only one body source: --body, --body-file, or --body-stdin.",
    );
  }
  if (opts.bodyFile) {
    return readUtf8File(opts.bodyFile, "Body file");
  }
  if (opts.bodyStdin) {
    return readStdin();
  }
  return opts.body;
}

function normalizeAttachmentPaths(paths) {
  return paths.map(resolveInputFilePath);
}

const program = new Command();

program
  .name("ms365-email-cli")
  .description(
    "CLI tool to manage MS365 mailbox via Microsoft Graph API, specially for AI agents",
  )
  .version(packageVersion)
  .addHelpText(
    "after",
    `
${buildConfigHelpIndicator()}
Output format (JSON-parseable line prefixes):
  [N]       = email index in list
  ID:       = Graph API message ID (use with mark-read)
  From:     = sender address
  Subject:  = email subject
  Received: = ISO 8601 timestamp
  Status:   = read | unread

Exit codes:
  0  success
  1  error (credentials, API failure, file not found)

Examples:
  $ ms365-email-cli init                  # configure credentials via wizard
  $ ms365-email-cli unread                # list 10 newest unread emails
  $ ms365-email-cli unread -n 5             # list 5 newest unread emails
  $ ms365-email-cli list                    # list 10 newest emails (all)
  $ ms365-email-cli list -n 20              # list 20 newest emails
  $ ms365-email-cli mark-read <ID>          # mark single email as read
  $ ms365-email-cli read <ID>               # read full email body
  $ ms365-email-cli thread <ID>             # show full conversation thread
  $ ms365-email-cli attachment <ID>           # list attachments
  $ ms365-email-cli attachment <ID> -o ./out  # download attachments
  $ ms365-email-cli search -q "invoice"      # full-text search
  $ ms365-email-cli search --from a@b.com    # search by sender
  $ ms365-email-cli reply <ID> -b "Thanks"  # reply to sender
  $ ms365-email-cli reply-all <ID> -b "Thanks"  # reply to all
  $ ms365-email-cli send -t a@b.com -s Hi -b "hello"
  $ ms365-email-cli send -t a@b.com -c manager@b.com -s Hi -b "hello"
  $ ms365-email-cli send -t a@b.com -s Hi -b "<h1>Hi</h1>" --html
  $ ms365-email-cli send -t a@b.com -s Hi --body-file email.html --html
  $ ms365-email-cli send -t a@b.com -s Hi --body-stdin --html < email.html
  $ ms365-email-cli reply <ID> --body-file reply.html --html
  $ ms365-email-cli send -t a@b.com -s Hi -b "see attached" -a file.pdf -a img.png

HTML body tip:
  Use --body-file or --body-stdin for multiline HTML or HTML containing shell-sensitive characters.
`,
  );

program
  .command("init")
  .description("Run config wizard to set up .env credentials")
  .action(async () => {
    await runWizard();
  });

program
  .command("unread")
  .description(
    "List unread emails (newest first)\n" +
      "  Output: ID, From, Subject, Received for each email\n" +
      "  [+attachments] shown for file attachments (inline images only visible via 'read')\n" +
      "  Use the ID value with 'mark-read' to mark as read",
  )
  .option("-n, --number <count>", "number of emails to fetch", "10")
  .addHelpText("after", "\nExample:\n  ms365-email-cli unread -n 5\n")
  .action(async (opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      const result = await getUnreadEmails(token, parseInt(opts.number, 10));
      const messages = result.value || [];

      if (messages.length === 0) {
        console.log("No unread emails.");
        return;
      }

      messages.forEach((msg, i) => {
        console.log(`\n[${i + 1}] ID: ${msg.id}`);
        console.log(
          `    From: ${safeDisplay(msg.from?.emailAddress?.address)}`,
        );
        console.log(
          `    Subject: ${safeDisplay(msg.subject, "(no subject)")}${msg.hasAttachments ? " [+attachments]" : ""}`,
        );
        console.log(
          `    Received: ${safeDisplay(msg.receivedDateTime, "(unknown)")}`,
        );
      });
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("list")
  .description(
    "List recent emails, all statuses (newest first)\n" +
      "  Output: ID, From, Subject, Received, Status (read/unread) for each email",
  )
  .option("-n, --number <count>", "number of emails to fetch", "10")
  .addHelpText("after", "\nExample:\n  ms365-email-cli list -n 20\n")
  .action(async (opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      const result = await getAllEmails(token, parseInt(opts.number, 10));
      const messages = result.value || [];

      if (messages.length === 0) {
        console.log("No emails found.");
        return;
      }

      messages.forEach((msg, i) => {
        const readFlag = msg.isRead ? "read" : "unread";
        console.log(`\n[${i + 1}] ID: ${msg.id}`);
        console.log(
          `    From: ${safeDisplay(msg.from?.emailAddress?.address)}`,
        );
        console.log(
          `    Subject: ${safeDisplay(msg.subject, "(no subject)")}${msg.hasAttachments ? " [+attachments]" : ""}`,
        );
        console.log(
          `    Received: ${safeDisplay(msg.receivedDateTime, "(unknown)")}`,
        );
        console.log(`    Status: ${readFlag}`);
      });
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("search")
  .description(
    "Search emails by various criteria\n" +
      "  Use --query for full-text search or combine --from/--subject/--since filters",
  )
  .option("-q, --query <text>", "full-text search across subject, body, sender")
  .option("-f, --from <email>", "search by sender address")
  .option("-s, --subject <text>", "search by subject text")
  .option("--since <date>", "filter emails since date (YYYY-MM-DD)")
  .option("--folder <inbox|sent>", "search folder: inbox or sent", "inbox")
  .option("-n, --number <count>", "max results", "20")
  .addHelpText(
    "after",
    "\nExamples:\n" +
      '  ms365-email-cli search -q "invoice"              # full-text search\n' +
      "  ms365-email-cli search --from user@example.com    # by sender\n" +
      '  ms365-email-cli search --subject "hello"          # by subject\n' +
      "  ms365-email-cli search --since 2026-03-01         # since date\n" +
      '  ms365-email-cli search --folder sent -q "report"  # search sent items\n' +
      "  ms365-email-cli search --from a@b.com --since 2026-03-01  # combined\n",
  )
  .action(async (opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();

      if (!opts.query && !opts.from && !opts.subject && !opts.since) {
        console.error(
          "At least one search parameter is required. Use -q, --from, --subject, or --since",
        );
        process.exit(1);
      }

      const result = await searchEmails(token, {
        folder: opts.folder,
        from: opts.from,
        subject: opts.subject,
        query: opts.query,
        since: opts.since,
        top: parseInt(opts.number, 10),
      });

      const messages = result.value || [];

      if (messages.length === 0) {
        console.log("No emails found.");
        return;
      }

      console.log(`Found ${messages.length} email(s):\n`);
      messages.forEach((msg, i) => {
        const readFlag = msg.isRead ? "read" : "unread";
        console.log(`[${i + 1}] ID: ${msg.id}`);
        console.log(
          `    From: ${safeDisplay(msg.from?.emailAddress?.address)}`,
        );
        if (msg.toRecipients?.length) {
          console.log(
            `    To: ${sanitizeTerminalOutput(msg.toRecipients.map((r) => safeDisplay(r.emailAddress?.address)).join(", "))}`,
          );
        }
        console.log(
          `    Subject: ${safeDisplay(msg.subject, "(no subject)")}${msg.hasAttachments ? " [+attachments]" : ""}`,
        );
        console.log(
          `    Received: ${safeDisplay(msg.receivedDateTime, "(unknown)")}`,
        );
        console.log(`    Status: ${readFlag}`);
        console.log();
      });
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("mark-read <messageId>")
  .description(
    "Mark an email as read by message ID\n" +
      "  Get the ID from 'unread' or 'list' command output",
  )
  .addHelpText(
    "after",
    "\nExample:\n  ms365-email-cli mark-read AAMkAGI2TG93AAA=\n",
  )
  .action(async (messageId) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      await markAsRead(token, messageId);
      console.log(`Email ${messageId} marked as read.`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("read <messageId>")
  .description(
    "Read full email content by message ID\n" +
      "  Shows: From, To, CC, Subject, Date, Body, Attachments",
  )
  .addHelpText("after", "\nExample:\n  ms365-email-cli read AAMkAGI2TG93AAA=\n")
  .action(async (messageId) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      const msg = await getEmail(token, messageId);

      console.log(
        `From: ${safeDisplay(msg.from?.emailAddress?.name, "")} <${safeDisplay(msg.from?.emailAddress?.address)}>`,
      );
      if (msg.toRecipients?.length) {
        console.log(
          `To: ${sanitizeTerminalOutput(msg.toRecipients.map((r) => `${safeDisplay(r.emailAddress?.name, "")} <${safeDisplay(r.emailAddress?.address)}>`).join(", "))}`,
        );
      }
      if (msg.ccRecipients?.length) {
        console.log(
          `CC: ${sanitizeTerminalOutput(msg.ccRecipients.map((r) => `${safeDisplay(r.emailAddress?.name, "")} <${safeDisplay(r.emailAddress?.address)}>`).join(", "))}`,
        );
      }
      console.log(`Subject: ${safeDisplay(msg.subject, "(no subject)")}`);
      console.log(
        `Date: ${safeDisplay(msg.sentDateTime || msg.receivedDateTime, "(unknown)")}`,
      );

      const attResult = await getEmailAttachments(token, messageId);
      const atts = attResult.value || [];
      const fileAtts = atts.filter((a) => !a.isInline);
      const inlineAtts = atts.filter((a) => a.isInline);

      if (atts.length > 0) {
        console.log(
          `Attachments (${fileAtts.length} files, ${inlineAtts.length} inline):`,
        );
        atts.forEach((a, i) => {
          const tag = a.isInline ? "inline" : "file";
          console.log(
            `  [${i + 1}] ${safeDisplay(a.name, "attachment")} (${formatSize(a.size)}) [${tag}]`,
          );
        });
        if (fileAtts.length > 0) {
          console.log(
            `  Download: ms365-email-cli attachment ${messageId} -o <dir>`,
          );
        }
      }

      console.log(`\n--- Body (${msg.body?.contentType || "unknown"}) ---\n`);
      console.log(formatMessageBody(msg.body));
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("thread <messageId>")
  .description(
    "Show full conversation thread for an email\n" +
      "  Fetches all messages in the same conversation, sorted oldest first",
  )
  .addHelpText(
    "after",
    "\nExample:\n  ms365-email-cli thread AAMkAGI2TG93AAA=\n",
  )
  .action(async (messageId) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();

      const rootMsg = await getEmail(token, messageId);
      const conversationId = rootMsg.conversationId;
      if (!conversationId) {
        console.error("Email has no conversationId.");
        process.exit(1);
      }

      const result = await getThreadMessages(token, conversationId);
      const messages = result.value || [];

      if (messages.length === 0) {
        console.log("No messages found in thread.");
        return;
      }

      console.log(
        `Thread: ${safeDisplay(rootMsg.subject, "(no subject)")} (${messages.length} messages)\n`,
      );

      messages.forEach((msg, i) => {
        const from = safeDisplay(msg.from?.emailAddress?.address);
        const date = safeDisplay(msg.receivedDateTime, "(unknown)");
        const bodyPreview = sanitizeTerminalOutput(
          (msg.body?.content || "")
            .replace(/<[^>]*>/g, "")
            .trim()
            .slice(0, 200),
        );

        console.log(`--- [${i + 1}] ${from} | ${date} ---`);
        console.log(`    ID: ${msg.id}`);
        console.log(`    Subject: ${safeDisplay(msg.subject, "(no subject)")}`);
        console.log(
          `    Preview: ${bodyPreview}${bodyPreview.length >= 200 ? "..." : ""}`,
        );
        console.log();
      });
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("attachment <messageId>")
  .description(
    "Download attachments from an email\n" +
      "  Lists attachments if no -o flag, saves to directory if -o is set",
  )
  .option(
    "-o, --output <dir>",
    "output directory (default: current directory)",
    ".",
  )
  .addHelpText(
    "after",
    "\nExamples:\n" +
      "  ms365-email-cli attachment <ID>                  # list attachments\n" +
      "  ms365-email-cli attachment <ID> -o ./downloads   # download all to dir\n",
  )
  .action(async (messageId, opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      const attResult = await getEmailAttachments(token, messageId);
      const atts = attResult.value || [];

      if (atts.length === 0) {
        console.log("No attachments found.");
        return;
      }

      if (
        opts.output === "." &&
        process.argv.indexOf("-o") === -1 &&
        process.argv.indexOf("--output") === -1
      ) {
        console.log(`Attachments (${atts.length}):\n`);
        atts.forEach((a, i) => {
          const tag = a.isInline ? "inline" : "file";
          console.log(
            `  [${i + 1}] ${safeDisplay(a.name, "attachment")} (${safeDisplay(a.contentType, "unknown")}, ${formatSize(a.size)}) [${tag}]`,
          );
          console.log(`      ID: ${a.id}`);
        });
        console.log(
          `\nTo download: ms365-email-cli attachment ${messageId} -o <dir>`,
        );
        return;
      }

      const outDir = path.resolve(opts.output);
      if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

      let savedCount = 0;
      for (const att of atts) {
        const full = await getAttachmentContent(token, messageId, att.id);
        if (!full.contentBytes) {
          console.log(
            `Skipped: ${safeDisplay(att.name, "attachment")} (attachment content is not a downloadable file)`,
          );
          continue;
        }
        const buf = Buffer.from(full.contentBytes, "base64");
        const outPath = resolveAttachmentOutputPath(outDir, att.name);
        fs.writeFileSync(outPath, buf);
        savedCount += 1;
        console.log(`Saved: ${outPath} (${formatSize(buf.length)})`);
      }

      console.log(`\nDone. ${savedCount} file(s) downloaded to ${outDir}`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("send")
  .description(
    "Send an email via MS365 mailbox\n" +
      "  Requires -t, -s, and one body source. Optional --cc, --html and -a for attachments",
  )
  .requiredOption("-t, --to <email>", "recipient email address")
  .option(
    "-c, --cc <email>",
    "CC recipient email address (repeatable, supports comma-separated values)",
    collectEmailList,
    [],
  )
  .requiredOption("-s, --subject <subject>", "email subject")
  .option("-b, --body <body>", "email body (plain text or HTML)")
  .option("--body-file <file>", "read email body from a UTF-8 file")
  .option("--body-stdin", "read email body from standard input", false)
  .option("--html", "send body as HTML", false)
  .option(
    "-a, --attachment <file>",
    "attach file (repeatable)",
    (v, a) => a.concat(v),
    [],
  )
  .addHelpText(
    "after",
    "\nExamples:\n" +
      '  ms365-email-cli send -t user@example.com -s Hello -b "plain text body"\n' +
      '  ms365-email-cli send -t user@example.com -c manager@example.com -s Hello -b "cc included"\n' +
      '  ms365-email-cli send -t user@example.com -c a@example.com,b@example.com -s Hello -b "multi cc"\n' +
      '  ms365-email-cli send -t user@example.com -s Hello -b "<b>bold</b>" --html\n' +
      "  ms365-email-cli send -t user@example.com -s Hello --body-file email.html --html\n" +
      "  ms365-email-cli send -t user@example.com -s Hello --body-stdin --html < email.html\n" +
      '  ms365-email-cli send -t user@example.com -s Report -b "see attached" -a report.pdf\n' +
      '  ms365-email-cli send -t user@example.com -s Files -b "attached" -a a.pdf -a b.xlsx\n',
  )
  .action(async (opts) => {
    try {
      const body = await resolveMessageBody(opts);
      const attachments = normalizeAttachmentPaths(opts.attachment);
      await ensureConfig();
      const token = await getAccessToken();
      await sendEmail(
        token,
        opts.to,
        opts.subject,
        body,
        opts.html,
        attachments,
        opts.cc,
      );
      console.log(`Email sent to ${opts.to}`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("reply <messageId>")
  .description(
    "Reply to sender of an email\n" +
      "  Sends reply only to the original sender",
  )
  .option("-b, --body <body>", "reply body (plain text or HTML)")
  .option("--body-file <file>", "read reply body from a UTF-8 file")
  .option("--body-stdin", "read reply body from standard input", false)
  .option("--html", "send body as HTML", false)
  .option(
    "-a, --attachment <file>",
    "attach file (repeatable)",
    (v, a) => a.concat(v),
    [],
  )
  .addHelpText(
    "after",
    "\nExamples:\n" +
      '  ms365-email-cli reply <ID> -b "Thanks for your email"\n' +
      '  ms365-email-cli reply <ID> -b "<p>Thanks</p>" --html\n' +
      "  ms365-email-cli reply <ID> --body-file reply.html --html\n" +
      "  ms365-email-cli reply <ID> --body-stdin --html < reply.html\n" +
      '  ms365-email-cli reply <ID> -b "See attached" -a report.pdf\n',
  )
  .action(async (messageId, opts) => {
    try {
      const body = await resolveMessageBody(opts);
      const attachments = normalizeAttachmentPaths(opts.attachment);
      await ensureConfig();
      const token = await getAccessToken();
      await replyEmail(
        token,
        messageId,
        body,
        opts.html,
        false,
        attachments,
      );
      console.log(`Reply sent to message ${messageId}`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("reply-all <messageId>")
  .description(
    "Reply to all recipients of an email\n" +
      "  Sends reply to original sender, To, and CC recipients",
  )
  .option("-b, --body <body>", "reply body (plain text or HTML)")
  .option("--body-file <file>", "read reply body from a UTF-8 file")
  .option("--body-stdin", "read reply body from standard input", false)
  .option("--html", "send body as HTML", false)
  .option(
    "-a, --attachment <file>",
    "attach file (repeatable)",
    (v, a) => a.concat(v),
    [],
  )
  .addHelpText(
    "after",
    "\nExamples:\n" +
      '  ms365-email-cli reply-all <ID> -b "Thanks everyone"\n' +
      '  ms365-email-cli reply-all <ID> -b "<p>Thanks</p>" --html\n' +
      "  ms365-email-cli reply-all <ID> --body-file reply.html --html\n" +
      "  ms365-email-cli reply-all <ID> --body-stdin --html < reply.html\n" +
      '  ms365-email-cli reply-all <ID> -b "See attached" -a report.pdf\n',
  )
  .action(async (messageId, opts) => {
    try {
      const body = await resolveMessageBody(opts);
      const attachments = normalizeAttachmentPaths(opts.attachment);
      await ensureConfig();
      const token = await getAccessToken();
      await replyEmail(
        token,
        messageId,
        body,
        opts.html,
        true,
        attachments,
      );
      console.log(`Reply-all sent to message ${messageId}`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program.parse(process.argv.map((arg) => (arg === "-help" ? "--help" : arg)));

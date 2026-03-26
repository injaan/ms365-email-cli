#!/usr/bin/env node

const { Command } = require("commander");
const fs = require("fs");
const path = require("path");
const { getAccessToken } = require("./auth");
const { getUnreadEmails, getAllEmails, searchEmails, getEmail, getThreadMessages, getEmailAttachments, getAttachmentContent, markAsRead, sendEmail, replyEmail } = require("./graph");
const { runWizard, checkConfig } = require("./config");

async function ensureConfig() {
  const missing = checkConfig();
  if (missing.length > 0) {
    console.error(`Missing config: ${missing.map((v) => v.label).join(", ")}`);
    console.error('Run: ms365-email-cli init');
    process.exit(1);
  }
}

function sanitizeTerminalOutput(value) {
  return String(value ?? "").replace(/[\u0000-\u0008\u000b-\u001f\u007f-\u009f\u001b]/g, "");
}

function safeDisplay(value, fallback = "Unknown") {
  const sanitized = sanitizeTerminalOutput(value).trim();
  return sanitized || fallback;
}

function sanitizeAttachmentName(name) {
  const baseName = path.basename(String(name || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")).trim();
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

const program = new Command();

program
  .name("ms365-email-cli")
  .description(
    "CLI tool to manage MS365 mailbox via Microsoft Graph API\n" +
    "  Requires .env with MS365_EMAIL_CLIENT_ID, MS365_EMAIL_TENANT_ID, MS365_EMAIL_CLIENT_SECRET\n" +
    "  Azure app needs Mail.ReadWrite + Mail.Send permissions with admin consent"
  )
  .version("1.0.0")
  .addHelpText(
    "after",
    `
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
  $ ms365-email-cli send -t a@b.com -s Hi -b "<h1>Hi</h1>" --html
  $ ms365-email-cli send -t a@b.com -s Hi -b "see attached" -a file.pdf -a img.png
`
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
    "  Use the ID value with 'mark-read' to mark as read"
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
        console.log(`    From: ${safeDisplay(msg.from?.emailAddress?.address)}`);
        console.log(`    Subject: ${safeDisplay(msg.subject, "(no subject)")}${msg.hasAttachments ? " [+attachments]" : ""}`);
        console.log(`    Received: ${safeDisplay(msg.receivedDateTime, "(unknown)")}`);
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
    "  Output: ID, From, Subject, Received, Status (read/unread) for each email"
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
        console.log(`    From: ${safeDisplay(msg.from?.emailAddress?.address)}`);
        console.log(`    Subject: ${safeDisplay(msg.subject, "(no subject)")}${msg.hasAttachments ? " [+attachments]" : ""}`);
        console.log(`    Received: ${safeDisplay(msg.receivedDateTime, "(unknown)")}`);
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
    "  Use --query for full-text search or combine --from/--subject/--since filters"
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
    "  ms365-email-cli search -q \"invoice\"              # full-text search\n" +
    "  ms365-email-cli search --from user@example.com    # by sender\n" +
    "  ms365-email-cli search --subject \"hello\"          # by subject\n" +
    "  ms365-email-cli search --since 2026-03-01         # since date\n" +
    "  ms365-email-cli search --folder sent -q \"report\"  # search sent items\n" +
    "  ms365-email-cli search --from a@b.com --since 2026-03-01  # combined\n"
  )
  .action(async (opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();

      if (!opts.query && !opts.from && !opts.subject && !opts.since) {
        console.error("At least one search parameter is required. Use -q, --from, --subject, or --since");
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
        console.log(`    From: ${safeDisplay(msg.from?.emailAddress?.address)}`);
        if (msg.toRecipients?.length) {
          console.log(`    To: ${sanitizeTerminalOutput(msg.toRecipients.map((r) => safeDisplay(r.emailAddress?.address)).join(", "))}`);
        }
        console.log(`    Subject: ${safeDisplay(msg.subject, "(no subject)")}${msg.hasAttachments ? " [+attachments]" : ""}`);
        console.log(`    Received: ${safeDisplay(msg.receivedDateTime, "(unknown)")}`);
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
    "  Get the ID from 'unread' or 'list' command output"
  )
  .addHelpText(
    "after",
    "\nExample:\n  ms365-email-cli mark-read AAMkAGI2TG93AAA=\n"
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
    "  Shows: From, To, CC, Subject, Date, Body, Attachments"
  )
  .addHelpText(
    "after",
    "\nExample:\n  ms365-email-cli read AAMkAGI2TG93AAA=\n"
  )
  .action(async (messageId) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      const msg = await getEmail(token, messageId);

      console.log(`From: ${safeDisplay(msg.from?.emailAddress?.name, "")} <${safeDisplay(msg.from?.emailAddress?.address)}>`);
      if (msg.toRecipients?.length) {
        console.log(`To: ${sanitizeTerminalOutput(msg.toRecipients.map((r) => `${safeDisplay(r.emailAddress?.name, "")} <${safeDisplay(r.emailAddress?.address)}>`).join(", "))}`);
      }
      if (msg.ccRecipients?.length) {
        console.log(`CC: ${sanitizeTerminalOutput(msg.ccRecipients.map((r) => `${safeDisplay(r.emailAddress?.name, "")} <${safeDisplay(r.emailAddress?.address)}>`).join(", "))}`);
      }
      console.log(`Subject: ${safeDisplay(msg.subject, "(no subject)")}`);
      console.log(`Date: ${safeDisplay(msg.sentDateTime || msg.receivedDateTime, "(unknown)")}`);

      const attResult = await getEmailAttachments(token, messageId);
      const atts = attResult.value || [];
      const fileAtts = atts.filter(a => !a.isInline);
      const inlineAtts = atts.filter(a => a.isInline);

      if (atts.length > 0) {
        console.log(`Attachments (${fileAtts.length} files, ${inlineAtts.length} inline):`);
        atts.forEach((a, i) => {
          const tag = a.isInline ? "inline" : "file";
          console.log(`  [${i + 1}] ${safeDisplay(a.name, "attachment")} (${formatSize(a.size)}) [${tag}]`);
        });
        if (fileAtts.length > 0) {
          console.log(`  Download: ms365-email-cli attachment ${messageId} -o <dir>`);
        }
      }

      console.log(`\n--- Body (${msg.body?.contentType || "unknown"}) ---\n`);
      console.log(sanitizeTerminalOutput(msg.body?.content || "(empty body)"));
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("thread <messageId>")
  .description(
    "Show full conversation thread for an email\n" +
    "  Fetches all messages in the same conversation, sorted oldest first"
  )
  .addHelpText(
    "after",
    "\nExample:\n  ms365-email-cli thread AAMkAGI2TG93AAA=\n"
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

      console.log(`Thread: ${safeDisplay(rootMsg.subject, "(no subject)")} (${messages.length} messages)\n`);

      messages.forEach((msg, i) => {
        const from = safeDisplay(msg.from?.emailAddress?.address);
        const date = safeDisplay(msg.receivedDateTime, "(unknown)");
        const bodyPreview = sanitizeTerminalOutput((msg.body?.content || "").replace(/<[^>]*>/g, "").trim().slice(0, 200));

        console.log(`--- [${i + 1}] ${from} | ${date} ---`);
        console.log(`    ID: ${msg.id}`);
        console.log(`    Subject: ${safeDisplay(msg.subject, "(no subject)")}`);
        console.log(`    Preview: ${bodyPreview}${bodyPreview.length >= 200 ? "..." : ""}`);
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
    "  Lists attachments if no -o flag, saves to directory if -o is set"
  )
  .option("-o, --output <dir>", "output directory (default: current directory)", ".")
  .addHelpText(
    "after",
    "\nExamples:\n" +
    "  ms365-email-cli attachment <ID>                  # list attachments\n" +
    "  ms365-email-cli attachment <ID> -o ./downloads   # download all to dir\n"
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

      if (opts.output === "." && process.argv.indexOf("-o") === -1 && process.argv.indexOf("--output") === -1) {
        console.log(`Attachments (${atts.length}):\n`);
        atts.forEach((a, i) => {
          const tag = a.isInline ? "inline" : "file";
          console.log(`  [${i + 1}] ${safeDisplay(a.name, "attachment")} (${safeDisplay(a.contentType, "unknown")}, ${formatSize(a.size)}) [${tag}]`);
          console.log(`      ID: ${a.id}`);
        });
        console.log(`\nTo download: ms365-email-cli attachment ${messageId} -o <dir>`);
        return;
      }

      const outDir = path.resolve(opts.output);
      if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

      for (const att of atts) {
        const full = await getAttachmentContent(token, messageId, att.id);
        const buf = Buffer.from(full.contentBytes, "base64");
        const outPath = resolveAttachmentOutputPath(outDir, att.name);
        fs.writeFileSync(outPath, buf);
        console.log(`Saved: ${outPath} (${formatSize(buf.length)})`);
      }

      console.log(`\nDone. ${atts.length} file(s) downloaded to ${outDir}`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program
  .command("send")
  .description(
    "Send an email via MS365 mailbox\n" +
    "  Requires -t, -s, -b flags. Optional --html and -a for attachments"
  )
  .requiredOption("-t, --to <email>", "recipient email address")
  .requiredOption("-s, --subject <subject>", "email subject")
  .requiredOption("-b, --body <body>", "email body (plain text or HTML)")
  .option("--html", "send body as HTML", false)
  .option("-a, --attachment <file>", "attach file (repeatable)", (v, a) => a.concat(v), [])
  .addHelpText(
    "after",
    "\nExamples:\n" +
    "  ms365-email-cli send -t user@example.com -s Hello -b \"plain text body\"\n" +
    "  ms365-email-cli send -t user@example.com -s Hello -b \"<b>bold</b>\" --html\n" +
    "  ms365-email-cli send -t user@example.com -s Report -b \"see attached\" -a report.pdf\n" +
    "  ms365-email-cli send -t user@example.com -s Files -b \"attached\" -a a.pdf -b b.xlsx\n"
  )
  .action(async (opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      await sendEmail(token, opts.to, opts.subject, opts.body, opts.html, opts.attachment);
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
    "  Sends reply only to the original sender"
  )
  .requiredOption("-b, --body <body>", "reply body (plain text or HTML)")
  .option("--html", "send body as HTML", false)
  .option("-a, --attachment <file>", "attach file (repeatable)", (v, a) => a.concat(v), [])
  .addHelpText(
    "after",
    "\nExamples:\n" +
    "  ms365-email-cli reply <ID> -b \"Thanks for your email\"\n" +
    "  ms365-email-cli reply <ID> -b \"<p>Thanks</p>\" --html\n" +
    "  ms365-email-cli reply <ID> -b \"See attached\" -a report.pdf\n"
  )
  .action(async (messageId, opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      await replyEmail(token, messageId, opts.body, opts.html, false, opts.attachment);
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
    "  Sends reply to original sender, To, and CC recipients"
  )
  .requiredOption("-b, --body <body>", "reply body (plain text or HTML)")
  .option("--html", "send body as HTML", false)
  .option("-a, --attachment <file>", "attach file (repeatable)", (v, a) => a.concat(v), [])
  .addHelpText(
    "after",
    "\nExamples:\n" +
    "  ms365-email-cli reply-all <ID> -b \"Thanks everyone\"\n" +
    "  ms365-email-cli reply-all <ID> -b \"<p>Thanks</p>\" --html\n" +
    "  ms365-email-cli reply-all <ID> -b \"See attached\" -a report.pdf\n"
  )
  .action(async (messageId, opts) => {
    try {
      await ensureConfig();
      const token = await getAccessToken();
      await replyEmail(token, messageId, opts.body, opts.html, true, opts.attachment);
      console.log(`Reply-all sent to message ${messageId}`);
    } catch (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
  });

program.parse();

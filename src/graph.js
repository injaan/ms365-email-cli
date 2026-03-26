const https = require("https");
const fs = require("fs");
const path = require("path");
const { normalizeAuthMode } = require("./auth");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const ISO_DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;

function encodeGraphPathSegment(value, label) {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${label} is required`);
  }
  return encodeURIComponent(value.trim());
}

function escapeODataString(value) {
  return value.replace(/'/g, "''");
}

function normalizeTop(top, defaultValue) {
  const parsed = Number.parseInt(top, 10);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    return defaultValue;
  }
  return Math.min(parsed, 100);
}

function normalizeSinceDate(since) {
  if (!since) return null;
  if (!ISO_DATE_PATTERN.test(since)) {
    throw new Error("Invalid --since value. Use YYYY-MM-DD.");
  }
  return `${since}T00:00:00Z`;
}

function getUserPath() {
  const authMode = normalizeAuthMode(process.env.AUTH_MODE);
  if (authMode === "delegated") {
    return "/me";
  }

  const email = process.env.MS365_EMAIL_ADDRESS;
  if (!email) {
    throw new Error(
      "MS365_EMAIL_ADDRESS is required in .env for client_credentials flow",
    );
  }
  return `/users/${encodeGraphPathSegment(email, "MS365_EMAIL_ADDRESS")}`;
}

function request(method, urlPath, token, body = null) {
  return new Promise((resolve, reject) => {
    const url = new URL(GRAPH_BASE + urlPath);

    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json",
      },
    };

    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(data ? JSON.parse(data) : {});
          } catch {
            resolve({});
          }
        } else {
          let errMsg = data;
          try {
            const parsed = JSON.parse(data);
            errMsg = parsed.error?.message || data;
          } catch {}
          reject(new Error(`Graph API ${res.statusCode}: ${errMsg}`));
        }
      });
    });

    req.on("error", (e) => reject(new Error(`Request failed: ${e.message}`)));

    if (body) {
      req.write(JSON.stringify(body));
    }
    req.end();
  });
}

async function getUnreadEmails(token, top = 10) {
  const userPath = getUserPath();
  const safeTop = normalizeTop(top, 10);
  const apiPath = `${userPath}/messages?$filter=isRead eq false&$orderby=receivedDateTime desc&$top=${safeTop}&$select=id,subject,from,receivedDateTime,isRead,hasAttachments`;
  return request("GET", apiPath, token);
}

async function getAllEmails(token, top = 10) {
  const userPath = getUserPath();
  const safeTop = normalizeTop(top, 10);
  const apiPath = `${userPath}/messages?$orderby=receivedDateTime desc&$top=${safeTop}&$select=id,subject,from,receivedDateTime,isRead,hasAttachments`;
  return request("GET", apiPath, token);
}

async function searchEmails(token, opts = {}) {
  const userPath = getUserPath();
  const { folder = "inbox", from, subject, query, since, top = 20 } = opts;

  let folderPath;
  if (folder === "sent") {
    folderPath = `${userPath}/mailFolders/sentitems/messages`;
  } else {
    folderPath = `${userPath}/mailFolders/inbox/messages`;
  }

  const params = new URLSearchParams();
  params.set("$top", String(normalizeTop(top, 20)));
  params.set(
    "$select",
    "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments",
  );

  const searchText =
    [query, from, subject].find(
      (value) => typeof value === "string" && value.trim(),
    ) || null;

  if (searchText) {
    const normalizedSearchText = searchText.replace(/"/g, "");
    params.set("$search", `"${normalizedSearchText}"`);
  } else if (since) {
    params.set("$filter", `receivedDateTime ge ${normalizeSinceDate(since)}`);
    params.set("$orderby", "receivedDateTime desc");
  }

  const apiPath = `${folderPath}?${params.toString()}`;
  return request("GET", apiPath, token);
}

async function getEmail(token, messageId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${encodeGraphPathSegment(messageId, "messageId")}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,hasAttachments,internetMessageId,conversationId`;
  return request("GET", apiPath, token);
}

async function getThreadMessages(token, conversationId) {
  const userPath = getUserPath();
  const escapedConversationId = escapeODataString(conversationId);
  const apiPath = `${userPath}/messages?$filter=conversationId eq '${escapedConversationId}'&$select=id,subject,from,toRecipients,receivedDateTime,isRead,body&$top=50`;
  const result = await request("GET", apiPath, token);
  if (result.value) {
    result.value.sort(
      (a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime),
    );
  }
  return result;
}

async function getEmailAttachments(token, messageId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${encodeGraphPathSegment(messageId, "messageId")}/attachments?$select=id,name,contentType,size,isInline`;
  return request("GET", apiPath, token);
}

async function getAttachmentContent(token, messageId, attachmentId) {
  const userPath = getUserPath();
  const safeMessageId = encodeGraphPathSegment(messageId, "messageId");
  const safeAttachmentId = encodeGraphPathSegment(attachmentId, "attachmentId");
  const apiPath = `${userPath}/messages/${safeMessageId}/attachments/${safeAttachmentId}`;
  return request("GET", apiPath, token);
}

async function markAsRead(token, messageId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${encodeGraphPathSegment(messageId, "messageId")}`;
  return request("PATCH", apiPath, token, { isRead: true });
}

async function sendEmail(
  token,
  to,
  subject,
  body,
  html = false,
  attachmentPaths = [],
) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/sendMail`;

  const attachments = attachmentPaths.map((filePath) => {
    if (!fs.existsSync(filePath)) {
      throw new Error(`Attachment not found: ${filePath}`);
    }
    return {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: path.basename(filePath),
      contentBytes: fs.readFileSync(filePath, { encoding: "base64" }),
    };
  });

  const mail = {
    message: {
      subject,
      body: {
        contentType: html ? "HTML" : "Text",
        content: body,
      },
      toRecipients: [
        {
          emailAddress: {
            address: to,
          },
        },
      ],
      ...(attachments.length > 0 && { attachments }),
    },
  };
  return request("POST", apiPath, token, mail);
}

async function replyEmail(
  token,
  messageId,
  body,
  html = false,
  replyAll = false,
  attachmentPaths = [],
) {
  const userPath = getUserPath();
  const action = replyAll ? "replyAll" : "reply";
  const apiPath = `${userPath}/messages/${encodeGraphPathSegment(messageId, "messageId")}/${action}`;

  const attachments = attachmentPaths.map((filePath) => {
    if (!fs.existsSync(filePath)) {
      throw new Error(`Attachment not found: ${filePath}`);
    }
    return {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: path.basename(filePath),
      contentBytes: fs.readFileSync(filePath, { encoding: "base64" }),
    };
  });

  const payload = {
    message: {
      body: {
        contentType: html ? "HTML" : "Text",
        content: body,
      },
      ...(attachments.length > 0 && { attachments }),
    },
  };
  return request("POST", apiPath, token, payload);
}

module.exports = {
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
};

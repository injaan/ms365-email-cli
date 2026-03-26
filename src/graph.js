const https = require("https");
const fs = require("fs");
const path = require("path");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function getUserPath() {
  const email = process.env.MS365_FROM_EMAIL;
  if (!email) {
    throw new Error("MS365_FROM_EMAIL is required in .env for client_credentials flow");
  }
  return `/users/${email}`;
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
  const apiPath = `${userPath}/messages?$filter=isRead eq false&$orderby=receivedDateTime desc&$top=${top}&$select=id,subject,from,receivedDateTime,isRead,hasAttachments`;
  return request("GET", apiPath, token);
}

async function getAllEmails(token, top = 10) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages?$orderby=receivedDateTime desc&$top=${top}&$select=id,subject,from,receivedDateTime,isRead,hasAttachments`;
  return request("GET", apiPath, token);
}

async function searchEmails(token, opts = {}) {
  const userPath = getUserPath();
  const { folder = "inbox", from, to, subject, query, since, top = 20 } = opts;

  let folderPath;
  if (folder === "sent") {
    folderPath = `${userPath}/mailFolders/sentitems/messages`;
  } else {
    folderPath = `${userPath}/mailFolders/inbox/messages`;
  }

  const params = new URLSearchParams();
  params.set("$top", top);
  params.set("$select", "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments");

  const searchText = query || from || subject || null;

  if (searchText) {
    params.set("$search", `"${searchText}"`);
  } else if (since) {
    params.set("$filter", `receivedDateTime ge ${since}`);
    params.set("$orderby", "receivedDateTime desc");
  }

  const apiPath = `${folderPath}?${params.toString()}`;
  return request("GET", apiPath, token);
}

async function getEmail(token, messageId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,hasAttachments,internetMessageId,conversationId`;
  return request("GET", apiPath, token);
}

async function getThreadMessages(token, conversationId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages?$filter=conversationId eq '${conversationId}'&$select=id,subject,from,toRecipients,receivedDateTime,isRead,body&$top=50`;
  const result = await request("GET", apiPath, token);
  if (result.value) {
    result.value.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));
  }
  return result;
}

async function getEmailAttachments(token, messageId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${messageId}/attachments?$select=id,name,contentType,size,isInline`;
  return request("GET", apiPath, token);
}

async function getAttachmentContent(token, messageId, attachmentId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${messageId}/attachments/${attachmentId}`;
  return request("GET", apiPath, token);
}

async function markAsRead(token, messageId) {
  const userPath = getUserPath();
  const apiPath = `${userPath}/messages/${messageId}`;
  return request("PATCH", apiPath, token, { isRead: true });
}

async function sendEmail(token, to, subject, body, html = false, attachmentPaths = []) {
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

async function replyEmail(token, messageId, body, html = false, replyAll = false, attachmentPaths = []) {
  const userPath = getUserPath();
  const action = replyAll ? "replyAll" : "reply";
  const apiPath = `${userPath}/messages/${messageId}/${action}`;

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

module.exports = { getUnreadEmails, getAllEmails, searchEmails, getEmail, getThreadMessages, getEmailAttachments, getAttachmentContent, markAsRead, sendEmail, replyEmail };

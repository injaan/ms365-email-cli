const fs = require("fs");
const path = require("path");
const { ENV_PATH, LOCAL_ENV } = require("./paths");

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, "utf-8").split("\n");
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx === -1) continue;
    const key = trimmed.slice(0, idx).trim();
    const val = trimmed.slice(idx + 1).trim();
    if (!process.env[key]) process.env[key] = val;
  }
}

loadEnvFile(ENV_PATH);
loadEnvFile(LOCAL_ENV);

const https = require("https");
const querystring = require("querystring");

function getCredentials() {
  const clientId = process.env.MS365_EMAIL_CLIENT_ID;
  const tenantId = process.env.MS365_EMAIL_TENANT_ID;
  const clientSecret = process.env.MS365_EMAIL_CLIENT_SECRET;

  if (!clientId || !tenantId || !clientSecret) {
    throw new Error(
      "Missing MS365 credentials. Run: ms365-email-cli init"
    );
  }

  return { clientId, tenantId, clientSecret };
}

function getAccessToken() {
  return new Promise((resolve, reject) => {
    const { clientId, tenantId, clientSecret } = getCredentials();

    const postData = querystring.stringify({
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
    });

    const options = {
      hostname: "login.microsoftonline.com",
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Content-Length": Buffer.byteLength(postData),
      },
    };

    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) {
            resolve(json.access_token);
          } else {
            reject(
              new Error(
                `Token error: ${json.error_description || json.error || "Unknown error"}`
              )
            );
          }
        } catch (e) {
          reject(new Error(`Failed to parse token response: ${e.message}`));
        }
      });
    });

    req.on("error", (e) => reject(new Error(`Token request failed: ${e.message}`)));
    req.write(postData);
    req.end();
  });
}

module.exports = { getAccessToken };

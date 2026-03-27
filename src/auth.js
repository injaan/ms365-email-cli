const fs = require("fs");
const path = require("path");
const { CONFIG_DIR, ENV_PATH, LOCAL_ENV } = require("./paths");

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

const TOKEN_CACHE_PATH = path.resolve(CONFIG_DIR, "delegated-token.json");
const TOKEN_EXPIRY_SKEW_SECONDS = 60;

function normalizeAuthMode(value) {
  const mode = String(value || "client_credentials")
    .trim()
    .toLowerCase();
  if (mode === "delegated") return "delegated";
  return "client_credentials";
}

function getCredentials() {
  const authMode = normalizeAuthMode(process.env.AUTH_MODE);
  const clientId = process.env.MS365_CLIENT_ID;
  const tenantId =
    process.env.MS365_TENANT_ID ||
    (authMode === "delegated" ? "consumers" : "");
  const clientSecret = process.env.MS365_CLIENT_SECRET;

  if (!clientId || !tenantId) {
    throw new Error("Missing MS365 credentials. Run: ms365-email-cli init");
  }

  if (authMode === "client_credentials" && !clientSecret) {
    throw new Error(
      "MS365_CLIENT_SECRET is required when AUTH_MODE=client_credentials",
    );
  }

  return { authMode, clientId, tenantId, clientSecret };
}

function postForm(hostname, path, formData) {
  return new Promise((resolve, reject) => {
    const postData = querystring.stringify(formData);

    const options = {
      hostname,
      path,
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
          resolve(JSON.parse(data));
        } catch (e) {
          reject(new Error(`Failed to parse token response: ${e.message}`));
        }
      });
    });

    req.on("error", (e) =>
      reject(new Error(`Token request failed: ${e.message}`)),
    );
    req.write(postData);
    req.end();
  });
}

function ensureConfigDir() {
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  }
  fs.chmodSync(CONFIG_DIR, 0o700);
}

function readDelegatedTokenCache() {
  try {
    if (!fs.existsSync(TOKEN_CACHE_PATH)) return null;
    const parsed = JSON.parse(fs.readFileSync(TOKEN_CACHE_PATH, "utf-8"));
    if (!parsed || typeof parsed !== "object") return null;
    return parsed;
  } catch {
    return null;
  }
}

function writeDelegatedTokenCache(payload) {
  ensureConfigDir();
  fs.writeFileSync(TOKEN_CACHE_PATH, JSON.stringify(payload, null, 2) + "\n", {
    encoding: "utf-8",
    mode: 0o600,
  });
  fs.chmodSync(TOKEN_CACHE_PATH, 0o600);
}

function isCachedTokenUsable(cache) {
  if (!cache?.access_token || !cache?.expires_at) return false;
  return (
    Date.now() < Number(cache.expires_at) - TOKEN_EXPIRY_SKEW_SECONDS * 1000
  );
}

async function tryRefreshDelegatedToken(cache, credentials) {
  if (!cache?.refresh_token) return null;

  const tenantForRefresh =
    cache.delegated_tenant || credentials.tenantId || "consumers";
  const refreshForm = {
    grant_type: "refresh_token",
    client_id: credentials.clientId,
    refresh_token: cache.refresh_token,
    scope: "offline_access openid profile email Mail.ReadWrite Mail.Send",
  };

  if (credentials.clientSecret) {
    refreshForm.client_secret = credentials.clientSecret;
  }

  const refreshed = await postForm(
    "login.microsoftonline.com",
    `/${tenantForRefresh}/oauth2/v2.0/token`,
    refreshForm,
  );

  if (!refreshed.access_token) {
    return null;
  }

  const refreshedCache = {
    auth_mode: "delegated",
    client_id: credentials.clientId,
    delegated_tenant: tenantForRefresh,
    access_token: refreshed.access_token,
    refresh_token: refreshed.refresh_token || cache.refresh_token,
    expires_at: Date.now() + Number(refreshed.expires_in || 3600) * 1000,
  };
  writeDelegatedTokenCache(refreshedCache);
  return refreshedCache.access_token;
}

async function getClientCredentialsToken() {
  const { clientId, tenantId, clientSecret } = getCredentials();
  const json = await postForm(
    "login.microsoftonline.com",
    `/${tenantId}/oauth2/v2.0/token`,
    {
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
    },
  );

  if (!json.access_token) {
    throw new Error(
      `Token error: ${json.error_description || json.error || "Unknown error"}`,
    );
  }

  return json.access_token;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getDelegatedTokenViaDeviceCode() {
  const credentials = getCredentials();
  const { clientId, tenantId, clientSecret } = credentials;

  const cached = readDelegatedTokenCache();
  if (cached && cached.client_id === clientId && isCachedTokenUsable(cached)) {
    return cached.access_token;
  }

  const refreshedToken = await tryRefreshDelegatedToken(cached, credentials);
  if (refreshedToken) {
    return refreshedToken;
  }

  let delegatedTenant = tenantId;
  let deviceCode = await postForm(
    "login.microsoftonline.com",
    `/${delegatedTenant}/oauth2/v2.0/devicecode`,
    {
      client_id: clientId,
      scope: "offline_access openid profile email Mail.ReadWrite Mail.Send",
    },
  );

  const deviceError = deviceCode.error_description || deviceCode.error || "";
  if (!deviceCode.device_code && /AADSTS9002346/.test(deviceError)) {
    delegatedTenant = "consumers";
    console.log(
      "\nDelegated mode: switching token endpoint to /consumers for Microsoft personal account sign-in.",
    );
    deviceCode = await postForm(
      "login.microsoftonline.com",
      `/${delegatedTenant}/oauth2/v2.0/devicecode`,
      {
        client_id: clientId,
        scope: "offline_access openid profile email Mail.ReadWrite Mail.Send",
      },
    );
  }

  if (!deviceCode.device_code) {
    const detailedError =
      deviceCode.error_description || deviceCode.error || "Unknown error";
    if (/AADSTS70002/.test(detailedError)) {
      throw new Error(
        "Device code error: this app registration is not enabled for device-code/public client usage. " +
          "In Azure App Registration, go to Authentication and enable public client flows (mobile/desktop). " +
          `Details: ${detailedError}`,
      );
    }
    throw new Error(`Device code error: ${detailedError}`);
  }

  console.log("\nSign in required for delegated mode:");
  console.log(
    deviceCode.message ||
      "Open the provided URL and enter the code to continue.",
  );

  const intervalMs = (Number(deviceCode.interval) || 5) * 1000;
  const expiresAt = Date.now() + (Number(deviceCode.expires_in) || 900) * 1000;

  while (Date.now() < expiresAt) {
    await sleep(intervalMs);

    const tokenForm = {
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: clientId,
      device_code: deviceCode.device_code,
    };

    // Some app registrations enforce confidential-client auth even for device flow.
    if (clientSecret) {
      tokenForm.client_secret = clientSecret;
    }

    const tokenResponse = await postForm(
      "login.microsoftonline.com",
      `/${delegatedTenant}/oauth2/v2.0/token`,
      tokenForm,
    );

    if (tokenResponse.access_token) {
      writeDelegatedTokenCache({
        auth_mode: "delegated",
        client_id: clientId,
        delegated_tenant: delegatedTenant,
        access_token: tokenResponse.access_token,
        refresh_token: tokenResponse.refresh_token || null,
        expires_at:
          Date.now() + Number(tokenResponse.expires_in || 3600) * 1000,
      });
      return tokenResponse.access_token;
    }

    const errorCode = tokenResponse.error;
    if (errorCode === "authorization_pending") {
      continue;
    }
    if (errorCode === "slow_down") {
      await sleep(5000);
      continue;
    }
    if (errorCode === "expired_token") {
      break;
    }

    const detailedError =
      tokenResponse.error_description || errorCode || "Unknown error";

    if (errorCode === "invalid_client" || /AADSTS7000218/.test(detailedError)) {
      throw new Error(
        "Token error: delegated sign-in is using an app registration that requires client authentication. " +
          "Either set MS365_CLIENT_SECRET in .env, or enable public client flows on the Azure app registration. " +
          `Details: ${detailedError}`,
      );
    }

    if (/AADSTS70002/.test(detailedError)) {
      throw new Error(
        "Token error: this app registration is not enabled for device-code/public client usage. " +
          "In Azure App Registration, enable public client flows (mobile/desktop). " +
          `Details: ${detailedError}`,
      );
    }

    throw new Error(`Token error: ${detailedError}`);
  }

  throw new Error("Device code authorization timed out. Please try again.");
}

async function getAccessToken() {
  const { authMode } = getCredentials();
  if (authMode === "delegated") {
    return getDelegatedTokenViaDeviceCode();
  }
  return getClientCredentialsToken();
}

module.exports = { getAccessToken, normalizeAuthMode };

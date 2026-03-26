const fs = require("fs");
const path = require("path");
const readline = require("readline");
const { CONFIG_DIR, ENV_PATH, LOCAL_ENV } = require("./paths");

const BASE_VARS = [
  {
    key: "AUTH_MODE",
    label: "Auth mode (client_credentials or delegated)",
  },
  { key: "MS365_EMAIL_CLIENT_ID", label: "Client ID" },
];

const CLIENT_CREDENTIALS_VARS = [
  { key: "MS365_EMAIL_TENANT_ID", label: "Tenant ID" },
  { key: "MS365_EMAIL_CLIENT_SECRET", label: "Client Secret", secret: true },
  { key: "MS365_EMAIL_ADDRESS", label: "Mailbox email address" },
];

const CONFIG_DIR_MODE = 0o700;
const ENV_FILE_MODE = 0o600;

function parseEnvFile(filePath) {
  const env = {};
  if (!fs.existsSync(filePath)) return env;
  const lines = fs.readFileSync(filePath, "utf-8").split("\n");
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx === -1) continue;
    env[trimmed.slice(0, idx).trim()] = trimmed.slice(idx + 1).trim();
  }
  return env;
}

function loadEnv() {
  const local = parseEnvFile(LOCAL_ENV);
  const home = parseEnvFile(ENV_PATH);
  return { ...local, ...home };
}

function normalizeAuthMode(value) {
  const mode = String(value || "client_credentials")
    .trim()
    .toLowerCase();
  return mode === "delegated" ? "delegated" : "client_credentials";
}

function getRequiredVars(env) {
  const authMode = normalizeAuthMode(env.AUTH_MODE);
  if (authMode === "delegated") {
    return BASE_VARS;
  }
  return [...BASE_VARS, ...CLIENT_CREDENTIALS_VARS];
}

function getMissingVars(env) {
  return getRequiredVars(env).filter((v) => !env[v.key]);
}

function ask(question, mask = false) {
  return new Promise((resolve) => {
    if (!mask) {
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
      });
      rl.question(question, (ans) => {
        rl.close();
        resolve(ans.trim());
      });
    } else {
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
        terminal: true,
      });
      process.stdout.write(question);
      let buf = "";
      const onData = (char) => {
        char = char + "";
        if (char === "\n" || char === "\r" || char === "\u0004") {
          process.stdin.removeListener("data", onData);
          rl.close();
          process.stdout.write("\n");
          resolve(buf.trim());
        } else {
          buf += char;
          process.stdout.write("*");
        }
      };
      process.stdin.on("data", onData);
    }
  });
}

async function runWizard() {
  const existing = loadEnv();
  existing.AUTH_MODE = normalizeAuthMode(existing.AUTH_MODE);
  const missing = getMissingVars(existing);

  if (missing.length === 0) {
    console.log("All config variables are set in .env");
    const ans = await ask("Reconfigure all? (y/N): ");
    if (ans.toLowerCase() !== "y") {
      console.log("No changes made.");
      return;
    }
  } else {
    console.log(`Missing: ${missing.map((v) => v.label).join(", ")}\n`);
  }

  console.log("=== MS365 Mail Manager Config ===\n");
  const values = { ...existing };
  if (!values.AUTH_MODE) {
    values.AUTH_MODE = "client_credentials";
  }

  const modeInput = await ask(
    `Auth mode (${values.AUTH_MODE}) [client_credentials/delegated]: `,
  );
  if (modeInput) {
    values.AUTH_MODE = normalizeAuthMode(modeInput);
  }

  const requiredVars = getRequiredVars(values).filter(
    (v) => v.key !== "AUTH_MODE",
  );

  for (const v of requiredVars) {
    const hasCurrent = !!existing[v.key];
    const masked =
      hasCurrent && v.secret
        ? "****" + existing[v.key].slice(-4)
        : existing[v.key] || "(not set)";

    if (hasCurrent && !missing.some((m) => m.key === v.key)) {
      const ans = await ask(`${v.label} [${masked}]: `);
      if (ans) values[v.key] = ans;
      continue;
    }

    const ans = await ask(`${v.label}: `, v.secret);
    if (!ans && hasCurrent) continue;
    if (!ans) {
      console.error(`${v.label} is required.`);
      process.exit(1);
    }
    values[v.key] = ans;
  }

  if (values.AUTH_MODE === "delegated") {
    values.MS365_EMAIL_TENANT_ID = "consumers";
    delete values.MS365_EMAIL_CLIENT_SECRET;
    delete values.MS365_EMAIL_ADDRESS;
  }

  const content =
    Object.entries(values)
      .filter(
        ([, value]) =>
          value !== undefined && value !== null && String(value).trim() !== "",
      )
      .map(([key, value]) => `${key}=${String(value).trim()}`)
      .join("\n") + "\n";
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: CONFIG_DIR_MODE });
  }
  fs.chmodSync(CONFIG_DIR, CONFIG_DIR_MODE);
  fs.writeFileSync(ENV_PATH, content, {
    encoding: "utf-8",
    mode: ENV_FILE_MODE,
  });
  fs.chmodSync(ENV_PATH, ENV_FILE_MODE);
  console.log(`\nSaved to ${ENV_PATH}`);
}

function checkConfig() {
  return getMissingVars(loadEnv());
}

module.exports = { runWizard, checkConfig };

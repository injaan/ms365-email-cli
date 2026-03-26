const fs = require("fs");
const path = require("path");
const readline = require("readline");
const { CONFIG_DIR, ENV_PATH, LOCAL_ENV } = require("./paths");

const REQUIRED_VARS = [
  { key: "MS365_EMAIL_CLIENT_ID", label: "Client ID" },
  { key: "MS365_EMAIL_TENANT_ID", label: "Tenant ID" },
  { key: "MS365_EMAIL_CLIENT_SECRET", label: "Client Secret", secret: true },
  { key: "MS365_FROM_EMAIL", label: "Mailbox email address" },
];

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

function getMissingVars(env) {
  return REQUIRED_VARS.filter((v) => !env[v.key]);
}

function ask(question, mask = false) {
  return new Promise((resolve) => {
    if (!mask) {
      const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
      rl.question(question, (ans) => { rl.close(); resolve(ans.trim()); });
    } else {
      const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true });
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
  const missing = getMissingVars(existing);

  if (missing.length === 0) {
    console.log("All config variables are set in .env");
    const ans = await ask("Reconfigure all? (y/N): ");
    if (ans.toLowerCase() !== "y") { console.log("No changes made."); return; }
  } else {
    console.log(`Missing: ${missing.map((v) => v.label).join(", ")}\n`);
  }

  console.log("=== MS365 Mail Manager Config ===\n");
  const values = { ...existing };

  for (const v of REQUIRED_VARS) {
    const hasCurrent = !!existing[v.key];
    const masked = hasCurrent && v.secret ? "****" + existing[v.key].slice(-4) : existing[v.key] || "(not set)";

    if (hasCurrent && !missing.some((m) => m.key === v.key)) {
      const ans = await ask(`${v.label} [${masked}]: `);
      if (ans) values[v.key] = ans;
      continue;
    }

    const ans = await ask(`${v.label}: `, v.secret);
    if (!ans && hasCurrent) continue;
    if (!ans) { console.error(`${v.label} is required.`); process.exit(1); }
    values[v.key] = ans;
  }

  const content = REQUIRED_VARS.map((v) => `${v.key}=${values[v.key]}`).join("\n") + "\n";
  if (!fs.existsSync(CONFIG_DIR)) fs.mkdirSync(CONFIG_DIR, { recursive: true });
  fs.writeFileSync(ENV_PATH, content, "utf-8");
  console.log(`\nSaved to ${ENV_PATH}`);
}

function checkConfig() {
  return getMissingVars(loadEnv());
}

module.exports = { runWizard, checkConfig };

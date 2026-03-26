const fs = require("fs");
const path = require("path");
const os = require("os");

const pkg = JSON.parse(fs.readFileSync(path.resolve(__dirname, "..", "package.json"), "utf-8"));
const PROJECT_NAME = pkg.name;

function getConfigDir() {
  if (process.platform === "win32") {
    const appData = process.env.APPDATA || path.resolve(os.homedir(), "AppData", "Roaming");
    return path.resolve(appData, PROJECT_NAME);
  }
  return path.resolve(os.homedir(), `.${PROJECT_NAME}`);
}

const CONFIG_DIR = getConfigDir();
const ENV_PATH = path.resolve(CONFIG_DIR, ".env");
const LOCAL_ENV = path.resolve(__dirname, "..", ".env");

module.exports = { PROJECT_NAME, CONFIG_DIR, ENV_PATH, LOCAL_ENV };

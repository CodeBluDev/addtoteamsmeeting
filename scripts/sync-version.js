const fs = require("fs");
const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const versionPath = path.join(rootDir, "version.json");
const manifestPath = path.join(rootDir, "manifest.xml");
const commandsPath = path.join(rootDir, "src", "commands", "commands.js");

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function writeFile(filePath, content) {
  fs.writeFileSync(filePath, content, "utf8");
}

function updateCommands(version, buildMarker, cacheBuster, baseUrl) {
  let content = fs.readFileSync(commandsPath, "utf8");
  content = content.replace(
    /const BUILD_TAG = ".*?";/g,
    `const BUILD_TAG = "${version}";`
  );
  content = content.replace(
    /const BUILD_MARKER = ".*?";/g,
    `const BUILD_MARKER = "${buildMarker}";`
  );
  if (cacheBuster) {
    const base = baseUrl || "https://mvteamsmeetinglink.netlify.app";
    content = content.replace(
      /const DIALOG_URL = ".*?";/g,
      `const DIALOG_URL = "${base}/create-event.html?v=${cacheBuster}";`
    );
    content = content.replace(
      /const AUTH_DIALOG_URL = ".*?";/g,
      `const AUTH_DIALOG_URL = "${base}/auth.html?v=${cacheBuster}";`
    );
  }
  writeFile(commandsPath, content);
}

function updateManifest(version, cacheBuster, manifestVersion, baseUrl) {
  let content = fs.readFileSync(manifestPath, "utf8");
  const labelText = `Add Teams Meeting to Location (${version})`;

  content = content.replace(
    /<DisplayName DefaultValue="[^"]*"\s*\/>/g,
    `<DisplayName DefaultValue="${labelText}"/>`
  );
  content = content.replace(
    /<bt:String id="ActionButton\.Label" DefaultValue="[^"]*"\s*\/>/g,
    `<bt:String id="ActionButton.Label" DefaultValue="${labelText}"/>`
  );

  if (manifestVersion) {
    content = content.replace(
      /<Version>[^<]*<\/Version>/g,
      `<Version>${manifestVersion}</Version>`
    );
  }

  if (baseUrl && cacheBuster) {
    const escapedBase = baseUrl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const baseRegex = new RegExp(`(${escapedBase}\\/[^"]+?)(\\?v=[^"]*)?"`, "g");
    content = content.replace(baseRegex, `$1?v=${cacheBuster}"`);
  }

  writeFile(manifestPath, content);
}

function main() {
  const versionInfo = readJson(versionPath);
  updateCommands(
    versionInfo.version,
    versionInfo.buildMarker,
    versionInfo.cacheBuster,
    versionInfo.baseUrl
  );
  updateManifest(
    versionInfo.version,
    versionInfo.cacheBuster,
    versionInfo.manifestVersion,
    versionInfo.baseUrl
  );
  console.log(`Synced version ${versionInfo.version}.`);
}

main();

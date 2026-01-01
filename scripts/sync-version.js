const fs = require("fs");
const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const versionPath = path.join(rootDir, "version.json");
const manifestPath = path.join(rootDir, "manifest.xml");
const manifestDevPath = path.join(rootDir, "manifest.dev.xml");
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
    content = content.replace(
      /const CACHE_BUSTER = ".*?";/g,
      `const CACHE_BUSTER = "${cacheBuster}";`
    );
  }
  if (baseUrl) {
    content = content.replace(
      /const DEFAULT_BASE_URL = ".*?";/g,
      `const DEFAULT_BASE_URL = "${baseUrl}";`
    );
  }
  writeFile(commandsPath, content);
}

function updateManifest(filePath, version, cacheBuster, manifestVersion, baseUrl) {
  let content = fs.readFileSync(filePath, "utf8");
  const labelText = `[${version}] Add Teams Meeting to Location`;

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

  writeFile(filePath, content);
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
    manifestPath,
    versionInfo.version,
    versionInfo.cacheBuster,
    versionInfo.manifestVersion,
    versionInfo.baseUrl
  );
  updateManifest(
    manifestDevPath,
    versionInfo.version,
    versionInfo.cacheBuster,
    versionInfo.manifestVersion,
    versionInfo.devBaseUrl || "https://127.0.0.1:3000"
  );
  console.log(`Synced version ${versionInfo.version}.`);
}

main();

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

function updateCommands(version, buildMarker, cacheBuster) {
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
      /const NOTIFICATION_ICON_URL = ".*?";/g,
      `const NOTIFICATION_ICON_URL = "https://mvteamsmeetinglink.netlify.app/assets/codeblu-teams-16.png?v=${cacheBuster}";`
    );
    content = content.replace(
      /const DIALOG_URL = ".*?";/g,
      `const DIALOG_URL = "https://mvteamsmeetinglink.netlify.app/create-event.html?v=${cacheBuster}";`
    );
  }
  writeFile(commandsPath, content);
}

function updateManifest(version, cacheBuster, manifestVersion) {
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

  content = content.replace(
    /(https:\/\/mvteamsmeetinglink\.netlify\.app\/[^"]+?)(\?v=[^"]*)?"/g,
    `$1?v=${cacheBuster}"`
  );

  writeFile(manifestPath, content);
}

function main() {
  const versionInfo = readJson(versionPath);
  updateCommands(versionInfo.version, versionInfo.buildMarker, versionInfo.cacheBuster);
  updateManifest(
    versionInfo.version,
    versionInfo.cacheBuster,
    versionInfo.manifestVersion
  );
  console.log(`Synced version ${versionInfo.version}.`);
}

main();

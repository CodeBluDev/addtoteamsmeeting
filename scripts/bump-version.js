const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const versionPath = path.resolve(__dirname, "..", "version.json");

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function writeJson(filePath, data) {
  fs.writeFileSync(filePath, `${JSON.stringify(data, null, 2)}\n`, "utf8");
}

function formatBuildMarker(date) {
  return date.toISOString().replace(/:\d{2}\.\d{3}Z$/, "Z");
}

function bumpPatch(version) {
  const match = version.match(/^v(\d+)\.(\d+)\.(\d+)$/);
  if (!match) {
    throw new Error(`Unsupported version format: ${version}`);
  }
  const major = Number(match[1]);
  const minor = Number(match[2]);
  const patch = Number(match[3]) + 1;
  return `v${major}.${minor}.${patch}`;
}

function toCacheBuster(version) {
  return version.replace(/^v/, "");
}

function toManifestVersion(cacheBuster) {
  return `${cacheBuster}.0`;
}

function main() {
  const data = readJson(versionPath);
  const nextVersion = bumpPatch(data.version);
  const cacheBuster = toCacheBuster(nextVersion);
  data.version = nextVersion;
  data.cacheBuster = cacheBuster;
  data.manifestVersion = toManifestVersion(cacheBuster);
  data.buildMarker = formatBuildMarker(new Date());
  writeJson(versionPath, data);
  execSync("node scripts/sync-version.js", { stdio: "inherit" });
}

main();

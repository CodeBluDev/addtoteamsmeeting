const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");

const rootDir = path.resolve(__dirname, "..");
const versionPath = path.join(rootDir, "version.json");
const manifestPath = path.join(rootDir, "manifest.xml");
const manifestDevPath = path.join(rootDir, "manifest.dev.xml");
const manifestTemplatePath = path.join(rootDir, "manifest.template.xml");
const manifestDevTemplatePath = path.join(rootDir, "manifest.dev.template.xml");
const commandsPath = path.join(rootDir, "src", "commands", "commands.js");

dotenv.config();

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function writeFile(filePath, content) {
  fs.writeFileSync(filePath, content, "utf8");
}

function requireEnv(name, fallback) {
  const value = process.env[name];
  if (value !== undefined && value !== "") {
    return value;
  }
  if (fallback !== undefined && fallback !== "") {
    return fallback;
  }
  throw new Error(`Missing required environment variable: ${name}`);
}

function normalizeBaseUrl(value) {
  if (!value) {
    return value;
  }
  return value.replace(/\/+$/, "");
}

function formatAppDomains(domains) {
  const entries = String(domains || "")
    .split(",")
    .map((domain) => domain.trim())
    .filter(Boolean);
  return entries.map((domain) => `<AppDomain>${domain}</AppDomain>`).join("\n    ");
}

function applyReplacements(content, replacements) {
  let output = content;
  Object.keys(replacements).forEach((key) => {
    output = output.split(key).join(replacements[key]);
  });
  return output;
}

function renderManifestTemplate(templatePath, outputPath, replacements) {
  const template = fs.readFileSync(templatePath, "utf8");
  const rendered = applyReplacements(template, replacements);
  writeFile(outputPath, rendered);
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
      /const CACHE_BUSTER = ".*?";/g,
      `const CACHE_BUSTER = "${cacheBuster}";`
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
    content = content.replace(/https:\/\/127\.0\.0\.1:3000/g, baseUrl);
    const escapedBase = baseUrl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const baseRegex = new RegExp(`(${escapedBase}\\/[^"]+?)(\\?v=[^"]*)?"`, "g");
    content = content.replace(baseRegex, `$1?v=${cacheBuster}"`);
  }

  writeFile(filePath, content);
}

function main() {
  const versionInfo = readJson(versionPath);
  const appId = requireEnv("MANIFEST_APP_ID");

  const baseUrl = normalizeBaseUrl(requireEnv("APP_BASE_URL"));
  const devBaseUrl = normalizeBaseUrl(requireEnv("APP_BASE_URL_DEV", baseUrl));
  const appDomains = formatAppDomains(requireEnv("APP_DOMAINS"));
  const appDomainsDev = formatAppDomains(requireEnv("APP_DOMAINS_DEV", requireEnv("APP_DOMAINS")));

  const templateReplacements = {
    __APP_ID__: appId,
    __APP_BASE_URL__: baseUrl || "",
    __APP_BASE_URL_DEV__: devBaseUrl || "",
    __APP_DOMAINS__: appDomains,
    __APP_DOMAINS_DEV__: appDomainsDev,
    __CACHE_BUSTER__: versionInfo.cacheBuster
  };

  renderManifestTemplate(manifestTemplatePath, manifestPath, templateReplacements);
  renderManifestTemplate(manifestDevTemplatePath, manifestDevPath, templateReplacements);

  updateCommands(
    versionInfo.version,
    versionInfo.buildMarker,
    versionInfo.cacheBuster
  );
  updateManifest(
    manifestPath,
    versionInfo.version,
    versionInfo.cacheBuster,
    versionInfo.manifestVersion,
    baseUrl
  );
  updateManifest(
    manifestDevPath,
    versionInfo.version,
    versionInfo.cacheBuster,
    versionInfo.manifestVersion,
    devBaseUrl
  );
  console.log(`Synced version ${versionInfo.version}.`);
}

main();

const path = require("path");
const fs = require("fs");

const WORD_EXTENSIONS = new Set([
  ".doc",
  ".docx",
  ".docm",
  ".dot",
  ".dotx",
  ".dotm",
  ".odt",
  ".rtf"
]);

const EXCEL_EXTENSIONS = new Set([
  ".xls",
  ".xlsx",
  ".xlsm",
  ".xlsb",
  ".xlt",
  ".xltx",
  ".xltm",
  ".xlam",
  ".csv",
  ".ods"
]);

const POWERPOINT_EXTENSIONS = new Set([
  ".ppt",
  ".pptx",
  ".pptm",
  ".pps",
  ".ppsx",
  ".pot",
  ".potx",
  ".potm",
  ".ppam",
  ".odp"
]);

const ONENOTE_EXTENSIONS = new Set([
  ".one",
  ".onetoc2",
  ".onepkg"
]);

function cleanArgument(argument) {
  if (!argument || typeof argument !== "string") {
    return null;
  }

  let cleaned = argument.trim();
  if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
    cleaned = cleaned.slice(1, -1);
  }

  if (cleaned === "@@" || cleaned === "@@u" || cleaned === "@@U") {
    return null;
  }

  if (cleaned.startsWith("file://")) {
    try {
      return decodeURIComponent(new URL(cleaned).pathname);
    } catch (_) {
      return cleaned.replace(/^file:\/\//, "");
    }
  }

  return cleaned;
}

function detectService(filePath) {
  const extension = path.extname(filePath).toLowerCase();

  if (WORD_EXTENSIONS.has(extension)) {
    return "word";
  }
  if (EXCEL_EXTENSIONS.has(extension)) {
    return "excel";
  }
  if (POWERPOINT_EXTENSIONS.has(extension)) {
    return "powerpoint";
  }
  if (ONENOTE_EXTENSIONS.has(extension)) {
    return "onenote";
  }

  return null;
}

function findOfficeFiles(argumentsList) {
  const files = [];

  for (const argument of argumentsList) {
    const candidate = cleanArgument(argument);

    if (!candidate) continue;
    if (candidate === "--background" || candidate === "--quit-ui") continue;

    try {
      if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) {
        const service = detectService(candidate);

        if (service) {
          const resolvedPath = path.resolve(candidate);
          files.push({
            path: resolvedPath,
            originalPath: resolvedPath,
            name: path.basename(resolvedPath),
            extension: path.extname(resolvedPath).toLowerCase(),
            service
          });
        }
      }
    } catch (_) {}
  }

  return files;
}

module.exports = {
  detectService,
  findOfficeFiles
};

const path = require("path");
const fs = require("fs");
const { importExternalFile } = require("./external-file-sync");

function detectService(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if ([".doc", ".docx", ".odt", ".rtf", ".txt"].includes(ext)) return "word";
  if ([".xls", ".xlsx", ".ods", ".csv"].includes(ext)) return "excel";
  if ([".ppt", ".pptx", ".odp"].includes(ext)) return "powerpoint";
  if ([".one"].includes(ext)) return "onenote";
  return "word";
}

function findOfficeFiles(argv) {
  const files = [];
  for (let i = 1; i < argv.length; i++) {
    const arg = argv[i];
    if (arg && !arg.startsWith("-")) {
      const resolved = path.resolve(arg);
      if (fs.existsSync(resolved) && fs.statSync(resolved).isFile()) {
        const service = detectService(resolved);
        const managedPath = importExternalFile(resolved) || resolved;

        files.push({
          path: managedPath,
          originalPath: resolved,
          name: path.basename(resolved),
          service: service
        });
      }
    }
  }
  return files;
}

module.exports = { findOfficeFiles, detectService };

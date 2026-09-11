const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { execFile } = require("child_process");
const oneDrive = require("./onedrive-manager");

const ONEDRIVE_BIN = "/app/bin/onedrive";

function getImportsDirectory() {
  const syncDir = oneDrive.syncDirectory || path.join(os.homedir(), "OneDrive");
  return path.join(syncDir, "Microsoft 365 for Linux Imports");
}

function extractWebUrl(rawOutput) {
  if (!rawOutput || typeof rawOutput !== "string") return null;
  const clean = rawOutput.replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, "");
  const urls = clean.match(/https?:\/\/[^\s"'`<>]+/gi);
  if (!urls) return null;

  for (const u of urls) {
    const lower = u.toLowerCase();
    if (
      !lower.includes("graph.microsoft.com") &&
      !lower.includes("login.microsoftonline.com") &&
      !lower.includes("account.live.com")
    ) {
      return u;
    }
  }
  return null;
}

/**
 * Consulta o binário CLI como fallback secundário
 */
function getFileWebUrlCliFallback(targetPath) {
  return new Promise((resolve) => {
    const confDir = oneDrive.configDirectory || path.join(os.homedir(), ".config", "onedrive");
    const syncDir = oneDrive.syncDirectory || path.join(os.homedir(), "OneDrive");
    const relPath = path.relative(syncDir, targetPath);

    const args = [
      "--confdir", confDir,
      "--syncdir", syncDir,
      "--get-file-link", relPath
    ];

    const options = {
      cwd: syncDir,
      env: { ...process.env, HOME: os.homedir() },
      timeout: 5000
    };

    execFile(ONEDRIVE_BIN, args, options, (error, stdout, stderr) => {
      const output = (stdout || "") + "\n" + (stderr || "");
      resolve(extractWebUrl(output));
    });
  });
}

/**
 * Prepara o documento e obtém seu link oficial online via Microsoft Graph API
 */
async function getFileWebUrl(filePath) {
  if (!filePath || !fs.existsSync(filePath)) {
    return { insideOneDrive: false, webUrl: null };
  }

  const absPath = path.resolve(filePath);
  const syncDir = path.resolve(oneDrive.syncDirectory || path.join(os.homedir(), "OneDrive"));

  let workingPath = absPath;
  let isImported = false;

  if (!absPath.startsWith(syncDir)) {
    const importsDir = getImportsDirectory();
    try {
      if (!fs.existsSync(importsDir)) {
        fs.mkdirSync(importsDir, { recursive: true, mode: 0o700 });
      }

      const fileHash = crypto.createHash("md5").update(absPath).digest("hex").substring(0, 12);
      const subDir = path.join(importsDir, fileHash);
      if (!fs.existsSync(subDir)) {
        fs.mkdirSync(subDir, { recursive: true, mode: 0o700 });
      }

      const fileName = path.basename(absPath);
      workingPath = path.join(subDir, fileName);

      if (!fs.existsSync(workingPath)) {
        fs.copyFileSync(absPath, workingPath);
        console.log(`[Office Import] Copied to OneDrive: ${workingPath}`);
      }
      isImported = true;
    } catch (err) {
      console.warn("[Office Online] Fallback to original file:", err.message);
      workingPath = absPath;
    }
  }

  const inside = workingPath.startsWith(syncDir);
  let webUrl = null;

  if (inside && oneDrive.isAuthenticated()) {
    const relPath = path.relative(syncDir, workingPath);
    
    // 1. Tenta obter o link instantâneo pela Microsoft Graph API
    webUrl = await oneDrive.fetchGraphFileWebUrl(relPath);

    // 2. Se a Graph API ainda não encontrou o item (em processo de upload), tenta o fallback CLI
    if (!webUrl) {
      webUrl = await getFileWebUrlCliFallback(workingPath);
    }
  }

  return {
    insideOneDrive: inside,
    webUrl: webUrl,
    originalPath: absPath,
    localOneDrivePath: workingPath,
    imported: isImported
  };
}

module.exports = {
  getFileWebUrl,
  getImportsDirectory
};

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

/**
 * Extrai qualquer link web válido da Microsoft contido na saída do terminal,
 * removendo códigos de cor ANSI e prefixos de texto.
 */
function extractWebUrl(rawOutput) {
  if (!rawOutput || typeof rawOutput !== "string") return null;

  // Remove caracteres de escape ANSI de cor do terminal
  const clean = rawOutput.replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, "");

  // Captura URLs completas http/https
  const urls = clean.match(/https?:\/\/[^\s"'`<>]+/gi);
  if (!urls) return null;

  for (const u of urls) {
    const lower = u.toLowerCase();
    // Filtra URLs internas da API de autenticação/Graph
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
 * Consulta o cliente onedrive para obter o link oficial web do documento.
 * Argumentos compatíveis com a versão compilada do abraunegg/onedrive.
 */
function getFileWebUrlAsync(targetPath) {
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
      timeout: 8000
    };

    execFile(ONEDRIVE_BIN, args, options, (error, stdout, stderr) => {
      const output = (stdout || "") + "\n" + (stderr || "");
      const link = extractWebUrl(output);

      if (link) {
        console.log(`[Office Online] Official Web URL found: ${link}`);
        return resolve(link);
      }

      // Log resumido para acompanhamento no terminal
      const cleanOut = output.replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, "").trim();
      if (cleanOut) {
        const firstLine = cleanOut.split("\n")
          .map(l => l.trim())
          .filter(l => l && !l.startsWith("Reading configuration") && !l.startsWith("Configuration file") && !l.startsWith("Using IPv"))
          .join(" | ");
        if (firstLine) {
          console.log(`[Office Online Link Query] ${firstLine.substring(0, 140)}`);
        }
      }

      resolve(null);
    });
  });
}

/**
 * Prepara o arquivo externo para ser acessado pelo OneDrive e obtém sua URL web.
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
        console.log(`[Office Import] Copied to OneDrive folder: ${workingPath}`);
      }
      isImported = true;
    } catch (err) {
      console.warn("[Office Online] Failed to import file, falling back to original:", err.message);
      workingPath = absPath;
    }
  }

  const inside = workingPath.startsWith(syncDir);
  let webUrl = null;

  if (inside) {
    webUrl = await getFileWebUrlAsync(workingPath);
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

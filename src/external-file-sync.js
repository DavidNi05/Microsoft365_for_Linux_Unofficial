const fs = require("fs");
const fsp = require("fs").promises;
const path = require("path");
const os = require("os");
const crypto = require("crypto");

function getOneDriveManager() {
  return require("./onedrive-manager");
}

const homeDirectory = os.homedir();
const configBase = process.env.XDG_CONFIG_HOME || path.join(homeDirectory, ".config");
const configDirectory = path.join(configBase, "onedrive");
const mappingsFile = path.join(configDirectory, "external-mappings.json");

let statusReporter = null;

function setStatusReporter(fn) {
  statusReporter = fn;
}

function notifyStatus(status) {
  if (typeof statusReporter === "function") {
    statusReporter(status);
  }
}

function loadMappings() {
  try {
    if (fs.existsSync(mappingsFile)) {
      const data = JSON.parse(fs.readFileSync(mappingsFile, "utf8"));
      if (data && typeof data === "object") return data;
    }
  } catch (_) {}
  return {};
}

function saveMappings(mappings) {
  try {
    fs.mkdirSync(path.dirname(mappingsFile), { recursive: true, mode: 0o700 });
    fs.writeFileSync(mappingsFile, JSON.stringify(mappings, null, 2), { encoding: "utf8", mode: 0o600 });
  } catch (_) {}
}

function registerExternalFile(originalPath, oneDrivePath, imported = true) {
  if (!originalPath || !oneDrivePath) return;
  const absOriginal = path.resolve(originalPath);
  const absOneDrive = path.resolve(oneDrivePath);

  const mappings = loadMappings();
  mappings[absOriginal] = {
    oneDrivePath: absOneDrive,
    imported,
    updatedAt: new Date().toISOString()
  };
  saveMappings(mappings);
  console.log(`[Mapping] Registered mapping: "${absOriginal}" -> "${absOneDrive}"`);
}

/**
 * Busca flexível e bidirecional de mapeamento:
 * Encontra tanto pelo caminho original quanto pelo caminho do OneDrive.
 */
function getExternalMapping(targetPath) {
  if (!targetPath) return null;
  const absTarget = path.resolve(targetPath);
  const mappings = loadMappings();

  // 1. Busca direta por chave original
  if (mappings[absTarget]) {
    return { originalPath: absTarget, ...mappings[absTarget] };
  }

  // 2. Busca reversa (quando targetPath é o caminho de dentro do OneDrive)
  for (const [orig, info] of Object.entries(mappings)) {
    if (info && info.oneDrivePath) {
      const absInfoOD = path.resolve(info.oneDrivePath);
      // Correspondência exata ou mesmo diretório de hash
      if (absInfoOD === absTarget || path.dirname(absInfoOD) === path.dirname(absTarget)) {
        return { originalPath: path.resolve(orig), ...info };
      }
    }
  }

  // 3. Busca por correspondência no diretório de imports pelo hash
  const oneDrive = getOneDriveManager();
  const syncDir = oneDrive.syncDirectory || path.join(os.homedir(), "OneDrive");
  const importsDir = path.join(syncDir, "Microsoft 365 for Linux Imports");

  if (absTarget.startsWith(importsDir)) {
    const rel = path.relative(importsDir, absTarget);
    const targetHash = rel.split(path.sep)[0];

    for (const [orig, info] of Object.entries(mappings)) {
      const origHash = crypto.createHash("md5").update(path.resolve(orig)).digest("hex").substring(0, 12);
      if (origHash === targetHash) {
        return { originalPath: path.resolve(orig), ...info };
      }
    }
  }

  return null;
}

function importExternalFile(originalPath) {
  if (!originalPath || !fs.existsSync(originalPath)) return null;

  const oneDrive = getOneDriveManager();
  const syncDir = oneDrive.syncDirectory || path.join(os.homedir(), "OneDrive");
  const absOriginal = path.resolve(originalPath);

  if (absOriginal.startsWith(syncDir)) {
    return absOriginal;
  }

  const importsDir = path.join(syncDir, "Microsoft 365 for Linux Imports");
  fs.mkdirSync(importsDir, { recursive: true, mode: 0o700 });

  const fileHash = crypto.createHash("md5").update(absOriginal).digest("hex").substring(0, 12);
  const targetSubDir = path.join(importsDir, fileHash);
  fs.mkdirSync(targetSubDir, { recursive: true, mode: 0o700 });

  const fileName = path.basename(absOriginal);
  const targetPath = path.join(targetSubDir, fileName);

  try {
    fs.copyFileSync(absOriginal, targetPath);
    registerExternalFile(absOriginal, targetPath, true);
    console.log(`[Office Import] Copied to OneDrive folder: ${targetPath}`);
    return targetPath;
  } catch (err) {
    console.warn("[Office Import] Error copying file:", err.message);
    return absOriginal;
  }
}

function cleanDocTitle(rawTitle, originalExtension) {
  if (!rawTitle || typeof rawTitle !== "string") return null;
  let title = rawTitle.trim();

  if (title.includes("…") || title.includes("...")) {
    return null;
  }

  const invalidKeywords = [
    "Word na Web", "Excel na Web", "PowerPoint na Web",
    "Microsoft Word", "Microsoft Excel", "Microsoft PowerPoint",
    "Criar e editar", "Documento", "Document", "Continuar",
    "Continue", "Salvo", "Saved", "|", "-"
  ];

  const lower = title.toLowerCase();
  for (const kw of invalidKeywords) {
    if (lower === kw.toLowerCase() || lower.startsWith(kw.toLowerCase() + " -")) {
      return null;
    }
  }

  title = title.replace(/[/\\?%*:|"<>]/g, "").trim();
  if (!title) return null;

  const ext = originalExtension || ".docx";
  if (!title.toLowerCase().endsWith(ext.toLowerCase())) {
    title += ext;
  }

  return title;
}

/**
 * Salva as alterações de forma atômica e segura de volta para o caminho original.
 */
async function saveLocalNow(currentPath, expectedWebTitle = null) {
  if (!currentPath) {
    console.error("[Save Local] Aborted: No currentPath provided.");
    return { ok: false, error: "No path specified." };
  }

  const mapping = getExternalMapping(currentPath);

  if (!mapping || !mapping.originalPath) {
    console.error(`[Save Local] Failed: No mapping found for "${currentPath}".`);
    return { ok: false, error: "No OneDrive mapping found for this file." };
  }

  const absOriginal = path.resolve(mapping.originalPath);
  let absOneDrive = path.resolve(mapping.oneDrivePath);

  // Se o arquivo foi renomeado no OneDrive, descobre o arquivo atual dentro da subpasta de hash
  if (!fs.existsSync(absOneDrive)) {
    const parentDir = path.dirname(absOneDrive);
    if (fs.existsSync(parentDir)) {
      const files = fs.readdirSync(parentDir).filter(f => !f.startsWith("."));
      if (files.length > 0) {
        absOneDrive = path.join(parentDir, files[0]);
        console.log(`[Save Local] Resolved active OneDrive working file: "${absOneDrive}"`);
      }
    }
  }

  if (!fs.existsSync(absOneDrive)) {
    console.error(`[Save Local] Working file not found in OneDrive directory: "${absOneDrive}"`);
    return { ok: false, fileName: path.basename(absOriginal), error: "Working file not found in OneDrive directory." };
  }

  const originalDir = path.dirname(absOriginal);
  const originalExt = path.extname(absOriginal);

  let targetFileName = path.basename(absOneDrive); // Nome herdado do OneDrive
  let wasRenamed = false;

  if (expectedWebTitle) {
    const cleaned = cleanDocTitle(expectedWebTitle, originalExt);
    if (cleaned) {
      targetFileName = cleaned;
    }
  }

  if (targetFileName.toLowerCase() !== path.basename(absOriginal).toLowerCase()) {
    wasRenamed = true;
    console.log(`[Save Local] Rename confirmed: "${path.basename(absOriginal)}" -> "${targetFileName}"`);
  }

  const newOriginalPath = path.join(originalDir, targetFileName);
  const tempSavePath = path.join(originalDir, `.${targetFileName}.tmp-${Date.now()}`);

  try {
    // 1. Copia de OneDrive para arquivo temporário no mesmo diretório de destino
    await fsp.copyFile(absOneDrive, tempSavePath);

    // 2. Transação atômica (rename substitui o arquivo com segurança)
    try {
      await fsp.rename(tempSavePath, newOriginalPath);
    } catch (renameErr) {
      if (renameErr.code === "EXDEV") {
        await fsp.copyFile(tempSavePath, newOriginalPath);
        await fsp.unlink(tempSavePath);
      } else {
        throw renameErr;
      }
    }

    console.log(`[Save Local] Successfully saved to disk: "${newOriginalPath}"`);

    // 3. Se foi renomeado, exclui o arquivo antigo caso o nome tenha mudado no mesmo diretório
    if (wasRenamed && newOriginalPath !== absOriginal && fs.existsSync(absOriginal)) {
      try {
        await fsp.unlink(absOriginal);
        console.log(`[Save Local] Removed old original file: "${absOriginal}"`);
      } catch (_) {}
    }

    // 4. Atualiza os registros do mapeamento
    const mappings = loadMappings();
    delete mappings[absOriginal];
    mappings[newOriginalPath] = {
      oneDrivePath: absOneDrive,
      imported: mapping.imported,
      updatedAt: new Date().toISOString()
    };
    saveMappings(mappings);

    notifyStatus("Saved ✓");
    return {
      ok: true,
      renamed: wasRenamed,
      fileName: targetFileName,
      originalPath: newOriginalPath,
      oneDrivePath: absOneDrive,
      message: wasRenamed ? `Saved and renamed to ${targetFileName}` : `Saved successfully: ${targetFileName}`
    };

  } catch (error) {
    console.error("[Save Local] Error during atomic write/rename:", error);
    try {
      if (fs.existsSync(tempSavePath)) {
        await fsp.unlink(tempSavePath);
      }
    } catch (_) {}

    notifyStatus("Save error");
    return { ok: false, fileName: targetFileName, error: error.message };
  }
}

function syncExternalCopiesBack() {
  const mappings = loadMappings();
  for (const [origPath, info] of Object.entries(mappings)) {
    if (info && info.oneDrivePath && fs.existsSync(info.oneDrivePath)) {
      if (fs.existsSync(origPath)) {
        try {
          const origStat = fs.statSync(origPath);
          const oneDriveStat = fs.statSync(info.oneDrivePath);
          if (oneDriveStat.mtimeMs > origStat.mtimeMs) {
            fs.copyFileSync(info.oneDrivePath, origPath);
          }
        } catch (_) {}
      }
    }
  }
}

function cleanupImportsTemp() {
  try {
    const oneDrive = getOneDriveManager();
    const syncDir = oneDrive.syncDirectory || path.join(os.homedir(), "OneDrive");
    const importsDir = path.join(syncDir, "Microsoft 365 for Linux Imports");
    if (fs.existsSync(importsDir)) {
      fs.rmSync(importsDir, { recursive: true, force: true });
    }
    if (fs.existsSync(mappingsFile)) {
      fs.unlinkSync(mappingsFile);
    }
  } catch (_) {}
}

module.exports = {
  importExternalFile,
  registerExternalFile,
  getExternalMapping,
  cleanDocTitle,
  saveLocalNow,
  syncExternalCopiesBack,
  cleanupImportsTemp,
  setStatusReporter
};

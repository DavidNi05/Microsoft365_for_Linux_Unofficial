const { spawn, exec } = require("child_process");
const { Tray, Menu, nativeImage, app, shell } = require("electron");
const fs = require("fs");
const path = require("path");
const os = require("os");

const ONEDRIVE_BIN = "/app/bin/onedrive";
const M365_BIN = "/app/bin/m365-linux";
const TRAY_ICON = "/app/share/microsoft365/assets/m365-tray.png";
const FOLDER_ICON_SOURCE = "/app/share/microsoft365/assets/onedrive-folder.png";
const CLIENT_ID = "d50ca740-c83f-4d1b-b616-12c519384f0c";

function getExternalSync() {
  return require("./external-file-sync");
}

const homeDirectory = os.homedir();
const configBase = process.env.XDG_CONFIG_HOME || path.join(homeDirectory, ".config");
const configDirectory = path.join(configBase, "onedrive");
const syncDirectory = path.join(homeDirectory, "OneDrive");

const refreshTokenFile = path.join(configDirectory, "refresh_token");
const heartbeatFile = path.join(configDirectory, "m365-background-heartbeat");
const lastSyncFile = path.join(configDirectory, "m365-last-sync");
const quotaCacheFile = path.join(configDirectory, "m365-quota-cache.json");

let monitorProcess = null;
let syncProcess = null;
let heartbeatTimer = null;
let tray = null;
let trayCurrentStatus = "OneDrive Sync: Active";
let memoryQuotaCache = null;
let webScrapedQuota = null;

let cachedAccessToken = null;
let tokenExpiresAt = 0;
let cachedAvatarData = null;

let uiOpenHandler = null;
let soundStateGetter = () => true;
let soundStateSetter = () => {};

function isAuthenticated() {
  try {
    if (!fs.existsSync(refreshTokenFile)) return false;
    const content = fs.readFileSync(refreshTokenFile, "utf8").trim();
    return content.length > 10;
  } catch (_) {
    return false;
  }
}

function getStoredRefreshToken() {
  try {
    if (!fs.existsSync(refreshTokenFile)) return null;
    const raw = fs.readFileSync(refreshTokenFile, "utf8").trim();
    if (raw.startsWith("{")) {
      const parsed = JSON.parse(raw);
      return parsed.refresh_token || raw;
    }
    return raw;
  } catch (_) {
    return null;
  }
}

function saveRefreshToken(newToken) {
  if (!newToken) return;
  try {
    ensureDirectories();
    fs.writeFileSync(refreshTokenFile, newToken.trim(), { encoding: "utf8", mode: 0o600 });
  } catch (_) {}
}

async function getGraphAccessToken() {
  const now = Date.now();
  if (cachedAccessToken && now < tokenExpiresAt - 60000) {
    return cachedAccessToken;
  }

  const rToken = getStoredRefreshToken();
  if (!rToken) return null;

  try {
    const params = new URLSearchParams();
    params.append("client_id", CLIENT_ID);
    params.append("grant_type", "refresh_token");
    params.append("refresh_token", rToken);

    const res = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString()
    });

    if (!res.ok) return null;
    const data = await res.json();

    if (data.access_token) {
      cachedAccessToken = data.access_token;
      tokenExpiresAt = Date.now() + (Number(data.expires_in) || 3600) * 1000;
      if (data.refresh_token && data.refresh_token !== rToken) {
        saveRefreshToken(data.refresh_token);
      }
      return cachedAccessToken;
    }
  } catch (err) {
    console.error("[Graph API] Error refreshing access token:", err.message);
  }
  return null;
}

async function fetchGraphUserProfilePicture() {
  if (cachedAvatarData) return cachedAvatarData;
  const token = await getGraphAccessToken();
  if (!token) return null;

  try {
    const res = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", {
      headers: { "Authorization": "Bearer " + token }
    });
    if (res.ok) {
      const buffer = await res.arrayBuffer();
      const base64 = Buffer.from(buffer).toString("base64");
      const contentType = res.headers.get("content-type") || "image/jpeg";
      cachedAvatarData = "data:" + contentType + ";base64," + base64;
      return cachedAvatarData;
    }
  } catch (_) {}
  return null;
}

async function fetchGraphQuota() {
  const token = await getGraphAccessToken();
  if (!token) return null;

  try {
    const res = await fetch("https://graph.microsoft.com/v1.0/me/drive", {
      headers: { "Authorization": "Bearer " + token }
    });
    if (!res.ok) return null;

    const data = await res.json();
    if (data && data.quota) {
      const totalBytes = data.quota.total || 0;
      const usedBytes = data.quota.used || 0;
      const remainingBytes = data.quota.remaining || Math.max(0, totalBytes - usedBytes);
      
      const formatGB = (b) => (b / (1024 ** 3)).toFixed(1).replace(".", ",") + " GB";
      const formatTB = (b) => (b / (1024 ** 4)).toFixed(1).replace(".", ",") + " TB";

      const totalFormatted = totalBytes >= (1024 ** 4) ? formatTB(totalBytes) : formatGB(totalBytes);
      const usedFormatted = formatGB(usedBytes);
      const remainingFormatted = remainingBytes >= (1024 ** 4) ? formatTB(remainingBytes) : formatGB(remainingBytes);
      const percent = totalBytes > 0 ? Math.round((usedBytes / totalBytes) * 100) : 0;

      const quotaResult = {
        total: totalFormatted,
        used: usedFormatted,
        remaining: remainingFormatted,
        free: remainingFormatted,
        percentUsed: percent,
        percent: percent,
        timestamp: Date.now()
      };

      saveDiskQuotaCache(quotaResult);
      return quotaResult;
    }
  } catch (err) {
    console.warn("[Graph API] Error querying drive quota:", err.message);
  }
  return null;
}

async function fetchGraphFileWebUrl(relativePath) {
  const token = await getGraphAccessToken();
  if (!token || !relativePath) return null;

  try {
    const normalized = relativePath.split(path.sep).join("/");
    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(normalized)}`;
    const res = await fetch(endpoint, {
      headers: { "Authorization": "Bearer " + token }
    });
    if (res.ok) {
      const item = await res.json();
      if (item && item.webUrl) return item.webUrl;
    }
  } catch (_) {}
  return null;
}

function setOpenUiHandler(handler) {
  if (typeof handler === "function") uiOpenHandler = handler;
}

function setSoundHandlers(getter, setter) {
  if (typeof getter === "function") soundStateGetter = getter;
  if (typeof setter === "function") soundStateSetter = setter;
}

function setupFolderIcon() {
  try {
    const iconDestination = path.join(syncDirectory, ".folder-icon.png");
    const oldSvg = path.join(syncDirectory, ".folder-icon.svg");
    if (fs.existsSync(oldSvg)) {
      try { fs.unlinkSync(oldSvg); } catch (_) {}
    }
    if (fs.existsSync(FOLDER_ICON_SOURCE) && !fs.existsSync(iconDestination)) {
      fs.copyFileSync(FOLDER_ICON_SOURCE, iconDestination);
    }
    if (fs.existsSync(iconDestination)) {
      const fileUri = `file://${iconDestination}`;
      exec(`/usr/bin/gio set "${syncDirectory}" metadata::custom-icon "${fileUri}"`, () => {});
    }
  } catch (_) {}
}

function ensureDirectories() {
  fs.mkdirSync(configDirectory, { recursive: true, mode: 0o700 });
  fs.mkdirSync(syncDirectory, { recursive: true, mode: 0o700 });
  setupFolderIcon();
}

function commonArgs() {
  return ["--confdir", configDirectory, "--syncdir", syncDirectory];
}

function processEnvironment() {
  return { ...process.env, HOME: homeDirectory };
}

function spawnM365(args = []) {
  try {
    const child = spawn(M365_BIN, args, { detached: true, stdio: "ignore", env: processEnvironment() });
    child.unref();
    return true;
  } catch (_) {
    return false;
  }
}

function openMicrosoft365() { 
  if (typeof uiOpenHandler === "function") {
    uiOpenHandler();
    return true;
  }
  return spawnM365([]); 
}

function requestInterfaceQuit() { return spawnM365(["--quit-ui"]); }

async function openOneDriveFolder() {
  ensureDirectories();
  await shell.openPath(syncDirectory);
}

function rebuildTrayMenu() {
  if (!tray || tray.isDestroyed()) return;
  const menuTemplate = [
    { label: "Office 365 Suite (Community)", enabled: false },
    { label: trayCurrentStatus, enabled: false },
    { type: "separator" },
    { 
      label: "Notification Sound", 
      type: "checkbox", 
      checked: soundStateGetter(), 
      click: (menuItem) => soundStateSetter(menuItem.checked)
    },
    { type: "separator" },
    { label: "Open Office 365 Suite", click: () => openMicrosoft365() },
    { label: "Open Local OneDrive", click: () => openOneDriveFolder() },
    { type: "separator" },
    { label: "Quit Office 365 Suite", click: () => { requestInterfaceQuit(); app.quit(); } }
  ];
  tray.setContextMenu(Menu.buildFromTemplate(menuTemplate));
}

function updateTrayStatus(status) {
  if (!tray || tray.isDestroyed()) return;
  trayCurrentStatus = status === "error" ? "OneDrive Sync: Error" : "OneDrive Sync: Active";
  tray.setToolTip("Office 365 — OneDrive Active");
  rebuildTrayMenu();
}

function createBackgroundTray() {
  if (tray && !tray.isDestroyed()) return tray;
  if (!fs.existsSync(TRAY_ICON)) return null;
  let image = nativeImage.createFromPath(TRAY_ICON);
  if (image.isEmpty()) return null;
  image = image.resize({ width: 16, height: 16, quality: "best" });
  tray = new Tray(image);
  const extSync = getExternalSync();
  if (typeof extSync.setStatusReporter === "function") {
    extSync.setStatusReporter(updateTrayStatus);
  }
  tray.setToolTip("Office 365 — OneDrive Active");
  rebuildTrayMenu();
  tray.on("click", () => openMicrosoft365());
  tray.on("double-click", () => openMicrosoft365());
  return tray;
}

function destroyBackgroundTray() {
  const extSync = getExternalSync();
  if (typeof extSync.setStatusReporter === "function") {
    extSync.setStatusReporter(null);
  }
  if (tray && !tray.isDestroyed()) { 
    tray.destroy(); 
    tray = null; 
  }
}

function writeHeartbeat() {
  if (!isAuthenticated()) return;
  ensureDirectories();
  fs.writeFileSync(heartbeatFile, String(Date.now()), { encoding: "utf8", mode: 0o600 });
}

function startHeartbeat() {
  if (!isAuthenticated()) return;
  writeHeartbeat();
  if (heartbeatTimer) clearInterval(heartbeatTimer);
  heartbeatTimer = setInterval(writeHeartbeat, 5000);
  createBackgroundTray();
}

function stopHeartbeat() {
  if (heartbeatTimer) { clearInterval(heartbeatTimer); heartbeatTimer = null; }
  try { fs.rmSync(heartbeatFile, { force: true }); } catch (_) {}
}

function backgroundAlive() {
  try {
    if (!fs.existsSync(heartbeatFile)) return false;
    const timestamp = Number(fs.readFileSync(heartbeatFile, "utf8").trim());
    return (Date.now() - timestamp) < 15000;
  } catch (_) { return false; }
}

function recordSuccessfulSync() {
  ensureDirectories();
  fs.writeFileSync(lastSyncFile, new Date().toISOString(), { encoding: "utf8", mode: 0o600 });
}

function getLastSync() {
  try {
    if (!fs.existsSync(lastSyncFile)) return null;
    return fs.readFileSync(lastSyncFile, "utf8").trim();
  } catch (_) { return null; }
}

function triggerImmediateSync() {
  if (!isAuthenticated()) return false;
  if (monitorProcess && monitorProcess.pid && monitorProcess.exitCode === null) {
    return true;
  }
  if (!syncProcess) {
    runInitialSync(() => {}).catch(() => {});
  }
  return true;
}

function authenticateWithResponseUrl(responseUrl) {
  ensureDirectories();
  return new Promise((resolve) => {
    const child = spawn(ONEDRIVE_BIN, [
      ...commonArgs(),
      "--auth-response",
      responseUrl
    ], { env: processEnvironment() });

    let output = "";
    child.stdout.on("data", d => { output += d.toString(); });
    child.stderr.on("data", d => { output += d.toString(); });

    child.on("close", async (code) => {
      if (code === 0 || isAuthenticated()) {
        recordSuccessfulSync();
        await getGraphAccessToken();
        await fetchGraphUserProfilePicture();
        await fetchGraphQuota();
        resolve({ ok: true, output });
      } else {
        resolve({ ok: false, output, code });
      }
    });

    child.on("error", (err) => {
      resolve({ ok: false, error: err.message });
    });
  });
}

function runInitialSync(onOutput) {
  if (!isAuthenticated()) {
    return Promise.resolve({ code: -1, output: "Not authenticated" });
  }
  ensureDirectories();
  return new Promise((resolve, reject) => {
    syncProcess = spawn(ONEDRIVE_BIN, [...commonArgs(), "--sync"], { env: processEnvironment() });
    let completeOutput = "";
    syncProcess.stdout.on("data", d => { 
      completeOutput += d.toString(); 
      if (onOutput) onOutput(d.toString()); 
    });
    syncProcess.stderr.on("data", d => { 
      completeOutput += d.toString(); 
      if (onOutput) onOutput(d.toString()); 
    });
    syncProcess.on("close", code => { 
      syncProcess = null; 
      if (code === 0) recordSuccessfulSync(); 
      resolve({ code, output: completeOutput }); 
    });
    syncProcess.on("error", err => { 
      syncProcess = null; 
      reject(err); 
    });
  });
}

function startMonitor(onOutput) {
  if (!isAuthenticated()) return null;
  ensureDirectories();
  if (monitorProcess && monitorProcess.exitCode === null) return monitorProcess;

  monitorProcess = spawn(ONEDRIVE_BIN, [...commonArgs(), "--monitor"], { env: processEnvironment() });

  monitorProcess.stdout.on("data", data => {
    const text = data.toString();
    if (text.includes("Sync with Microsoft OneDrive is complete")) {
      recordSuccessfulSync();
      try {
        const extSync = getExternalSync();
        if (typeof extSync.syncExternalCopiesBack === "function") {
          extSync.syncExternalCopiesBack();
        }
      } catch (_) {}
    }
    if (onOutput) onOutput(text);
  });

  monitorProcess.on("close", () => {
    monitorProcess = null;
  });

  return monitorProcess;
}

function stopMonitor() {
  if (monitorProcess && monitorProcess.exitCode === null) monitorProcess.kill("SIGTERM");
  monitorProcess = null;
}

function monitorRunning() {
  return Boolean(monitorProcess && monitorProcess.exitCode === null);
}

function stopSync() {
  if (syncProcess && syncProcess.exitCode === null) syncProcess.kill("SIGTERM");
  syncProcess = null;
}

function loadDiskQuotaCache() {
  if (memoryQuotaCache) return memoryQuotaCache;
  try {
    if (fs.existsSync(quotaCacheFile)) {
      const data = JSON.parse(fs.readFileSync(quotaCacheFile, "utf8"));
      if (data && data.total && data.used) {
        memoryQuotaCache = data;
        return memoryQuotaCache;
      }
    }
  } catch (_) {}
  return null;
}

function saveDiskQuotaCache(quotaData) {
  if (!quotaData) return;
  memoryQuotaCache = quotaData;
  try {
    ensureDirectories();
    fs.writeFileSync(quotaCacheFile, JSON.stringify(quotaData, null, 2), { encoding: "utf8", mode: 0o600 });
  } catch (_) {}
}

async function getQuota() {
  if (webScrapedQuota) return webScrapedQuota;
  const graphQ = await fetchGraphQuota();
  if (graphQ) return graphQ;
  return loadDiskQuotaCache() || {
    total: "1 TB",
    used: "0 GB",
    remaining: "1 TB",
    free: "1 TB",
    percentUsed: 0,
    percent: 0,
    timestamp: Date.now()
  };
}

async function getStatus() {
  ensureDirectories();
  if (!isAuthenticated()) {
    return {
      connected: false,
      background: false,
      monitor: false,
      lastSync: null,
      syncDirectory,
      syncStatus: "Not Authenticated",
      quota: null
    };
  }

  return {
    connected: true,
    background: backgroundAlive(),
    monitor: monitorRunning(),
    lastSync: getLastSync(),
    syncDirectory,
    syncStatus: monitorRunning() ? "Active (Monitor)" : "Active",
    quota: await getQuota()
  };
}

async function syncNow(onOutput) {
  if (!isAuthenticated()) {
    return { code: -1, skipped: true, output: "Account not authenticated." };
  }
  triggerImmediateSync();
  if (backgroundAlive()) {
    return { code: 0, skipped: true, output: "Monitor active, inotify handling changes." };
  }
  return runInitialSync(onOutput);
}

module.exports = {
  CLIENT_ID,
  configDirectory,
  syncDirectory,
  ensureDirectories,
  isAuthenticated,
  getGraphAccessToken,
  fetchGraphUserProfilePicture,
  fetchGraphQuota,
  fetchGraphFileWebUrl,
  authenticateWithResponseUrl,
  runInitialSync,
  startMonitor,
  stopMonitor,
  monitorRunning,
  stopSync,
  startHeartbeat,
  stopHeartbeat,
  backgroundAlive,
  createBackgroundTray,
  destroyBackgroundTray,
  rebuildTrayMenu,
  setOpenUiHandler,
  setSoundHandlers,
  openMicrosoft365,
  requestInterfaceQuit,
  updateTrayStatus,
  getQuota,
  getStatus,
  syncNow,
  triggerImmediateSync
};

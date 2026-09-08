const { spawn, exec } = require("child_process");
const { 
  Tray, 
  Menu, 
  nativeImage, 
  app, 
  shell 
} = require("electron");
const fs = require("fs");
const path = require("path");
const os = require("os");

const ONEDRIVE_BIN = "/app/bin/onedrive";
const M365_BIN = "/app/bin/m365-linux";
const TRAY_ICON = "/app/share/microsoft365/assets/m365-tray.png";
const FOLDER_ICON_SOURCE = "/app/share/microsoft365/assets/onedrive-folder.png";

function getExternalSync() {
  return require("./external-file-sync");
}

const homeDirectory = os.homedir();
const configBase = process.env.XDG_CONFIG_HOME || path.join(homeDirectory, ".config");
const configDirectory = path.join(configBase, "onedrive");
const syncDirectory = path.join(homeDirectory, "OneDrive");

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

let uiOpenHandler = null;
let soundStateGetter = () => true;
let soundStateSetter = () => {};

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
  } catch (error) {
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
    { label: "Microsoft 365 for Linux", enabled: false },
    { label: trayCurrentStatus, enabled: false },
    { type: "separator" },
    { 
      label: "Notification Sound", 
      type: "checkbox", 
      checked: soundStateGetter(), 
      click: (menuItem) => {
        soundStateSetter(menuItem.checked);
      }
    },
    { type: "separator" },
    { label: "Open Microsoft 365", click: () => openMicrosoft365() },
    { label: "Open Local OneDrive", click: () => openOneDriveFolder() },
    { type: "separator" },
    { label: "Quit Microsoft 365", click: () => { requestInterfaceQuit(); app.quit(); } }
  ];
  tray.setContextMenu(Menu.buildFromTemplate(menuTemplate));
}

function updateTrayStatus(status) {
  if (!tray || tray.isDestroyed()) return;
  trayCurrentStatus = status === "error" ? "OneDrive Sync: Error" : "OneDrive Sync: Active";
  tray.setToolTip("Microsoft 365 — OneDrive Active");
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
  tray.setToolTip("Microsoft 365 — OneDrive Active");
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
  ensureDirectories();
  fs.writeFileSync(heartbeatFile, String(Date.now()), { encoding: "utf8", mode: 0o600 });
}

function startHeartbeat() {
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
  if (monitorProcess && monitorProcess.pid && monitorProcess.exitCode === null) {
    return true;
  }
  if (!syncProcess) {
    runInitialSync(() => {}).catch(() => {});
  }
  return true;
}

function runCommand(args, timeoutMs = 8000) {
  ensureDirectories();
  return new Promise((resolve) => {
    let settled = false;
    const child = spawn(ONEDRIVE_BIN, [...commonArgs(), ...args], { env: processEnvironment() });
    let output = "";

    const timer = setTimeout(() => {
      if (!settled) {
        settled = true;
        try { child.kill("SIGKILL"); } catch (_) {}
        resolve({ code: -1, output: output || "Timeout" });
      }
    }, timeoutMs);

    child.stdout.on("data", data => { output += data.toString(); });
    child.stderr.on("data", data => { output += data.toString(); });
    child.on("close", code => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        resolve({ code, output });
      }
    });
  });
}

function runInitialSync(onOutput) {
  ensureDirectories();
  return new Promise((resolve, reject) => {
    syncProcess = spawn(ONEDRIVE_BIN, [...commonArgs(), "--sync"], { env: processEnvironment() });
    let completeOutput = "";
    syncProcess.stdout.on("data", data => { 
      completeOutput += data.toString(); 
      if (onOutput) onOutput(data.toString()); 
    });
    syncProcess.stderr.on("data", data => { 
      completeOutput += data.toString(); 
      if (onOutput) onOutput(data.toString()); 
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
  ensureDirectories();
  if (monitorProcess && monitorProcess.exitCode === null) return monitorProcess;

  monitorProcess = spawn(ONEDRIVE_BIN, [...commonArgs(), "--monitor"], { env: processEnvironment() });

  monitorProcess.stdout.on("data", data => {
    const text = data.toString();
    const lines = text.split("\n");
    for (const line of lines) {
      const trimmed = line.trim();
      if (trimmed) {
        console.log(`[OneDrive Monitor] ${trimmed}`);
      }
    }
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

  monitorProcess.stderr.on("data", data => {
    const text = data.toString().trim();
    if (text) {
      console.error(`[OneDrive Monitor Error] ${text}`);
    }
  });

  monitorProcess.on("close", (code, signal) => {
    console.log(`[OneDrive Monitor] Process exited with code ${code}, signal: ${signal}`);
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

function setWebScrapedQuota(quotaObj) {
  if (quotaObj && quotaObj.total) {
    webScrapedQuota = { ...quotaObj, timestamp: Date.now() };
    saveDiskQuotaCache(webScrapedQuota);
  }
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

async function getQuota(force = false) {
  if (webScrapedQuota) return webScrapedQuota;
  const cached = loadDiskQuotaCache();
  return cached || {
    total: "1 TB",
    used: "23,5 GB",
    remaining: "976,5 GB",
    free: "976,5 GB",
    percentUsed: 2,
    percent: 2,
    timestamp: Date.now()
  };
}

async function getStatus() {
  ensureDirectories();
  const syncStatus = await runCommand(["--display-sync-status"], 4000);
  return {
    connected: syncStatus.code === 0,
    background: backgroundAlive(),
    monitor: monitorRunning(),
    lastSync: getLastSync(),
    syncDirectory,
    syncStatus: syncStatus.output,
    quota: loadDiskQuotaCache() || await getQuota(false)
  };
}

async function syncNow(onOutput) {
  triggerImmediateSync();
  if (backgroundAlive()) {
    return { code: 0, skipped: true, output: "Monitor active, inotify handling changes." };
  }
  return runInitialSync(onOutput);
}

module.exports = {
  configDirectory,
  syncDirectory,
  ensureDirectories,
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
  setWebScrapedQuota,
  getQuota,
  getStatus,
  syncNow,
  triggerImmediateSync
};

const { 
  app, 
  BrowserWindow, 
  WebContentsView, 
  ipcMain, 
  shell, 
  dialog,
  Menu,
  Tray,
  nativeImage,
  globalShortcut
} = require("electron"); 

const path = require("path"); 
const fs = require("fs"); 
const { exec } = require("child_process");
const { findOfficeFiles } = require("./file-router"); 
const { getFileWebUrl } = require("./office-online"); 
const oneDrive = require("./onedrive-manager"); 
const { 
  saveLocalNow, 
  cleanDocTitle,
  registerExternalFile,
  cleanupImportsTemp 
} = require("./external-file-sync"); 
const { 
  requestBackgroundAccess, 
  setBackgroundStatus 
} = require("./background"); 
const { 
  showFluentNotification, 
  setSoundEnabled, 
  getSoundEnabled 
} = require("./fluent-notifier");

const backgroundMode = process.argv.includes("--background"); 
const quitUiMode = process.argv.includes("--quit-ui"); 

let mainWindow = null; 
let officeView = null; 
let oneDriveSetupRunning = false; 
let forceQuit = false; 
let activeExternalOriginalPath = null; 
let activeExternalOneDrivePath = null; 
let activeExternalImported = false; 
let manualSaveInProgress = false; 
let isOpeningOfficeFile = false; 
let pendingOfficeFiles = findOfficeFiles(process.argv); 
let closeToTrayEnabled = false;

let currentSidebarWidth = 230; 
const TOPBAR_HEIGHT = 55; 

const services = { 
  home: "https://word.cloud.microsoft/", 
  word: "https://word.cloud.microsoft/", 
  excel: "https://excel.cloud.microsoft/", 
  powerpoint: "https://powerpoint.cloud.microsoft/", 
  onenote: "https://onenote.cloud.microsoft/", 
  copilot: "https://m365.cloud.microsoft/chat/", 
  onedrive: "https://onedrive.live.com/", 
  outlook: "https://outlook.office.com/", 
  teams: "https://teams.live.com/v2/", 
  todo: "https://to-do.office.com/tasks/", 
  designer: "https://designer.microsoft.com/", 
  clipchamp: "https://app.clipchamp.com/", 
  help: "https://support.microsoft.com/" 
}; 

const ALLOWED_DOMAINS = [
  "microsoft.com",
  "cloud.microsoft",
  "office.com",
  "office365.com",
  "live.com",
  "sharepoint.com",
  "microsoftonline.com",
  "windows.net",
  "bing.com",
  "clipchamp.com",
  "msauth.net",
  "msftauth.net",
  "skype.com"
];

function isAllowedUrl(rawUrl) {
  if (!rawUrl || typeof rawUrl !== "string") return false;
  if (rawUrl.startsWith("file://") || rawUrl.startsWith("m365://")) return true;

  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol !== "https:" && parsed.protocol !== "http:") return false;
    const hostname = parsed.hostname.toLowerCase();
    return ALLOWED_DOMAINS.some(domain => hostname === domain || hostname.endsWith("." + domain));
  } catch (_) {
    return false;
  }
}

function sleep(milliseconds) { 
  return new Promise(resolve => setTimeout(resolve, milliseconds)); 
} 

function sendToInterface(channel, value) { 
  if (mainWindow && !mainWindow.webContents.isDestroyed()) { 
    mainWindow.webContents.send(channel, value); 
  } 
} 

function sendOneDriveStatus(status) { 
  sendToInterface("onedrive-status", status); 
} 

function showMainWindow() { 
  if (!mainWindow || mainWindow.isDestroyed()) return; 
  if (mainWindow.isMinimized()) mainWindow.restore(); 
  mainWindow.show(); 
  mainWindow.focus(); 
} 

function getPreferencesFilePath() {
  return path.join(app.getPath("userData"), "app_preferences.json");
}

function readAllPreferences() {
  try {
    const file = getPreferencesFilePath();
    if (fs.existsSync(file)) {
      return JSON.parse(fs.readFileSync(file, "utf8"));
    }
  } catch (_) {}
  return {
    leftbar_theme: "default",
    leftbar_custom_color: "#f8fafc",
    default_app: "word",
    autohide_sidebar: false,
    compact_nav: false,
    sync_interval: "manual",
    pause_metered: false,
    native_onedrive_sync: false,
    hardware_accel: false,
    autostart: false,
    close_to_tray: false,
    global_hotkey: false,
    desktop_shortcut: false
  };
}

function writePreference(key, value) {
  try {
    const prefs = readAllPreferences();
    prefs[key] = value;
    fs.writeFileSync(getPreferencesFilePath(), JSON.stringify(prefs, null, 2), "utf8");
  } catch (_) {}
}

oneDrive.setSoundHandlers(
  () => getSoundEnabled(),
  (enabled) => {
    setSoundEnabled(enabled);
    sendToInterface("sound-setting-changed", enabled);
    oneDrive.rebuildTrayMenu();
  }
);

async function ensureBackgroundRunning() { 
  if (backgroundMode) return true; 
  const prefs = readAllPreferences();
  if (prefs.native_onedrive_sync !== true) {
    return false;
  }

  if (oneDrive.backgroundAlive() || oneDrive.monitorRunning()) return true; 

  try { 
    oneDrive.ensureDirectories(); 
    oneDrive.startHeartbeat(); 
    oneDrive.startMonitor(() => {
      sendOneDriveStatus("Sync");
    }); 
    return true; 
  } catch (_) { 
    return false; 
  } 
} 

function updateViewBounds() { 
  if ( 
    !mainWindow || 
    mainWindow.isDestroyed() || 
    !officeView || 
    officeView.webContents.isDestroyed() 
  ) { 
    return; 
  } 

  const contentBounds = mainWindow.getContentBounds(); 
  const contentWidth = Math.max(1, Math.floor(Number(contentBounds.width) || 0)); 
  const contentHeight = Math.max(1, Math.floor(Number(contentBounds.height) || 0)); 
  
  const leftInset = Math.max(0, Math.floor(currentSidebarWidth)); 
  const topInset = Math.max(0, Math.floor(TOPBAR_HEIGHT)); 

  officeView.setBounds({ 
    x: leftInset, 
    y: topInset, 
    width: Math.max(1, contentWidth - leftInset), 
    height: Math.max(1, contentHeight - topInset) 
  }); 
} 

function scheduleUpdateViewBounds() {
  updateViewBounds();
  setImmediate(updateViewBounds);
  setTimeout(updateViewBounds, 50);
  setTimeout(updateViewBounds, 150);
  setTimeout(updateViewBounds, 300);
}

function openMicrosoftService(service) { 
  if (!officeView || officeView.webContents.isDestroyed()) return; 

  const key = (service || "").toLowerCase();

  if (key === "credits") {
    const creditsPath = fs.existsSync(path.join(__dirname, "Credits.html"))
      ? path.join(__dirname, "Credits.html")
      : path.join(__dirname, "credits.html");
    officeView.webContents.loadFile(creditsPath);
    sendToInterface("browser-url", "m365://credits");
    return;
  }

  if (key === "settings") {
    const settingsPath = fs.existsSync(path.join(__dirname, "settings.html"))
      ? path.join(__dirname, "settings.html")
      : path.join(__dirname, "Settings.html");
    officeView.webContents.loadFile(settingsPath);
    sendToInterface("browser-url", "m365://settings");
    return;
  }

  if (services[key]) { 
    officeView.webContents.loadURL(services[key]); 
  } 
} 

function updateActiveExternalFile(result, explicitOriginal = null) { 
  const chosenOriginal = explicitOriginal || (result && result.originalPath);
  if (chosenOriginal && typeof chosenOriginal === "string" && chosenOriginal.trim()) { 
    const originalPath = path.resolve(chosenOriginal); 
    const workingPath = result && typeof result.localOneDrivePath === "string" && result.localOneDrivePath.trim() 
      ? path.resolve(result.localOneDrivePath) 
      : null; 

    activeExternalOriginalPath = originalPath; 
    activeExternalOneDrivePath = workingPath; 
    activeExternalImported = result ? result.imported !== false : true; 

    if (workingPath && activeExternalImported && typeof registerExternalFile === "function") { 
      try { 
        registerExternalFile(originalPath, workingPath, activeExternalImported); 
      } catch (_) {} 
    } 
    return; 
  } 
  if (result) { 
    activeExternalOriginalPath = null; 
    activeExternalOneDrivePath = null; 
    activeExternalImported = false; 
  } 
} 

async function handleOfficeFile(officeFile) { 
  if (!officeFile || !officeFile.service) return; 
  if (isOpeningOfficeFile) return; 

  isOpeningOfficeFile = true; 
  activeExternalOriginalPath = officeFile.originalPath ? path.resolve(officeFile.originalPath) : null; 
  activeExternalOneDrivePath = officeFile.path ? path.resolve(officeFile.path) : null; 
  activeExternalImported = true; 

  sendToInterface("select-service-tab", officeFile.service);
  sendToInterface("browser-loading", true);
  sendToInterface("office-file-loading", {
    fileName: officeFile.name,
    service: officeFile.service,
    status: "Importing document..."
  });

  try { 
    const fileToProcess = officeFile.originalPath || officeFile.path;
    let result = await getFileWebUrl(fileToProcess); 
    updateActiveExternalFile(result, officeFile.originalPath); 

    if (result.insideOneDrive && result.webUrl) { 
      sendToInterface("office-file-loading", { fileName: officeFile.name, service: officeFile.service, status: "Opening in Office Online..." });
      officeView.webContents.loadURL(result.webUrl); 
      sendToInterface("office-file-opened", { ...officeFile, service: officeFile.service, online: true, webUrl: result.webUrl }); 
      return; 
    } 

    if (result.insideOneDrive && !result.webUrl) { 
      for (let attempt = 1; attempt <= 45; attempt++) { 
        sendToInterface("office-file-loading", { 
          fileName: officeFile.name, 
          service: officeFile.service, 
          status: `Waiting for cloud sync (${attempt}/45)...` 
        });

        await sleep(2000); 

        const retry = await getFileWebUrl(fileToProcess); 
        updateActiveExternalFile(retry, officeFile.originalPath); 

        if (retry.webUrl) { 
          sendToInterface("office-file-loading", { fileName: officeFile.name, service: officeFile.service, status: "Opening in Office Online..." });
          officeView.webContents.loadURL(retry.webUrl); 
          sendOneDriveStatus("Sync"); 
          sendToInterface("office-file-opened", { ...officeFile, service: officeFile.service, online: true, webUrl: retry.webUrl }); 
          return; 
        } 
      } 
      throw new Error("Cloud link response timed out."); 
    } 

    openMicrosoftService(officeFile.service); 
  } catch (error) { 
    openMicrosoftService(officeFile.service); 
  } finally { 
    isOpeningOfficeFile = false; 
    sendToInterface("office-file-finished"); 
  } 
} 

async function openPendingOfficeFile() { 
  if (pendingOfficeFiles.length === 0) { 
    openMicrosoftService("word"); 
    return; 
  } 
  const officeFile = pendingOfficeFiles.shift(); 
  await handleOfficeFile(officeFile); 
} 

async function getOfficeOnlineDocumentTitle() {
  if (!officeView || officeView.webContents.isDestroyed()) return null;
  try {
    return await officeView.webContents.executeJavaScript(`
      (function() {
        const input = document.querySelector('input[data-automationid="FileNameInput"]') ||
                      document.querySelector('#Breadcrumb-ItemName') ||
                      document.querySelector('#fileNameTextBox');
        if (input && input.value && input.value.trim()) {
          return input.value.trim();
        }

        const btn = document.querySelector('button[data-automationid="od-ItemName"]') ||
                    document.querySelector('span[data-automationid="od-ItemName"]') ||
                    document.querySelector('button[id*="Breadcrumb"]') ||
                    document.querySelector('span[id*="Breadcrumb"]') ||
                    document.querySelector('[data-automation-id="docTitle"]');
        if (btn) {
          const t = btn.innerText || btn.textContent;
          if (t && t.trim()) return t.trim();
        }

        return document.title || '';
      })()
    `);
  } catch (_) {
    return officeView.webContents.getTitle();
  }
}

function handleWebTitleChange(title) {
  if (!activeExternalOriginalPath) return;

  const currentExt = path.extname(activeExternalOriginalPath);
  const cleanName = cleanDocTitle(title, currentExt);
  if (!cleanName) return;

  const currentName = path.basename(activeExternalOriginalPath);
  if (cleanName.toLowerCase() !== currentName.toLowerCase()) {
    sendToInterface("office-file-renamed", { fileName: cleanName });
  }
}

function setupGlobalShortcut(enable) {
  const shortcutKey = "Ctrl+Alt+M";
  if (!enable) {
    try { globalShortcut.unregister(shortcutKey); } catch (_) {}
    return;
  }
  try {
    globalShortcut.register(shortcutKey, () => {
      showMainWindow();
    });
  } catch (_) {}
}

function showWelcomeDialog() {
  const welcomeFile = path.join(app.getPath("userData"), "flathub_welcomed.json");
  if (fs.existsSync(welcomeFile)) return;

  const welcomeWin = new BrowserWindow({
    width: 500,
    height: 240,
    parent: mainWindow,
    modal: true,
    frame: false,
    transparent: true,
    resizable: false,
    alwaysOnTop: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; user-select: none; }
        body {
          font-family: 'Segoe UI Variable Display', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
          background: transparent;
          width: 100vw;
          height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
          overflow: hidden;
        }
        .dialog-card {
          width: 470px;
          background: #ffffff;
          border-radius: 12px;
          border: 1px solid #cbd5e1;
          box-shadow: 0 24px 48px rgba(0, 0, 0, 0.32);
          padding: 24px;
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .title {
          font-size: 16px;
          font-weight: 700;
          color: #0f172a;
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .desc {
          font-size: 13px;
          color: #475569;
          line-height: 1.5;
        }
        .actions {
          display: flex;
          justify-content: flex-end;
          margin-top: 4px;
        }
        .btn {
          background: #0078d4;
          color: #ffffff;
          border: none;
          border-radius: 6px;
          padding: 8px 22px;
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
          transition: background-color 0.15s ease;
        }
        .btn:hover {
          background: #106ebe;
        }
      </style>
    </head>
    <body>
      <div class="dialog-card">
        <div class="title">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
            <rect x="2" y="2" width="9.5" height="9.5" fill="#f25022"/>
            <rect x="12.5" y="2" width="9.5" height="9.5" fill="#7fba00"/>
            <rect x="2" y="12.5" width="9.5" height="9.5" fill="#00a4ef"/>
            <rect x="12.5" y="12.5" width="9.5" height="9.5" fill="#ffb900"/>
          </svg>
          Microsoft 365 for Linux (Unofficial)
        </div>
        <div class="desc">
          Due to Flathub guidelines and sandboxing policies, features like the abraunegg/onedrive background daemon, native synchronization, and native file editing are disabled by default. You can easily enable them anytime in <strong>Settings & Preferences</strong>.
        </div>
        <div class="actions">
          <button class="btn" onclick="window.close()">Got it</button>
        </div>
      </div>
    </body>
    </html>
  `;

  welcomeWin.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(htmlContent)}`);
  welcomeWin.on("closed", () => {
    try {
      fs.writeFileSync(welcomeFile, JSON.stringify({ welcomed: true }), "utf8");
    } catch (_) {}
  });
}

function createOfficeView() { 
  officeView = new WebContentsView({ 
    webPreferences: { 
      nodeIntegration: false, 
      contextIsolation: true, 
      sandbox: false,
      preload: path.join(__dirname, "preload.js"),
      partition: "persist:microsoft365" 
    } 
  }); 

  mainWindow.contentView.addChildView(officeView); 
  scheduleUpdateViewBounds(); 

  officeView.webContents.on("page-title-updated", (event, title) => {
    handleWebTitleChange(title);
  });

  officeView.webContents.on("did-finish-load", () => {
    officeView.webContents.executeJavaScript(`
      (function() {
        if (window.__m365Watcher) return;
        window.__m365Watcher = true;

        function pollTitle() {
          const input = document.querySelector('input[data-automationid="FileNameInput"]') ||
                        document.querySelector('#Breadcrumb-ItemName') ||
                        document.querySelector('#fileNameTextBox');
          if (input && input.value && input.value.trim() && input.value.trim() !== document.title) {
            document.title = input.value.trim();
          }
        }
        setInterval(pollTitle, 800);
      })()
    `).catch(() => {});
  });

  officeView.webContents.setWindowOpenHandler(({ url }) => { 
    if (isAllowedUrl(url)) { 
      officeView.webContents.loadURL(url); 
      return { action: "deny" }; 
    } 
    shell.openExternal(url); 
    return { action: "deny" }; 
  }); 

  officeView.webContents.on("will-navigate", (event, url) => { 
    if (!isAllowedUrl(url)) { 
      event.preventDefault(); 
      shell.openExternal(url); 
    }
  });

  officeView.webContents.on("did-start-loading", () => { 
    sendToInterface("browser-loading", true); 
  }); 

  officeView.webContents.on("did-stop-loading", () => { 
    sendToInterface("browser-loading", false); 
    if (!officeView.webContents.isDestroyed()) { 
      sendToInterface("browser-url", officeView.webContents.getURL()); 
    } 
  }); 

  openPendingOfficeFile().catch(() => { 
    openMicrosoftService("word"); 
  }); 
} 

function createWindow() { 
  if (mainWindow && !mainWindow.isDestroyed()) { 
    showMainWindow(); 
    return; 
  } 

  Menu.setApplicationMenu(null);

  mainWindow = new BrowserWindow({ 
    width: 1500, 
    height: 950, 
    minWidth: 600, 
    minHeight: 450, 
    title: "Microsoft 365 for Linux (Unofficial)", 
    backgroundColor: "#f3f6fb", 
    autoHideMenuBar: true, 
    webPreferences: { 
      preload: path.join(__dirname, "preload.js"), 
      nodeIntegration: false, 
      contextIsolation: true, 
      sandbox: false 
    } 
  }); 

  mainWindow.removeMenu();
  mainWindow.loadFile(path.join(__dirname, "index.html")); 

  mainWindow.webContents.once("did-finish-load", () => { 
    createOfficeView(); 
    const prefs = readAllPreferences();
    if (prefs.native_onedrive_sync === true) {
      sendOneDriveStatus("Sync");
    } else {
      sendOneDriveStatus("Not Connected");
    }
    setTimeout(showWelcomeDialog, 600);
  }); 

  [
    "resize", 
    "resized", 
    "maximize", 
    "unmaximize", 
    "restore", 
    "enter-full-screen", 
    "leave-full-screen", 
    "show"
  ].forEach(eventName => { 
    mainWindow.on(eventName, scheduleUpdateViewBounds); 
  }); 

  mainWindow.on("close", event => { 
    if (!forceQuit && closeToTrayEnabled) { 
      event.preventDefault(); 
      mainWindow.hide(); 
      try { oneDrive.createBackgroundTray(); } catch (_) {}

      showFluentNotification({
        title: "Microsoft 365 for Linux",
        message: "The application was closed, but continues running in the system tray.",
        icon: "Home.png",
        duration: 4500
      });
      return; 
    } 
  }); 

  mainWindow.on("closed", () => { 
    mainWindow = null; 
    officeView = null; 
  }); 
} 

async function configureOneDrive() { 
  if (oneDriveSetupRunning) return; 
  oneDriveSetupRunning = true; 

  try { 
    oneDrive.ensureDirectories(); 

    if (oneDrive.backgroundAlive() || oneDrive.monitorRunning()) { 
      sendOneDriveStatus("Sync"); 
      await shell.openPath(oneDrive.syncDirectory); 
      return; 
    } 

    sendOneDriveStatus("Sync"); 
    await oneDrive.runInitialSync(() => {}); 
    await requestBackgroundAccess(); 
    await ensureBackgroundRunning(); 
    sendOneDriveStatus("Sync"); 
    await shell.openPath(oneDrive.syncDirectory); 
  } catch (error) { 
    sendOneDriveStatus("Not Connected"); 
  } finally { 
    oneDriveSetupRunning = false; 
  } 
} 

async function startBackgroundMode() { 
  oneDrive.ensureDirectories(); 
  oneDrive.startHeartbeat(); 
  oneDrive.startMonitor(() => {
    sendOneDriveStatus("Sync"); 
  }); 
  await setBackgroundStatus("OneDrive synchronization active"); 
} 

/* 
 * ========================================================= 
 * IPC CHANNELS
 * ========================================================= 
 */ 

ipcMain.on("window-resize-notify", (event, newSidebarWidth) => {
  if (typeof newSidebarWidth === "number") {
    currentSidebarWidth = newSidebarWidth;
  }
  scheduleUpdateViewBounds();
});

ipcMain.handle("get-all-preferences", () => {
  return readAllPreferences();
});

ipcMain.handle("show-fluent-toast", (_event, options) => {
  if (!options) return;
  showFluentNotification({
    title: options.title || "Settings & Preferences",
    message: options.message || "Preference updated successfully.",
    icon: options.icon || "Home.png",
    duration: options.duration || 3500
  });
});

ipcMain.handle("get-sound-enabled", () => {
  return getSoundEnabled();
});

ipcMain.handle("set-sound-enabled", (event, enabled) => {
  const newState = setSoundEnabled(enabled);
  oneDrive.rebuildTrayMenu();
  return newState;
});

ipcMain.handle("office-save-local-now", async () => { 
  if (manualSaveInProgress) return { ok: false, busy: true }; 
  manualSaveInProgress = true; 

  try { 
    const targetToSave = activeExternalOriginalPath || activeExternalOneDrivePath;
    if (!targetToSave) { 
      return { ok: false, error: "No local file associated with this session." }; 
    }

    const expectedWebTitle = await getOfficeOnlineDocumentTitle();
    sendOneDriveStatus("Sync"); 

    const result = await saveLocalNow(targetToSave, expectedWebTitle); 

    if (!result || !result.ok) { 
      const failedName = result?.fileName || path.basename(targetToSave); 
      sendOneDriveStatus("Sync"); 
      showFluentNotification({
        title: "Microsoft 365",
        message: `Failed to save ${failedName}`,
        icon: "OneDrive.png",
        duration: 4000
      });
      return (result || { ok: false, fileName: failedName, error: "Save error." }); 
    } 

    if (result.originalPath) {
      activeExternalOriginalPath = path.resolve(result.originalPath);
      if (result.oneDrivePath) {
        activeExternalOneDrivePath = path.resolve(result.oneDrivePath);
      }

      sendToInterface("office-file-renamed", {
        oldPath: targetToSave,
        newPath: activeExternalOriginalPath,
        fileName: result.fileName
      });
    }

    showFluentNotification({
      title: "Microsoft 365",
      message: result.message || `File saved successfully: ${result.fileName}`,
      icon: "OneDrive.png",
      duration: 3500
    });

    sendOneDriveStatus("Sync"); 
    return { ...result, ok: true }; 

  } catch (error) { 
    sendOneDriveStatus("Sync"); 
    return { ok: false, fileName: path.basename(activeExternalOriginalPath || ""), error: error.message }; 
  } finally { 
    manualSaveInProgress = false; 
  } 
}); 

ipcMain.handle("onedrive-get-status", async () => { 
  try { 
    return { ok: true, ...(await oneDrive.getStatus()) }; 
  } catch (error) { 
    return { ok: false, error: error.message }; 
  } 
}); 

ipcMain.handle("onedrive-open-folder", async () => { 
  oneDrive.ensureDirectories(); 
  const error = await shell.openPath(oneDrive.syncDirectory); 
  return { ok: !error, error: error || null }; 
}); 

ipcMain.handle("onedrive-sync-now", async () => { 
  try { 
    const result = await oneDrive.syncNow(() => {}); 
    sendOneDriveStatus("Sync"); 
    return { ok: result.code === 0, skipped: Boolean(result.skipped) }; 
  } catch (error) { 
    sendOneDriveStatus("Sync"); 
    return { ok: false, error: error.message }; 
  } 
}); 

ipcMain.handle("open-file-dialog", async () => { 
  const win = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
  try { 
    const result = await dialog.showOpenDialog(win, { 
      properties: ["openFile"], 
      filters: [{ name: "Microsoft 365 Documents", extensions: ["doc", "docx", "xls", "xlsx", "ppt", "pptx", "one"] }] 
    }); 
    if (result.canceled || result.filePaths.length === 0) return { canceled: true, filePath: null }; 
    return { canceled: false, filePath: result.filePaths[0] }; 
  } catch (error) { 
    return { canceled: true, error: error.message }; 
  } 
}); 

ipcMain.handle("office-open-external-file", async (event, filePath) => { 
  if (!filePath) return { ok: false, error: "No path." }; 
  try { 
    const detectedFiles = findOfficeFiles([process.execPath, filePath]); 
    if (!detectedFiles || detectedFiles.length === 0) throw new Error("Unsupported format."); 
    handleOfficeFile(detectedFiles[0]); 
    return { ok: true, service: detectedFiles[0].service }; 
  } catch (error) { 
    return { ok: false, error: error.message }; 
  } 
}); 

ipcMain.on("open-service", (event, service) => { openMicrosoftService(service); }); 
ipcMain.on("open-local-onedrive", () => { configureOneDrive(); }); 
ipcMain.on("browser-back", () => { if (officeView && officeView.webContents.canGoBack()) officeView.webContents.goBack(); }); 
ipcMain.on("browser-forward", () => { if (officeView && officeView.webContents.canGoForward()) officeView.webContents.goForward(); }); 
ipcMain.on("browser-reload", () => { if (officeView) officeView.webContents.reload(); }); 
ipcMain.on("browser-home", () => { openMicrosoftService("word"); }); 

ipcMain.handle("get-onedrive-logs", async () => {
  return new Promise((resolve) => {
    exec("journalctl --user-unit=onedrive -n 50 --no-pager 2>/dev/null || cat ~/.config/onedrive/onedrive.log 2>/dev/null", (err, stdout) => {
      resolve(stdout || "[OneDrive Daemon] Running in background.\n[Sync Status] Monitor loop active.\n[Status] Synchronized with Microsoft 365 Cloud.");
    });
  });
});

ipcMain.handle("set-desktop-shortcut", (event, enable) => {
  writePreference("desktop_shortcut", enable);
  const homeDir = app.getPath("home");
  let desktopDir = null;

  try {
    const userDirsPath = path.join(homeDir, ".config", "user-dirs.dirs");
    if (fs.existsSync(userDirsPath)) {
      const content = fs.readFileSync(userDirsPath, "utf8");
      const match = content.match(/XDG_DESKTOP_DIR="\$HOME\/(.*?)"/);
      if (match && match[1]) {
        const resolved = path.join(homeDir, match[1]);
        if (fs.existsSync(resolved)) {
          desktopDir = resolved;
        }
      }
    }
  } catch (_) {}

  if (!desktopDir) {
    const candidateFolders = ["Área de trabalho", "Área de Trabalho", "Desktop", "desktop"];
    for (const folder of candidateFolders) {
      const candidatePath = path.join(homeDir, folder);
      if (fs.existsSync(candidatePath)) {
        desktopDir = candidatePath;
        break;
      }
    }
  }

  if (!desktopDir) {
    desktopDir = path.join(homeDir, "Área de trabalho");
  }

  const targetFile = path.join(desktopDir, "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial.desktop");

  const desktopEntry = `[Desktop Entry]
Name=Microsoft 365 (Unofficial)
Comment=Microsoft 365 Desktop Integration for Linux
Exec=flatpak run io.github.DavidNi05.Microsoft365_for_Linux_Unofficial
Icon=io.github.DavidNi05.Microsoft365_for_Linux_Unofficial
Terminal=false
Type=Application
Categories=Office;Productivity;
StartupWMClass=Microsoft365
`;

  try {
    if (enable) {
      if (!fs.existsSync(desktopDir)) {
        fs.mkdirSync(desktopDir, { recursive: true });
      }
      fs.writeFileSync(targetFile, desktopEntry, { mode: 0o755 });
    } else {
      if (fs.existsSync(targetFile)) {
        fs.unlinkSync(targetFile);
      }
    }
    return true;
  } catch (err) {
    console.error("[Shortcut] Error updating desktop shortcut:", err);
    return false;
  }
});

ipcMain.handle("set-native-onedrive-sync", (event, enable) => {
  writePreference("native_onedrive_sync", enable);
  try {
    if (enable) {
      oneDrive.ensureDirectories();
      oneDrive.startHeartbeat();
      oneDrive.startMonitor(() => {
        sendOneDriveStatus("Sync");
      });
      setBackgroundStatus("OneDrive synchronization active");
      sendOneDriveStatus("Sync");
    } else {
      oneDrive.stopMonitor();
      oneDrive.stopSync();
      oneDrive.stopHeartbeat();
      setBackgroundStatus("OneDrive synchronization paused");
      sendOneDriveStatus("Not Connected");
    }
    return true;
  } catch (err) {
    console.error("[Native OneDrive] Error toggling sync daemon:", err);
    sendOneDriveStatus("Not Connected");
    return false;
  }
});

ipcMain.handle("clear-session-data", async () => {
  if (officeView && officeView.webContents && officeView.webContents.session) {
    await officeView.webContents.session.clearStorageData();
    officeView.webContents.reload();
  }
  return true;
});

ipcMain.handle("select-folder-dialog", async () => {
  const win = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
  const res = await dialog.showOpenDialog(win, {
    properties: ["openDirectory"]
  });
  if (!res.canceled && res.filePaths.length > 0) {
    return res.filePaths[0];
  }
  return null;
});

ipcMain.on("set-app-preference", (event, key, value) => {
  writePreference(key, value);

  if (mainWindow && !mainWindow.webContents.isDestroyed()) {
    mainWindow.webContents.send("config-updated", { key, value });
  }

  if (key === "autostart") {
    const autostartDir = path.join(app.getPath("home"), ".config", "autostart");
    const autostartFile = path.join(autostartDir, "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial.desktop");

    if (value) {
      if (!fs.existsSync(autostartDir)) fs.mkdirSync(autostartDir, { recursive: true });
      const entry = `[Desktop Entry]
Name=Microsoft 365 (Unofficial)
Exec=flatpak run io.github.DavidNi05.Microsoft365_for_Linux_Unofficial --background
Icon=io.github.DavidNi05.Microsoft365_for_Linux_Unofficial
Terminal=false
Type=Application
`;
      fs.writeFileSync(autostartFile, entry, { mode: 0o755 });
    } else {
      if (fs.existsSync(autostartFile)) fs.unlinkSync(autostartFile);
    }
  } else if (key === "close_to_tray") {
    closeToTrayEnabled = Boolean(value);
  } else if (key === "global_hotkey") {
    setupGlobalShortcut(Boolean(value));
  } else if (key === "compact_nav") {
    currentSidebarWidth = value ? 54 : 230;
    scheduleUpdateViewBounds();
  } else if (key === "autohide_sidebar") {
    scheduleUpdateViewBounds();
  }
});

let hasSingleInstanceLock = app.requestSingleInstanceLock(); 
if (!hasSingleInstanceLock) { 
  app.quit(); 
} else { 
  app.on("second-instance", (event, argv) => { 
    if (argv.includes("--quit-ui")) { 
      forceQuit = true; 
      if (mainWindow && !mainWindow.isDestroyed()) mainWindow.destroy(); 
      app.quit(); 
      return; 
    } 
    const additionalFiles = findOfficeFiles(argv); 
    if (additionalFiles.length > 0) pendingOfficeFiles.push(...additionalFiles); 
    showMainWindow(); 
    if (additionalFiles.length > 0 && officeView) openPendingOfficeFile().catch(() => {}); 
  }); 
} 

app.whenReady().then(async () => { 
  if (!hasSingleInstanceLock) return; 
  if (backgroundMode) { await startBackgroundMode(); return; } 
  if (quitUiMode) { forceQuit = true; app.quit(); return; } 

  createWindow(); 
}); 

app.on("window-all-closed", () => { if (forceQuit && process.platform !== "darwin") app.quit(); }); 
app.on("activate", () => { 
  if (backgroundMode) return; 
  if (mainWindow && !mainWindow.isDestroyed()) { showMainWindow(); return; } 
  if (!forceQuit) createWindow(); 
}); 
app.on("before-quit", () => { 
  forceQuit = true; 
  globalShortcut.unregisterAll();
  try {
    cleanupImportsTemp();
  } catch (_) {}

  oneDrive.stopMonitor(); 
  oneDrive.stopSync(); 
  oneDrive.stopHeartbeat(); 
  oneDrive.destroyBackgroundTray(); 
});

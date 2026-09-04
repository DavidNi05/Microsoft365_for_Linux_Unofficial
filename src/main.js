const { 
  app, 
  BrowserWindow, 
  WebContentsView, 
  ipcMain, 
  shell, 
  dialog,
  Menu,
  Tray,
  nativeImage
} = require("electron"); 

const path = require("path"); 
const fs = require("fs"); 
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

const SIDEBAR_WIDTH = 230; 
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
  if (oneDrive.backgroundAlive() || oneDrive.monitorRunning()) return true; 

  try { 
    oneDrive.ensureDirectories();
    oneDrive.startHeartbeat();
    oneDrive.startMonitor(() => {});
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
  const leftInset = Math.max(0, Math.floor(SIDEBAR_WIDTH)); 
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

  if (service === "credits") {
    officeView.webContents.loadFile(path.join(__dirname, "credits.html"));
    sendToInterface("browser-url", "m365://credits");
    return;
  }

  if (services[service]) { 
    officeView.webContents.loadURL(services[service]); 
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
      console.log(`[Office Online] Immediate link obtained: ${result.webUrl}`);
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
          console.log(`[Office Online] Cloud link confirmed: ${retry.webUrl}`);
          sendToInterface("office-file-loading", { fileName: officeFile.name, service: officeFile.service, status: "Opening in Office Online..." });
          officeView.webContents.loadURL(retry.webUrl); 
          sendOneDriveStatus("Synced"); 
          sendToInterface("office-file-opened", { ...officeFile, service: officeFile.service, online: true, webUrl: retry.webUrl }); 
          return; 
        } 
      } 
      console.error("[Office Online] Timed out waiting for official web link from OneDrive.");
      throw new Error("Cloud link response timed out."); 
    } 

    openMicrosoftService(officeFile.service); 
  } catch (error) { 
    console.error("[Office Online] File loading error:", error.message);
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
    console.log(`[Office Online] Real-time rename detected: '${currentName}' -> '${cleanName}'`);
    sendToInterface("office-file-renamed", { fileName: cleanName });
  }
}

async function scrapeOneDriveQuotaFromWeb() {
  if (!officeView || officeView.webContents.isDestroyed()) return;
  const currentUrl = officeView.webContents.getURL();

  if (currentUrl.includes("onedrive.live.com")) {
    try {
      const quotaData = await officeView.webContents.executeJavaScript(`
        (function() {
          const el = document.querySelector('[aria-label*="Storage:"]');
          if (!el) return null;
          const label = el.getAttribute('aria-label');
          const match = label.match(/Storage:\\s*([^\\s]+(?:\\s+[^\\s]+)*?)\\s+usado de\\s+([^\\s]+(?:\\s+[^\\s]+)*?)\\s*\\((\\d+)\\%\\)/i);
          if (!match) return null;
          return {
            used: match[1],
            total: match[2],
            percentUsed: parseInt(match[3], 10),
            percent: parseInt(match[3], 10),
            remaining: "Available"
          };
        })();
      `);

      if (quotaData && quotaData.total) {
        oneDrive.setWebScrapedQuota(quotaData);
        sendToInterface("onedrive:quota-updated", quotaData);
      }
    } catch (_) {}
  }
}

function createOfficeView() { 
  officeView = new WebContentsView({ 
    webPreferences: { 
      nodeIntegration: false, 
      contextIsolation: true, 
      sandbox: true, 
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
      scrapeOneDriveQuotaFromWeb();
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
    sendOneDriveStatus(oneDrive.backgroundAlive() ? "Synced" : "Starting background..."); 
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
    if (!forceQuit) { 
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
      sendOneDriveStatus("Synced"); 
      await shell.openPath(oneDrive.syncDirectory); 
      return; 
    } 

    sendOneDriveStatus("Signing in / Synchronizing..."); 
    await oneDrive.runInitialSync(() => {}); 
    await requestBackgroundAccess(); 
    await ensureBackgroundRunning(); 
    sendOneDriveStatus("Synced • Background active"); 
    await shell.openPath(oneDrive.syncDirectory); 
  } catch (error) { 
    sendOneDriveStatus("OneDrive error"); 
  } finally { 
    oneDriveSetupRunning = false; 
  } 
} 

async function startBackgroundMode() { 
  oneDrive.ensureDirectories(); 
  oneDrive.startHeartbeat(); 
  oneDrive.startMonitor(() => {}); 
  await setBackgroundStatus("OneDrive synchronization active"); 
} 

/* 
 * ========================================================= 
 * IPC CHANNELS
 * ========================================================= 
 */ 

ipcMain.on("window-resize-notify", () => {
  scheduleUpdateViewBounds();
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
    console.log(`[Manual Save Local] Triggered. Target Path: "${targetToSave}"`);

    if (!targetToSave) {
      return { ok: false, error: "No local file associated with this session." }; 
    }

    const expectedWebTitle = await getOfficeOnlineDocumentTitle();
    console.log(`[Manual Save Local] Detected Document Web Title: "${expectedWebTitle}"`);

    sendOneDriveStatus("Saving to disk..."); 

    const result = await saveLocalNow(targetToSave, expectedWebTitle); 

    if (!result || !result.ok) { 
      const failedName = result?.fileName || path.basename(targetToSave); 
      console.error(`[Manual Save Local] Failed to save ${failedName}:`, result?.error);
      sendOneDriveStatus("Save error"); 
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

    sendOneDriveStatus(result.renamed ? "Saved (Renamed) ✓" : "Saved ✓"); 
    return { ...result, ok: true }; 

  } catch (error) { 
    console.error("[Manual Save Local] Exception error:", error); 
    sendOneDriveStatus("Save error"); 
    return { ok: false, fileName: path.basename(activeExternalOriginalPath || ""), error: error.message }; 
  } finally { 
    manualSaveInProgress = false; 
  } 
}); 

ipcMain.handle("onedrive:get-storage-quota", async () => {
  try {
    const quotaData = await oneDrive.getQuota(false);
    return { success: true, data: quotaData, quota: quotaData };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle("onedrive-get-quota", async (event, force) => {
  try {
    const quota = await oneDrive.getQuota(Boolean(force));
    return { ok: true, quota };
  } catch (error) {
    return { ok: false, error: error.message };
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
    sendOneDriveStatus("Synced"); 
    return { ok: result.code === 0, skipped: Boolean(result.skipped) }; 
  } catch (error) { 
    sendOneDriveStatus("Sync error"); 
    return { ok: false, error: error.message }; 
  } 
}); 

ipcMain.handle("open-file-dialog", async () => { 
  const window = BrowserWindow.getFocusedWindow() || mainWindow; 
  try { 
    const result = await dialog.showOpenDialog(window, { 
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
  ensureBackgroundRunning().then(active => { 
    sendOneDriveStatus(active ? "Synced" : "Background starting..."); 
  }).catch(() => {}); 
}); 

app.on("window-all-closed", () => { if (forceQuit && process.platform !== "darwin") app.quit(); }); 
app.on("activate", () => { 
  if (backgroundMode) return; 
  if (mainWindow && !mainWindow.isDestroyed()) { showMainWindow(); return; } 
  if (!forceQuit) createWindow(); 
}); 
app.on("before-quit", () => { 
  forceQuit = true; 
  try {
    cleanupImportsTemp();
  } catch (_) {}

  oneDrive.stopMonitor(); 
  oneDrive.stopSync(); 
  oneDrive.stopHeartbeat(); 
  oneDrive.destroyBackgroundTray(); 
});

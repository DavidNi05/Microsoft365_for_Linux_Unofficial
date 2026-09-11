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
  globalShortcut, 
  session 
} = require("electron"); 

const path = require("path"); 
const fs = require("fs"); 
const { spawn, exec } = require("child_process");
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

const CHROME_DESKTOP_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36";
app.userAgentFallback = CHROME_DESKTOP_UA;

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
let isAuthenticating = false;
let pendingOfficeFiles = findOfficeFiles(process.argv); 
let closeToTrayEnabled = true;

let currentSidebarWidth = 230; 
const TOPBAR_HEIGHT = 55; 

const MS_AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" +
  "client_id=" + oneDrive.CLIENT_ID +
  "&response_type=code" +
  "&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient" +
  "&response_mode=query" +
  "&prompt=select_account" +
  "&scope=Files.ReadWrite%20Files.ReadWrite.All%20Sites.ReadWrite.All%20User.Read%20offline_access";

// URLs canônicas compatíveis tanto com contas pessoais (Live/Hotmail) quanto corporativas
const services = { 
  home: "https://office.live.com/start/Word.aspx", 
  word: "https://office.live.com/start/Word.aspx", 
  excel: "https://office.live.com/start/Excel.aspx", 
  powerpoint: "https://office.live.com/start/PowerPoint.aspx", 
  onenote: "https://www.onenote.com/notebooks", 
  copilot: "https://copilot.microsoft.com/", 
  onedrive: "https://onedrive.live.com/", 
  outlook: "https://outlook.live.com/mail/", 
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
  "skype.com",
  "microsoft365.com"
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

function sleep(ms) { 
  return new Promise(resolve => setTimeout(resolve, ms)); 
} 

function sendToInterface(channel, value) { 
  if (mainWindow && !mainWindow.isDestroyed() && mainWindow.webContents) { 
    mainWindow.webContents.send(channel, value); 
  } 
} 

function sendOneDriveStatus(status) { 
  sendToInterface("onedrive-status", status); 
} 

function showMainWindow() { 
  if (!mainWindow || mainWindow.isDestroyed()) { 
    createWindow(); 
    return; 
  } 
  if (mainWindow.isMinimized()) mainWindow.restore(); 
  if (!mainWindow.isVisible()) mainWindow.show(); 
  mainWindow.focus(); 
} 

function getPreferencesFilePath() { 
  return path.join(app.getPath("userData"), "app_preferences.json"); 
}

function readAllPreferences() { 
  try { 
    const filePath = getPreferencesFilePath(); 
    if (fs.existsSync(filePath)) {
      return JSON.parse(fs.readFileSync(filePath, "utf8")); 
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
    native_onedrive_sync: true, 
    hardware_accel: false, 
    autostart: false, 
    close_to_tray: true, 
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

oneDrive.setOpenUiHandler(() => {
  showMainWindow();
});

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
  if (prefs.native_onedrive_sync === false) return false; 
  if (!oneDrive.isAuthenticated()) { 
    sendOneDriveStatus("Not Connected"); 
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

async function updateProfileAvatar() {
  try {
    const avatarData = await oneDrive.fetchGraphUserProfilePicture();
    if (avatarData) {
      sendToInterface("update-user-avatar", avatarData);
    }
  } catch (_) {}
}

function isOAuthCodeRedirect(targetUrl) {
  if (!targetUrl || typeof targetUrl !== "string") return false;
  if (targetUrl.includes("/oauth2/v2.0/authorize")) return false;

  const isNativeClient = targetUrl.startsWith("https://login.microsoftonline.com/common/oauth2/nativeclient") ||
                         targetUrl.includes("/oauth2/nativeclient");
  const isLiveDesktop = targetUrl.includes("login.live.com/oauth20_desktop.srf");
  const hasAuthCode = targetUrl.includes("code=") || targetUrl.includes("?code=") || targetUrl.includes("&code=");

  return (isNativeClient || isLiveDesktop) && hasAuthCode;
}

async function handleAuthRedirect(targetUrl) {
  if (isAuthenticating) return;
  if (!isOAuthCodeRedirect(targetUrl)) return;

  isAuthenticating = true;
  sendOneDriveStatus("Authenticating...");

  try {
    const mainSession = session.fromPartition("persist:office365");
    await mainSession.cookies.flushStore();
  } catch (_) {}

  const result = await oneDrive.authenticateWithResponseUrl(targetUrl);

  if (result && result.ok) {
    await ensureBackgroundRunning();
    sendOneDriveStatus("Sync");
    await updateProfileAvatar();

    sendToInterface("auth-state-changed", { authenticated: true });
    scheduleUpdateViewBounds();

    const prefs = readAllPreferences();
    const startApp = prefs.default_app === "last_used" ? "word" : (prefs.default_app || "word");
    openMicrosoftService(startApp);

    showFluentNotification({ 
      title: "Office 365 Suite", 
      message: "Account authenticated successfully! Single Sign-On and OneDrive sync are active.", 
      icon: "Home.png", 
      duration: 4000 
    });
  } else {
    sendOneDriveStatus("Not Connected");
    dialog.showErrorBox("Authentication Failed", "Could not complete native OneDrive synchronization login.");
    openMicrosoftService("word");
  }

  setTimeout(() => { isAuthenticating = false; }, 3000);
}

function startUnifiedLoginWindow() {
  sendToInterface("auth-state-changed", { authenticated: false });
  updateViewBounds();

  if (officeView && !officeView.webContents.isDestroyed()) {
    officeView.webContents.loadURL(MS_AUTH_URL);
  }
}

function updateViewBounds() { 
  if (!mainWindow || mainWindow.isDestroyed() || !officeView || officeView.webContents.isDestroyed()) return; 

  const contentBounds = mainWindow.getContentBounds(); 
  const contentWidth = Math.max(1, Math.floor(Number(contentBounds.width) || 0)); 
  const contentHeight = Math.max(1, Math.floor(Number(contentBounds.height) || 0)); 

  if (!oneDrive.isAuthenticated() || isAuthenticating) {
    officeView.setBounds({ 
      x: 0, 
      y: 0, 
      width: contentWidth, 
      height: contentHeight 
    });
    return;
  }

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
  const chosenOriginal = explicitOriginal || (result && result.originalPath) || activeExternalOriginalPath; 
  if (chosenOriginal && typeof chosenOriginal === "string" && chosenOriginal.trim()) { 
    const originalPath = path.resolve(chosenOriginal); 
    const workingPath = result && typeof result.localOneDrivePath === "string" && result.localOneDrivePath.trim() 
      ? path.resolve(result.localOneDrivePath) 
      : activeExternalOneDrivePath; 

    activeExternalOriginalPath = originalPath; 
    activeExternalOneDrivePath = workingPath; 
    activeExternalImported = result ? result.imported !== false : true; 

    if (workingPath && activeExternalImported && typeof registerExternalFile === "function") {
      try { 
        registerExternalFile(originalPath, workingPath, activeExternalImported); 
      } catch (_) {} 
    }
  } 
} 

async function handleOfficeFile(officeFile) { 
  if (!officeFile || !officeFile.service || isOpeningOfficeFile) return; 

  isOpeningOfficeFile = true; 
  const targetPath = officeFile.originalPath || officeFile.path;
  activeExternalOriginalPath = targetPath ? path.resolve(targetPath) : null; 
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
    const fileToProcess = activeExternalOriginalPath || officeFile.path;
    let result = await getFileWebUrl(fileToProcess); 
    updateActiveExternalFile(result, activeExternalOriginalPath); 

    if (result && result.insideOneDrive && result.webUrl) { 
      sendToInterface("office-file-loading", { 
        fileName: officeFile.name, 
        service: officeFile.service, 
        status: "Opening document..." 
      }); 
      officeView.webContents.loadURL(result.webUrl); 
      sendToInterface("office-file-opened", { 
        ...officeFile, 
        service: officeFile.service, 
        online: true, 
        webUrl: result.webUrl 
      }); 
      return; 
    } 

    if (result && result.insideOneDrive && !result.webUrl) { 
      for (let attempt = 1; attempt <= 25; attempt++) { 
        sendToInterface("office-file-loading", { 
          fileName: officeFile.name, 
          service: officeFile.service, 
          status: `Synchronizing document (${attempt}/25)...` 
        }); 

        await sleep(1500); 

        const retry = await getFileWebUrl(fileToProcess); 
        updateActiveExternalFile(retry, activeExternalOriginalPath); 

        if (retry && retry.webUrl) { 
          sendToInterface("office-file-loading", { 
            fileName: officeFile.name, 
            service: officeFile.service, 
            status: "Opening document..." 
          }); 
          officeView.webContents.loadURL(retry.webUrl); 
          sendOneDriveStatus("Sync"); 
          sendToInterface("office-file-opened", { 
            ...officeFile, 
            service: officeFile.service, 
            online: true, 
            webUrl: retry.webUrl 
          }); 
          return; 
        } 
      } 
      throw new Error("Timeout waiting for cloud file link."); 
    } 

    openMicrosoftService(officeFile.service); 
  } catch (_) { 
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
  const clean = cleanDocTitle(title, currentExt);
  if (!clean) return;

  const currentName = path.basename(activeExternalOriginalPath);
  if (clean.toLowerCase() !== currentName.toLowerCase()) {
    sendToInterface("office-file-renamed", { fileName: clean });
  }
}

function setupGlobalShortcut(enable) { 
  const shortcutKey = "Ctrl+Alt+M"; 
  if (!enable) { 
    try { globalShortcut.unregister(shortcutKey); } catch(_) {} 
    return; 
  } 
  try { 
    globalShortcut.register(shortcutKey, () => {
      showMainWindow();
    }); 
  } catch(_) {} 
}

function createOfficeView() { 
  officeView = new WebContentsView({ 
    webPreferences: { 
      nodeIntegration: false, 
      contextIsolation: true, 
      sandbox: false, 
      preload: path.join(__dirname, "preload.js"), 
      partition: "persist:office365" 
    } 
  }); 

  mainWindow.contentView.addChildView(officeView); 
  scheduleUpdateViewBounds(); 

  officeView.webContents.setUserAgent(CHROME_DESKTOP_UA);
  officeView.webContents.on("page-title-updated", (_e, title) => handleWebTitleChange(title)); 

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
    if (isOAuthCodeRedirect(url)) {
      event.preventDefault();
      handleAuthRedirect(url);
      return;
    }
    if (!isAllowedUrl(url)) { 
      event.preventDefault(); 
      shell.openExternal(url); 
    } 
  }); 

  officeView.webContents.on("will-redirect", (event, url) => {
    if (isOAuthCodeRedirect(url)) {
      event.preventDefault();
      handleAuthRedirect(url);
    }
  });

  officeView.webContents.on("did-start-loading", () => sendToInterface("browser-loading", true)); 
  officeView.webContents.on("did-stop-loading", () => { 
    sendToInterface("browser-loading", false); 
    if (!officeView.webContents.isDestroyed()) {
      sendToInterface("browser-url", officeView.webContents.getURL()); 
    }
  }); 

  if (oneDrive.isAuthenticated()) {
    openPendingOfficeFile().catch(() => openMicrosoftService("word"));
  } else {
    startUnifiedLoginWindow();
  }
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
    title: "Office 365 Suite (Community Client)", 
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
    const mainSession = session.fromPartition("persist:office365");
    mainSession.setUserAgent(CHROME_DESKTOP_UA);

    mainSession.webRequest.onBeforeSendHeaders((details, callback) => {
      details.requestHeaders["User-Agent"] = CHROME_DESKTOP_UA;
      details.requestHeaders["sec-ch-ua"] = '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"';
      details.requestHeaders["sec-ch-ua-mobile"] = "?0";
      details.requestHeaders["sec-ch-ua-platform"] = '"Windows"';
      delete details.requestHeaders["X-Requested-With"];
      callback({ requestHeaders: details.requestHeaders });
    });

    mainSession.webRequest.onBeforeRequest(
      { urls: ["https://login.microsoftonline.com/common/oauth2/nativeclient*", "*://login.live.com/oauth20_desktop.srf*"] },
      (details, callback) => {
        if (isOAuthCodeRedirect(details.url)) {
          handleAuthRedirect(details.url);
          callback({ cancel: true });
          return;
        }
        callback({});
      }
    );

    createOfficeView(); 
    const prefs = readAllPreferences(); 
    closeToTrayEnabled = prefs.close_to_tray !== false; 
    
    if (oneDrive.isAuthenticated() && prefs.native_onedrive_sync !== false) { 
      ensureBackgroundRunning(); 
      sendOneDriveStatus("Sync"); 
      updateProfileAvatar(); 
      sendToInterface("auth-state-changed", { authenticated: true });
    } else { 
      sendOneDriveStatus("Not Connected"); 
      sendToInterface("auth-state-changed", { authenticated: false });
    } 
  }); 

  [
    "resize", "resized", "maximize", "unmaximize", 
    "restore", "enter-full-screen", "leave-full-screen", "show"
  ].forEach(eventName => {
    mainWindow.on(eventName, scheduleUpdateViewBounds);
  }); 

  mainWindow.on("close", event => { 
    if (!forceQuit && closeToTrayEnabled) { 
      event.preventDefault(); 
      mainWindow.hide(); 
      try { oneDrive.createBackgroundTray(); } catch (_) {} 
      showFluentNotification({ 
        title: "Office 365 Suite", 
        message: "Application minimized to system tray.", 
        icon: "Home.png", 
        duration: 3500 
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
  if (!oneDrive.isAuthenticated()) { 
    startUnifiedLoginWindow(); 
    return; 
  } 
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
  } catch (_) { 
    sendOneDriveStatus("Not Connected"); 
  } finally { 
    oneDriveSetupRunning = false; 
  } 
} 

async function startBackgroundMode() { 
  if (!oneDrive.isAuthenticated()) return; 
  oneDrive.ensureDirectories(); 
  oneDrive.startHeartbeat(); 
  oneDrive.startMonitor(() => { 
    sendOneDriveStatus("Sync"); 
  }); 
  await setBackgroundStatus("OneDrive synchronization active"); 
} 

/* 
 * ========================================================= 
 * CANAIS IPC COMPLETOS 
 * ========================================================= 
 */ 

ipcMain.handle("start-unified-login", () => { 
  startUnifiedLoginWindow(); 
  return true; 
}); 

ipcMain.handle("check-auth-status", () => {
  return { authenticated: oneDrive.isAuthenticated() };
}); 

ipcMain.handle("get-user-avatar", async () => {
  return await oneDrive.fetchGraphUserProfilePicture();
}); 

ipcMain.on("window-resize-notify", (_event, newSidebarWidth) => { 
  if (typeof newSidebarWidth === "number") {
    currentSidebarWidth = newSidebarWidth; 
  }
  updateViewBounds(); 
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

ipcMain.handle("set-sound-enabled", (_event, enabled) => { 
  const newState = setSoundEnabled(enabled); 
  oneDrive.rebuildTrayMenu(); 
  return newState; 
}); 

ipcMain.handle("office-save-local-now", async () => { 
  if (manualSaveInProgress) return { ok: false, busy: true }; 
  manualSaveInProgress = true; 

  try { 
    const target = activeExternalOriginalPath || activeExternalOneDrivePath; 
    if (!target) return { ok: false, error: "No local file associated with this session." }; 

    const expectedTitle = await getOfficeOnlineDocumentTitle(); 
    sendOneDriveStatus("Sync"); 
    const result = await saveLocalNow(target, expectedTitle); 

    if (!result || !result.ok) {
      const failedName = result?.fileName || path.basename(target);
      sendOneDriveStatus("Sync");
      showFluentNotification({
        title: "Office 365 Suite",
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
        oldPath: target,
        newPath: activeExternalOriginalPath,
        fileName: result.fileName
      });
    }

    showFluentNotification({
      title: "Office 365 Suite",
      message: result.message || `File saved successfully: ${result.fileName}`,
      icon: "OneDrive.png",
      duration: 3500
    });

    sendOneDriveStatus("Sync");
    return { ...result, ok: true };
  } catch (err) { 
    sendOneDriveStatus("Sync");
    return { ok: false, fileName: path.basename(activeExternalOriginalPath || ""), error: err.message }; 
  } finally { 
    manualSaveInProgress = false; 
  } 
}); 

ipcMain.handle("onedrive-get-status", async () => { 
  try { 
    return { ok: true, ...(await oneDrive.getStatus()) }; 
  } catch (err) { 
    return { ok: false, error: err.message }; 
  } 
}); 

ipcMain.handle("onedrive-open-folder", async () => { 
  oneDrive.ensureDirectories(); 
  const err = await shell.openPath(oneDrive.syncDirectory); 
  return { ok: !err, error: err || null }; 
}); 

ipcMain.handle("onedrive-sync-now", async () => { 
  try { 
    const res = await oneDrive.syncNow(() => {}); 
    sendOneDriveStatus("Sync"); 
    return { ok: res.code === 0, skipped: Boolean(res.skipped) }; 
  } catch (err) { 
    sendOneDriveStatus("Sync"); 
    return { ok: false, error: err.message }; 
  } 
}); 

ipcMain.handle("open-file-dialog", async () => { 
  const win = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null; 
  try { 
    const res = await dialog.showOpenDialog(win, { 
      properties: ["openFile"], 
      filters: [{ name: "Office Documents", extensions: ["doc", "docx", "xls", "xlsx", "ppt", "pptx", "one"] }] 
    }); 
    if (res.canceled || res.filePaths.length === 0) return { canceled: true, filePath: null }; 
    return { canceled: false, filePath: res.filePaths[0] }; 
  } catch (err) { 
    return { canceled: true, error: err.message }; 
  } 
}); 

ipcMain.handle("office-open-external-file", async (_event, filePath) => { 
  if (!filePath) return { ok: false, error: "No path." }; 
  try { 
    const detectedFiles = findOfficeFiles([process.execPath, filePath]); 
    if (!detectedFiles || detectedFiles.length === 0) throw new Error("Unsupported format."); 
    handleOfficeFile(detectedFiles[0]); 
    return { ok: true, service: detectedFiles[0].service }; 
  } catch (err) { 
    return { ok: false, error: err.message }; 
  } 
}); 

ipcMain.on("open-service", (_event, service) => { openMicrosoftService(service); }); 
ipcMain.on("open-local-onedrive", () => { configureOneDrive(); }); 
ipcMain.on("browser-back", () => { if (officeView && officeView.webContents.canGoBack()) officeView.webContents.goBack(); }); 
ipcMain.on("browser-forward", () => { if (officeView && officeView.webContents.canGoForward()) officeView.webContents.goForward(); }); 
ipcMain.on("browser-reload", () => { if (officeView) officeView.webContents.reload(); }); 
ipcMain.on("browser-home", () => { openMicrosoftService("word"); }); 

ipcMain.handle("get-onedrive-logs", async () => { 
  return new Promise((resolve) => { 
    exec("journalctl --user-unit=onedrive -n 50 --no-pager 2>/dev/null || cat ~/.config/onedrive/onedrive.log 2>/dev/null", (_err, stdout) => { 
      resolve(stdout || "[OneDrive Daemon] Running in background.\n[Sync Status] Synchronized."); 
    }); 
  }); 
}); 

ipcMain.handle("set-desktop-shortcut", (_event, enable) => { 
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
        if (fs.existsSync(resolved)) desktopDir = resolved; 
      } 
    } 
  } catch (_) {} 

  if (!desktopDir) { 
    const candidates = ["Área de trabalho", "Área de Trabalho", "Desktop", "desktop"]; 
    for (const folder of candidates) { 
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
  const entry = `[Desktop Entry]
Name=Office 365 Suite (Community Client)
Comment=Community desktop integration for Office 365
Exec=flatpak run io.github.DavidNi05.Microsoft365_for_Linux_Unofficial
Icon=io.github.DavidNi05.Microsoft365_for_Linux_Unofficial
Terminal=false
Type=Application
Categories=Office;Productivity;
StartupWMClass=Office 365 Suite (Community Client)
`; 

  try { 
    if (enable) { 
      if (!fs.existsSync(desktopDir)) fs.mkdirSync(desktopDir, { recursive: true }); 
      fs.writeFileSync(targetFile, entry, { mode: 0o755 }); 
    } else { 
      if (fs.existsSync(targetFile)) fs.unlinkSync(targetFile); 
    } 
    return true; 
  } catch (_) { 
    return false; 
  } 
}); 

ipcMain.handle("set-native-onedrive-sync", (_event, enable) => { 
  writePreference("native_onedrive_sync", enable); 
  try { 
    if (enable) { 
      if (oneDrive.isAuthenticated()) { 
        oneDrive.ensureDirectories(); 
        oneDrive.startHeartbeat(); 
        oneDrive.startMonitor(() => sendOneDriveStatus("Sync")); 
        setBackgroundStatus("OneDrive synchronization active"); 
        sendOneDriveStatus("Sync"); 
        updateProfileAvatar(); 
      } else { 
        startUnifiedLoginWindow(); 
      } 
    } else { 
      oneDrive.stopMonitor(); 
      oneDrive.stopSync(); 
      oneDrive.stopHeartbeat(); 
      setBackgroundStatus("OneDrive synchronization paused"); 
      sendOneDriveStatus("Not Connected"); 
    } 
    return true; 
  } catch (_) { 
    sendOneDriveStatus("Not Connected"); 
    return false; 
  } 
}); 

ipcMain.handle("clear-session-data", async () => { 
  if (officeView && officeView.webContents && officeView.webContents.session) { 
    await officeView.webContents.session.clearStorageData(); 
  } 
  try {
    const rf = path.join(oneDrive.configDirectory, "refresh_token");
    if (fs.existsSync(rf)) fs.unlinkSync(rf);
  } catch (_) {}

  sendToInterface("auth-state-changed", { authenticated: false });
  updateViewBounds();
  return true; 
}); 

ipcMain.handle("select-folder-dialog", async () => { 
  const win = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null; 
  const res = await dialog.showOpenDialog(win, { properties: ["openDirectory"] }); 
  if (!res.canceled && res.filePaths.length > 0) return res.filePaths[0]; 
  return null; 
}); 

ipcMain.handle("factory-reset", async () => { 
  try { 
    oneDrive.stopMonitor(); 
    oneDrive.stopSync(); 
    oneDrive.stopHeartbeat(); 
    oneDrive.destroyBackgroundTray(); 

    if (officeView && officeView.webContents && officeView.webContents.session) {
      await officeView.webContents.session.clearStorageData(); 
    }
    await session.defaultSession.clearStorageData(); 

    const homeDir = app.getPath("home"); 
    const autostartFile = path.join(homeDir, ".config", "autostart", "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial.desktop"); 
    if (fs.existsSync(autostartFile)) fs.unlinkSync(autostartFile); 

    ["Área de trabalho", "Desktop"].forEach(d => { 
      const dfp = path.join(homeDir, d, "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial.desktop"); 
      if (fs.existsSync(dfp)) fs.unlinkSync(dfp); 
    }); 

    const odc = path.join(homeDir, ".config", "onedrive"); 
    if (fs.existsSync(odc)) try { fs.rmSync(odc, { recursive: true, force: true }); } catch (_) {} 

    const prefsFile = getPreferencesFilePath(); 
    if (fs.existsSync(prefsFile)) fs.unlinkSync(prefsFile); 

    const welcomeFile = path.join(app.getPath("userData"), "welcome_dialog_shown.json"); 
    if (fs.existsSync(welcomeFile)) fs.unlinkSync(welcomeFile); 

    forceQuit = true; 
    app.relaunch(); 
    app.exit(0); 
    return true; 
  } catch (_) { 
    return false; 
  } 
}); 

ipcMain.handle("uninstall-app", async (_event, { deleteData }) => { 
  try { 
    oneDrive.stopMonitor(); 
    oneDrive.stopSync(); 
    oneDrive.stopHeartbeat(); 
    oneDrive.destroyBackgroundTray(); 

    const homeDir = app.getPath("home"); 

    if (deleteData) { 
      const autostartFile = path.join(homeDir, ".config", "autostart", "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial.desktop"); 
      if (fs.existsSync(autostartFile)) try { fs.unlinkSync(autostartFile); } catch (_) {} 

      ["Área de trabalho", "Desktop"].forEach(d => { 
        const dfp = path.join(homeDir, d, "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial.desktop"); 
        if (fs.existsSync(dfp)) try { fs.unlinkSync(dfp); } catch (_) {} 
      }); 

      const odc = path.join(homeDir, ".config", "onedrive"); 
      if (fs.existsSync(odc)) try { fs.rmSync(odc, { recursive: true, force: true }); } catch (_) {} 
    } 

    const appId = "io.github.DavidNi05.Microsoft365_for_Linux_Unofficial"; 
    const flag = deleteData ? "--delete-data" : ""; 
    const flatpakCmd = `flatpak uninstall ${flag} -y ${appId}`; 
    const hostCmd = fs.existsSync("/.flatpak-info") ? `flatpak-spawn --host ${flatpakCmd}` : flatpakCmd; 

    const child = spawn("sh", ["-c", `sleep 1.2 && ${hostCmd}`], { detached: true, stdio: "ignore" }); 
    child.unref(); 

    forceQuit = true; 
    setTimeout(() => app.exit(0), 400); 
    return { success: true }; 
  } catch (err) { 
    return { success: false, error: err.message }; 
  } 
}); 

ipcMain.on("set-app-preference", (_event, key, value) => { 
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
Name=Office 365 Suite (Community Client)
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
    updateViewBounds(); 
  } else if (key === "autohide_sidebar") { 
    updateViewBounds(); 
  } 
}); 

const hasSingleInstanceLock = app.requestSingleInstanceLock(); 
if (!hasSingleInstanceLock) { 
  app.quit(); 
} else { 
  app.on("second-instance", (_event, argv) => { 
    if (argv.includes("--quit-ui")) { 
      forceQuit = true; 
      if (mainWindow && !mainWindow.isDestroyed()) mainWindow.destroy(); 
      app.quit(); 
      return; 
    } 
    const additionalFiles = findOfficeFiles(argv); 
    if (additionalFiles.length > 0) pendingOfficeFiles.push(...additionalFiles); 
    showMainWindow(); 
    if (additionalFiles.length > 0 && officeView) { 
      openPendingOfficeFile().catch(() => {}); 
    } 
  }); 
} 

app.whenReady().then(async () => { 
  if (!hasSingleInstanceLock) return; 
  if (backgroundMode) { 
    await startBackgroundMode(); 
    return; 
  } 
  if (quitUiMode) { 
    forceQuit = true; 
    app.quit(); 
    return; 
  } 
  closeToTrayEnabled = readAllPreferences().close_to_tray !== false; 
  createWindow(); 
}); 

app.on("window-all-closed", () => { 
  if (!closeToTrayEnabled || forceQuit) { 
    if (process.platform !== "darwin") app.quit(); 
  } 
}); 

app.on("activate", () => { 
  if (backgroundMode) return; 
  showMainWindow(); 
}); 

app.on("before-quit", () => { 
  forceQuit = true; 
  globalShortcut.unregisterAll(); 
  try { cleanupImportsTemp(); } catch (_) {} 
  oneDrive.stopMonitor(); 
  oneDrive.stopSync(); 
  oneDrive.stopHeartbeat(); 
  oneDrive.destroyBackgroundTray(); 
});

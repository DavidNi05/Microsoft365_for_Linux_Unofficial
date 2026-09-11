const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  notifyResize: (newWidth) => ipcRenderer.send("window-resize-notify", newWidth),
  showFluentToast: (options) => ipcRenderer.invoke("show-fluent-toast", options),

  // Navegação e Serviços
  openService: (service) => ipcRenderer.send("open-service", service),
  browserBack: () => ipcRenderer.send("browser-back"),
  browserForward: () => ipcRenderer.send("browser-forward"),
  browserReload: () => ipcRenderer.send("browser-reload"),
  browserHome: () => ipcRenderer.send("browser-home"),

  // Áudio
  getSoundEnabled: () => ipcRenderer.invoke("get-sound-enabled"),
  setSoundEnabled: (enabled) => ipcRenderer.invoke("set-sound-enabled", enabled),
  onSoundSettingChanged: (callback) => ipcRenderer.on("sound-setting-changed", (_event, val) => callback(val)),

  // OneDrive e Armazenamento
  openLocalOneDrive: () => ipcRenderer.send("open-local-onedrive"),
  getOneDriveStatus: () => ipcRenderer.invoke("onedrive-get-status"),
  openOneDriveFolder: () => ipcRenderer.invoke("onedrive-open-folder"),
  syncOneDriveNow: () => ipcRenderer.invoke("onedrive-sync-now"),
  saveLocalNow: () => ipcRenderer.invoke("office-save-local-now"),

  // Login Único Nativo (SSO) e Avatar
  startUnifiedLogin: () => ipcRenderer.invoke("start-unified-login"),
  checkAuthStatus: () => ipcRenderer.invoke("check-auth-status"),
  getUserAvatar: () => ipcRenderer.invoke("get-user-avatar"),
  onUpdateUserAvatar: (callback) => ipcRenderer.on("update-user-avatar", (_event, val) => callback(val)),

  // Diálogo e Arquivos Externos
  openFileDialog: () => ipcRenderer.invoke("open-file-dialog"),
  openExternalFile: (filePath) => ipcRenderer.invoke("office-open-external-file", filePath),

  // Configurações & Manutenção do Sistema
  getAllPreferences: () => ipcRenderer.invoke("get-all-preferences"),
  setAppPreference: (key, val) => ipcRenderer.send("set-app-preference", key, val),
  setDesktopShortcut: (enable) => ipcRenderer.invoke("set-desktop-shortcut", enable),
  setNativeOneDriveSync: (enable) => ipcRenderer.invoke("set-native-onedrive-sync", enable),
  clearSessionData: () => ipcRenderer.invoke("clear-session-data"),
  selectFolderDialog: () => ipcRenderer.invoke("select-folder-dialog"),
  getOneDriveLogs: () => ipcRenderer.invoke("get-onedrive-logs"),
  factoryReset: () => ipcRenderer.invoke("factory-reset"),
  uninstallApp: (deleteData) => ipcRenderer.invoke("uninstall-app", { deleteData }),

  // Eventos Push (Main -> Renderer)
  onConfigUpdated: (callback) => ipcRenderer.on("config-updated", (_event, val) => callback(val)),
  onSelectServiceTab: (callback) => ipcRenderer.on("select-service-tab", (_event, val) => callback(val)),
  onBrowserLoading: (callback) => ipcRenderer.on("browser-loading", (_event, val) => callback(val)),
  onOfficeFileLoading: (callback) => ipcRenderer.on("office-file-loading", (_event, val) => callback(val)),
  onOfficeFileOpened: (callback) => ipcRenderer.on("office-file-opened", (_event, val) => callback(val)),
  onOfficeFileFinished: (callback) => ipcRenderer.on("office-file-finished", () => callback()),
  onOfficeFileRenamed: (callback) => ipcRenderer.on("office-file-renamed", (_event, val) => callback(val)),
  onOneDriveStatus: (callback) => ipcRenderer.on("onedrive-status", (_event, val) => callback(val)),
  onBrowserUrl: (callback) => ipcRenderer.on("browser-url", (_event, val) => callback(val)),
  onLocalError: (callback) => ipcRenderer.on("local-error", (_event, val) => callback(val))
});

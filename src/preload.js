const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  notifyResize: () => ipcRenderer.send("window-resize-notify"),

  // Navegação e Serviços
  openService: (service) => ipcRenderer.send("open-service", service),
  browserBack: () => ipcRenderer.send("browser-back"),
  browserForward: () => ipcRenderer.send("browser-forward"),
  browserReload: () => ipcRenderer.send("browser-reload"),
  browserHome: () => ipcRenderer.send("browser-home"),

  // Configuração de Áudio das Notificações Fluent
  getSoundEnabled: () => ipcRenderer.invoke("get-sound-enabled"),
  setSoundEnabled: (enabled) => ipcRenderer.invoke("set-sound-enabled", enabled),
  onSoundSettingChanged: (callback) => ipcRenderer.on("sound-setting-changed", (_event, val) => callback(val)),

  // OneDrive e Armazenamento
  openLocalOneDrive: () => ipcRenderer.send("open-local-onedrive"),
  getOneDriveStatus: () => ipcRenderer.invoke("onedrive-get-status"),
  getOneDriveQuota: (force) => ipcRenderer.invoke("onedrive-get-quota", force),
  getStorageQuota: () => ipcRenderer.invoke("onedrive:get-storage-quota"),
  openOneDriveFolder: () => ipcRenderer.invoke("onedrive-open-folder"),
  syncOneDriveNow: () => ipcRenderer.invoke("onedrive-sync-now"),
  saveLocalNow: () => ipcRenderer.invoke("office-save-local-now"),

  // Diálogo e Arquivos Externos
  openFileDialog: () => ipcRenderer.invoke("open-file-dialog"),
  openExternalFile: (filePath) => ipcRenderer.invoke("office-open-external-file", filePath),

  // Eventos Push (Main -> Renderer)
  onSelectServiceTab: (callback) => ipcRenderer.on("select-service-tab", (_event, val) => callback(val)),
  onBrowserLoading: (callback) => ipcRenderer.on("browser-loading", (_event, val) => callback(val)),
  onOfficeFileLoading: (callback) => ipcRenderer.on("office-file-loading", (_event, val) => callback(val)),
  onOfficeFileOpened: (callback) => ipcRenderer.on("office-file-opened", (_event, val) => callback(val)),
  onOfficeFileFinished: (callback) => ipcRenderer.on("office-file-finished", () => callback()),
  onOfficeFileRenamed: (callback) => ipcRenderer.on("office-file-renamed", (_event, val) => callback(val)),
  onOneDriveStatus: (callback) => ipcRenderer.on("onedrive-status", (_event, val) => callback(val)),
  onOneDriveQuota: (callback) => ipcRenderer.on("onedrive-quota-updated", (_event, val) => callback(val)),
  onStorageQuotaUpdated: (callback) => ipcRenderer.on("onedrive:quota-updated", (_event, val) => callback(val)),
  onBrowserUrl: (callback) => ipcRenderer.on("browser-url", (_event, val) => callback(val)),
  onLocalError: (callback) => ipcRenderer.on("local-error", (_event, val) => callback(val))
});

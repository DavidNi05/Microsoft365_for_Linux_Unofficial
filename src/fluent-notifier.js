const { BrowserWindow, screen, app } = require("electron");
const path = require("path");
const fs = require("fs");

let activeToastWindow = null;
let toastTimeout = null;

function getSettingsPath() {
  try {
    return path.join(app.getPath("userData"), "notification-settings.json");
  } catch (_) {
    return path.join(process.env.HOME || "/tmp", ".m365-notification-settings.json");
  }
}

let isAudioEnabled = true;

// Carrega o estado salvo previamente
try {
  const cfgPath = getSettingsPath();
  if (fs.existsSync(cfgPath)) {
    const raw = JSON.parse(fs.readFileSync(cfgPath, "utf8"));
    if (typeof raw.sound === "boolean") {
      isAudioEnabled = raw.sound;
    }
  }
} catch (_) {}

function setSoundEnabled(enabled) {
  isAudioEnabled = Boolean(enabled);
  try {
    const cfgPath = getSettingsPath();
    fs.mkdirSync(path.dirname(cfgPath), { recursive: true });
    fs.writeFileSync(cfgPath, JSON.stringify({ sound: isAudioEnabled }, null, 2), "utf8");
  } catch (err) {
    console.warn("[Fluent Notifier] Could not save sound setting:", err.message);
  }
  return isAudioEnabled;
}

function getSoundEnabled() {
  return isAudioEnabled;
}

/**
 * Exibe a notificação flutuante com padrão Fluent 2 no canto superior direito da tela
 */
function showFluentNotification({ 
  title, 
  message, 
  icon = "Home.png", 
  duration = 4500,
  playSound = null,
  soundPath = null
}) {
  if (activeToastWindow && !activeToastWindow.isDestroyed()) {
    activeToastWindow.destroy();
    activeToastWindow = null;
  }
  if (toastTimeout) {
    clearTimeout(toastTimeout);
    toastTimeout = null;
  }

  const shouldPlayAudio = playSound !== null ? Boolean(playSound) : isAudioEnabled;

  const primaryDisplay = screen.getPrimaryDisplay();
  const { x: workAreaX, y: workAreaY, width: workAreaWidth } = primaryDisplay.workArea;

  const toastWidth = 360;
  const toastHeight = 90;
  const margin = 20;

  // Canto superior direito (Top-Right)
  const x = Math.floor(workAreaX + workAreaWidth - toastWidth - margin);
  const y = Math.floor(workAreaY + margin);

  activeToastWindow = new BrowserWindow({
    width: toastWidth,
    height: toastHeight,
    x,
    y,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    skipTaskbar: true,
    resizable: false,
    movable: false,
    focusable: false,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      autoplayPolicy: "no-user-gesture-required"
    }
  });

  const iconPath = `assets/icons/${icon}`;

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
        .fluent-card {
          width: ${toastWidth - 10}px;
          height: ${toastHeight - 10}px;
          background-color: rgba(255, 255, 255, 0.92);
          backdrop-filter: blur(24px) saturate(180%);
          -webkit-backdrop-filter: blur(24px) saturate(180%);
          border: 1px solid rgba(0, 0, 0, 0.08);
          border-left: 4px solid #0078d4;
          border-radius: 8px;
          box-shadow: 0 8px 22px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.06);
          display: flex;
          align-items: center;
          padding: 12px 14px;
          gap: 12px;
          position: relative;
          opacity: 0;
          transform: translateY(-16px) scale(0.98);
          animation: fluentIn 0.28s cubic-bezier(0.1, 0.9, 0.2, 1) forwards;
        }
        @keyframes fluentIn {
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .fluent-icon {
          width: 36px;
          height: 36px;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
        }
        .fluent-icon img {
          width: 32px;
          height: 32px;
          object-fit: contain;
        }
        .fluent-body {
          flex: 1;
          display: flex;
          flex-direction: column;
          gap: 2px;
          overflow: hidden;
        }
        .fluent-title {
          font-size: 13px;
          font-weight: 600;
          color: #1e293b;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .fluent-message {
          font-size: 12px;
          color: #64748b;
          line-height: 1.35;
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }
        .fluent-close {
          width: 22px;
          height: 22px;
          border-radius: 4px;
          border: none;
          background: transparent;
          color: #64748b;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 14px;
          transition: background-color 0.15s ease;
        }
        .fluent-close:hover {
          background-color: rgba(0, 0, 0, 0.08);
          color: #0f172a;
        }
      </style>
    </head>
    <body>
      <div class="fluent-card">
        <div class="fluent-icon">
          <img src="${iconPath}" alt="Icon" onerror="this.src='assets/icons/Home.png'">
        </div>
        <div class="fluent-body">
          <div class="fluent-title">${title}</div>
          <div class="fluent-message">${message}</div>
        </div>
        <button class="fluent-close" onclick="window.close()">✕</button>
      </div>

      <script>
        (function() {
          const shouldPlay = ${Boolean(shouldPlayAudio)};
          const customSound = ${soundPath ? JSON.stringify(soundPath) : "null"};

          if (!shouldPlay) return;

          if (customSound) {
            try {
              const audio = new Audio(customSound);
              audio.volume = 0.5;
              audio.play().catch(() => playSyntheticFluentChime());
              return;
            } catch (_) {}
          }

          playSyntheticFluentChime();

          function playSyntheticFluentChime() {
            try {
              const AudioCtx = window.AudioContext || window.webkitAudioContext;
              if (!AudioCtx) return;
              const ctx = new AudioCtx();
              const now = ctx.currentTime;

              const osc1 = ctx.createOscillator();
              const gain1 = ctx.createGain();
              osc1.type = "sine";
              osc1.frequency.setValueAtTime(587.33, now);
              gain1.gain.setValueAtTime(0.001, now);
              gain1.gain.linearRampToValueAtTime(0.18, now + 0.02);
              gain1.gain.exponentialRampToValueAtTime(0.001, now + 0.35);
              osc1.connect(gain1);
              gain1.connect(ctx.destination);
              osc1.start(now);
              osc1.stop(now + 0.36);

              const osc2 = ctx.createOscillator();
              const gain2 = ctx.createGain();
              osc2.type = "sine";
              osc2.frequency.setValueAtTime(880.00, now + 0.07);
              gain2.gain.setValueAtTime(0.001, now + 0.07);
              gain2.gain.linearRampToValueAtTime(0.14, now + 0.09);
              gain2.gain.exponentialRampToValueAtTime(0.001, now + 0.48);
              osc2.connect(gain2);
              gain2.connect(ctx.destination);
              osc2.start(now + 0.07);
              osc2.stop(now + 0.49);
            } catch (_) {}
          }
        })();
      </script>
    </body>
    </html>
  `;

  activeToastWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(htmlContent)}`, {
    baseURLForDataURL: `file://${__dirname}/`
  });

  activeToastWindow.once("ready-to-show", () => {
    if (activeToastWindow && !activeToastWindow.isDestroyed()) {
      activeToastWindow.showInactive();
    }
  });

  toastTimeout = setTimeout(() => {
    if (activeToastWindow && !activeToastWindow.isDestroyed()) {
      activeToastWindow.close();
      activeToastWindow = null;
    }
  }, duration);
}

module.exports = { 
  showFluentNotification,
  setSoundEnabled,
  getSoundEnabled
};

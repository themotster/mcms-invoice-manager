const { app, BrowserWindow, ipcMain, dialog, shell, clipboard, globalShortcut } = require('electron');
const { execFile, spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const isDev = !app.isPackaged || String(process.env.NODE_ENV || '').toLowerCase() !== 'production';
const FORCE_DEVTOOLS = String(process.env.DEVTOOLS || '') === '1';

if (String(process.env.ELECTRON_DISABLE_GPU || '') === '1') {
  app.disableHardwareAcceleration();
}

const DEFAULT_WINDOW_STATE = { width: 1200, height: 800 };

let mainWindow = null;
let db = null;
let documentService = null;
let isQuitting = false;

function ensureServices() {
  try {
    if (!db) db = require('./db');
    if (!documentService) documentService = require('./documentService');
  } catch (err) {
    console.error('Failed to load services', err);
    throw err;
  }
  return { db, documentService };
}

const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.show();
      mainWindow.focus();
    } else {
      createWindow();
    }
  });
}

function broadcastDocumentsChange(payload = {}) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('documents-change', payload);
  }
}

function getStateStorePath() {
  return path.join(app.getPath('userData'), 'window-state.json');
}

const PRELOAD_PATH = path.join(__dirname, 'mcms', 'preload.js');

function restoreWindowState() {
  try {
    const raw = fs.readFileSync(getStateStorePath(), 'utf-8');
    const state = JSON.parse(raw);
    return { ...DEFAULT_WINDOW_STATE, ...state };
  } catch (err) {
    return { ...DEFAULT_WINDOW_STATE };
  }
}

function persistWindowState(state) {
  try {
    fs.writeFileSync(getStateStorePath(), JSON.stringify(state), 'utf-8');
  } catch (err) {
    console.warn('Failed to persist window state', err);
  }
}

function showMainWindow() {
  if (!mainWindow || mainWindow.isDestroyed()) {
    createWindow();
  } else {
    mainWindow.show();
    mainWindow.focus();
  }
}

function createWindow() {
  const state = restoreWindowState();
  const browserOptions = {
    ...DEFAULT_WINDOW_STATE,
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      enableRemoteModule: false,
      nativeWindowOpen: true,
      webSecurity: true,
      allowRunningInsecureContent: false,
      webviewTag: true
    }
  };
  if (typeof state.x === 'number' && typeof state.y === 'number') {
    browserOptions.x = state.x;
    browserOptions.y = state.y;
  }
  if (typeof state.width === 'number') browserOptions.width = state.width;
  if (typeof state.height === 'number') browserOptions.height = state.height;

  const win = new BrowserWindow(browserOptions);
  mainWindow = win;

  const attachDiagnostics = (bw) => {
    if (!bw || bw.isDestroyed()) return;
    const wc = bw.webContents;
    wc.on('did-fail-load', (_e, code, desc, _url, isMainFrame) => {
      console.error('[electron] did-fail-load', { code, desc, isMainFrame });
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
    let lastCrashAt = 0;
    wc.on('render-process-gone', (_e, details) => {
      console.error('[electron] render-process-gone', details);
      if (!bw || bw.isDestroyed()) return;
      const now = Date.now();
      if (now - lastCrashAt > 3000) {
        lastCrashAt = now;
        try { bw.reload(); } catch (_) {}
      }
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
    wc.on('unresponsive', () => {
      console.error('[electron] window unresponsive');
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
    wc.on('dom-ready', () => {
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
  };
  attachDiagnostics(win);

  if (state.isMaximized) win.maximize();

  const saveState = () => {
    if (win.isDestroyed()) return;
    persistWindowState({
      ...win.getBounds(),
      isMaximized: win.isMaximized()
    });
  };

  win.on('close', (event) => {
    saveState();
  });
  win.on('move', () => {
    if (!win.isMaximized() && !win.isMinimized()) saveState();
  });
  win.on('resize', () => {
    if (!win.isMaximized() && !win.isMinimized()) saveState();
  });
  win.on('closed', () => {
    if (mainWindow === win) mainWindow = null;
  });

  win.loadFile(path.join('mcms', 'renderer', 'index.html'));
}

app.whenReady().then(() => {
  try {
    globalShortcut.register('CommandOrControl+Alt+I', () => {
      const bw = BrowserWindow.getFocusedWindow() || mainWindow;
      if (bw && !bw.isDestroyed()) {
        try { bw.webContents.openDevTools({ mode: 'detach' }); } catch (_) {}
      }
    });
    globalShortcut.register('CommandOrControl+R', () => {
      const bw = BrowserWindow.getFocusedWindow() || mainWindow;
      if (bw && !bw.isDestroyed()) { try { bw.reload(); } catch (_) {} }
    });
    globalShortcut.register('F12', () => {
      const bw = BrowserWindow.getFocusedWindow() || mainWindow;
      if (bw && !bw.isDestroyed()) { try { bw.webContents.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
  } catch (err) {
    console.warn('Failed to register global shortcuts', err);
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    } else if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.show();
    }
  });

  app.on('web-contents-created', (_event, contents) => {
    const type = (typeof contents.getType === 'function') ? contents.getType() : 'window';
    if (type === 'webview') {
      contents.setWindowOpenHandler(() => ({ action: 'deny' }));
      return;
    }
    contents.setWindowOpenHandler(({ url }) => {
      try {
        const u = new URL(url);
        if (u.protocol === 'http:' || u.protocol === 'https:') {
          shell.openExternal(url);
          return { action: 'deny' };
        }
      } catch (_) {}
      return { action: 'deny' };
    });
    contents.on('will-navigate', (e, url) => {
      try {
        const u = new URL(url);
        if (u.protocol === 'http:' || u.protocol === 'https:') {
          e.preventDefault();
          shell.openExternal(url);
        }
      } catch (_) {}
    });
  });
});

app.on('will-quit', () => {
  isQuitting = true;
  try { globalShortcut.unregisterAll(); } catch (_) {}
});

// --- IPC handlers (MCMS only) ---

ipcMain.handle('copy-text-to-clipboard', async (_event, text) => {
  try {
    if (typeof text !== 'string' || !text) return { ok: false, message: 'No text to copy' };
    clipboard.writeText(text);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('choose-directory', async (event, args = {}) => {
  const browser = BrowserWindow.fromWebContents(event.sender);
  if (browser && !browser.isDestroyed()) {
    browser.focus();
    browser.show();
  }
  const dialogOptions = {
    title: args.title || 'Select folder',
    defaultPath: args.defaultPath || undefined,
    properties: ['openDirectory', 'createDirectory']
  };
  const result = browser && !browser.isDestroyed()
    ? await dialog.showOpenDialog(browser, dialogOptions)
    : await dialog.showOpenDialog(dialogOptions);
  if (result.canceled || !result.filePaths?.length) return null;
  return result.filePaths[0] || null;
});

ipcMain.handle('choose-file', async (event, args = {}) => {
  const browser = BrowserWindow.fromWebContents(event.sender);
  const dialogOptions = {
    title: args.title || 'Select file',
    defaultPath: args.defaultPath || undefined,
    properties: ['openFile']
  };
  if (args.multiple) dialogOptions.properties.push('multiSelections');
  if (Array.isArray(args.filters) && args.filters.length) dialogOptions.filters = args.filters;
  const result = browser && !browser.isDestroyed()
    ? await dialog.showOpenDialog(browser, dialogOptions)
    : await dialog.showOpenDialog(dialogOptions);
  if (result.canceled || !result.filePaths?.length) return null;
  return args.multiple ? result.filePaths : (result.filePaths[0] || null);
});

ipcMain.handle('open-path', async (_event, targetPath) => {
  if (!targetPath || typeof targetPath !== 'string') return { ok: false, message: 'Missing path' };
  try {
    const result = await shell.openPath(targetPath);
    return result ? { ok: false, message: result } : { ok: true };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('show-item-in-folder', async (_event, targetPath) => {
  if (!targetPath || typeof targetPath !== 'string') return { ok: false, message: 'Missing path' };
  try {
    shell.showItemInFolder(targetPath);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('quick-look-path', async (_event, targetPath) => {
  try {
    if (!targetPath || typeof targetPath !== 'string') return { ok: false, message: 'Missing path' };
    const abs = path.resolve(targetPath);
    try { await fs.promises.access(abs, fs.constants.R_OK); } catch (_) { return { ok: false, message: 'File not found' }; }
    if (process.platform === 'darwin') {
      try {
        const child = spawn('qlmanage', ['-p', abs], { detached: true, stdio: 'ignore' });
        child.unref();
        return { ok: true, mode: 'qlmanage' };
      } catch (_) {
        const esc = abs.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
        const script = `tell application "Finder"
  try
    set theItem to POSIX file "${esc}"
    reveal theItem
    set selection to {theItem}
  end try
end tell
tell application "System Events"
  keystroke space
end tell`;
        try {
          await new Promise((resolve, reject) => {
            execFile('osascript', ['-e', script], (err) => (err ? reject(err) : resolve()));
          });
          return { ok: true, mode: 'finder-quicklook' };
        } catch (err) {
          try {
            const res = await shell.openPath(abs);
            return res ? { ok: false, message: res } : { ok: true, mode: 'default' };
          } catch (e) {
            return { ok: false, message: err?.message || String(err) };
          }
        }
      }
    }
    const res = await shell.openPath(abs);
    return res ? { ok: false, message: res } : { ok: true };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('watch-documents', async (_event, args = {}) => {
  try {
    const { documentService } = ensureServices();
    return await documentService.watchDocumentsFolder({
      ...args,
      onChange: broadcastDocumentsChange
    });
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('unwatch-documents', async (_event, args = {}) => {
  try {
    const { documentService } = ensureServices();
    return await documentService.unwatchDocumentsFolder(args || {});
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

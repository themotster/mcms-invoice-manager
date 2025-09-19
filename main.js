const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

const DEFAULT_WINDOW_STATE = {
  width: 1200,
  height: 800
};

function getStateStorePath() {
  return path.join(app.getPath('userData'), 'window-state.json');
}

const CHILD_WINDOW_OPTIONS = {
  width: 1280,
  height: 900,
  webPreferences: {
    preload: path.join(__dirname, 'preload.js'),
    contextIsolation: true,
    nodeIntegration: true,
    enableRemoteModule: false,
    nativeWindowOpen: true
  }
};

function restoreWindowState() {
  try {
    const raw = fs.readFileSync(getStateStorePath(), 'utf-8');
    const state = JSON.parse(raw);
    return {
      ...DEFAULT_WINDOW_STATE,
      ...state
    };
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

function createWindow() {
  const state = restoreWindowState();
  const browserOptions = {
    ...DEFAULT_WINDOW_STATE,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: true,
      enableRemoteModule: false,
      nativeWindowOpen: true
    }
  };

  if (typeof state.x === 'number' && typeof state.y === 'number') {
    browserOptions.x = state.x;
    browserOptions.y = state.y;
  }
  if (typeof state.width === 'number') browserOptions.width = state.width;
  if (typeof state.height === 'number') browserOptions.height = state.height;

  const win = new BrowserWindow(browserOptions);

  if (state.isMaximized) {
    win.maximize();
  }

  const saveState = () => {
    if (win.isDestroyed()) return;
    const bounds = win.getBounds();
    persistWindowState({
      ...bounds,
      isMaximized: win.isMaximized()
    });
  };

  win.on('close', saveState);
  win.on('move', () => {
    if (!win.isMaximized() && !win.isMinimized()) {
      saveState();
    }
  });
  win.on('resize', () => {
    if (!win.isMaximized() && !win.isMinimized()) {
      saveState();
    }
  });

  win.loadFile('renderer/index.html');
}

function createJobsheetWindow(parent, { businessId, businessName, jobsheetId }) {
  const child = new BrowserWindow({
    ...CHILD_WINDOW_OPTIONS,
    parent
  });

  const query = {
    mode: 'jobsheet'
  };

  if (businessId != null) query.businessId = String(businessId);
  if (businessName) query.businessName = businessName;
  if (jobsheetId != null) query.jobsheetId = String(jobsheetId);

  child.loadFile('renderer/index.html', { query });
  child.on('closed', () => {
    if (parent && !parent.isDestroyed()) {
      parent.focus();
    }
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

ipcMain.handle('open-jobsheet-window', (event, args = {}) => {
  const parent = BrowserWindow.fromWebContents(event.sender);
  createJobsheetWindow(parent || null, args || {});
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

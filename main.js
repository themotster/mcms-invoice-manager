const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

const DEFAULT_WINDOW_STATE = {
  width: 1200,
  height: 800
};

const DEFAULT_JOBSHEET_WINDOW_STATE = {
  width: 1280,
  height: 900
};

let mainWindow = null;
let jobsheetWindow = null;

function getStateStorePath() {
  return path.join(app.getPath('userData'), 'window-state.json');
}

const CHILD_WINDOW_OPTIONS = {
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

function getJobsheetStateStorePath() {
  return path.join(app.getPath('userData'), 'jobsheet-window-state.json');
}

function restoreJobsheetWindowState() {
  try {
    const raw = fs.readFileSync(getJobsheetStateStorePath(), 'utf-8');
    const state = JSON.parse(raw);
    return {
      ...DEFAULT_JOBSHEET_WINDOW_STATE,
      ...state
    };
  } catch (err) {
    return { ...DEFAULT_JOBSHEET_WINDOW_STATE };
  }
}

function persistJobsheetWindowState(state) {
  try {
    fs.writeFileSync(getJobsheetStateStorePath(), JSON.stringify(state), 'utf-8');
  } catch (err) {
    console.warn('Failed to persist jobsheet window state', err);
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
  mainWindow = win;

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

  win.on('closed', () => {
    if (mainWindow === win) {
      mainWindow = null;
    }
  });

  win.loadFile('renderer/index.html');
}

function createJobsheetWindow(parent, { businessId, businessName, jobsheetId }) {
  const savedState = restoreJobsheetWindowState();
  const childOptions = {
    ...CHILD_WINDOW_OPTIONS,
    width: savedState.width,
    height: savedState.height,
    parent
  };

  if (typeof savedState.x === 'number' && typeof savedState.y === 'number') {
    childOptions.x = savedState.x;
    childOptions.y = savedState.y;
  }

  const child = new BrowserWindow(childOptions);
  jobsheetWindow = child;
  child.__jobsheetBusinessId = businessId != null ? Number(businessId) : null;
  child.__jobsheetId = jobsheetId != null ? Number(jobsheetId) : null;

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
    if (jobsheetWindow === child) {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('jobsheet-change', {
          type: 'jobsheet-editor-focus',
          businessId,
          jobsheetId,
          active: false
        });
      }
      jobsheetWindow = null;
    }
  });

  const saveChildState = () => {
    if (child.isDestroyed()) return;
    const bounds = child.getBounds();
    persistJobsheetWindowState({
      ...bounds,
      isMaximized: child.isMaximized()
    });
  };

  child.on('close', saveChildState);
  child.on('move', () => {
    if (!child.isMaximized() && !child.isMinimized()) {
      saveChildState();
    }
  });
  child.on('resize', () => {
    if (!child.isMaximized() && !child.isMinimized()) {
      saveChildState();
    }
  });

  if (savedState.isMaximized) {
    child.maximize();
  }

  return child;
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
  const targetBusinessId = args.businessId != null ? Number(args.businessId) : null;
  const targetJobsheetId = args.jobsheetId != null ? Number(args.jobsheetId) : null;

  if (jobsheetWindow && !jobsheetWindow.isDestroyed()) {
    const currentBusinessId = jobsheetWindow.__jobsheetBusinessId;
    if (currentBusinessId != null && targetBusinessId != null && currentBusinessId !== targetBusinessId) {
      jobsheetWindow.destroy();
      jobsheetWindow = null;
    } else {
      if (jobsheetWindow.isMinimized()) jobsheetWindow.restore();
      jobsheetWindow.focus();
      jobsheetWindow.__jobsheetBusinessId = targetBusinessId;
      jobsheetWindow.__jobsheetId = targetJobsheetId;
      jobsheetWindow.webContents.send('jobsheet-change', {
        type: 'jobsheet-load-request',
        businessId: targetBusinessId,
        businessName: args.businessName || '',
        jobsheetId: targetJobsheetId
      });
      return;
    }
  }

  jobsheetWindow = createJobsheetWindow(parent || null, args || {});
});

ipcMain.on('jobsheet-change', (event, payload = {}) => {
  if (
    payload &&
    payload.type === 'jobsheet-editor-focus' &&
    payload.active &&
    jobsheetWindow &&
    !jobsheetWindow.isDestroyed() &&
    event.sender === jobsheetWindow.webContents
  ) {
    jobsheetWindow.__jobsheetId = payload.jobsheetId != null ? Number(payload.jobsheetId) : null;
    jobsheetWindow.__jobsheetBusinessId = payload.businessId != null ? Number(payload.businessId) : jobsheetWindow.__jobsheetBusinessId;
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('jobsheet-change', payload);
  }
  if (jobsheetWindow && !jobsheetWindow.isDestroyed() && event.sender !== jobsheetWindow.webContents) {
    jobsheetWindow.webContents.send('jobsheet-change', payload);
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

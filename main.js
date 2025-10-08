const { app, BrowserWindow, ipcMain, dialog, shell, clipboard } = require('electron');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');
const documentService = require('./documentService');
const db = require('./db');

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

function broadcastDocumentsChange(payload = {}) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('documents-change', payload);
  }
  if (jobsheetWindow && !jobsheetWindow.isDestroyed()) {
    jobsheetWindow.webContents.send('documents-change', payload);
  }
}

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

ipcMain.handle('create-gig-info-pdf', async (event, args = {}) => {
  try {
    const businessId = Number(args.businessId ?? args.business_id);
    const jobsheetId = Number(args.jobsheetId ?? args.jobsheet_id);
    if (!Number.isInteger(businessId) || !Number.isInteger(jobsheetId)) {
      throw new Error('businessId and jobsheetId are required');
    }
    const gigInfoOverride = (args && typeof args.gigInfo === 'object') ? args.gigInfo : null;
    const { html, targetPath } = await documentService.buildGigInfoHtml({ businessId, jobsheetId, gigInfo: gigInfoOverride });

    const win = new BrowserWindow({ show: false });
    await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
    const pdfBuffer = await win.webContents.printToPDF({ printBackground: true, pageSize: 'A4', marginsType: 2, landscape: false });
    await fs.promises.writeFile(targetPath, pdfBuffer);
    try { if (!win.isDestroyed()) win.close(); } catch (_) {}

    // Insert document row
    const now = new Date().toISOString().slice(0, 19).replace('T', ' ');
    await db.addDocument({
      business_id: businessId,
      jobsheet_id: jobsheetId,
      doc_type: 'pdf_export',
      status: 'ready',
      file_path: targetPath,
      document_date: now,
      definition_key: 'gig_info'
    });

    broadcastDocumentsChange({ type: 'documents-updated', businessId, jobsheetId });
    return { ok: true, file_path: targetPath };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
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

ipcMain.handle('choose-directory', async (event, args = {}) => {
  const browser = BrowserWindow.fromWebContents(event.sender);
  const dialogOptions = {
    title: args.title || 'Select folder',
    defaultPath: args.defaultPath || undefined,
    properties: ['openDirectory', 'createDirectory']
  };

  const result = browser && !browser.isDestroyed()
    ? await dialog.showOpenDialog(browser, dialogOptions)
    : await dialog.showOpenDialog(dialogOptions);
  if (result.canceled || !result.filePaths || !result.filePaths.length) {
    return null;
  }
  return result.filePaths[0] || null;
});

ipcMain.handle('choose-file', async (event, args = {}) => {
  const browser = BrowserWindow.fromWebContents(event.sender);
  const dialogOptions = {
    title: args.title || 'Select file',
    defaultPath: args.defaultPath || undefined,
    properties: ['openFile']
  };
  if (args.multiple) {
    dialogOptions.properties.push('multiSelections');
  }

  if (Array.isArray(args.filters) && args.filters.length) {
    dialogOptions.filters = args.filters;
  }

  const result = browser && !browser.isDestroyed()
    ? await dialog.showOpenDialog(browser, dialogOptions)
    : await dialog.showOpenDialog(dialogOptions);

  if (result.canceled || !result.filePaths || !result.filePaths.length) {
    return null;
  }
  return args.multiple ? result.filePaths : (result.filePaths[0] || null);
});

// New: choose multiple files (returns an array)

ipcMain.handle('open-path', async (_event, targetPath) => {
  if (!targetPath || typeof targetPath !== 'string') return { ok: false, message: 'Missing path' };
  try {
    const result = await shell.openPath(targetPath);
    if (result) {
      return { ok: false, message: result };
    }
    return { ok: true };
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

// Copy a file reference to the OS clipboard to allow pasting into apps (e.g. WhatsApp)
ipcMain.handle('copy-file-to-clipboard', async (_event, targetPath) => {
  try {
    if (!targetPath || typeof targetPath !== 'string') return { ok: false, message: 'Missing path' };
    const abs = path.resolve(targetPath);
    try { await fs.promises.access(abs, fs.constants.R_OK); } catch (_) { return { ok: false, message: 'File not found' }; }

    // macOS: set a true file reference on the clipboard via AppleScript
    if (process.platform === 'darwin') {
      await new Promise((resolve, reject) => {
        const script = `set the clipboard to (POSIX file \"${abs.replace(/\\/g, '\\\\').replace(/\"/g, '\\\"')}\")`;
        execFile('osascript', ['-e', script], (err) => {
          if (err) reject(err); else resolve();
        });
      });
      return { ok: true };
    }

    // Fallback (non-macOS): write file URL formats
    const fileUrl = 'file://' + abs.split(path.sep).map(encodeURIComponent).join('/');
    try { clipboard.clear(); } catch (_) {}
    try { clipboard.writeBuffer('public.file-url', Buffer.from(fileUrl)); } catch (_) {}
    try { clipboard.writeBuffer('text/uri-list', Buffer.from(fileUrl)); } catch (_) {}
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('normalize-template', async (_event, args = {}) => {
  try {
    const result = await documentService.normalizeTemplate(args);
    return result;
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('open-template', async (_event, args = {}) => {
  try {
    const providedPath = typeof args?.templatePath === 'string' && args.templatePath.trim()
      ? args.templatePath.trim()
      : null;
    const templatePath = providedPath
      ? path.resolve(providedPath)
      : path.resolve(__dirname, 'AhMen Client Data and Docs Template.xlsx');
    const result = await shell.openPath(templatePath);
    if (result) {
      return { ok: false, message: result };
    }
    return { ok: true, templatePath };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('watch-documents', async (_event, args = {}) => {
  try {
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
    return await documentService.unwatchDocumentsFolder(args || {});
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
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

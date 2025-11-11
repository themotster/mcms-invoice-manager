const { app, BrowserWindow, ipcMain, dialog, shell, clipboard, globalShortcut } = require('electron');
const { execFile, spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const documentService = require('./documentService');
const db = require('./db');

const isMCMS = (() => {
  try { return /mcms/i.test(app.getName() || '') || String(process.env.APP_MODE || '').toLowerCase() === 'mcms'; } catch (_) { return String(process.env.APP_MODE || '').toLowerCase() === 'mcms'; }
})();

const isDev = !app.isPackaged || String(process.env.NODE_ENV || '').toLowerCase() !== 'production';
const FORCE_DEVTOOLS = isDev || String(process.env.DEVTOOLS || '') === '1';

// Mitigate “white screen” on some GPUs
if (String(process.env.ELECTRON_DISABLE_GPU || '') === '1') {
  app.disableHardwareAcceleration();
}

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

const PRELOAD_PATH = isMCMS
  ? path.join(__dirname, 'mcms', 'preload.js')
  : path.join(__dirname, 'preload.js');

const CHILD_WINDOW_OPTIONS = {
  webPreferences: {
    preload: PRELOAD_PATH,
    contextIsolation: true,
    nodeIntegration: false,
    sandbox: false,
    enableRemoteModule: false,
    nativeWindowOpen: true,
    webSecurity: true,
    allowRunningInsecureContent: false
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
      preload: PRELOAD_PATH,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      enableRemoteModule: false,
      nativeWindowOpen: true,
      webSecurity: true,
      allowRunningInsecureContent: false
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

  // Developer tools and crash diagnostics
  const attachDiagnostics = (bw) => {
    if (!bw || bw.isDestroyed()) return;
    const wc = bw.webContents;
    wc.on('did-fail-load', (_e, code, desc, _url, isMainFrame) => {
      console.error('[electron] did-fail-load', { code, desc, isMainFrame });
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
    // Avoid intrusive modals; log and optionally auto-reload
    let lastCrashAt = 0;
    wc.on('render-process-gone', (_e, details) => {
      console.error('[electron] render-process-gone', details);
      if (!bw || bw.isDestroyed()) return;
      const now = Date.now();
      // Simple cooldown to avoid reload loops
      if (now - lastCrashAt > 3000) {
        lastCrashAt = now;
        try { bw.reload(); } catch (_) {}
      }
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
    wc.on('unresponsive', () => {
      console.error('[electron] window unresponsive');
      // No modal; try to open devtools for diagnostics
      if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
    wc.on('dom-ready', () => {
      if (FORCE_DEVTOOLS) {
        try { wc.openDevTools({ mode: 'detach' }); } catch (_) {}
      }
    });
  };
  attachDiagnostics(win);

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

  if (isMCMS) {
    win.loadFile(path.join('mcms', 'renderer', 'index.html'));
  } else {
    win.loadFile('renderer/index.html');
  }
}

app.whenReady().then(() => {
  // Global shortcuts to aid diagnostics even on a white screen
  try {
    globalShortcut.register('CommandOrControl+Alt+I', () => {
      const bw = BrowserWindow.getFocusedWindow() || mainWindow || jobsheetWindow;
      if (bw && !bw.isDestroyed()) {
        try { bw.webContents.openDevTools({ mode: 'detach' }); } catch (_) {}
      }
    });
    globalShortcut.register('CommandOrControl+R', () => {
      const bw = BrowserWindow.getFocusedWindow() || mainWindow;
      if (bw && !bw.isDestroyed()) { try { bw.reload(); } catch (_) {} }
    });
    globalShortcut.register('F12', () => {
      const bw = BrowserWindow.getFocusedWindow() || mainWindow || jobsheetWindow;
      if (bw && !bw.isDestroyed()) { try { bw.webContents.openDevTools({ mode: 'detach' }); } catch (_) {} }
    });
  } catch (err) {
    console.warn('Failed to register global shortcuts', err);
  }
});

app.on('will-quit', () => {
  try { globalShortcut.unregisterAll(); } catch (_) {}
});

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
  // Attach diagnostics and devtools to child windows too
  try { if (child && !child.isDestroyed()) { const wc = child.webContents; wc.on('dom-ready', () => { if (FORCE_DEVTOOLS) { try { wc.openDevTools({ mode: 'detach' }); } catch (_) {} } }); } } catch (_) {}
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

  // Security: open external links in default browser; block in‑app navigation to remote content
  app.on('web-contents-created', (_event, contents) => {
    contents.setWindowOpenHandler(({ url }) => {
      try { const u = new URL(url); if (u.protocol === 'http:' || u.protocol === 'https:') { shell.openExternal(url); return { action: 'deny' }; } } catch (_) {}
      return { action: 'deny' };
    });
    contents.on('will-navigate', (e, url) => {
      try {
        const u = new URL(url);
        if (u.protocol === 'http:' || u.protocol === 'https:') {
          e.preventDefault();
          shell.openExternal(url);
        }
      } catch (_) { /* non-URL targets (like file) are allowed */ }
    });
  });
});

// Email via Microsoft Graph – handle in main process to avoid renderer CORS
ipcMain.handle('send-mail-via-graph', async (_event, args = {}) => {
  try {
    const res = await documentService.sendMailViaGraph(args || {});
    // Maintain original return shape
    return { ok: true, ...(res || {}) };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('schedule-mail-via-graph', async (_event, args = {}) => {
  try {
    const res = await documentService.scheduleMailViaGraph(args || {});
    return { ok: true, ...(res || {}) };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
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

// Build and export a PDF of upcoming events and required personnel
ipcMain.handle('create-personnel-log-pdf', async (_event, args = {}) => {
  try {
    const businessId = Number(args.businessId ?? args.business_id);
    if (!Number.isInteger(businessId)) {
      throw new Error('businessId is required');
    }
    const { html, targetPath } = await documentService.buildPersonnelLogHtml({
      businessId,
      fromDate: args.fromDate || args.from_date,
      toDate: args.toDate || args.to_date,
      includeArchived: args.includeArchived || args.include_archived,
      columns: Array.isArray(args.columns) ? args.columns : undefined
    });

    const win = new BrowserWindow({ show: false });
    await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
    const pdfBuffer = await win.webContents.printToPDF({ printBackground: true, pageSize: 'A4', marginsType: 2, landscape: false });
    await fs.promises.writeFile(targetPath, pdfBuffer);
    try { if (!win.isDestroyed()) win.close(); } catch (_) {}
    return { ok: true, file_path: targetPath };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});


ipcMain.handle('compose-mail-draft', async (_event, args = {}) => {
  try {
    const res = await documentService.composeMailDraft(args || {});
    return { ok: true, ...(res || {}) };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('create-personnel-log-text', async (_event, args = {}) => {
  try {
    const businessId = Number(args.businessId ?? args.business_id);
    if (!Number.isInteger(businessId)) {
      throw new Error('businessId is required');
    }
    const { text } = await documentService.buildPersonnelLogText({
      businessId,
      fromDate: args.fromDate || args.from_date,
      toDate: args.toDate || args.to_date,
      includeArchived: args.includeArchived || args.include_archived,
      columns: Array.isArray(args.columns) ? args.columns : undefined,
      singleLine: args.singleLine !== false,
      bullet: args.bullet !== false
    });
    return { ok: true, text };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

// List Apple Contacts (JXA) for import UI
ipcMain.handle('list-apple-contacts', async () => {
  try {
    const jxa = `
      var app = Application('Contacts');
      app.includeStandardAdditions = true;
      var people = app.people();
      function g(x) { try { return (typeof x === 'function') ? x() : x; } catch (e) { return ''; } }
      var out = [];
      for (var i = 0; i < people.length; i++) {
        var p = people[i];
        var emails = [];
        try { var es = p.emails(); for (var j = 0; j < es.length; j++) emails.push({ label: g(es[j].label), value: g(es[j].value) }); } catch (_) {}
        var phones = [];
        try { var phs = p.phones(); for (var k = 0; k < phs.length; k++) phones.push({ label: g(phs[k].label), value: g(phs[k].value) }); } catch (_) {}
        var addrs = [];
        try { var as = p.addresses(); for (var m = 0; m < as.length; m++) addrs.push({ label: g(as[m].label), street: g(as[m].street), city: g(as[m].city), state: g(as[m].state), zip: g(as[m].zip), country: g(as[m].country) }); } catch (_) {}
        out.push({ id: g(p.id), name: g(p.name), firstName: g(p.firstName), lastName: g(p.lastName), organization: g(p.organization), emails: emails, phones: phones, addresses: addrs });
      }
      JSON.stringify(out);
    `;
    const args = ['-l', 'JavaScript', '-e', jxa];
    const json = await new Promise((resolve, reject) => {
      execFile('osascript', args, { timeout: 30000, maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
        if (error) {
          const msg = (stderr || stdout || error.message || '').toString();
          return reject(new Error(msg.trim() || 'Contacts fetch failed'));
        }
        resolve((stdout || '').toString());
      });
    });
    let list = [];
    try { list = JSON.parse(json); } catch (_) { list = []; }
    if (Array.isArray(list) && list.length > 0) {
      return { ok: true, contacts: list };
    }

    // Fallback: AppleScript TSV (name, emailsCSV, phonesCSV)
    const as = `
      on join_list(L, d)
        set {tids, AppleScript's text item delimiters} to {AppleScript's text item delimiters, d}
        set res to L as text
        set AppleScript's text item delimiters to tids
        return res
      end join_list
      tell application "Contacts"
        set thePeople to every person
        set out to {}
        repeat with p in thePeople
          set nm to name of p as text
          set ems to {}
          repeat with e in (emails of p)
            try
              set end of ems to (value of e as text)
            end try
          end repeat
          set phs to {}
          repeat with ph in (phones of p)
            try
              set end of phs to (value of ph as text)
            end try
          end repeat
          set emText to my join_list(ems, ",")
          set phText to my join_list(phs, ",")
          set end of out to (nm & tab & emText & tab & phText)
        end repeat
      end tell
      return out as text
    `;
    const tsv = await new Promise((resolve, reject) => {
      execFile('osascript', ['-e', as], { timeout: 30000, maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
        if (error) {
          const msg = (stderr || stdout || error.message || '').toString();
          return reject(new Error(msg.trim() || 'Contacts AppleScript failed'));
        }
        resolve((stdout || '').toString());
      });
    });
    const rows = tsv.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const parsed = rows.map(line => {
      const parts = line.split(/\t/);
      const name = parts[0] || '';
      const emails = (parts[1] || '').split(',').map(s => s.trim()).filter(Boolean).map((v,i) => ({ label: i === 0 ? 'Primary' : '', value: v }));
      const phones = (parts[2] || '').split(',').map(s => s.trim()).filter(Boolean).map((v,i) => ({ label: i === 0 ? 'Mobile' : '', value: v }));
      return { id: name, name, emails, phones, addresses: [] };
    });
    return { ok: true, contacts: parsed };
  } catch (err) {
    return { ok: false, message: err?.message || String(err), contacts: [] };
  }
});

ipcMain.handle('copy-text-to-clipboard', async (_event, text) => {
  try {
    if (typeof text !== 'string' || !text) return { ok: false, message: 'No text to copy' };
    clipboard.writeText(text);
    return { ok: true };
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
      // Prefer writing a file URL UTI that many apps (e.g., WhatsApp) accept
      try {
        const fileUrl = 'file://' + abs.split(path.sep).map(encodeURIComponent).join('/');
        try { clipboard.clear(); } catch (_) {}
        try { clipboard.writeBuffer('public.file-url', Buffer.from(fileUrl)); } catch (_) {}
        try { clipboard.writeBuffer('text/uri-list', Buffer.from(fileUrl)); } catch (_) {}
      } catch (_) {}
      // Fallback to AppleScript alias on macOS for broader compatibility
      try {
        await new Promise((resolve, reject) => {
          const script = `set the clipboard to (POSIX file \"${abs.replace(/\\/g, '\\\\').replace(/\"/g, '\\\"')}\")`;
          execFile('osascript', ['-e', script], (err) => {
            if (err) reject(err); else resolve();
          });
        });
      } catch (_) {}
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

// Quick Look a file (macOS). Non-macOS falls back to opening the file.
ipcMain.handle('quick-look-path', async (_event, targetPath) => {
  try {
    if (!targetPath || typeof targetPath !== 'string') return { ok: false, message: 'Missing path' };
    const abs = path.resolve(targetPath);
    try { await fs.promises.access(abs, fs.constants.R_OK); } catch (_) { return { ok: false, message: 'File not found' }; }
    if (process.platform === 'darwin') {
      // Prefer qlmanage to avoid opening the Finder window
      try {
        const child = spawn('qlmanage', ['-p', abs], { detached: true, stdio: 'ignore' });
        child.unref();
        return { ok: true, mode: 'qlmanage' };
      } catch (_qmErr) {
        // Fallback: Finder-based Quick Look via AppleScript
        const esc = abs.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
        const script = `
tell application "Finder"
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
            execFile('osascript', ['-e', script], (err) => {
              if (err) reject(err); else resolve();
            });
          });
          return { ok: true, mode: 'finder-quicklook' };
        } catch (err) {
          // Final fallback: default handler (e.g., Preview)
          try {
            const res = await shell.openPath(abs);
            if (res) return { ok: false, message: res };
            return { ok: true, mode: 'default' };
          } catch (e) {
            return { ok: false, message: err?.message || String(err) };
          }
        }
      }
    }
    const res = await shell.openPath(abs);
    if (res) return { ok: false, message: res };
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

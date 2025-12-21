const { app, BrowserWindow, ipcMain, dialog, shell, clipboard, globalShortcut, Tray, Menu, Notification, nativeImage } = require('electron');
const { execFile, spawn } = require('child_process');
const os = require('os');
const path = require('path');
const fs = require('fs');

const isMCMS = (() => {
  // 1) Explicit env override
  const mode = String(process.env.APP_MODE || '').toLowerCase();
  if (mode === 'mcms') return true;

  const candidates = [];
  try { candidates.push(String(app.getName ? app.getName() : '')); } catch (_) {}
  try { candidates.push(String(process.execPath || '')); } catch (_) {}
  try { candidates.push(String(app.getPath ? app.getPath('exe') : '')); } catch (_) {}
  try { candidates.push(String(app.getAppPath ? app.getAppPath() : '')); } catch (_) {}
  try { candidates.push(String(process.resourcesPath || '')); } catch (_) {}

  const joined = candidates.filter(Boolean).join(' | ').toLowerCase();
  if (joined.includes('mcms') || joined.includes('m.c.m.s') || /mcms\s+invoice/.test(joined)) {
    return true;
  }
  return false;
})();

const isDev = !app.isPackaged || String(process.env.NODE_ENV || '').toLowerCase() !== 'production';
// Only force devtools when explicitly requested (DEVTOOLS=1)
const FORCE_DEVTOOLS = String(process.env.DEVTOOLS || '') === '1';
const FORCE_MAIN_SCHEDULER = String(process.env.AHMEN_SCHEDULED_WORKER || '') === '1';

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
let tray = null;
let plannerSchedulerStarted = false;
let plannerSchedulerRunning = false;
let loginPreference = null;
let db = null;
let documentService = null;
let pendingUiAction = null;
let backgroundModeEnabled = false;
let startHiddenOnLogin = false;
let isQuitting = false;

const SHARED_SUPPORT_DIR = path.join(os.homedir(), 'Library', 'Application Support', 'AhMen Booking Manager');
const PENDING_UI_ACTION_PATH = path.join(SHARED_SUPPORT_DIR, 'pending-ui-action.json');

function setBackgroundMode(enabled, { startHidden = false } = {}) {
  backgroundModeEnabled = !!enabled;
  startHiddenOnLogin = !!startHidden;
  if (backgroundModeEnabled) {
    ensureTray();
    startPlannerScheduler();
    if (startHiddenOnLogin) {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.hide();
      }
      try { if (app.dock) app.dock.hide(); } catch (_) {}
    }
  } else if (tray) {
    try { tray.destroy(); } catch (_) {}
    tray = null;
  }
}

function requestQuit() {
  isQuitting = true;
  app.quit();
}

function ensureServices() {
  try {
    if (!db) {
      db = require('./db');
    }
    if (!documentService) {
      documentService = require('./documentService');
    }
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
    showMainWindow();
  });
}

function broadcastDocumentsChange(payload = {}) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('documents-change', payload);
  }
  if (jobsheetWindow && !jobsheetWindow.isDestroyed()) {
    jobsheetWindow.webContents.send('documents-change', payload);
  }
}

function dispatchUiAction(payload = null) {
  if (!payload) return;
  if (mainWindow && !mainWindow.isDestroyed()) {
    try {
      mainWindow.webContents.send('ui-action', payload);
      pendingUiAction = null;
      return;
    } catch (_err) {}
  }
  pendingUiAction = payload;
}

function writePendingUiAction(payload = {}) {
  try {
    fs.mkdirSync(SHARED_SUPPORT_DIR, { recursive: true });
    const data = { ...payload, ts: Date.now() };
    fs.writeFileSync(PENDING_UI_ACTION_PATH, JSON.stringify(data), 'utf8');
    return true;
  } catch (err) {
    console.warn('Failed to write pending UI action', err);
    return false;
  }
}

function consumePendingUiAction() {
  try {
    if (!fs.existsSync(PENDING_UI_ACTION_PATH)) return null;
    const raw = fs.readFileSync(PENDING_UI_ACTION_PATH, 'utf8');
    fs.unlinkSync(PENDING_UI_ACTION_PATH);
    const parsed = JSON.parse(raw || '{}');
    if (parsed && parsed.type) {
      dispatchUiAction(parsed);
      return parsed;
    }
  } catch (err) {
    console.warn('Failed to consume pending UI action', err);
  }
  return null;
}

function watchPendingUiAction() {
  try {
    fs.mkdirSync(SHARED_SUPPORT_DIR, { recursive: true });
    consumePendingUiAction();
    const watcher = fs.watch(SHARED_SUPPORT_DIR, (event, filename) => {
      if (!filename) return;
      if (String(filename) === path.basename(PENDING_UI_ACTION_PATH)) {
        consumePendingUiAction();
      }
    });
    watcher.on('error', (err) => {
      console.warn('Pending UI action watcher error', err);
    });
  } catch (err) {
    console.warn('Failed to watch pending UI action file', err);
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
    allowRunningInsecureContent: false,
    webviewTag: true
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

function parseDateKey(value) {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  return { year, month, day };
}

function toLocalDateFromKey(dateKey, hour = 9, minute = 0) {
  const parts = parseDateKey(dateKey);
  if (!parts) return null;
  return new Date(parts.year, parts.month - 1, parts.day, hour, minute, 0, 0);
}

function parseSqlDateTime(value) {
  if (!value) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(raw)) {
    const parsed = new Date(`${raw.replace(' ', 'T')}Z`);
    return Number.isNaN(parsed.valueOf()) ? null : parsed;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

function createTrayIcon() {
  const trayPath = path.join(__dirname, 'icons', 'ahmen-tray.png');
  let img = nativeImage.createFromPath(trayPath);
  if (img.isEmpty()) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 2L12.4 14H10.8L9.9 11.4H6.1L5.2 14H3.6L8 2Z" fill="black"/><rect x="6.7" y="8.6" width="2.6" height="1.2" fill="white"/></svg>`;
    const dataUrl = `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
    img = nativeImage.createFromDataURL(dataUrl);
  }
  img.setTemplateImage(true);
  return img;
}

function showMainWindow({ openPlanner = false } = {}) {
  if (backgroundModeEnabled && app.dock && process.platform === 'darwin') {
    try { app.dock.show(); } catch (_) {}
  }
  if (!mainWindow || mainWindow.isDestroyed()) {
    createWindow();
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.once('did-finish-load', () => {
        if (openPlanner) dispatchUiAction({ type: 'open-planner' });
        if (pendingUiAction) {
          dispatchUiAction(pendingUiAction);
        }
      });
    }
  } else {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.show();
    mainWindow.focus();
    if (openPlanner) dispatchUiAction({ type: 'open-planner' });
  }
}

function ensureTray() {
  if (tray) return;
  tray = new Tray(createTrayIcon());
  tray.setToolTip('AhMen Reminders');
  const showTrayMenu = () => {
    try {
      tray.popUpContextMenu();
    } catch (_) {}
  };
  tray.on('click', () => {
    showTrayMenu();
  });
  tray.on('right-click', showTrayMenu);
  tray.setContextMenu(Menu.buildFromTemplate([
    { label: 'Open AhMen', click: () => showMainWindow() },
    { type: 'separator' },
    { label: 'Quit AhMen', click: () => requestQuit() }
  ]));
}

function updateTrayBadge(count) {
  if (!tray) return;
  const numeric = Number(count) || 0;
  if (process.platform === 'darwin') {
    tray.setTitle(numeric > 0 ? ` ${numeric}` : '');
  }
  tray.setToolTip(numeric > 0 ? `AhMen Reminders (${numeric} due)` : 'AhMen Reminders');
}

function getLoginPreferencePath() {
  try {
    return path.join(SHARED_SUPPORT_DIR, 'login-item.json');
  } catch (_err) {
    return null;
  }
}

function readLoginPreference() {
  const prefPath = getLoginPreferencePath();
  const legacyPaths = [
    path.join(os.homedir(), 'Library', 'Application Support', 'invoice-master-dashboard', 'login-item.json')
  ];
  const paths = [prefPath, ...legacyPaths].filter(Boolean);
  for (const candidate of paths) {
    if (!candidate) continue;
    try {
      if (!fs.existsSync(candidate)) continue;
      const raw = fs.readFileSync(candidate, 'utf-8');
      const parsed = JSON.parse(raw || '{}');
      if (typeof parsed?.openAtLogin === 'boolean') return parsed.openAtLogin;
    } catch (_err) {}
  }
  return null;
}

function writeLoginPreference(openAtLogin) {
  const prefPath = getLoginPreferencePath();
  if (!prefPath) return;
  try {
    fs.mkdirSync(path.dirname(prefPath), { recursive: true });
    fs.writeFileSync(prefPath, JSON.stringify({ openAtLogin: !!openAtLogin }, null, 2), 'utf-8');
  } catch (_err) {}
}

function applyLoginItemSetting(openAtLogin) {
  try {
    app.setLoginItemSettings({
      openAtLogin: !!openAtLogin,
      openAsHidden: true
    });
  } catch (_err) {}
}

function ensureLoginItemDefault() {
  if (isMCMS) return;
  if (loginPreference !== null) return;
  loginPreference = readLoginPreference();
  if (loginPreference === null) {
    loginPreference = true;
    writeLoginPreference(true);
  }
  applyLoginItemSetting(loginPreference);
}

async function startPlannerScheduler() {
  if (plannerSchedulerStarted) return;
  plannerSchedulerStarted = true;
  if (isMCMS) return;
  try {
    ensureServices();
  } catch (err) {
    console.error('Planner scheduler failed to start', err);
    return;
  }

  const tick = async () => {
    if (plannerSchedulerRunning) return;
    plannerSchedulerRunning = true;
    try {
      const businesses = await db.businessSettings();
      let dueCount = 0;
      for (const business of businesses || []) {
        const businessId = Number(business?.id);
        if (!Number.isInteger(businessId)) continue;
        const result = await documentService.listPlannerItems({
          businessId,
          includeCompleted: true,
          sync: true
        });
        const items = Array.isArray(result?.items) ? result.items : [];
        const now = new Date();
        for (const item of items) {
          const status = String(item.status || 'pending').toLowerCase();
          const actionKey = String(item.action_key || '').toLowerCase();
          const scheduledFor = item.scheduled_for || '';
          const dueAt = toLocalDateFromKey(scheduledFor, 9, 0);
          if (!dueAt) continue;
          const isDue = now >= dueAt;
          if (actionKey === 'balance_send') {
            if (['sent', 'done', 'completed', 'dismissed'].includes(status)) continue;
            const msToDue = dueAt.getTime() - now.getTime();
            const isDueSoon = msToDue > 0 && msToDue <= 24 * 60 * 60 * 1000;
            if (item.scheduled_email_at) {
              continue;
            }
            if (!isDue && !isDueSoon) continue;
            dueCount += 1;
            const lastNotified = parseSqlDateTime(item.last_notified_at);
            if (!lastNotified || now - lastNotified >= 24 * 60 * 60 * 1000) {
              const missingParts = [];
              if (item.needs_email) missingParts.push('client email');
              if (item.needs_invoice) missingParts.push('invoice PDF');
              const missingLabel = missingParts.length
                ? `Missing ${missingParts.join(' and ')}`
                : (item.can_send ? 'Ready to send (manual)' : 'Missing details');
              try {
                new Notification({
                  title: 'Balance invoice reminder',
                  body: `${item.client_name || 'Client'} · ${scheduledFor} · ${missingLabel}`
                }).show();
              } catch (_) {}
              await documentService.updatePlannerAction({
                businessId,
                jobsheetId: item.jobsheet_id,
                actionKey: item.action_key,
                scheduled_for: scheduledFor,
                last_notified_at: new Date().toISOString()
              });
            }
            continue;
          }

          if (actionKey === 'payment_check') {
            if (!isDue) continue;
            if (['done', 'completed', 'dismissed'].includes(status)) continue;
            const lastNotified = parseSqlDateTime(item.last_notified_at);
            if (lastNotified && now - lastNotified < 12 * 60 * 60 * 1000) {
              dueCount += 1;
              continue;
            }
            dueCount += 1;
            try {
              new Notification({
                title: 'Payment check due',
                body: `${item.client_name || 'Client'} · ${item.event_date || ''}`
              }).show();
            } catch (_) {}
            await documentService.updatePlannerAction({
              businessId,
              jobsheetId: item.jobsheet_id,
              actionKey: item.action_key,
              scheduled_for: scheduledFor,
              last_notified_at: new Date().toISOString()
            });
          }
        }
      }
      updateTrayBadge(dueCount);
    } catch (err) {
      console.error('Planner scheduler error', err);
    } finally {
      plannerSchedulerRunning = false;
    }
  };

  tick();
  setInterval(tick, 5 * 60 * 1000);
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

  win.on('close', (event) => {
    if (backgroundModeEnabled && !isQuitting) {
      event.preventDefault();
      saveState();
      win.hide();
      try { if (app.dock) app.dock.hide(); } catch (_) {}
      return;
    }
    saveState();
  });
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

  win.webContents.once('did-finish-load', () => {
    if (pendingUiAction) {
      dispatchUiAction(pendingUiAction);
    }
  });
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

app.on('before-quit', (event) => {
  if (backgroundModeEnabled && !isQuitting) {
    event.preventDefault();
    if (mainWindow && !mainWindow.isDestroyed()) {
      try { mainWindow.hide(); } catch (_) {}
    }
    try { if (app.dock) app.dock.hide(); } catch (_) {}
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
  watchPendingUiAction();
  if (isMCMS) {
    setBackgroundMode(false, { startHidden: false });
    createWindow();
  } else {
    ensureLoginItemDefault();
    try {
      const pref = loginPreference === null ? readLoginPreference() : loginPreference;
      const settings = app.getLoginItemSettings();
      const openedAtLogin = !!(pref && settings?.wasOpenedAtLogin);
      setBackgroundMode(!!pref, { startHidden: openedAtLogin });
      if (!openedAtLogin) {
        createWindow();
      } else {
        try { if (app.dock) app.dock.hide(); } catch (_) {}
      }
    } catch (_err) {
      createWindow();
    }
  }

  if (FORCE_MAIN_SCHEDULER && !backgroundModeEnabled) {
    try {
      startPlannerScheduler();
    } catch (err) {
      console.error('Failed to start main scheduler', err);
    }
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      showMainWindow();
    } else if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.show();
    }
  });

  // Security: open external links in default browser for windows; allow webviews to navigate to permitted hosts
  app.on('web-contents-created', (_event, contents) => {
    const type = (typeof contents.getType === 'function') ? contents.getType() : 'window';
    if (type === 'webview') {
      // Block popups from webviews; keep navigation within the webview itself
      contents.setWindowOpenHandler(({ url }) => {
        try {
          const u = new URL(url);
          const isHttp = u.protocol === 'http:' || u.protocol === 'https:';
          if (!isHttp) return { action: 'deny' };
          // Allow Google hosts to remain in the webview; deny creating new windows
          const host = (u.host || '').toLowerCase();
          const isGoogle = host.includes('google.');
          return { action: isGoogle ? 'deny' : 'deny' };
        } catch (_) {
          return { action: 'deny' };
        }
      });
      // Do NOT redirect webview navigation to external browser; keep it in the webview
      return;
    }

    // Default behavior for normal BrowserWindow/webContents
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
      } catch (_) { /* non-URL targets (like file) are allowed */ }
    });
  });
});

// Email via Microsoft Graph – handle in main process to avoid renderer CORS
ipcMain.handle('send-mail-via-graph', async (_event, args = {}) => {
  try {
    const { documentService } = ensureServices();
    const res = await documentService.sendMailViaGraph(args || {});
    // Maintain original return shape
    return { ok: true, ...(res || {}) };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('schedule-mail-via-graph', async (_event, args = {}) => {
  try {
    const { documentService } = ensureServices();
    const res = await documentService.scheduleMailViaGraph(args || {});
    return { ok: true, ...(res || {}) };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('get-login-item-settings', async () => {
  try {
    const pref = readLoginPreference();
    const settings = app.getLoginItemSettings();
    if (pref === null || pref === undefined) return settings;
    return { ...(settings || {}), openAtLogin: !!pref };
  } catch (err) {
    return { error: err?.message || String(err) };
  }
});

ipcMain.handle('set-login-item-settings', async (_event, args = {}) => {
  try {
    const openAtLogin = args?.openAtLogin ?? args?.open_at_login;
    const next = !!openAtLogin;
    loginPreference = next;
    writeLoginPreference(next);
    applyLoginItemSetting(next);
    setBackgroundMode(next, { startHidden: false });
    const settings = app.getLoginItemSettings();
    return { ...(settings || {}), openAtLogin: next };
  } catch (err) {
    return { error: err?.message || String(err) };
  }
});

ipcMain.handle('test-notification', async () => {
  if (!Notification.isSupported()) {
    return { ok: false, error: 'Notifications not supported' };
  }
  try {
    new Notification({
      title: 'AhMen test notification',
      body: 'Notifications are enabled.'
    }).show();
    return { ok: true };
  } catch (err) {
    console.error('Failed to show test notification', err);
    return { ok: false, error: err?.message || 'Unable to show notification' };
  }
});

ipcMain.handle('create-gig-info-pdf', async (event, args = {}) => {
  try {
    const { db, documentService } = ensureServices();
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
    const { documentService } = ensureServices();
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
    const { documentService } = ensureServices();
    const res = await documentService.composeMailDraft(args || {});
    return { ok: true, ...(res || {}) };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
});

ipcMain.handle('create-personnel-log-text', async (_event, args = {}) => {
  try {
    const { documentService } = ensureServices();
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
    const { documentService } = ensureServices();
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
    return;
  }
  if (backgroundModeEnabled) {
    try { if (app.dock) app.dock.hide(); } catch (_) {}
    return;
  }
  app.quit();
});
